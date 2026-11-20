# -*- coding: utf-8 -*-
#
# DOYDL's Microsoft Graph Integration Toolkit — m365
#
# The `m365` package provides reusable Python clients for Microsoft Graph.
# It supports SharePoint document libraries, OneDrive files and folders,
# and Outlook mailboxes through the `m365.graph` namespace.
#
# Applications can search for files and folders, transfer documents, move
# or rename items within a drive, and work with messages and attachments.
# The toolkit separates these operations from tenant-specific authentication,
# configuration, and business logic.
#
# Access is subject to Microsoft Graph capabilities, tenant policies, and
# the permissions granted to the calling application or signed-in user.
# Installing the package does not grant access to Microsoft 365 resources.
#
# Microsoft, Microsoft Graph, Microsoft 365, Outlook, SharePoint, and OneDrive
# are trademarks of Microsoft Corporation. This project is not an official
# Microsoft product and does not imply affiliation, sponsorship, or
# endorsement. Third-party assets remain subject to their own terms and
# notices; the license below does not grant rights to third-party trademarks
# or logos.
#
# Copyright (c) 2026 by DOYDL Technologies. All rights reserved.
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the “Software”), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#
"""Organization-neutral Graph transport; capability URLs never receive Graph tokens."""
import time
import math
import re
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import urlsplit, unquote, parse_qsl, urlencode, urlunsplit

import requests


class _NoImplicitAuth(requests.auth.AuthBase):
    def __call__(self, request):
        return request


class GraphError(RuntimeError):
    """A redacted HTTP failure (response bodies and capability URLs are excluded)."""
    def __init__(self, status=None, code="request_failed"):
        self.status_code = status
        self.code = code
        super().__init__("Microsoft Graph request failed (status={}, code={})".format(status, code))


def read_response(request, *, resource_url, **kwargs):
    """Require a completed GET without changing mutation/session status semantics."""
    result = request(resource_url=resource_url, response_metadata=True, **kwargs)
    if not isinstance(result, dict) or "body" not in result:
        raise GraphError(code="invalid_read_response")
    status = result.get("status_code")
    if isinstance(status, bool) or not isinstance(status, int):
        raise GraphError(code="invalid_read_status")
    if status != 200:
        raise GraphError(status, "unexpected_read_status")
    return result["body"]


class _CollectionBudget:
    """Share finite logical-request and processed-row limits across one operation."""

    def __init__(self, max_requests, max_items, *, prefix="collection"):
        for name, value in (("max_requests", max_requests), ("max_items", max_items)):
            if isinstance(value, bool) or not isinstance(value, int) or value < 1:
                raise ValueError(name + " must be a positive integer")
        self.max_requests, self.max_items = max_requests, max_items
        self.requests = self.items = 0
        self.prefix = prefix

    def before_request(self):
        if self.requests >= self.max_requests:
            raise GraphError(code=self.prefix + "_request_limit")
        self.requests += 1

    def consume_items(self, count):
        if count > self.max_items - self.items:
            raise GraphError(code=self.prefix + "_item_limit")
        self.items += count


class CollectionScope:
    """Bind opaque continuations to one endpoint, not just one Graph origin."""

    def __init__(self, absolute):
        self.identity = self._identity(absolute)

    @staticmethod
    def _identity(absolute):
        parts = urlsplit(absolute)
        # Segment-wise decoding preserves the distinction between IDs containing
        # encoded slashes and actual endpoint separators. Query tokens stay opaque.
        return (parts.scheme.lower(), parts.netloc.lower(),
                tuple(unquote(segment) for segment in parts.path.split("/")))

    def check_page_url(self, absolute):
        if self._identity(absolute) != self.identity:
            raise GraphError(code="collection_context_changed")


def collection_next_link(page, code="invalid_collection_continuation"):
    """Absence ends a scan; a present malformed continuation never does."""
    if "@odata.nextLink" not in page:
        return None
    link = page["@odata.nextLink"]
    if not isinstance(link, str) or not link.strip() or any(ord(c) < 32 for c in link):
        raise GraphError(code=code)
    return link  # Opaque: validate, never rebuild or normalize tokens.


def resource_drive_id(resource):
    try:
        path = urlsplit(resource).path
    except (TypeError, ValueError, AttributeError):
        raise GraphError(code="invalid_drive_resource") from None
    match = re.match(r"(?:/v[0-9.]+/|/beta/|/)?drives/([^/]+)(?:/|$)", path)
    return unquote(match.group(1)) if match else None


def checked_drive_item(item, drive_id, code="invalid_drive_item"):
    """Reject malformed/contradictory metadata without inventing missing fields."""
    if not isinstance(item, dict) or not isinstance(item.get("id"), str) or not item["id"].strip():
        raise GraphError(code=code)
    if any(marker in item for marker in ("error", "deleted", "@removed")):
        raise GraphError(code=code)
    if "parentReference" in item:
        parent = item["parentReference"]
        if not isinstance(parent, dict):
            raise GraphError(code=code)
        if "driveId" in parent:
            observed = parent["driveId"]
            if not isinstance(observed, str) or not observed.strip():
                raise GraphError(code=code)
            if drive_id is not None and observed != drive_id:
                raise GraphError(code="foreign_drive_item")
        if "path" in parent:
            if not isinstance(parent["path"], str):
                raise GraphError(code=code)
            observed = resource_drive_id(parent["path"])
            if drive_id is not None and observed is not None and observed != drive_id:
                raise GraphError(code="foreign_drive_item")
    return item


def drive_item_kind(item, code="invalid_drive_item_facet"):
    """Identify supported file/folder facets without treating invalid facets as absent."""
    kinds = [kind for kind in ("file", "folder") if kind in item]
    if len(kinds) > 1 or any(not isinstance(item[kind], dict) for kind in kinds):
        raise GraphError(code=code)
    return kinds[0] if kinds else None


def drive_item_name(item, code="invalid_drive_item_name"):
    name = item.get("name")
    if (not isinstance(name, str) or not name.strip() or name in (".", "..")
            or "/" in name or "\\" in name or any(ord(c) < 32 for c in name)):
        raise GraphError(code=code)
    return name


_CHILDREN_ROUTE = re.compile(
    r"(?P<route>.+?/(?:drives/(?P<drive>[^/]+)|drive))/"
    r"(?P<ref>items/(?P<id>[^/]+)|root|root:/.*:)/children\Z"
)


class DriveChildrenScope:
    """Keep child pages bound to one drive and parent, including root aliases."""

    def __init__(self, absolute, *, prefix="collection"):
        route = _CHILDREN_ROUTE.fullmatch(urlsplit(absolute).path)
        self.route = self._route(route)
        self.drive_id = unquote(route["drive"]) if route["drive"] is not None else None
        self.parent_id = unquote(route["id"]) if route["id"] is not None else None
        self.canonical_item = route["drive"] is not None and route["id"] is not None
        self.prefix = prefix

    @staticmethod
    def _route(match):
        # Decode the resource segments, not the query/continuation token.
        return unquote(match["route"]), unquote(match["ref"])

    def check_page_url(self, absolute):
        match = _CHILDREN_ROUTE.fullmatch(urlsplit(absolute).path)
        if match is None:
            raise GraphError(code=self.prefix + "_foreign_parent")
        drive = unquote(match["drive"]) if match["drive"] is not None else None
        if drive is not None and self.drive_id is not None and drive != self.drive_id:
            raise GraphError(code=self.prefix + "_foreign_drive")
        route = self._route(match)
        if route == self.route:
            return
        # Graph may canonicalize root/path or a drive alias to drive/item IDs.
        # After the transition, do not allow another parent or alias switch.
        item_id = unquote(match["id"]) if match["id"] is not None else None
        if (drive is None or item_id is None or
                (self.parent_id is not None and item_id != self.parent_id) or self.canonical_item):
            raise GraphError(code=self.prefix + "_foreign_parent")
        self.drive_id, self.parent_id, self.route = drive, item_id, route
        self.canonical_item = True

    def check_item(self, item, code="invalid_collection_item"):
        checked_drive_item(item, self.drive_id, code)
        parent = item.get("parentReference", {})
        drive = parent.get("driveId")
        if drive is not None and self.drive_id is None:
            self.drive_id = drive
        if "id" in parent:
            parent_id = parent["id"]
            if not isinstance(parent_id, str) or not parent_id.strip():
                raise GraphError(code=code)
            if self.parent_id is not None and parent_id != self.parent_id:
                raise GraphError(code=self.prefix + "_foreign_parent")
            self.parent_id = parent_id


class GraphAPI:
    def __init__(self, parent):
        self.parent = parent
        self.session = requests.Session()
        # Separate sessions also prevent session-level Authorization from escaping.
        self.transfer_session = requests.Session()
        # Disable implicit .netrc authentication on both sessions.
        self.session.auth = _NoImplicitAuth()
        self.transfer_session.auth = _NoImplicitAuth()
        self.timeout = parent.request_timeout
        self.max_retries = getattr(parent, "max_retries", 3)

    def _url(self, resource):
        if not isinstance(resource, str) or not resource:
            raise ValueError("A Graph resource URL is required")
        base = self.parent.graph_url.rstrip("/")
        url = resource if urlsplit(resource).scheme or resource.startswith("//") else base + "/" + resource.lstrip("/")
        p, b = urlsplit(url), urlsplit(base)
        path = unquote(p.path)
        if (p.scheme != "https" or p.netloc.lower() != b.netloc.lower()
                or p.username or p.password or p.fragment
                or not p.path.startswith(b.path.rstrip("/") + "/")
                or any(part in (".", "..") for part in path.split("/")) or "\\" in path):
            raise ValueError("Graph resources must remain within the configured HTTPS Graph endpoint")
        return url

    @staticmethod
    def _capability(url):
        p = urlsplit(url)
        if p.scheme != "https" or not p.hostname or p.username or p.password or p.fragment:
            raise ValueError("Transfer URLs must be absolute HTTPS URLs without user information")
        return url

    @staticmethod
    def _delay(response, attempt):
        value = response.headers.get("Retry-After", "")
        try:
            seconds = float(value)
        except ValueError:
            try:
                when = parsedate_to_datetime(value)
                seconds = (when - datetime.now(timezone.utc)).total_seconds()
            except (ValueError, TypeError, OverflowError):
                seconds = 2 ** attempt
        if math.isfinite(seconds) and seconds > 60:
            return None  # Do not retry before the server permits it.
        return max(0, seconds) if math.isfinite(seconds) else min(2 ** attempt, 60)

    def _request(self, resource, method="GET", data=None, json=None, headers=None,
                 stream=False, preauthenticated=False, content_type=None, response_metadata=False):
        if response_metadata and stream:
            raise ValueError("response_metadata cannot be combined with streaming")
        if data is not None and json is not None:
            raise ValueError("Supply either data or json, not both")
        if isinstance(data, (dict, list)):
            json, data = data, None  # Legacy Graph callers used data= for JSON.
        method = method.upper()
        url = self._capability(resource) if preauthenticated else self._url(resource)
        session = self.transfer_session if preauthenticated else self.session
        request_headers = {k: v for k, v in (headers or {}).items()
                           if k.lower() not in ("authorization", "cookie", "proxy-authorization")}
        if not preauthenticated:
            request_headers["Authorization"] = "Bearer " + self.parent.get_token()
        if content_type:
            request_headers.setdefault("Content-Type", content_type)
        # Decode the response independently of the request body type.
        for attempt in range(self.max_retries + 1):
            try:
                response = session.request(method, url, headers=request_headers, data=data,
                                           json=json, stream=stream, timeout=self.timeout,
                                           allow_redirects=False)
            except requests.RequestException:
                # Mutations with lost responses have an ambiguous outcome: do not replay blindly.
                if method in ("GET", "HEAD") and attempt < self.max_retries:
                    time.sleep(min(2 ** attempt, 60))
                    continue
                raise GraphError(code="transport_error") from None
            if response.status_code == 429 or (response.status_code in (502, 503, 504) and method in ("GET", "HEAD")):
                if attempt < self.max_retries:
                    delay = self._delay(response, attempt)
                    response.close()
                    if delay is None:
                        raise GraphError(response.status_code, "retry_after_exceeds_budget")
                    time.sleep(delay)
                    continue
            if 300 <= response.status_code < 400:
                # Graph content endpoints return a short-lived preauthenticated download URL.
                location = response.headers.get("Location")
                response.close()
                if method != "GET" or not stream or not location:
                    raise GraphError(code="unexpected_redirect")
                return self._download_redirect(location, request_headers)
            if response.status_code >= 400:
                status = response.status_code
                response.close()
                raise GraphError(status)
            if stream:
                return response
            try:
                if response.status_code == 204 or not response.content:
                    body = {}
                elif "json" in response.headers.get("Content-Type", "").lower():
                    try:
                        body = response.json()
                    except ValueError:
                        raise GraphError(code="invalid_json_response") from None
                else:
                    body = response.content
                # Opt-in status without headers: signed capability/Location
                # values must not be exposed as routine diagnostic metadata.
                return {"status_code": response.status_code, "body": body} if response_metadata else body
            finally:
                response.close()

    def _download_redirect(self, url, headers):
        safe_headers = {k: v for k, v in headers.items() if k.lower() not in
                        ("authorization", "cookie", "proxy-authorization")}
        for _ in range(5):
            self._capability(url)
            try:
                response = self.transfer_session.request("GET", url, headers=safe_headers,
                    stream=True, timeout=self.timeout, allow_redirects=False)
            except requests.RequestException:
                raise GraphError(code="download_transport_error") from None
            if response.status_code in (301, 302, 303, 307, 308):
                url = response.headers.get("Location", "")
                response.close()
                continue
            if response.status_code >= 400:
                status = response.status_code
                response.close()
                raise GraphError(status)
            return response
        raise GraphError(code="redirect_limit")

    def graph_call(self, verbose=False, resource_url=None, method="GET", batch_requests=None,
                   data=None, json=None, headers=None, stream=False, batch_size=20,
                   content_type=None, *, preauthenticated=False, response_metadata=False):
        """Raise GraphError on failure; explicitly opt in to capability URL transfers."""
        if batch_requests is not None:
            if response_metadata:
                raise ValueError("response_metadata is for individual requests")
            return self._batch(batch_requests, batch_size)
        return self._request(resource_url, method, data, json, headers, stream,
                             preauthenticated, content_type, response_metadata)

    def _collection_start(self, resource):
        """Stabilize initial drive-child ordering; never rewrite a nextLink."""
        absolute = self._url(resource)
        p = urlsplit(absolute)
        children = _CHILDREN_ROUTE.fullmatch(p.path) is not None
        if not children:
            return resource, False
        pairs = parse_qsl(p.query, keep_blank_values=True)
        if any(k.lower() in ("$skiptoken", "$skip") for k, _ in pairs):
            return resource, True  # An explicit continuation is already opaque.
        if sum(k.lower() == "$orderby" for k, _ in pairs) > 1:
            raise ValueError("Supply only one collection orderby")
        changed = False
        for index, (key, value) in enumerate(pairs):
            if key.lower() == "$orderby":
                fields = [term.strip().split()[0].lower() for term in value.split(",") if term.strip()]
                if not fields:
                    raise ValueError("Collection orderby must not be empty")
                if "name" not in fields:
                    pairs[index] = (key, value + ",name asc")
                    changed = True
            elif key.lower() == "$select":
                if not any(field.strip().lower() in ("id", "*") for field in value.split(",")):
                    pairs[index] = (key, value + ",id")
                    changed = True
        if changed:
            return urlunsplit((p.scheme, p.netloc, p.path, urlencode(pairs), "")), True
        return resource, True

    def iter_collection(self, resource_url, *, _budget=None):
        """Follow nextLinks without hiding duplicates or replaying a scan.

        Drive-child scans get a name tie-break and identity validation. Delta and
        other collections retain their semantics (delta may repeat an item).
        A later error can follow previously yielded pages; consume successfully
        before committing work that depends on completeness. Internal callers can
        share an operation budget across collections, including empty pages.
        """
        resource_url, children = self._collection_start(resource_url)
        scope = (DriveChildrenScope(self._url(resource_url)) if children
                 else CollectionScope(self._url(resource_url)))
        seen, item_ids = set(), set()
        while resource_url:
            absolute = self._url(resource_url)
            scope.check_page_url(absolute)
            if absolute in seen:
                raise GraphError(code="pagination_cycle")
            seen.add(absolute)
            if _budget is not None:
                _budget.before_request()
            page = read_response(self.graph_call, resource_url=resource_url)
            if not isinstance(page, dict) or "error" in page or not isinstance(page.get("value"), list):
                raise GraphError(code="invalid_collection")
            if _budget is not None:
                _budget.consume_items(len(page["value"]))
            next_link = collection_next_link(page)
            if children:
                # Validate the entire current page before yielding any of it.
                for item in page["value"]:
                    scope.check_item(item)
                    if item["id"] in item_ids:
                        raise GraphError(code="collection_duplicate_item")
                    item_ids.add(item["id"])
            if next_link is not None:
                scope.check_page_url(self._url(next_link))
            yield from page["value"]
            resource_url = next_link

    def _batch(self, requests_, batch_size):
        if isinstance(batch_size, bool) or not isinstance(batch_size, int) or not 1 <= batch_size <= 20:
            raise ValueError("Graph batch_size must be between 1 and 20")
        ids = [str(r["id"]) for r in requests_]
        if len(set(ids)) != len(ids):
            raise ValueError("Batch request IDs must be unique")
        prepared = []
        for r in requests_:
            # Batch URLs are Graph-relative and must not contain authorization headers.
            if urlsplit(r["url"]).scheme or r["url"].startswith("//"):
                raise ValueError("Batch URLs must be Graph-relative")
            self._url(r["url"])
            if any(k.lower() in ("authorization", "cookie") for k in r.get("headers", {})):
                raise ValueError("Batch authorization headers are not accepted")
            prepared.append(dict(r, id=str(r["id"]), url="/" + r["url"].lstrip("/")))
        results = {}
        for start in range(0, len(prepared), batch_size):
            chunk = prepared[start:start + batch_size]
            payload = self._request("$batch", "POST", json={"requests": chunk})
            responses = payload.get("responses", []) if isinstance(payload, dict) else []
            expected = {r["id"] for r in chunk}
            if (not isinstance(responses, list) or not all(isinstance(r, dict) for r in responses)
                    or {str(r.get("id")) for r in responses} != expected or len(responses) != len(chunk)):
                raise GraphError(code="incomplete_batch")
            for r in responses:
                status = r.get("status")
                if isinstance(status, bool) or not isinstance(status, int) or not 200 <= status <= 599:
                    raise GraphError(code="invalid_batch_status")
                if 300 <= status < 400:
                    raise GraphError(status, "batch_subrequest_redirect")
                if status >= 400:
                    # Do not hide partial failures or replay non-idempotent subrequests.
                    raise GraphError(status, "batch_subrequest_failed")
                results[str(r["id"])] = r
        return results

    def close(self):
        self.session.close()
        self.transfer_session.close()
