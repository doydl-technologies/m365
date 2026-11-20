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
"""Mailbox identifiers and fail-closed reads; no tenant or auth assumptions."""
from urllib.parse import quote
from ._graph import CollectionScope, GraphAPI, GraphError, read_response


def mail_id(value):
    if not isinstance(value, str) or not value.strip() or any(ord(c) < 32 for c in value):
        raise ValueError("Mailbox resource IDs must be nonempty strings without controls")
    return quote(value, safe="")


def checked_mail_result(response, *, delta=False):
    if isinstance(response, dict) and "error" in response:
        status = response.get("status_code")
        if isinstance(status, bool) or not isinstance(status, int) or not 100 <= status <= 599:
            status = None
        code = "delta_resync_required" if delta and status == 410 else "mailbox_request_failed"
        # Never include the provider's error body, tokens or URLs in exceptions.
        raise GraphError(status, code)
    return response


def mail_read(request, *, delta=False, **kwargs):
    try:
        return checked_mail_result(read_response(request, **kwargs), delta=delta)
    except GraphError as exc:
        if delta and exc.status_code == 410:
            raise GraphError(410, "delta_resync_required") from None
        raise


def checked_mail_item(item, *, delta=False, require_id=True):
    """Validate fields consumed by queries while retaining legal null addresses."""
    if not isinstance(item, dict) or "error" in item or "deleted" in item:
        raise GraphError(code="invalid_mail_item")
    if require_id or "id" in item:
        if not isinstance(item.get("id"), str) or not item["id"].strip():
            raise GraphError(code="invalid_mail_item")
    if "@removed" in item and (not delta or not isinstance(item["@removed"], dict)):
        raise GraphError(code="invalid_mail_item")
    if "displayName" in item and (not isinstance(item["displayName"], str) or not item["displayName"].strip()):
        raise GraphError(code="invalid_mail_folder")
    for field in ("subject", "bodyPreview"):
        if item.get(field) is not None and not isinstance(item[field], str):
            raise GraphError(code="invalid_mail_item")
    for field in ("from", "sender"):
        value = item.get(field)
        if value is None:
            continue
        if not isinstance(value, dict):
            raise GraphError(code="invalid_mail_address")
        address = value.get("emailAddress")
        if address is None:
            continue
        if not isinstance(address, dict):
            raise GraphError(code="invalid_mail_address")
        if any(address.get(key) is not None and not isinstance(address[key], str) for key in ("name", "address")):
            raise GraphError(code="invalid_mail_address")
    return item


def checked_mail_page(response, *, delta=False, require_id=True):
    response = checked_mail_result(response, delta=delta)
    if not isinstance(response, dict) or not isinstance(response.get("value"), list):
        raise GraphError(code="invalid_mail_collection")
    for item in response["value"]:
        checked_mail_item(item, delta=delta, require_id=require_id)
    for key in ("@odata.nextLink", "@odata.deltaLink"):
        if key in response and (not isinstance(response[key], str) or not response[key].strip()
                                or any(ord(c) < 32 for c in response[key])):
            raise GraphError(code="invalid_mail_continuation")
    if delta and (("@odata.nextLink" in response) == ("@odata.deltaLink" in response)):
        raise GraphError(code="invalid_delta_cursor")
    return response


def mail_pages(request, endpoint, *, headers=None, auto_paginate=True, page_limit=None, require_id=True):
    """Yield valid pages; later failures raise, never become partial success.

    Explicit page_limit is a caller-requested subset. Otherwise a 200-page
    safety budget raises. A repeated opaque nextLink is always an error.
    """
    if page_limit is not None and (isinstance(page_limit, bool) or not isinstance(page_limit, int) or page_limit < 1):
        raise ValueError("page_limit must be a positive integer")
    if not isinstance(auto_paginate, bool):
        raise ValueError("auto_paginate must be a boolean")
    api = getattr(request, "__self__", None)
    resolve = api._url if isinstance(api, GraphAPI) else lambda url: url
    scope = CollectionScope(resolve(endpoint))
    seen, seen_items = set(), set()
    url = endpoint
    limit = page_limit if page_limit is not None else 200
    count = 0
    while url:
        if not isinstance(url, str) or not url:
            raise GraphError(code="invalid_mail_continuation")
        absolute = resolve(url)
        scope.check_page_url(absolute)
        if count >= limit:
            if page_limit is not None:
                return
            raise GraphError(code="mailbox_page_limit")
        if absolute in seen:
            raise GraphError(code="pagination_cycle")
        seen.add(absolute)
        page = checked_mail_page(mail_read(request, resource_url=url, headers=headers), require_id=require_id)
        for item in page["value"]:
            if "id" in item:
                if item["id"] in seen_items:
                    raise GraphError(code="mailbox_duplicate_item")
                seen_items.add(item["id"])
        next_link = page.get("@odata.nextLink")
        if next_link is not None:
            scope.check_page_url(resolve(next_link))
        count += 1
        yield page
        if not auto_paginate:
            return
        url = next_link
