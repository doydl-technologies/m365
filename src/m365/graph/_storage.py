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
"""Drive-scoped storage engine shared by OneDrive and SharePoint facades."""
from collections import deque
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
import math
import os
import re
import threading
from types import SimpleNamespace
from urllib.parse import unquote, urlsplit

from ._graph import (
    DriveChildrenScope, GraphAPI, GraphError, checked_drive_item,
    collection_next_link, drive_item_kind, drive_item_name, read_response,
)
from ._site_resolution import graph_id, drive_path
from ._local_paths import native_path, public_path
from .operations.doc_sync import Upload
from .operations.doc_transfer import Download, _child, _compress, _workers, _write_download

from ._storage_auth import _DEFAULT_GRAPH, _profile_provider, _legacy_provider


def _identifier(value, name):
    if not isinstance(value, str) or not value.strip() or any(ord(c) < 32 or ord(c) == 127 for c in value):
        raise ValueError(name + " must be a nonempty opaque identifier")
    try:
        value.encode("utf-8")
    except UnicodeEncodeError:
        raise ValueError(name + " must be valid Unicode") from None
    return value


def _name(value):
    if (not isinstance(value, str) or not value or value in (".", "..")
            or value.endswith((".", " ")) or any(c in value for c in '/\\:<>"|?*')
            or any(ord(c) < 32 or ord(c) == 127 for c in value)
            or re.match(r"^(CON|PRN|AUX|NUL|COM[1-9\u00b9\u00b2\u00b3]|LPT[1-9\u00b9\u00b2\u00b3])(?:\.|$)", value, re.I)):
        raise ValueError("A single safe OneDrive filename or folder name is required")
    try:
        if len(value.encode("utf-16-le")) > 510:
            raise ValueError("A OneDrive name exceeds 255 UTF-16 code units")
    except UnicodeEncodeError:
        raise ValueError("OneDrive names must be valid Unicode") from None
    return value


def _parts(path):
    if not isinstance(path, str):
        raise ValueError("Paths must be relative strings, not URLs, IDs or metadata")
    if path == "":
        return []
    return [_name(part) for part in path.split("/")]


def _integer(value, name, minimum=1):
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ValueError(name + " must be an integer >= " + str(minimum))
    return value


def _options(drive_id, folder_id=None, *, graph_url=_DEFAULT_GRAPH,
             request_timeout=(10, 120), max_retries=3, max_requests=200,
             max_items=5000, max_depth=30):
    _identifier(drive_id, "drive_id")
    if folder_id is not None:
        _identifier(folder_id, "folder_id")
    if not isinstance(graph_url, str):
        raise ValueError("graph_url must be a trusted HTTPS Graph endpoint")
    try:
        parsed = urlsplit(graph_url)
        valid = (parsed.scheme == "https" and parsed.hostname and not parsed.username
                 and not parsed.password and not parsed.query and not parsed.fragment
                 and parsed.path.rstrip("/") and "\\" not in graph_url
                 and not any(ord(c) < 33 for c in graph_url)
                 and not any(p in (".", "..") for p in parsed.path.split("/")))
        parsed.port
        graph_url.encode("utf-8")
        unquote(parsed.path, errors="strict")
        if re.search(r"%(?![0-9a-fA-F]{2})", parsed.path):
            valid = False
    except ValueError:
        valid = False
    if not valid:
        raise ValueError("graph_url must be a trusted HTTPS Graph endpoint")
    values = request_timeout if isinstance(request_timeout, (list, tuple)) else (request_timeout,)
    if (isinstance(request_timeout, (tuple, list)) and len(values) != 2) or any(isinstance(v, bool) or not isinstance(v, (int, float))
                                      or not math.isfinite(v) or v <= 0 for v in values):
        raise ValueError("request_timeout must be positive finite seconds or a connect/read pair")
    _integer(max_retries, "max_retries", 0)
    if max_retries > 10:
        raise ValueError("max_retries must be <= 10")
    _integer(max_requests, "max_requests")
    _integer(max_items, "max_items")
    _integer(max_depth, "max_depth", 0)
    return dict(drive_id=drive_id, folder_id=folder_id, graph_url=graph_url.rstrip("/"),
                request_timeout=tuple(values) if isinstance(request_timeout, (tuple, list)) else request_timeout, max_retries=max_retries,
                max_requests=max_requests, max_items=max_items, max_depth=max_depth)


class _Budget:
    def __init__(self):
        self.requests = 0
        self.items = 0


class _StorageClient:
    """Bind to one drive and optionally one granted folder.

    Paths are relative to the selected folder (or drive root when folder_id=None).
    No drive discovery, share-link redemption or permission changes are attempted.
    Returned metadata cannot retarget this client. Server grants remain the
    authorization boundary; this facade does not restrict reuse of the provider.
    """

    def __init__(self, token_engine, drive_id, *, folder_id=None, graph_url=_DEFAULT_GRAPH,
                 request_timeout=(10, 120), max_retries=3, max_requests=200,
                 max_items=5000, max_depth=30):
        options = _options(drive_id, folder_id, graph_url=graph_url,
                           request_timeout=request_timeout, max_retries=max_retries,
                           max_requests=max_requests, max_items=max_items, max_depth=max_depth)
        if token_engine is None:
            raise ValueError("Supply an authenticated token provider")
        self._provider = token_engine
        self._options = options
        self._lock = threading.RLock()
        self._closed = False
        self._client = SimpleNamespace(graph_url=options["graph_url"], request_timeout=options["request_timeout"],
                                       max_retries=max_retries, drive_id=drive_id, get_token=self._get_token)
        self._api = GraphAPI(self._client)
        self._client.api = self._api
        self._prefix = "drives/" + graph_id(drive_id)
        endpoint = self._prefix + ("/items/" + graph_id(folder_id) if folder_id is not None else "/root")
        try:
            root = self._read(endpoint, _Budget())
            self._check_item(root, kind="folder")
            if folder_id is not None and root["id"] != folder_id:
                raise GraphError(code="onedrive_root_id_mismatch")
            self._root_id = root["id"]
        except Exception:
            self.close()
            raise

    @classmethod
    def from_profile(cls, profile, drive_id, *, folder_id=None, scopes=None,
                     allow_interactive=False, **client_options):
        """Load an existing profile without creating credentials or granting access."""
        options = _options(drive_id, folder_id, **client_options)
        auth = _profile_provider(profile, scopes=scopes, allow_interactive=allow_interactive,
                                 graph_url=options["graph_url"])
        return cls(auth, **options)

    @classmethod
    def _from_builder(cls, drive_id, folder_id, build, build_kwargs, *, scopes=None,
                      token_path=None, authority=None, **client_options):
        from azauthlib import Authentication
        options = _options(drive_id, folder_id, **client_options)
        auth = _legacy_provider(Authentication, scopes=scopes, token_path=token_path,
                                authority=authority, build=build, build_kwargs=build_kwargs)
        return cls(auth, **options)

    @classmethod
    def from_default_env(cls, drive_id, *, folder_id=None, scopes=None,
                         token_path=None, authority=None, **client_options):
        """Use the legacy default builder and silent delegated authentication."""
        return cls._from_builder(drive_id, folder_id, "Default", {}, scopes=scopes,
                                 token_path=token_path, authority=authority, **client_options)

    @classmethod
    def from_env_file(cls, drive_id, env_path, *, folder_id=None, key_map=None,
                      scopes=None, token_path=None, authority=None, **client_options):
        """Load explicit environment-file credentials; never prompt automatically."""
        return cls._from_builder(drive_id, folder_id, "WithEnvFile",
                                 {"env_path": env_path, "key_map": key_map}, scopes=scopes,
                                 token_path=token_path, authority=authority, **client_options)

    @classmethod
    def from_os_env(cls, drive_id, *, folder_id=None, key_map=None, scopes=None,
                    token_path=None, authority=None, **client_options):
        """Use process-environment credentials and the legacy silent auth contract."""
        return cls._from_builder(drive_id, folder_id, "WithOSEnv", {"key_map": key_map},
                                 scopes=scopes, token_path=token_path,
                                 authority=authority, **client_options)

    @classmethod
    def from_entry(cls, drive_id, *, folder_id=None, client_id=None, tenant_id=None,
                   client_secret=None, scopes=None, token_path=None, authority=None,
                   **client_options):
        """Build explicit credentials; use from_profile for app-only flow selection."""
        return cls._from_builder(drive_id, folder_id, "WithEntry",
                                 {"client_id": client_id, "tenant_id": tenant_id,
                                  "client_secret": client_secret}, scopes=scopes,
                                 token_path=token_path, authority=authority, **client_options)

    @property
    def drive_id(self):
        return self._options["drive_id"]

    @property
    def folder_id(self):
        """The resolved immutable anchor ID, including a resolved drive root."""
        return self._root_id

    def _ensure_open(self):
        if self._closed:
            raise ValueError("Drive client is closed")

    def _get_token(self):
        self._ensure_open()
        with self._lock:
            try:
                token = self._provider.access_token
            except AttributeError:
                raise ValueError("token_engine must provide access_token; pass an authenticated provider, not a raw token") from None
        if not isinstance(token, str) or not token.strip() or any(ord(c) < 32 or ord(c) == 127 for c in token):
            raise GraphError(code="missing_access_token")
        return token

    def _read(self, resource, budget):
        self._ensure_open()
        if budget.requests >= self._options["max_requests"]:
            raise GraphError(code="onedrive_request_limit")
        budget.requests += 1
        return read_response(self._api.graph_call, resource_url=resource)

    def _check_item(self, item, *, parent=None, name=None, kind=None):
        checked_drive_item(item, self.drive_id, "invalid_onedrive_item")
        if "remoteItem" in item:
            raise GraphError(code="onedrive_remote_item")
        try:
            _identifier(item["id"], "item ID")
        except ValueError:
            raise GraphError(code="invalid_onedrive_item") from None
        actual_kind = drive_item_kind(item)
        if actual_kind not in ("file", "folder") or kind is not None and actual_kind != kind:
            raise GraphError(code="onedrive_item_kind_mismatch")
        try:
            _name(drive_item_name(item))
        except ValueError:
            raise GraphError(code="invalid_onedrive_name") from None
        if parent is not None and item.get("parentReference", {}).get("id") != parent:
            raise GraphError(code="onedrive_parent_mismatch")
        if name is not None and item["name"].casefold() != name.casefold():
            raise GraphError(code="onedrive_name_mismatch")
        if actual_kind == "file":
            size = item.get("size")
            if isinstance(size, bool) or not isinstance(size, int) or size < 0:
                raise GraphError(code="invalid_onedrive_size")
        return item

    def _page_url(self, resource):
        absolute = self._api._url(resource)
        path = urlsplit(absolute).path
        try:
            unquote(path, errors="strict")
        except UnicodeDecodeError:
            raise GraphError(code="onedrive_invalid_continuation_encoding") from None
        if re.search(r"%(?![0-9a-fA-F]{2})", path):
            raise GraphError(code="onedrive_invalid_continuation_encoding")
        return absolute

    def _children(self, folder, budget, *, level=0, ancestors=()):
        if drive_item_kind(folder) != "folder":
            raise ValueError("The selected path is not a folder")
        endpoint = self._prefix + "/items/" + graph_id(folder["id"]) + "/children?$top=200"
        scope = DriveChildrenScope(self._page_url(endpoint), prefix="onedrive")
        pages, ids, names = set(), set(), set()
        while endpoint:
            absolute = self._page_url(endpoint)
            scope.check_page_url(absolute)
            if absolute in pages:
                raise GraphError(code="onedrive_pagination_cycle")
            pages.add(absolute)
            page = self._read(endpoint, budget)
            if not isinstance(page, dict) or "error" in page or not isinstance(page.get("value"), list):
                raise GraphError(code="invalid_onedrive_page")
            if page["value"] and level >= self._options["max_depth"]:
                raise GraphError(code="onedrive_depth_limit")
            next_link = collection_next_link(page)
            if next_link is not None:
                scope.check_page_url(self._page_url(next_link))
            for item in page["value"]:
                self._check_item(item, parent=folder["id"])
                if (item["id"] in (folder["id"], self._root_id)
                        or item["id"] in ids or item["id"] in ancestors):
                    raise GraphError(code="onedrive_duplicate_item")
                if item["name"].casefold() in names:
                    raise GraphError(code="onedrive_name_collision")
                ids.add(item["id"])
                names.add(item["name"].casefold())
                budget.items += 1
                if budget.items > self._options["max_items"]:
                    raise GraphError(code="onedrive_item_limit")
            yield from page["value"]
            endpoint = next_link

    def _walk(self, folder, budget, recursive, *, level=0, ancestors=()):
        queue = deque([(folder, "", level)])
        seen = {folder["id"], self._root_id} | set(ancestors)
        while queue:
            parent, relative, depth = queue.popleft()
            for item in self._children(parent, budget, level=depth, ancestors=seen):
                if item["id"] in seen:
                    raise GraphError(code="onedrive_identity_cycle")
                seen.add(item["id"])
                path = relative + "/" + item["name"] if relative else item["name"]
                yield item, path
                if recursive and drive_item_kind(item) == "folder":
                    queue.append((item, path, depth + 1))

    def get_item(self, path=""):
        """Read metadata for a relative path; missing/forbidden paths raise GraphError."""
        parts = _parts(path)
        return deepcopy(self._resolve(parts, _Budget()))

    def list_children(self, folder_path=""):
        """List one folder completely; later errors never return partial success."""
        parts = _parts(folder_path)
        budget = _Budget()
        folder, ancestors = self._resolve_with_ancestors(parts, budget)
        return deepcopy(list(self._children(folder, budget, level=len(parts), ancestors=ancestors)))

    @staticmethod
    def _search_term(term):
        if not isinstance(term, str) or not term or any(ord(c) < 32 or ord(c) == 127 for c in term):
            raise ValueError("Search terms must be nonempty strings without controls")
        try:
            term.encode("utf-8")
        except UnicodeEncodeError:
            raise ValueError("Search terms must be valid Unicode") from None
        return term.casefold()

    def _find(self, term, kind, folder_path, exact_match, recursive, limit, sort_order=None):
        query = self._search_term(term)
        parts = _parts(folder_path)
        if not isinstance(exact_match, bool) or not isinstance(recursive, bool):
            raise ValueError("exact_match and recursive must be booleans")
        if sort_order not in (None, "recent", "oldest", "name", "name_desc"):
            raise ValueError("sort_order must be recent, oldest, name, name_desc or None")
        if limit is not None:
            _integer(limit, "limit", 0)
        self._ensure_open()
        if limit == 0:
            return []
        budget = _Budget()
        folder, ancestors = self._resolve_with_ancestors(parts, budget)
        if drive_item_kind(folder) != "folder":
            raise ValueError("Search must start at a folder")
        found = []
        for item, _ in self._walk(folder, budget, recursive, level=len(parts), ancestors=ancestors):
            name = item["name"].casefold()
            if drive_item_kind(item) == kind and (name == query if exact_match else query in name):
                found.append(deepcopy(item))
                if sort_order is None and limit is not None and len(found) >= limit:
                    return found
        if sort_order in ("recent", "oldest"):
            from datetime import timezone
            from ._core_dates import parse_graph_datetime
            def created(item):
                if not item.get("createdDateTime"):
                    raise GraphError(code="missing_created_datetime")
                try:
                    return parse_graph_datetime(item["createdDateTime"]).astimezone(timezone.utc)
                except (TypeError, ValueError, AttributeError, OverflowError):
                    raise GraphError(code="invalid_created_datetime") from None
            found.sort(key=created, reverse=sort_order == "recent")
        elif sort_order is not None:
            found.sort(key=lambda item: item["name"].casefold(), reverse=sort_order == "name_desc")
        return found if limit is None else found[:limit]

    def find_files(self, term, *, folder_path="", exact_match=False, recursive=True,
                   limit=None, sort_order=None):
        """Search filenames only; sorting scans the complete bounded population.

        recent/oldest sort by createdDateTime, matching Site's legacy sort field.
        Unlike the legacy helper, sorting alone does not imply a one-item limit.
        """
        return self._find(term, "file", folder_path, exact_match, recursive, limit, sort_order)

    def find_folders(self, term, *, folder_path="", exact_match=False, recursive=True,
                     limit=None, sort_order=None):
        return self._find(term, "folder", folder_path, exact_match, recursive, limit, sort_order)

    def find_file_folder(self, file_term, folder_term, *, folder_path="",
                         exact_match=True, recursive=True):
        """Resolve one file and one folder in one complete traversal or raise."""
        queries = {"file": self._search_term(file_term), "folder": self._search_term(folder_term)}
        parts = _parts(folder_path)
        if not isinstance(exact_match, bool) or not isinstance(recursive, bool):
            raise ValueError("exact_match and recursive must be booleans")
        budget = _Budget()
        folder, ancestors = self._resolve_with_ancestors(parts, budget)
        if drive_item_kind(folder) != "folder":
            raise ValueError("Search must start at a folder")
        matches = {"file": [], "folder": []}
        for item, _ in self._walk(folder, budget, recursive, level=len(parts), ancestors=ancestors):
            kind = drive_item_kind(item)
            name, query = item["name"].casefold(), queries[kind]
            if (name == query if exact_match else query in name):
                matches[kind].append(deepcopy(item))
        if len(matches["file"]) != 1 or len(matches["folder"]) != 1:
            raise ValueError("Exactly one file and one destination folder must resolve")
        return matches["file"] + matches["folder"]

    @staticmethod
    def _etag(value):
        if (not isinstance(value, str) or not value.strip() or value == "*"
                or any(ord(c) < 33 or ord(c) > 126 for c in value)):
            raise ValueError("if_match must be a nonempty ASCII ETag, not a wildcard")
        return value

    def move(self, path, destination_folder="", new_name=None, *, if_match=None):
        """Move within this anchor/drive; never overwrite or replay an unknown commit.

        The source ETag is required (from metadata or explicit if_match). Tree
        resolution is not server snapshot isolation; grants remain authoritative.
        A contradictory final response can follow a committed move. Inspect state
        before retrying; this client never attempts an automatic inverse move.
        """
        parts, target = _parts(path), _parts(destination_folder)
        if not parts:
            raise ValueError("The bound anchor cannot be moved or renamed")
        if len(target) >= self._options["max_depth"]:
            raise ValueError("Move destination exceeds max_depth")
        if new_name is not None:
            _name(new_name)
        if if_match is not None:
            self._etag(if_match)
        budget = _Budget()
        item, source_ancestors = self._resolve_with_ancestors(parts, budget)
        parent, target_ancestors = self._resolve_with_ancestors(target, budget)
        if drive_item_kind(parent) != "folder":
            raise ValueError("Move destination must be a folder")
        if item["id"] in target_ancestors:
            raise ValueError("A folder cannot be moved into itself or its descendants")
        name = item["name"] if new_name is None else new_name
        old_parent = item["parentReference"]["id"]
        if parent["id"] == old_parent and name == item["name"]:
            if if_match is not None and item.get("eTag") != if_match:
                raise GraphError(412, "storage_etag_mismatch")
            return deepcopy(item)
        if if_match is None:
            try:
                if_match = self._etag(item.get("eTag"))
            except ValueError:
                raise GraphError(code="storage_missing_or_invalid_etag") from None
        # Resolve conflicts locally for an early error, then also request fail
        # behavior from Graph. There is no fallback to replace or auto-rename.
        for existing in self._children(parent, budget, level=len(target), ancestors=target_ancestors):
            if existing["id"] == item["id"] and parent["id"] != old_parent:
                raise GraphError(code="storage_source_context_changed")
            if existing["name"].casefold() == name.casefold() and existing["id"] != item["id"]:
                raise FileExistsError("A destination item already has that name")
        if drive_item_kind(item) == "folder":
            for child, relative in self._walk(item, budget, True, level=len(parts), ancestors=source_ancestors):
                if child["id"] in source_ancestors:
                    raise GraphError(code="storage_identity_cycle")
                if len(target) + 1 + len(relative.split("/")) > self._options["max_depth"]:
                    raise GraphError(code="onedrive_depth_limit")
        reply = self._api.graph_call(
            resource_url=self._prefix + "/items/" + graph_id(item["id"])
                         + "?@microsoft.graph.conflictBehavior=fail",
            method="PATCH", json={"parentReference": {"id": parent["id"]}, "name": name},
            headers={"If-Match": if_match}, response_metadata=True)
        if not isinstance(reply, dict) or type(reply.get("status_code")) is not int or reply["status_code"] != 200:
            raise GraphError(code="storage_move_not_confirmed")
        result = self._check_item(reply.get("body"), parent=parent["id"], name=name,
                                  kind=drive_item_kind(item))
        if (result["id"] != item["id"] or result.get("parentReference", {}).get("driveId") != self.drive_id
                or drive_item_kind(item) == "file" and result["size"] != item["size"]):
            raise GraphError(code="storage_move_confirmation_mismatch")
        return deepcopy(result)

    def rename(self, path, new_name, *, if_match=None):
        """Rename in the current parent; new_name is complete, with no suffix inference."""
        parts = _parts(path)
        _name(new_name)
        return self.move(path, "/".join(parts[:-1]), new_name, if_match=if_match)

    def upload(self, source, folder_path="", new_file_name=None, *, conflict_behavior="fail",
               chunk_size=10 * 327680):
        """Upload a nonempty file/BytesIO; overwrite only with explicit 'replace'.

        Metadata lookups are read-only. Session/chunk work uses the shared uploader.
        A lost final acknowledgement is ambiguous and raises instead of replaying a
        possible commit. The source position is restored for BytesIO inputs.
        """
        parts = _parts(folder_path)
        if len(parts) >= self._options["max_depth"]:
            raise ValueError("Upload destination exceeds max_depth")
        if new_file_name is not None:
            _name(new_file_name)
        uploader = Upload(self._client, conflict_behavior=conflict_behavior, chunk_size=chunk_size)
        parent, ancestors = self._resolve_with_ancestors(parts, _Budget())
        if drive_item_kind(parent) != "folder":
            raise ValueError("Upload destination must be a folder")
        source_name = new_file_name if new_file_name is not None else (
            os.path.basename(os.fspath(source)) if isinstance(source, (str, os.PathLike)) else None)
        _name(source_name)
        client = SimpleNamespace(drive_id=self.drive_id, api=self._api, upload_parent_id=parent["id"])
        uploader.sharepoint_instance = client
        result = uploader.upload(source, "", source_name)
        self._check_item(result, parent=parent["id"], kind="file")
        if result["id"] in ancestors:
            raise GraphError(code="onedrive_identity_cycle")
        return deepcopy(result)

    def download_file(self, path, directory):
        """Download one file atomically without overwriting an existing local file."""
        parts = _parts(path)
        item = self._resolve(parts, _Budget())
        if drive_item_kind(item) != "file":
            raise ValueError("The selected path is not a file")
        return Download(self._client, max_workers=1).file(directory, item)[0]

    def download_folder(self, folder_path, directory, *, compress=False, max_workers=4):
        """Plan a bounded tree, then download files; optional ZIP preserves originals.

        Each file and ZIP is published atomically without clobbering. A whole
        folder is not transactional: completed files remain if a later transfer
        fails. No unrelated local content is removed during recovery.
        """
        parts = _parts(folder_path)
        _workers(max_workers)
        if not isinstance(compress, bool):
            raise ValueError("compress must be a boolean")
        budget = _Budget()
        folder, ancestors = self._resolve_with_ancestors(parts, budget)
        if drive_item_kind(folder) != "folder":
            raise ValueError("The selected path is not a folder")
        # Validate the complete metadata plan before creating a local tree.
        plan = list(self._walk(folder, budget, True, level=len(parts), ancestors=ancestors))
        parent = native_path(directory)
        root = _child(parent, folder["name"])
        if compress and (root.with_name(root.name + ".zip").exists()
                         or root.with_name(root.name + ".zip").is_symlink()):
            raise FileExistsError("Download ZIP destination already exists")
        paths = [(item, root.joinpath(*relative.split("/"))) for item, relative in plan]
        keys = [str(path).casefold() for _, path in paths]
        if len(set(keys)) != len(keys):
            raise ValueError("Remote paths collide on the local filesystem")
        parent.mkdir(parents=True, exist_ok=True)
        root.mkdir()
        directories = [path for item, path in paths if drive_item_kind(item) == "folder"]
        for path in sorted(directories, key=lambda p: len(p.parts)):
            path.mkdir()
        tasks = [(item, path) for item, path in paths if drive_item_kind(item) == "file"]
        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            futures = [pool.submit(_write_download, self._api, item, path, self.drive_id) for item, path in tasks]
            completed = [f.result() for f in futures]
        if compress:
            # Explicit directory entries preserve empty folders, including an
            # empty selected root. Downloaded originals are never deleted.
            return public_path(_compress(root, completed, True, directories=directories or ([root] if not tasks else [])))
        long_paths = any(str(public_path(p)).startswith("\\\\?\\") for _, p in paths)
        return public_path(root, keep_extended=long_paths)

    def create_folder(self, name, folder_path=""):
        """Create one folder with conflictBehavior=fail; never merge into a namesake."""
        _name(name)
        parts = _parts(folder_path)
        if len(parts) >= self._options["max_depth"]:
            raise ValueError("New folder exceeds max_depth")
        parent, ancestors = self._resolve_with_ancestors(parts, _Budget())
        if drive_item_kind(parent) != "folder":
            raise ValueError("The destination path is not a folder")
        reply = self._api.graph_call(
            resource_url=self._prefix + "/items/" + graph_id(parent["id"]) + "/children",
            method="POST", json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
            response_metadata=True)
        if not isinstance(reply, dict) or type(reply.get("status_code")) is not int or reply["status_code"] != 201:
            raise GraphError(code="onedrive_folder_not_created")
        item = reply.get("body")
        self._check_item(item, parent=parent["id"], name=name, kind="folder")
        if item["id"] in ancestors:
            raise GraphError(code="onedrive_identity_cycle")
        return deepcopy(item)

    def close(self):
        if not self._closed:
            self._closed = True
            self._api.close()

    def __enter__(self):
        self._ensure_open()
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.close()
        return False


    def _resolve(self, parts, budget):
        return self._resolve_with_ancestors(parts, budget)[0]

    def _resolve_with_ancestors(self, parts, budget):
        self._ensure_open()
        if len(parts) > self._options["max_depth"]:
            raise ValueError("Path exceeds max_depth")
        root = self._read(self._prefix + "/items/" + graph_id(self._root_id), budget)
        self._check_item(root, kind="folder")
        if root["id"] != self._root_id:
            raise GraphError(code="onedrive_root_id_mismatch")
        item, seen = root, {root["id"]}
        for part in parts:
            if drive_item_kind(item) != "folder":
                raise ValueError("A path component is not a folder")
            parent = item["id"]
            item = self._read(self._prefix + "/items/" + graph_id(parent) + ":/" + drive_path(part), budget)
            self._check_item(item, parent=parent, name=part)
            if item["id"] in seen:
                raise GraphError(code="onedrive_identity_cycle")
            seen.add(item["id"])
        return item, seen


__all__ = []
