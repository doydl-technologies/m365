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
"""Search and path lookups scoped exclusively to the selected document library."""
from datetime import datetime, timezone
import re
from urllib.parse import quote, unquote

from .._graph import (
    GraphError, _CollectionBudget, checked_drive_item, drive_item_kind, drive_item_name, read_response,
)
from .._site_resolution import graph_id, drive_path, library_relative
from .._core_dates import parse_graph_datetime
from .._drive_walk import walk_drive


def _search_budget(client):
    return _CollectionBudget(getattr(client, "search_max_requests", 200),
                             getattr(client, "search_max_items", 5000), prefix="search")


def _enumerate(client, query, kind, exact, budget):
    """Bounded name-only traversal; never return a partial/ambiguous success."""
    results = []
    for item, _ in walk_drive(client, root=getattr(client, "search_root", None), _budget=budget):
        name = item.get("name", "").casefold()
        match = name == query.casefold() if exact else query.casefold() in name
        if kind in item and match:
            results.append(item)
    return results


def _lookup(client, query, kind, exact=True, *, _budget=None):
    if not isinstance(exact, bool):
        raise ValueError("exact_match must be a boolean")
    if not isinstance(query, str) or not query:
        raise ValueError("Search terms must be nonempty strings")
    budget = _search_budget(client) if _budget is None else _budget
    prefix = "drives/" + graph_id(client.drive_id)
    if "/" in query:
        path = library_relative(client, query, _budget=budget)
        resource = prefix + ("/root:/" + drive_path(path) if path else "/root")
        try:
            budget.before_request()
            item = read_response(client.api.graph_call, resource_url=resource)
        except GraphError as exc:
            if exc.status_code == 404:
                return []
            raise
        budget.consume_items(1)
        checked_drive_item(item, client.drive_id, "invalid_search_item")
        if "remoteItem" in item:
            raise GraphError(code="unsupported_remote_item")
        if path:
            folder, _, name = path.rpartition("/")
            if drive_item_name(item, "invalid_search_item").casefold() != name.casefold():
                raise GraphError(code="search_name_mismatch")
            parent_path = item.get("parentReference", {}).get("path", "")
            match = re.fullmatch(r"/(?:drives/[^/]+|drive)/root:(/.*)?", parent_path)
            if match and unquote(match[1] or "").strip("/").casefold() != folder.casefold():
                raise GraphError(code="search_parent_mismatch")
        return [item] if drive_item_kind(item, "invalid_search_item") == kind else []
    if getattr(client, "search_mode", "enumerate") == "enumerate":
        return _enumerate(client, query, kind, exact, budget)
    # Indexed search is explicit opt-in; it needs a compatible permission grant.
    resource = prefix + "/root/search(q='{}')".format(quote(query.replace("'", "''"), safe=""))
    items = client.api.iter_collection(resource, _budget=budget)
    results = {}
    seen_ids = set()
    for item in items:
        checked_drive_item(item, client.drive_id, "invalid_search_item")
        if item["id"] in seen_ids:
            raise GraphError(code="search_duplicate_item")
        seen_ids.add(item["id"])
        if "remoteItem" in item:
            continue
        if drive_item_kind(item, "invalid_search_item") == kind:
            name = drive_item_name(item, "invalid_search_item")
            if not (name.casefold() == query.casefold() if exact else query.casefold() in name.casefold()):
                continue
            results[item["id"]] = item
    return list(results.values())


class Folder:
    def __init__(self, sharepoint_instance):
        self.client = sharepoint_instance

    def find(self, search_term, include_children=False, exact_match=True):
        if not isinstance(include_children, bool):
            raise ValueError("include_children must be a boolean")
        budget = _search_budget(self.client)
        results = _lookup(self.client, search_term, "folder", exact_match, _budget=budget)
        if include_children:
            for item in results:
                item["children"] = list(self.client.api.iter_collection(
                    "drives/{}/items/{}/children".format(graph_id(self.client.drive_id), graph_id(item["id"])),
                    _budget=budget))
        return results


class File:
    def __init__(self, sharepoint_instance):
        self.client = sharepoint_instance

    def find(self, search_term, exact_match=True, sort_order=None, num_files=None):
        if sort_order not in (None, "recent", "oldest"):
            raise ValueError("sort_order must be recent or oldest")
        if num_files is not None and (isinstance(num_files, bool) or not isinstance(num_files, int) or num_files < 1):
            raise ValueError("num_files must be a positive integer")
        results = _lookup(self.client, search_term, "file", exact_match)
        if sort_order:
            def created(item):
                value = item.get("createdDateTime")
                if not value:
                    raise GraphError(code="missing_created_datetime")
                try:
                    return parse_graph_datetime(value).astimezone(timezone.utc)
                except (TypeError, ValueError, AttributeError, OverflowError):
                    raise GraphError(code="invalid_created_datetime") from None
            results.sort(key=created, reverse=sort_order == "recent")
            results = results[:num_files or 1]
        elif num_files is not None:
            results = results[:num_files]
        return results


class FileFolder:
    def __init__(self, sharepoint_instance):
        self.client = sharepoint_instance

    def find(self, file_search_term, folder_search_term):
        budget = _search_budget(self.client)
        files = _lookup(self.client, file_search_term, "file", _budget=budget)
        folders = _lookup(self.client, folder_search_term, "folder", _budget=budget)
        if len(files) != 1 or len(folders) != 1:
            raise ValueError("Exactly one file and one destination folder must resolve")
        return files + folders


__all__ = ["File", "Folder", "FileFolder"]
