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
"""Bounded, selected-drive traversal shared by name and document-GUID lookup."""
from collections import deque
from ._graph import (
    DriveChildrenScope, GraphError, collection_next_link, drive_item_kind, drive_item_name,
    read_response,
)
from ._site_resolution import graph_id


def walk_drive(client, *, root=None, document_ids=False, _budget=None):
    """Yield (item, relative_path); paths are known when starting at library root.

    Limits cover the whole traversal. Shortcuts are never followed; no partial
    traversal should be used to establish a unique organization-independent match.
    """
    prefix = "drives/" + graph_id(client.drive_id)
    endpoint = prefix + ("/items/" + graph_id(root) if root else "/root")
    queue = deque([(endpoint, 0, None if root else "")])
    seen_pages, seen_items = set(), set()
    requests = items = 0
    while queue:
        parent, depth, parent_path = queue.popleft()
        url = parent + "/children?$top=200"
        if document_ids:
            url += "&$select=id,name,file,folder,parentReference,sharepointIds"
        scope = DriveChildrenScope(client.api._url(url), prefix="search")
        while url:
            if requests >= getattr(client, "search_max_requests", 200):
                raise GraphError(code="search_request_limit")
            absolute = client.api._url(url)
            scope.check_page_url(absolute)
            if absolute in seen_pages:
                raise GraphError(code="search_pagination_cycle")
            seen_pages.add(absolute)
            if _budget is not None:
                _budget.before_request()
            page = read_response(client.api.graph_call, resource_url=url)
            requests += 1
            if not isinstance(page, dict) or "error" in page or not isinstance(page.get("value"), list):
                raise GraphError(code="invalid_search_page")
            if _budget is not None:
                _budget.consume_items(len(page["value"]))
            next_link = collection_next_link(page, "invalid_search_continuation")
            if next_link is not None:
                scope.check_page_url(client.api._url(next_link))
            for item in page["value"]:
                items += 1
                if items > getattr(client, "search_max_items", 5000):
                    raise GraphError(code="search_item_limit")
                if not isinstance(item, dict) or not isinstance(item.get("id"), str) or not item["id"]:
                    raise GraphError(code="invalid_search_item")
                if "remoteItem" in item:
                    continue
                scope.check_item(item, "invalid_search_item")
                if item["id"] in seen_items:
                    raise GraphError(code="search_duplicate_item")
                seen_items.add(item["id"])
                name = drive_item_name(item, "invalid_search_item")
                kind = drive_item_kind(item, "invalid_search_item")
                path = None if parent_path is None else (parent_path + "/" if parent_path else "") + name
                yield item, path
                if kind == "folder":
                    if depth >= getattr(client, "search_max_depth", 30):
                        raise GraphError(code="search_depth_limit")
                    queue.append((prefix + "/items/" + graph_id(item["id"]), depth + 1, path))
            url = next_link
