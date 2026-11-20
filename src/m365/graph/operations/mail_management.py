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

from copy import deepcopy
from typing import Dict, List, Optional, Sequence, Union
import logging
from urllib.parse import quote
from .._graph import GraphError
from .._mail_common import mail_id, checked_mail_result, checked_mail_page, checked_mail_item, mail_pages, mail_read

try:
    from .._preprocessing import filters
    from .._preprocessing import lists
except (ImportError, ModuleNotFoundError):
    from .._preprocessing import filters
    from .._preprocessing import lists

logger = logging.getLogger(__name__)

def _folder_reference(folder):
    if folder is None:
        return None
    value = folder.get("id") if isinstance(folder, dict) and "id" in folder else (
        folder.get("displayName") if isinstance(folder, dict) else folder)
    if not isinstance(value, str) or not value.strip():
        raise ValueError("folder must be a nonempty name/ID, a folder record, or None")
    value = value.strip()
    mail_id(value)
    return value

def _mail_folder_endpoint(mailbox_instance, folder: Optional[Union[str, Dict[str, str]]]) -> str:
    """
    Return a Graph messages endpoint for a mailbox folder.

    Notes
    -----
    - For performance, this intentionally does NOT resolve friendly folder names
      by crawling the mailbox tree.
    - Microsoft Graph accepts well-known folder names like "inbox" directly in
      /mailFolders/{folder}/messages.
    - Explicit folder IDs also work.
    """
    base = mailbox_instance.base_resource
    folder_str = _folder_reference(folder)
    if folder_str is None: return f"{base}/messages"
    return f"{base}/mailFolders/{mail_id(folder_str)}/messages"

def _count(value, name, default=None):
    if value is None:
        return default
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValueError(name + " must be a nonnegative integer or None")
    return value


def _clamp_top(top: Optional[int], default: Optional[int] = None) -> Optional[int]:
    top = _count(top, "top", default)
    return min(top, 1000) if top is not None else None


def _requires_id(select_clause):
    return any(field.strip() in ("id", "*") for field in select_clause.split(","))


def _delta_options(delta_link, top):
    top = _clamp_top(top)
    if top == 0:
        raise ValueError("delta top must be positive")
    if delta_link is not None and (not isinstance(delta_link, str) or not delta_link.strip()
                                   or any(ord(c) < 32 for c in delta_link)):
        raise ValueError("delta_link must be a nonempty cursor or None for an initial sync")
    return top

def _collect_paged_values(
    HTTP_request,
    endpoint: str,
    *,
    headers=None,
    auto_paginate: bool = True,
    max_items: Optional[int] = None,
    page_limit: Optional[int] = None,
):
    """
    Collect Graph collection results across @odata.nextLink pages.
    """
    values = []
    max_items = _count(max_items, "max_items")
    if max_items == 0:
        return []
    for response in mail_pages(HTTP_request, endpoint, headers=headers,
                               auto_paginate=auto_paginate, page_limit=page_limit):
        values.extend(response["value"])
        if max_items is not None and len(values) >= max_items: return values[:max_items]
    return values[:max_items] if max_items is not None else values

class Folder:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.base_resource = mailbox_instance.base_resource; self.HTTP_request = mailbox_instance.api.graph_call; self._deduplicate_lists = lists.deduplicate; self._filter_data_structure = filters.filter_dicts_by_key_value
    def _list_root_folders(self, include_hidden=False, top=None):
        endpoint = f"{self.base_resource}/mailFolders"; params = []
        if include_hidden: params.append("includeHiddenFolders=true")
        top = _clamp_top(top)
        if top == 0: return []
        if top: params.append(f"$top={top}")
        if params: endpoint = f"{endpoint}?{'&'.join(params)}"
        return _collect_paged_values(self.HTTP_request, endpoint, auto_paginate=True)
    def _list_child_folders(self, folder_id, include_hidden=False):
        endpoint = f"{self.base_resource}/mailFolders/{mail_id(folder_id)}/childFolders"
        if include_hidden: endpoint += "?includeHiddenFolders=true"
        return _collect_paged_values(self.HTTP_request, endpoint, auto_paginate=True)
    def _walk_folders(self, folders, include_hidden=False, recursive=True, parent_path="", _seen=None, _depth=0):
        seen = set() if _seen is None else _seen
        if _depth > 30:
            raise GraphError(code="mailbox_folder_depth_limit")
        collected = []
        for folder in folders:
            if not isinstance(folder.get("id"), str) or not folder["id"]:
                raise GraphError(code="invalid_mail_folder")
            if folder["id"] in seen:
                raise GraphError(code="mailbox_folder_cycle")
            if len(seen) >= 5000:
                raise GraphError(code="mailbox_folder_limit")
            seen.add(folder["id"])
            folder_copy = deepcopy(folder); display_name = folder_copy.get('displayName') or folder_copy.get('id'); folder_copy['path'] = f"{parent_path}/{display_name}".strip('/'); collected.append(folder_copy)
            if recursive:
                children = self._list_child_folders(folder_copy['id'], include_hidden=include_hidden)
                collected.extend(
                    self._walk_folders(
                        children,
                        include_hidden=include_hidden,
                        recursive=True,
                        parent_path=folder_copy['path'],
                        _seen=seen,
                        _depth=_depth + 1,
                    )
                )
        return collected
    def _attach_children(self, folder, include_hidden=False):
        children = self._list_child_folders(folder['id'], include_hidden=include_hidden); folder = deepcopy(folder); folder['children'] = children; return folder
    def resolve(self, folder: Union[str, Dict[str, str], None]):
        """
        Resolve a folder only when the caller explicitly needs folder metadata.

        Message listing intentionally avoids this for performance.
        """
        if folder is None: return None
        if isinstance(folder, dict) and folder.get('id'):
            mail_id(folder["id"])
            return folder
        folder_str = _folder_reference(folder)
        try:
            direct = mail_read(self.HTTP_request, resource_url=f"{self.base_resource}/mailFolders/{mail_id(folder_str)}")
        except GraphError as exc:
            if exc.status_code != 404:
                raise
            direct = None
        if direct is not None:
            return checked_mail_item(direct)
        matches = self.find(folder_str, exact_match=True, recursive=True)
        if len(matches) > 1: raise ValueError("Mail folder name is ambiguous; supply its exact ID")
        if matches: return matches[0]
        matches = self.find(folder_str, exact_match=False, recursive=True)
        if len(matches) > 1: raise ValueError("Mail folder name is ambiguous; supply its exact ID")
        if matches: return matches[0]
        raise ValueError(f"Unable to resolve mail folder: {folder}")
    def find(self, search_term=None, include_hidden=False, include_children=False, exact_match=True, recursive=True, top=None):
        top = _count(top, "top")
        if top == 0: return []
        root_folders = self._list_root_folders(include_hidden=include_hidden, top=top); all_folders = self._walk_folders(root_folders, include_hidden=include_hidden, recursive=recursive)
        if not search_term: results = all_folders
        else:
            search_term = str(search_term)
            exact_results = self._filter_data_structure(
                all_folders,
                key='displayName',
                value=search_term,
                exact=exact_match,
                search_nested=False,
                dynamic_search=False,
            )
            if not exact_results and search_term.lower() in {
                'inbox', 'drafts', 'sentitems', 'deleteditems', 'archive', 'junkemail', 'outbox'
            }:
                for item in all_folders:
                    if item.get('displayName', '').replace(' ', '').lower() == search_term.replace(' ', '').lower(): exact_results.append(item)
            results = exact_results
        if include_children: results = [self._attach_children(item, include_hidden=include_hidden) for item in results]
        if top is not None: results = results[:top]
        return results

class Message:
    DEFAULT_SELECT = [
        'id', 'subject', 'bodyPreview', 'receivedDateTime', 'sentDateTime',
        'from', 'sender', 'toRecipients', 'ccRecipients', 'bccRecipients',
        'conversationId', 'conversationIndex', 'internetMessageId',
        'isRead', 'hasAttachments', 'importance', 'parentFolderId', 'webLink'
    ]
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.base_resource = mailbox_instance.base_resource; self.HTTP_request = mailbox_instance.api.graph_call
    def _headers(self, body_type='html'): return {'Prefer': f'outlook.body-content-type="{body_type}"'} if body_type else None
    def _normalize_select(self, select: Optional[Sequence[str]]) -> str:
        selected = [select] if isinstance(select, str) else list(select) if select else list(self.DEFAULT_SELECT)
        if any(not isinstance(item, str) or not item.strip() for item in selected):
            raise ValueError("select fields must be nonempty strings")
        seen = []
        for item in selected:
            if item not in seen: seen.append(item)
        return ','.join(seen)
    def _build_query(self, folder=None, page_size=25, order_by=None, select=None, unread=None):
        if unread is not None and not isinstance(unread, bool):
            raise ValueError("unread must be a boolean or None")
        endpoint = _mail_folder_endpoint(self.mailbox_instance, folder); query_parts = []; select_clause = self._normalize_select(select) if select else self._normalize_select(None); query_parts.append(f"$select={quote(select_clause, safe=',')}")
        page_size = _clamp_top(page_size, default=25)
        if page_size: query_parts.append(f"$top={page_size}")
        if order_by: query_parts.append(f"$orderby={quote(str(order_by), safe=' ,')}")
        if unread is True: query_parts.append("$filter=isRead eq false")
        elif unread is False: query_parts.append("$filter=isRead eq true")
        return f"{endpoint}?{'&'.join(query_parts)}"
    def _client_side_query(self, messages: List[Dict], query: Optional[str]):
        if not query: return messages
        query_l = str(query).lower(); result = []
        for item in messages:
            values = [
                item.get('subject', ''),
                item.get('bodyPreview', ''),
                ((item.get('from') or {}).get('emailAddress') or {}).get('address', ''),
                ((item.get('sender') or {}).get('emailAddress') or {}).get('address', ''),
                ((item.get('from') or {}).get('emailAddress') or {}).get('name', ''),
            ]
            joined = ' '.join(v for v in values if isinstance(v, str)).lower()
            if query_l in joined: result.append(item)
        return result
    def find(
        self,
        query=None,
        folder=None,
        top=25,
        order_by="receivedDateTime desc",
        select=None,
        unread=None,
        body_type='html',
        auto_paginate=True,
        max_items=None,
        page_limit=None,
    ):
        """
        Return messages from the chosen folder.

        Important behavior
        ------------------
        `top` is treated as the desired maximum number of results returned to the
        caller, not as "scan the entire mailbox with pages of size `top`".
        Counts must be nonnegative integers (not booleans or coerced strings).
        A zero top or max_items returns an empty list without requesting data.

        This avoids a pathological case where:
          - query filtering is done client-side, and
          - auto_paginate=True, and
          - top is small (e.g. 3),
        which would otherwise walk the whole mailbox trying to exhaust all pages.
        """
        page_size = _clamp_top(top, default=25)
        desired_limit = _count(max_items, "max_items", page_size)
        if page_size == 0: desired_limit = 0
        select_clause = self._normalize_select(select)
        endpoint = self._build_query(
            folder=folder,
            page_size=page_size,
            order_by=order_by,
            select=select_clause,
            unread=unread,
        )
        headers = self._headers(body_type)
        collected = []
        if desired_limit == 0:
            return []
        for response in mail_pages(self.HTTP_request, endpoint, headers=headers,
                                   auto_paginate=auto_paginate, page_limit=page_limit,
                                   require_id=_requires_id(select_clause)):
            page_messages = self._client_side_query(response["value"], query)
            collected.extend(page_messages)
            if desired_limit is not None and len(collected) >= desired_limit: return collected[:desired_limit]
        return collected[:desired_limit] if desired_limit is not None else collected
    def get(self, message_id, select=None, body_type='html', mime=False):
        if mime:
            return mail_read(self.HTTP_request,
                resource_url=f"{self.base_resource}/messages/{mail_id(message_id)}/$value",
                headers=self._headers(body_type),
                content_type='text/plain',
            )
        endpoint = f"{self.base_resource}/messages/{mail_id(message_id)}"; select_clause = self._normalize_select(select)
        if select_clause: endpoint += f"?$select={quote(select_clause, safe=',')}"
        return checked_mail_item(mail_read(self.HTTP_request, resource_url=endpoint, headers=self._headers(body_type)),
                                 require_id=_requires_id(select_clause))

class Attachment:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.base_resource = mailbox_instance.base_resource; self.HTTP_request = mailbox_instance.api.graph_call
    def list(self, message_id, auto_paginate=True, max_items=None, page_limit=None):
        endpoint = f"{self.base_resource}/messages/{mail_id(message_id)}/attachments"
        return _collect_paged_values(
            self.HTTP_request,
            endpoint,
            auto_paginate=auto_paginate,
            max_items=max_items,
            page_limit=page_limit,
        )
    def get(self, message_id, attachment_id, expand_item=False):
        endpoint = f"{self.base_resource}/messages/{mail_id(message_id)}/attachments/{mail_id(attachment_id)}"
        if expand_item: endpoint += "?$expand=microsoft.graph.itemAttachment/item"
        return checked_mail_item(mail_read(self.HTTP_request, resource_url=endpoint))

class Delta:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.base_resource = mailbox_instance.base_resource; self.HTTP_request = mailbox_instance.api.graph_call
    def _headers(self): return {'Prefer': 'odata.maxpagesize=200'}
    def messages(self, folder='inbox', delta_link=None, select=None, top=None, order_by=None, change_type=None):
        top = _delta_options(delta_link, top)
        if delta_link is not None: endpoint = delta_link
        else:
            folder_ref = _folder_reference(folder) or 'inbox'; endpoint = f"{self.base_resource}/mailFolders/{mail_id(folder_ref)}/messages/delta"; parts = []
            if select: parts.append(f"$select={quote(Message(self.mailbox_instance)._normalize_select(select), safe=',')}")
            if top is not None: parts.append(f"$top={top}")
            if order_by: parts.append(f"$orderby={quote(str(order_by), safe=' ,')}")
            if change_type: parts.append(f"changeType={quote(str(change_type))}")
            if parts: endpoint += '?' + '&'.join(parts)
        response = checked_mail_page(mail_read(self.HTTP_request, resource_url=endpoint, headers=self._headers(), delta=True),
                                     delta=True, require_id=not select or _requires_id(Message(self.mailbox_instance)._normalize_select(select)))
        return {
            'value': response.get('value', []),
            'next_link': response.get('@odata.nextLink'),
            'delta_link': response.get('@odata.deltaLink'),
        }
    def folders(self, delta_link=None, top=None, include_hidden=False):
        top = _delta_options(delta_link, top)
        if delta_link is not None: endpoint = delta_link
        else:
            endpoint = f"{self.base_resource}/mailFolders/delta"; parts = []
            if top is not None: parts.append(f"$top={top}")
            if include_hidden: parts.append("includeHiddenFolders=true")
            if parts: endpoint += '?' + '&'.join(parts)
        response = checked_mail_page(mail_read(self.HTTP_request, resource_url=endpoint, headers=self._headers(), delta=True), delta=True)
        return {
            'value': response.get('value', []),
            'next_link': response.get('@odata.nextLink'),
            'delta_link': response.get('@odata.deltaLink'),
        }


__all__ = ["Folder", "Message", "Attachment", "Delta"]
def __dir__(): return __all__
