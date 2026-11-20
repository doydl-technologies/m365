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

import base64
from pathlib import Path
import logging
from .._mail_common import mail_id
from typing import Iterable, Optional, Sequence, Union

logger = logging.getLogger(__name__)


def _ensure_list(value):
    if value is None: return []
    if isinstance(value, (list, tuple, set)): return [item for item in value if item]
    return [value]

def _file_attachment_from_bytes(name: str, payload: bytes, content_type: str = "application/octet-stream"):
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name,
        "contentType": content_type,
        "contentBytes": base64.b64encode(payload).decode('ascii'),
    }

def _maybe_inline_attachment(item):
    if isinstance(item, dict): return item
    path = Path(item); payload = path.read_bytes(); return _file_attachment_from_bytes(name=path.name, payload=payload)


class Draft:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.HTTP_request = mailbox_instance.api.graph_call; self.tools = mailbox_instance.tools
    def _message_payload(
        self,
        subject: str,
        body: str,
        to=None,
        cc=None,
        bcc=None,
        reply_to=None,
        body_type: str = "HTML",
        importance: str = "normal",
        attachments=None,
        categories=None,
        internet_message_headers=None,
    ):
        message = {
            "subject": subject,
            "body": {
                "contentType": str(body_type).upper(),
                "content": body,
            },
            "toRecipients": self.tools.recipients(to),
            "ccRecipients": self.tools.recipients(cc),
            "bccRecipients": self.tools.recipients(bcc),
            "replyTo": self.tools.recipients(reply_to),
            "importance": importance,
        }
        if attachments: message["attachments"] = [_maybe_inline_attachment(item) for item in _ensure_list(attachments)]
        if categories: message["categories"] = list(categories)
        if internet_message_headers: message["internetMessageHeaders"] = list(internet_message_headers)
        return message
    def create(
        self,
        subject: str,
        body: str,
        to=None,
        cc=None,
        bcc=None,
        reply_to=None,
        body_type: str = "HTML",
        importance: str = "normal",
        attachments=None,
        categories=None,
        internet_message_headers=None,
    ):
        endpoint = f"{self.mailbox_instance.base_resource}/messages"
        payload = self._message_payload(
            subject=subject,
            body=body,
            to=to,
            cc=cc,
            bcc=bcc,
            reply_to=reply_to,
            body_type=body_type,
            importance=importance,
            attachments=attachments,
            categories=categories,
            internet_message_headers=internet_message_headers,
        )
        return self.HTTP_request(resource_url=endpoint, method='POST', json=payload)
    def update(self, message_id, **fields):
        endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}"; return self.HTTP_request(resource_url=endpoint, method='PATCH', json=fields)
    def create_reply(self, message_id, comment: Optional[str] = None, body: Optional[str] = None, body_type: str = "HTML"):
        endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}/createReply"; payload = {}
        if comment is not None: payload['comment'] = comment
        elif body is not None: payload['message'] = {"body": {"contentType": str(body_type).upper(), "content": body}}
        return self.HTTP_request(resource_url=endpoint, method='POST', json=payload if payload else None)
    def create_reply_all(self, message_id, comment: Optional[str] = None, body: Optional[str] = None, body_type: str = "HTML"):
        endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}/createReplyAll"; payload = {}
        if comment is not None: payload['comment'] = comment
        elif body is not None: payload['message'] = {"body": {"contentType": str(body_type).upper(), "content": body}}
        return self.HTTP_request(resource_url=endpoint, method='POST', json=payload if payload else None)
    def create_forward(self, message_id, comment: str = "", to=None):
        endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}/createForward"
        payload = {
            "comment": comment,
            "toRecipients": self.tools.recipients(to),
        }
        return self.HTTP_request(resource_url=endpoint, method='POST', json=payload)
    def send(self, message_id):
        endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}/send"; return self.HTTP_request(resource_url=endpoint, method='POST')
    def __dir__(self): return ["create", "update", "create_reply", "create_reply_all", "create_forward", "send"]


class Send:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.HTTP_request = mailbox_instance.api.graph_call; self.draft = Draft(mailbox_instance)
    def send(
        self,
        subject: str,
        body: str,
        to=None,
        cc=None,
        bcc=None,
        reply_to=None,
        body_type: str = "HTML",
        importance: str = "normal",
        attachments=None,
        save_to_sent_items: bool = True,
        categories=None,
        internet_message_headers=None,
    ):
        endpoint = f"{self.mailbox_instance.base_resource}/sendMail"
        message = self.draft._message_payload(
            subject=subject,
            body=body,
            to=to,
            cc=cc,
            bcc=bcc,
            reply_to=reply_to,
            body_type=body_type,
            importance=importance,
            attachments=attachments,
            categories=categories,
            internet_message_headers=internet_message_headers,
        )
        payload = {"message": message}
        if save_to_sent_items is False: payload["saveToSentItems"] = False
        return self.HTTP_request(resource_url=endpoint, method='POST', json=payload)
    def __dir__(self): return ["send"]


__all__ = ["Draft", "Send"]
def __dir__(): return __all__
