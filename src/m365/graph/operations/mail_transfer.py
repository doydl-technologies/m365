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

import os
from pathlib import Path
import re
import logging
from .doc_transfer import _child, _save_response
from .._graph import GraphAPI, GraphError
from .._mail_common import mail_id, checked_mail_result
from .._local_paths import native_path, public_path

try:
    from .mail_management import Attachment
except (ImportError, ModuleNotFoundError):
    from .mail_management import Attachment


logger = logging.getLogger(__name__)


class AttachmentDownload:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.HTTP_request = mailbox_instance.api.graph_call; self.attachment_lookup = Attachment(mailbox_instance)
    def _resolve_attachment(self, message_id, attachment_id=None, attachment_name=None):
        if attachment_id: return self.attachment_lookup.get(message_id, attachment_id)
        attachments = self.attachment_lookup.list(message_id)
        if attachment_name:
            matches = [item for item in attachments if item.get("name") == attachment_name]
            if len(matches) > 1:
                raise ValueError("Attachment name is ambiguous; supply its exact ID")
            if matches: return matches[0]
            raise ValueError(f"Unable to find attachment named {attachment_name!r}")
        if len(attachments) == 1: return attachments[0]
        if not attachments: raise ValueError("No attachments found on the message.")
        raise ValueError("Multiple attachments found. Provide attachment_id or attachment_name.")
    def download(self, message_id, save_dir='.', attachment_id=None, attachment_name=None):
        attachment = self._resolve_attachment(message_id, attachment_id=attachment_id, attachment_name=attachment_name)
        name = attachment.get('name') or attachment.get('id')
        save_dir = native_path(save_dir)
        save_dir.mkdir(parents=True, exist_ok=True)
        output_path = _child(save_dir, name)
        stream = self.HTTP_request(
            resource_url=f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}/attachments/{mail_id(attachment['id'])}/$value",
            method='GET',
            stream=True,
            content_type='application/octet-stream',
        )
        return str(public_path(_save_response(stream, output_path, attachment.get('size'))))
    def __dir__(self): return ["download"]

class AttachmentUpload:
    SMALL_LIMIT = 3 * 1024 * 1024; LARGE_LIMIT = 150 * 1024 * 1024; CHUNK_SIZE = 327680
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.HTTP_request = mailbox_instance.api.graph_call
    def _create_upload_session(self, message_id, file_name, file_size):
        endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}/attachments/createUploadSession"
        payload = {
            "AttachmentItem": {
                "attachmentType": "file",
                "name": file_name,
                "size": file_size,
            }
        }
        return checked_mail_result(self.HTTP_request(resource_url=endpoint, method='POST', json=payload))
    def _upload_large(self, upload_url, payload: bytes):
        total = len(payload); start = 0
        if not total: raise ValueError("A fragment upload needs a nonempty payload")
        while start < total:
            end = min(start + self.CHUNK_SIZE, total) - 1; chunk = payload[start:end + 1]
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {start}-{end}/{total}",
                "Content-Type": "application/octet-stream",
            }
            final_response = self.HTTP_request(
                resource_url=upload_url,
                method='PUT',
                headers=headers,
                data=chunk,
                content_type='application/octet-stream',
                preauthenticated=True,
                response_metadata=True,
            )
            if (not isinstance(final_response, dict) or not isinstance(final_response.get("body"), dict)
                    or "error" in final_response["body"]):
                raise GraphError(code="invalid_attachment_response")
            status, body = final_response.get("status_code"), final_response["body"]
            if end + 1 == total:
                # Outlook commits with 201; its documented final body is empty.
                if status != 201 or "nextExpectedRanges" in body:
                    raise GraphError(code="attachment_not_committed")
                if "size" in body and body["size"] != total:
                    raise GraphError(code="attachment_size_mismatch")
                return body or {"completed": True, "status_code": 201, "size": total}
            ranges = body.get("nextExpectedRanges")
            if status != 200 or not isinstance(ranges, list) or len(ranges) != 1 or not isinstance(ranges[0], str):
                raise GraphError(code="invalid_attachment_acknowledgement")
            # Outlook reports one start byte; tolerate the legacy "start-" form.
            if not re.fullmatch(r"\d+-?", ranges[0]) or int(ranges[0].rstrip("-")) != end + 1:
                raise GraphError(code="unexpected_attachment_acknowledgement")
            start = end + 1
        raise GraphError(code="attachment_not_committed")
    def upload(self, message_id, file_source, file_name=None, content_type="application/octet-stream"):
        if isinstance(file_source, (bytes, bytearray)):
            payload = bytes(file_source); resolved_name = file_name or 'attachment.bin'
        else:
            path = native_path(file_source); payload = path.read_bytes(); resolved_name = file_name or path.name
        file_size = len(payload)
        if file_size > self.LARGE_LIMIT: raise ValueError("Attachment exceeds the Microsoft Graph Outlook 150 MB limit.")
        if file_size < self.SMALL_LIMIT:
            endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}/attachments"
            body = {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": resolved_name,
                "contentType": content_type,
                "contentBytes": __import__('base64').b64encode(payload).decode('ascii'),
            }
            return self.HTTP_request(resource_url=endpoint, method='POST', json=body)
        session = self._create_upload_session(message_id, resolved_name, file_size)
        if not isinstance(session, dict) or not isinstance(session.get("uploadUrl"), str) or not session["uploadUrl"]:
            raise GraphError(code="invalid_attachment_session")
        upload_url = GraphAPI._capability(session["uploadUrl"])
        return self._upload_large(upload_url, payload)
    def __dir__(self): return ["upload"]



__all__ = ["AttachmentDownload", "AttachmentUpload"]
def __dir__(): return __all__
