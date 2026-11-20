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

import logging
from .._mail_common import mail_id

try:
    from .mail_management import Folder
except (ImportError, ModuleNotFoundError):
    from .mail_management import Folder

logger = logging.getLogger(__name__)


class Move:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.HTTP_request = mailbox_instance.api.graph_call; self.folder_lookup = Folder(mailbox_instance)
    def move(self, message_id, destination_folder):
        destination = self.folder_lookup.resolve(destination_folder); body = {"destinationId": destination['id']}; endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}/move"; return self.HTTP_request(resource_url=endpoint, method='POST', json=body)
    def __dir__(self): return ["move"]

class MessageState:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.HTTP_request = mailbox_instance.api.graph_call
    def update(self, message_id, **fields):
        endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}"; return self.HTTP_request(resource_url=endpoint, method='PATCH', json=fields)
    def mark_read(self, message_id, is_read=True): return self.update(message_id, isRead=bool(is_read))
    def mark_unread(self, message_id): return self.update(message_id, isRead=False)
    def __dir__(self): return ["update", "mark_read", "mark_unread"]

class Delete:
    def __init__(self, mailbox_instance):
        self.mailbox_instance = mailbox_instance; self.HTTP_request = mailbox_instance.api.graph_call
    def delete(self, message_id):
        endpoint = f"{self.mailbox_instance.base_resource}/messages/{mail_id(message_id)}"; return self.HTTP_request(resource_url=endpoint, method='DELETE')
    def __dir__(self): return ["delete"]


__all__ = ["Move", "MessageState", "Delete"]
def __dir__(): return __all__
