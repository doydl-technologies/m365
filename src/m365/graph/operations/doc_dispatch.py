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

import concurrent.futures  
from pathlib import Path
import logging
import sys

from ..utils import safe_tqdm as tqdm
from .._site_resolution import graph_id


logger = logging.getLogger(__name__)


class Move:
    def __init__(self, sharepoint_instance):
        self.sharepoint_instance = sharepoint_instance; self._batch_file_folder = self.sharepoint_instance.search_by.filefoldername; self.file_name = None; self.destination_folder_name = None; self.new_file_name = None; self.file_id = None; self.folder_id = None; self.url = None; self._retrieved_filename = None
    def prepare_move(self, file_name, destination_folder_name, new_file_name=None):
        self.file_name = file_name; self.destination_folder_name = destination_folder_name; self.new_file_name = new_file_name
        self.file_id = self.folder_id = self.url = self._retrieved_filename = None
        batch_results = self._batch_file_folder(file_name, destination_folder_name)
        if batch_results:
            if len([item for item in batch_results if 'file' in item]) != 1 or len([item for item in batch_results if 'folder' in item]) != 1:
                raise ValueError("Move requires exactly one file and one destination folder")
            self.folder_id = (next((item for item in batch_results if 'folder' in item), None) or {}).get('id'); self.file_id = (next((item for item in batch_results if 'file' in item), None) or {}).get('id'); self._retrieved_filename = (next((item for item in batch_results if 'file' in item), None) or {}).get('name')
        if self.file_id and self.folder_id: self.url = f"drives/{graph_id(self.sharepoint_instance.drive_id)}/items/{graph_id(self.file_id)}"
        if self.new_file_name is None:
            self.new_file_name = self._retrieved_filename
        if not isinstance(self.new_file_name, str) or not self.new_file_name or any(c in self.new_file_name for c in '/\\:<>"|?*') or self.new_file_name in ('.', '..'):
            raise ValueError("The destination name must be a single valid filename")
    def _move_file(self):
        """Submit the prepared move through the client transport."""
        if not self.url: raise ValueError("URL is not set. Call prepare_move first.")
        body = {'parentReference': {'id': self.folder_id}, 'name': self.new_file_name}; result = self.sharepoint_instance.api.graph_call(resource_url=self.url, method='PATCH', content_type='application/json', json=body); return result
    def _move(self, file_name, destination_folder_name, new_file_name=None):
        self.prepare_move(file_name, destination_folder_name, new_file_name)
        if self.new_file_name and Path(self.new_file_name).suffix == '': self.new_file_name += Path(self._retrieved_filename).suffix
        result = self._move_file(); return result
    def move(self, sharepoint_filename, sharepoint_foldername, new_sharepoint_filename):
        """Move one resolved file and verify the returned destination metadata."""; file_movements = [(sharepoint_filename, sharepoint_foldername, new_sharepoint_filename)]
        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = []
            for file_name, destination_folder_name, new_file_name in file_movements: futures.append(executor.submit(self._move, file_name, destination_folder_name, new_file_name))
            results = []
            for future in tqdm(
                    concurrent.futures.as_completed(futures),
                    total=len(futures),
                    desc="Moving Files",
            ):
                try:
                    result = future.result(); results.append(result)
                    parent = result.get("parentReference") or {}
                    if (result.get("id") != self.file_id or result.get("name") != self.new_file_name
                            or parent.get("id") != self.folder_id
                            or parent.get("driveId", self.sharepoint_instance.drive_id) != self.sharepoint_instance.drive_id):
                        raise RuntimeError("Graph move returned unexpected item, name or destination metadata; inspect the item before retrying")
                    else: logger.info(f"File {result.get('name')} moved successfully")
                except Exception:
                    raise
            return results[0]
    def __dir__(self): return ["move"]


__all__ = ["Move"]
def __dir__(): return __all__
