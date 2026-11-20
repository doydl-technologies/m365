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
"""Library-relative, conflict-explicit SharePoint uploads (no tenant assumptions)."""
import io
import os
import re
from contextlib import contextmanager
from urllib.parse import unquote

from .._graph import GraphError, checked_drive_item
from .._site_resolution import drive_path, graph_id
from .._local_paths import native_path


class Upload:
    """Upload one file/BytesIO using a sequential Graph upload session.

    Directories are rejected rather than partially uploading a tree.
    Folder creation/ZIP snapshots belong to caller orchestration.
    """
    def __init__(self, sharepoint_instance, chunk_verbose=False, max_file_size="4MB",
                 *, conflict_behavior="fail", chunk_size=10 * 327680):
        if conflict_behavior not in ("fail", "replace", "rename"):
            raise ValueError("conflict_behavior must be fail, replace, or rename")
        if isinstance(chunk_size, bool) or not isinstance(chunk_size, int) or chunk_size <= 0 or chunk_size % 327680 or chunk_size >= 60 * 1024 * 1024:
            raise ValueError("chunk_size must be a 320 KiB multiple below 60 MiB")
        self.sharepoint_instance = sharepoint_instance
        self.conflict_behavior = conflict_behavior
        self.chunk_size = chunk_size
        # max_file_size remains accepted for source compatibility. All nonempty
        # files use sessions so conflict semantics do not vary with file size.

    @contextmanager
    def _source(self, source):
        if isinstance(source, io.BytesIO):
            position = source.tell()
            try:
                source.seek(0)
                yield source
            finally:
                source.seek(position)
        else:
            source = native_path(source)
            if os.path.isdir(source):
                raise ValueError("Upload accepts a file, not a directory; archive or enumerate directories explicitly")
            with open(source, "rb") as stream:
                yield stream

    @staticmethod
    def _offset(response, size):
        if not isinstance(response, dict) or "error" in response:
            raise GraphError(code="invalid_upload_response")
        ranges = response.get("nextExpectedRanges")
        if not isinstance(ranges, list) or not ranges:
            raise GraphError(code="missing_upload_ranges")
        offsets = []
        for value in ranges:
            match = re.fullmatch(r"([0-9]{1,20})-([0-9]{1,20})?", value) if isinstance(value, str) else None
            if not match:
                raise GraphError(code="invalid_upload_ranges")
            start = int(match.group(1))
            end = int(match.group(2)) if match.group(2) is not None else size - 1
            if not 0 <= start <= end < size or start % 327680:
                raise GraphError(code="invalid_upload_offset")
            offsets.append(start)
        return min(offsets)

    def _completion(self, result, size, name, folder, drive_id):
        # A successful PUT is not proof that arbitrary response JSON identifies
        # the expected file. Reject contradictions, without retrying a commit.
        checked_drive_item(result, drive_id, "invalid_upload_completion")
        actual_size = result.get("size")
        if isinstance(actual_size, bool) or not isinstance(actual_size, int) or actual_size != size:
            raise GraphError(code="upload_size_mismatch")
        actual_name = result.get("name")
        if (not isinstance(actual_name, str) or not actual_name or actual_name in (".", "..")
                or any(c in actual_name for c in '/\\:<>"|?*')
                or any(ord(c) < 32 for c in actual_name)
                or not isinstance(result.get("file"), dict)
                or any(key in result for key in ("folder", "remoteItem", "deleted", "error", "nextExpectedRanges"))):
            raise GraphError(code="invalid_upload_completion")
        if self.conflict_behavior != "rename" and actual_name.casefold() != name.casefold():
            raise GraphError(code="upload_name_mismatch")
        parent = result.get("parentReference", {})
        scoped_parent = getattr(self.sharepoint_instance, "upload_parent_id", None)
        if scoped_parent is not None:
            if parent.get("id") != scoped_parent or parent.get("driveId") != drive_id:
                raise GraphError(code="upload_parent_mismatch")
        elif "path" in parent:
            match = re.fullmatch(r"/drives/([^/]+)/root:(/.*)?", parent["path"])
            if (not match or unquote(match.group(1)) != drive_id
                    or unquote(match.group(2) or "").rstrip("/").casefold() !=
                       ("/" + unquote(folder) if folder else "").casefold()):
                raise GraphError(code="upload_parent_mismatch")
        return result

    def upload(self, file_source, sharepoint_folder_path, new_file_name=None):
        name = new_file_name or (None if isinstance(file_source, io.BytesIO) else os.path.basename(file_source))
        if (not isinstance(name, str) or not name or name in (".", "..")
                or any(c in name for c in '/\\:<>"|?*') or any(ord(c) < 32 for c in name)):
            raise ValueError("A single destination filename is required (including for BytesIO)")
        folder = drive_path(sharepoint_folder_path)
        encoded_name = drive_path(name)
        drive_id = self.sharepoint_instance.drive_id
        scoped_parent = getattr(self.sharepoint_instance, "upload_parent_id", None)
        if scoped_parent is not None:
            if folder:
                raise ValueError("An ID-scoped upload requires an already resolved parent")
            target = "drives/{}/items/{}:/{}".format(graph_id(drive_id), graph_id(scoped_parent), encoded_name)
        else:
            target = "drives/{}/root:/{}{}".format(graph_id(drive_id),
                       folder + "/" if folder else "", encoded_name)
        api = self.sharepoint_instance.api
        with self._source(file_source) as stream:
            stream.seek(0, 2)
            size = stream.tell()
            stream.seek(0)
            if size == 0:
                # A fragment session cannot represent zero bytes. Do not fall back to
                # an unconditional overwrite request.
                raise ValueError("Empty-file uploads are not supported by the conflict-safe session uploader")
            options = {"resource_url": target + ":/createUploadSession", "method": "POST",
                       "json": {"item": {"@microsoft.graph.conflictBehavior": self.conflict_behavior, "name": name}}}
            if scoped_parent is not None:
                reply = api.graph_call(**options, response_metadata=True)
                if not isinstance(reply, dict) or type(reply.get("status_code")) is not int or reply["status_code"] != 200:
                    raise GraphError(code="invalid_upload_session_status")
                session = reply.get("body")
            else:
                session = api.graph_call(**options)
            if (not isinstance(session, dict) or not isinstance(session.get("uploadUrl"), str)
                    or not session["uploadUrl"] or "error" in session):
                raise GraphError(code="invalid_upload_session")
            upload_url = session["uploadUrl"]
            offset, recoveries = 0, 0
            while offset < size:
                stream.seek(offset)
                chunk = stream.read(min(self.chunk_size, size - offset))
                if len(chunk) != min(self.chunk_size, size - offset):
                    raise GraphError(code="source_changed_during_upload")
                try:
                    reply = api.graph_call(resource_url=upload_url, method="PUT", data=chunk,
                        headers={"Content-Length": str(len(chunk)),
                                 "Content-Range": "bytes {}-{}/{}".format(offset, offset + len(chunk) - 1, size)},
                        content_type="application/octet-stream", preauthenticated=True,
                        response_metadata=True)
                except GraphError as exc:
                    if exc.status_code not in (None, 416, 500, 502, 503, 504) or recoveries >= 3:
                        raise
                    # Query session state before replaying an ambiguously acknowledged PUT.
                    recoveries += 1
                    try:
                        state = api.graph_call(resource_url=upload_url, preauthenticated=True)
                    except GraphError:
                        raise GraphError(code="upload_completion_unknown") from None
                    offset = self._offset(state, size)
                    continue
                if not isinstance(reply, dict) or not isinstance(reply.get("body"), dict):
                    raise GraphError(code="invalid_upload_response")
                status, result = reply.get("status_code"), reply["body"]
                if type(status) is not int or "error" in result:
                    raise GraphError(code="invalid_upload_response")
                end = offset + len(chunk)
                if end == size:
                    if status not in (200, 201):
                        raise GraphError(status, "upload_not_committed")
                    if result.get("size") != size:
                        raise GraphError(code="upload_size_mismatch")
                    return self._completion(result, size, name, folder, drive_id)
                if status != 202 or any(key in result for key in ("id", "file", "folder", "remoteItem", "deleted")):
                    raise GraphError(status, "unexpected_upload_acknowledgement")
                next_offset = self._offset(result, size)
                if next_offset != end:
                    raise GraphError(code="unexpected_upload_acknowledgement")
                offset = next_offset
            raise GraphError(code="upload_not_committed")


__all__ = ["Upload"]
