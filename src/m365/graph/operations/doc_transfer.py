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
"""Paginated downloads, staged files, and non-destructive optional compression."""
import os
import re
import tempfile
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import requests
from requests.structures import CaseInsensitiveDict

from .._graph import GraphError, checked_drive_item, drive_item_kind, drive_item_name
from .._site_resolution import graph_id
from .._local_paths import native_path, public_path


def _workers(value):
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise ValueError("max_workers must be a positive integer")
    return value


def _child(root, name):
    # A remote name must never escape a destination or alias a Windows device.
    if (not isinstance(name, str) or not name or name in (".", "..")
            or any(c in name for c in '/\\:<>"|?*') or any(ord(c) < 32 for c in name)
            or name.endswith((".", " "))
            or re.match(r"^(CON|PRN|AUX|NUL|COM[1-9\u00b9\u00b2\u00b3]|LPT[1-9\u00b9\u00b2\u00b3])(?:\.|$)", name, re.I)):
        raise ValueError("Remote item has an unsafe local filename")
    root = Path(root)
    if root.is_symlink():
        raise ValueError("Download destinations must not be symbolic links")
    target = root / name
    if target.exists() or target.is_symlink():
        raise FileExistsError("Download destination already exists: " + str(target))
    return target


def _write_download(api, item, destination, drive_id):
    if item.get("id"):
        resource = "drives/{}/items/{}/content".format(graph_id(drive_id), graph_id(item["id"]))
        response = api.graph_call(resource_url=resource, stream=True)
    elif item.get("@microsoft.graph.downloadUrl"):
        response = api.graph_call(resource_url=item["@microsoft.graph.downloadUrl"], stream=True, preauthenticated=True)
    else:
        raise ValueError("File metadata needs an item ID or Graph-provided download URL")
    return _save_response(response, destination, item.get("size"))


def _full_download_size(response, expected):
    status = getattr(response, "status_code", None)
    if type(status) is not int or status != 200:
        raise GraphError(status if type(status) is int else None, "unexpected_download_status")
    headers = CaseInsensitiveDict(response.headers)
    if "Content-Range" in headers:
        raise GraphError(code="unexpected_content_range")
    if expected is not None and (isinstance(expected, bool) or not isinstance(expected, int) or expected < 0):
        raise GraphError(code="invalid_download_size")
    if "Content-Length" not in headers:
        return expected
    # Requests yields decoded content. A compressed wire length neither proves
    # nor contradicts the decoded size; the HTTP transport frames the wire.
    length_header = headers["Content-Length"]
    if "Transfer-Encoding" in headers or not isinstance(length_header, str):
        raise GraphError(code="invalid_content_length")
    values = [part.strip() for part in length_header.split(",")]
    if any(re.fullmatch(r"[0-9]{1,20}", part) is None for part in values):
        raise GraphError(code="invalid_content_length")
    lengths = {int(part) for part in values}
    if len(lengths) != 1:
        raise GraphError(code="invalid_content_length")
    wire_length = lengths.pop()
    if headers.get("Content-Encoding", "identity").strip().lower() not in ("", "identity"):
        return expected
    if expected is not None and wire_length != expected:
        raise GraphError(code="download_size_mismatch")
    return wire_length


def _save_response(response, destination, expected=None):
    """Shared stream staging: validate length, publish without clobber, always close."""
    original_destination = destination
    temporary = None
    try:
        expected = _full_download_size(response, expected)
        destination = native_path(destination)
        with tempfile.NamedTemporaryFile(prefix=".msftgraph-", dir=str(destination.parent), delete=False) as stream:
            temporary = stream.name
            length = 0
            for chunk in response.iter_content(chunk_size=64 * 1024):
                if chunk:
                    if expected is not None and length + len(chunk) > expected:
                        raise GraphError(code="download_size_mismatch")
                    stream.write(chunk)
                    length += len(chunk)
        if expected is not None and length != expected:
            raise GraphError(code="download_size_mismatch")
        # Atomic no-clobber publication of the complete file. Requires hard-link
        # support; fail explicitly on unsupported filesystems rather than clobber.
        os.link(temporary, destination)
    except requests.RequestException:
        raise GraphError(code="download_stream_failed") from None
    finally:
        try:
            response.close()
        finally:
            if temporary is not None:
                os.unlink(temporary)
    return original_destination


def _compress(root, files, preserve_structure, *, directories=()):
    root = native_path(root)
    files = [native_path(f) for f in files]
    directories = [native_path(d) for d in directories]
    archive = Path(str(root) + ".zip")
    temporary = None
    names = [f.relative_to(root).as_posix() if preserve_structure else f.name for f in files]
    directory_names = []
    for directory in directories:
        relative = directory.relative_to(root).as_posix()
        name = (relative if relative != "." else root.name) if preserve_structure else directory.name
        directory_names.append(name.rstrip("/") + "/")
    keys = [n.rstrip("/").casefold() for n in names + directory_names]
    if len(set(keys)) != len(keys):
        raise ValueError("Compression would create duplicate archive entry names")
    try:
        with tempfile.NamedTemporaryFile(prefix=".msftgraph-", dir=str(root.parent), delete=False) as stream:
            temporary = stream.name
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as z:
            for file, name in zip(files, names):
                z.write(file, name)
            for name in directory_names:
                z.writestr(name, b"")
        os.link(temporary, archive)
    finally:
        if temporary is not None:
            os.unlink(temporary)
    # Never delete original downloads or unrelated destination contents.
    return archive


class File:
    def __init__(self, sharepoint_instance, max_workers=20, compress=False, preserve_dir_structure=False):
        self.client = sharepoint_instance
        self.max_workers = _workers(max_workers)
        self.compress_files = compress
        self.preserve_dir_structure = preserve_dir_structure

    def download(self, folder_path, files):
        root = native_path(folder_path)
        root.mkdir(parents=True, exist_ok=True)
        files = files if isinstance(files, list) else [files]
        tasks = [(item, _child(root, item.get("name", item.get("file_name")))) for item in files]
        if len({str(p).casefold() for _, p in tasks}) != len(tasks):
            raise ValueError("Multiple remote files target the same local filename")
        with ThreadPoolExecutor(max_workers=self.max_workers) as pool:
            futures = [pool.submit(_write_download, self.client.api, item, path, self.client.drive_id) for item, path in tasks]
            paths = [f.result() for f in futures]
        return _compress(root, paths, self.preserve_dir_structure) if self.compress_files else paths


class Folder:
    def __init__(self, sharepoint_instance, max_workers=20, compress=False, preserve_dir_structure=True):
        self.client = sharepoint_instance
        self.max_workers = _workers(max_workers)
        self.compress = compress
        self.preserve_dir_structure = preserve_dir_structure

    def _get_folder_contents(self, item_id):
        return list(self.client.api.iter_collection("drives/{}/items/{}/children".format(
            graph_id(self.client.drive_id), graph_id(item_id))))

    def _plan(self, item_id, root, visited, empty_directories):
        if item_id in visited:
            raise GraphError(code="folder_cycle")
        visited.add(item_id)
        tasks = []
        names = set()
        children = self._get_folder_contents(item_id)
        if not children:
            empty_directories.append(root)
        for item in children:
            checked_drive_item(item, self.client.drive_id)
            if "remoteItem" in item or "deleted" in item:
                raise GraphError(code="unsupported_remote_item")
            kind = drive_item_kind(item)
            target = _child(root, drive_item_name(item))
            key = str(target).casefold()
            if key in names:
                raise ValueError("Remote children contain colliding local names")
            names.add(key)
            if kind == "folder":
                target.mkdir()
                tasks.extend(self._plan(item["id"], target, visited, empty_directories))
            elif kind == "file":
                tasks.append((item, target))
            else:
                raise GraphError(code="unsupported_remote_item")
        return tasks

    def download_folder(self, save_dir, item_id):
        metadata = self.client.api.graph_call(resource_url="drives/{}/items/{}".format(
            graph_id(self.client.drive_id), graph_id(item_id)))
        checked_drive_item(metadata, self.client.drive_id)
        if metadata["id"] != item_id:
            raise GraphError(code="download_folder_id_mismatch")
        if "remoteItem" in metadata or "deleted" in metadata or drive_item_kind(metadata) != "folder":
            raise ValueError("The selected item is not a folder")
        name = drive_item_name(metadata)
        parent = native_path(save_dir)
        parent.mkdir(parents=True, exist_ok=True)
        root = _child(parent, name)
        root.mkdir()
        empty_directories = []
        tasks = self._plan(item_id, root, set(), empty_directories)
        if len({str(p).casefold() for _, p in tasks}) != len(tasks):
            raise ValueError("Multiple remote files target the same local path")
        with ThreadPoolExecutor(max_workers=self.max_workers) as pool:
            futures = [pool.submit(_write_download, self.client.api, item, path, self.client.drive_id) for item, path in tasks]
            paths = [f.result() for f in futures]
        if self.compress:
            return public_path(_compress(root, paths, self.preserve_dir_structure, directories=empty_directories))
        # A short root can still contain long descendants. Keep its namespace
        # usable for callers walking the returned directory tree.
        long_descendants = any(str(public_path(p)).startswith("\\\\?\\") for p in paths + empty_directories)
        return public_path(root, keep_extended=long_descendants)


def _items(metadata, kind):
    """Accept items, lists, and Graph collection/batch wrappers."""
    if isinstance(metadata, list):
        for value in metadata:
            yield from _items(value, kind)
    elif isinstance(metadata, dict):
        if "error" in metadata:
            raise GraphError(code="download_metadata_error")
        if "status" in metadata:
            status = metadata["status"]
            if isinstance(status, bool) or not isinstance(status, int) or not 200 <= status <= 599:
                raise GraphError(code="invalid_download_metadata")
            if not 200 <= status < 300:
                raise GraphError(status, "download_metadata_error")
        keys = [key for key in ("value", "body", "responses") if key in metadata]
        facet = drive_item_kind(metadata)
        if len(keys) > 1 or (keys and facet):
            raise GraphError(code="invalid_download_metadata")
        if keys:
            key = keys[0]
            value = metadata[key]
            if ((key != "body" and not isinstance(value, list)) or
                    (key == "body" and not isinstance(value, (dict, list)))):
                raise GraphError(code="invalid_download_metadata")
            yield from _items(value, kind)
        elif facet == kind:
            yield metadata
    else:
        raise GraphError(code="invalid_download_metadata")


class Download:
    def __init__(self, sharepoint_instance, max_workers=10, preserve_dir_structure=False, compress=False):
        self.file_download = File(sharepoint_instance, max_workers, compress, preserve_dir_structure)
        self.folder_download = Folder(sharepoint_instance, max_workers, compress, preserve_dir_structure)

    def file(self, folder_path, file_metadata):
        files = list(_items(file_metadata, "file"))
        if not files:
            raise ValueError("No file items found in metadata")
        for item in files:
            if "remoteItem" in item or "deleted" in item:
                raise GraphError(code="unsupported_remote_item")
            if "id" in item:
                checked_drive_item(item, self.file_download.client.drive_id)
        result = self.file_download.download(folder_path, files)
        return [public_path(p) for p in result] if isinstance(result, list) else public_path(result)

    def folder(self, save_dir, folder_metadata):
        folders = list(_items(folder_metadata, "folder"))
        if not folders:
            raise ValueError("No folder items found in metadata")
        for item in folders:
            checked_drive_item(item, self.folder_download.client.drive_id)
            if "remoteItem" in item or "deleted" in item:
                raise GraphError(code="unsupported_remote_item")
        return [self.folder_download.download_folder(save_dir, item["id"]) for item in folders]


__all__ = ["Download"]
