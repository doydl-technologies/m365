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
"""In-memory Graph workflows and 400 deterministic adversarial OneDrive cases."""
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
import hashlib
import io
from pathlib import Path
import random
import re
import tempfile
import unittest
from urllib.parse import quote, unquote, urlsplit
import zipfile

from m365.graph import GraphError
from test_onedrive import BASE, CHUNK, DRIVE, ROOT, DriveFixture, file_item, folder, page
from test_sharepoint import response
from test_transfer_hardening import binary


class MemoryDrive:
    """A synthetic service, not a substitute for a live Graph integration test."""
    def __init__(self):
        self.nodes = {ROOT: folder()}
        self.sessions = {}
        self.records = []
        self.counter = 0
        self.denied = False

    def metadata(self, identifier):
        return deepcopy({k: v for k, v in self.nodes[identifier].items() if k != "_data"})

    def lookup(self, parent, name):
        return next((identifier for identifier, node in self.nodes.items()
                     if node.get("parentReference", {}).get("id") == parent
                     and node["name"].casefold() == name.casefold()), None)

    def add(self, name, parent=ROOT, data=None):
        self.counter += 1
        identifier = "node-" + str(self.counter)
        value = folder(identifier, name, parent) if data is None else file_item(identifier, name, parent, len(data))
        if data is not None:
            value["_data"] = bytes(data)
        self.nodes[identifier] = value
        return identifier

    def graph(self, method, url, **kwargs):
        self.records.append((method, url))
        path = urlsplit(url).path
        prefix = "/v1.0/drives/drive/items/"
        if not path.startswith(prefix):
            raise AssertionError("Broad discovery/foreign endpoint requested")
        tail = path[len(prefix):]
        if method == "GET":
            if tail.endswith("/children"):
                parent = unquote(tail[:-9], errors="strict")
                return page(*(self.metadata(identifier) for identifier, node in self.nodes.items()
                              if node.get("parentReference", {}).get("id") == parent))
            if tail.endswith("/content"):
                identifier = unquote(tail[:-8], errors="strict")
                return binary(data=self.nodes[identifier]["_data"])
            if ":/" in tail:
                parent, name = (unquote(part, errors="strict") for part in tail.split(":/", 1))
                identifier = self.lookup(parent, name)
            else:
                identifier = unquote(tail, errors="strict")
            if self.denied or identifier not in self.nodes:
                return response({}, status=403 if self.denied else 404)
            return response(self.metadata(identifier))
        if method != "POST":
            raise AssertionError("Unexpected mutation")
        if tail.endswith("/children"):
            parent = unquote(tail[:-9], errors="strict")
            body = kwargs["json"]
            assert body["@microsoft.graph.conflictBehavior"] == "fail"
            if self.lookup(parent, body["name"]):
                return response({}, status=409)
            identifier = self.add(body["name"], parent)
            return response(self.metadata(identifier), status=201)
        if tail.endswith(":/createUploadSession"):
            target = tail[:-len(":/createUploadSession")]
            parent, name = (unquote(part, errors="strict") for part in target.split(":/", 1))
            behavior = kwargs["json"]["item"]["@microsoft.graph.conflictBehavior"]
            existing = self.lookup(parent, name)
            if existing and behavior == "fail":
                return response({}, status=409)
            if existing and behavior == "rename":
                name += " (copy)"
                assert not self.lookup(parent, name)
                existing = None
            key = "session-" + str(len(self.sessions))
            self.sessions[key] = dict(parent=parent, name=name, existing=existing, data=bytearray())
            return response({"uploadUrl": "https://upload.invalid/" + key})
        raise AssertionError("Unknown synthetic Graph operation")

    def transfer(self, method, url, **kwargs):
        self.records.append((method, url))
        assert not any(k.lower() in ("authorization", "cookie", "proxy-authorization") for k in kwargs["headers"])
        key = urlsplit(url).path.strip("/")
        session = self.sessions[key]
        if method == "GET":
            return response({"nextExpectedRanges": [str(len(session["data"]))+"-"]})
        assert method == "PUT"
        match = re.fullmatch(r"bytes (\d+)-(\d+)/(\d+)", kwargs["headers"]["Content-Range"])
        assert match
        start, end, total = map(int, match.groups())
        assert start == len(session["data"]) and end-start+1 == len(kwargs["data"])
        assert kwargs["headers"]["Content-Length"] == str(len(kwargs["data"]))
        session["data"].extend(kwargs["data"])
        if end + 1 < total:
            return response({"nextExpectedRanges": [str(end+1)+"-"]}, status=202)
        assert len(session["data"]) == total
        identifier = session["existing"]
        if identifier:
            self.nodes[identifier]["_data"] = bytes(session["data"])
            self.nodes[identifier]["size"] = total
            return response(self.metadata(identifier), status=200)
        identifier = self.add(session["name"], session["parent"], bytes(session["data"]))
        return response(self.metadata(identifier), status=201)


class OneDriveModelWorkflows(DriveFixture):
    def test_search_never_returns_paths_rejected_by_the_same_depth_policy(self):
        client, service = self.model(max_depth=1)
        sub = service.add("A")
        service.add("report.txt", sub, b"abc")
        with self.assertRaises(GraphError):
            client.find_files("report")

    def test_explicit_search_subfolder_does_not_reset_global_depth(self):
        client, service = self.model(max_depth=1)
        sub = service.add("A")
        service.add("report.txt", sub, b"abc")
        with self.assertRaises(GraphError):
            client.find_files("report", folder_path="A")

    def test_listing_does_not_advertise_unreadable_children_at_depth_boundary(self):
        client, service = self.model(max_depth=1)
        sub = service.add("A")
        service.add("report.txt", sub, b"abc")
        with self.assertRaises(GraphError):
            client.list_children("A")

    def test_mutations_cannot_create_paths_outside_the_depth_policy(self):
        client, service = self.model(max_depth=1)
        service.add("A")
        before = len(service.records)
        with self.assertRaises(ValueError):
            client.upload(io.BytesIO(b"abc"), "A", "report.txt")
        with self.assertRaises(ValueError):
            client.create_folder("New", "A")
        self.assertEqual(len(service.records), before)

    def test_root_file_at_depth_boundary_still_uploads_and_downloads(self):
        client, service = self.model(max_depth=1)
        committed = client.upload(io.BytesIO(b"abc"), new_file_name="report.txt")
        self.assertEqual(client.find_files("report"), [committed])
        with tempfile.TemporaryDirectory() as local:
            self.assertEqual(client.download_file("report.txt", local).read_bytes(), b"abc")

    def test_empty_folder_at_depth_boundary_can_be_listed_and_zipped(self):
        client, service = self.model(max_depth=1)
        service.add("A")
        self.assertEqual(client.list_children("A"), [])
        with tempfile.TemporaryDirectory() as local:
            archive = client.download_folder("A", local, compress=True)
            with zipfile.ZipFile(archive) as z:
                self.assertEqual(z.namelist(), ["A/"])

    def model(self, **options):
        client = self.client(**options)
        service = MemoryDrive()
        self.api.session.request.side_effect = service.graph
        self.api.transfer_session.request.side_effect = service.transfer
        return client, service

    def test_create_upload_query_download_roundtrip_preserves_sha256(self):
        client, service = self.model()
        created = client.create_folder("Incoming")
        payload = "UTF-8 content café 🧪\n".encode() + bytes(range(256))*17
        item = client.upload(io.BytesIO(payload), "Incoming", "O'Brien #1.bin")
        self.assertEqual(item["parentReference"]["id"], created["id"])
        self.assertEqual(client.find_files("o'brien", folder_path="Incoming"), [item])
        self.assertEqual(client.find_folders("Incoming", recursive=False), [created])
        with tempfile.TemporaryDirectory() as local:
            path = client.download_file("Incoming/O'Brien #1.bin", local)
            self.assertEqual(hashlib.sha256(path.read_bytes()).digest(), hashlib.sha256(payload).digest())
        self.assertTrue(service.records)
        self.assertFalse(any("/sites/" in url or "/me/" in url for _, url in service.records))

    def test_fail_replace_and_rename_preserve_their_distinct_semantics(self):
        client, service = self.model()
        first = client.upload(io.BytesIO(b"ORIGINAL"), new_file_name="report.bin")
        with self.assertRaises(GraphError) as error:
            client.upload(io.BytesIO(b"NO"), new_file_name="report.bin")
        self.assertEqual(error.exception.status_code, 409)
        self.assertEqual(service.nodes[first["id"]]["_data"], b"ORIGINAL")
        replaced = client.upload(io.BytesIO(b"REPLACED"), new_file_name="report.bin", conflict_behavior="replace")
        self.assertEqual(replaced["id"], first["id"])
        self.assertEqual(service.nodes[first["id"]]["_data"], b"REPLACED")
        renamed = client.upload(io.BytesIO(b"COPY"), new_file_name="report.bin", conflict_behavior="rename")
        self.assertNotEqual(renamed["id"], first["id"])
        self.assertNotEqual(renamed["name"], first["name"])
        self.assertEqual(service.nodes[first["id"]]["_data"], b"REPLACED")

    def test_empty_upload_causes_no_remote_mutation(self):
        client, service = self.model()
        with self.assertRaises(ValueError):
            client.upload(io.BytesIO(b""), new_file_name="empty")
        self.assertTrue(all(method == "GET" for method, _ in service.records))
        self.assertEqual(set(service.nodes), {ROOT})

    def test_chunk_boundaries_roundtrip_without_loss_or_duplication(self):
        client, service = self.model(max_requests=500)
        for size in (1, CHUNK-1, CHUNK, CHUNK+1, 2*CHUNK, 2*CHUNK+1):
            payload = bytes(range(256))*(size//256) + bytes(range(size%256))
            name = "boundary-" + str(size) + ".bin"
            committed = client.upload(io.BytesIO(payload), new_file_name=name, chunk_size=CHUNK)
            self.assertEqual(hashlib.sha256(service.nodes[committed["id"]]["_data"]).digest(),
                             hashlib.sha256(payload).digest())
            with tempfile.TemporaryDirectory() as local:
                self.assertEqual(client.download_file(name, local).read_bytes(), payload)

    def test_tree_zip_keeps_nested_and_empty_directories(self):
        client, service = self.model()
        client.create_folder("A")
        client.create_folder("Empty", "A")
        client.upload(io.BytesIO(b"abc"), "A", "report.txt")
        with tempfile.TemporaryDirectory() as local:
            archive = client.download_folder("", local, compress=True, max_workers=1)
            with zipfile.ZipFile(archive) as z:
                self.assertIsNone(z.testzip())
                self.assertEqual(set(z.namelist()), {"A/", "A/Empty/", "A/report.txt"})
                self.assertEqual(z.read("A/report.txt"), b"abc")
            self.assertEqual((Path(local)/"Archive/A/report.txt").read_bytes(), b"abc")
            self.assertTrue((Path(local)/"Archive/A/Empty").is_dir())

    def test_permission_loss_is_not_recovered_by_wider_queries(self):
        client, service = self.model()
        service.denied = True
        with self.assertRaises(GraphError) as error:
            client.list_children()
        self.assertEqual(error.exception.status_code, 403)
        self.assertEqual(len(service.records), 1)
        self.assertTrue(service.records[0][1].endswith("/items/anchor"))

    def test_metadata_mutation_by_caller_does_not_retarget_the_client(self):
        client, service = self.model()
        client.create_folder("A")
        first = client.list_children()
        first[0]["id"] = "outside"
        first[0]["parentReference"]["id"] = "outside"
        actual = client.list_children()
        self.assertNotEqual(actual[0]["id"], "outside")
        self.assertEqual(actual[0]["parentReference"]["id"], ROOT)

    def test_empty_nested_folder_creation_cannot_merge_with_an_existing_name(self):
        client, service = self.model()
        client.create_folder("A")
        with self.assertRaises(GraphError) as error:
            client.create_folder("A")
        self.assertEqual(error.exception.status_code, 409)
        self.assertEqual(len(service.nodes), 2)


class OneDriveGenerated(DriveFixture):
    pass


def valid_name_case(name):
    def test(self):
        expected = file_item(name=name)
        client = self.client([response(folder()), response(expected)])
        self.assertEqual(client.get_item(name), expected)
        self.assertTrue(self.calls()[-1].args[1].endswith("anchor:/"+quote(name, safe="")))
    return test


def bad_path_case(path):
    def test(self):
        client = self.client()
        with self.assertRaises((ValueError, TypeError)):
            client.get_item(path)
        self.assertEqual(len(self.calls()), 1)
    return test


def foreign_case(target):
    def test(self):
        client = self.client([response(folder()), page(file_item(), link=target)])
        with self.assertRaises((ValueError, GraphError)):
            client.list_children()
        self.assertEqual(len(self.calls()), 3)
    return test


def ancestor_case(depth, ancestor_index, upload):
    def test(self):
        nodes = [folder()]
        path = []
        for i in range(depth):
            value = folder("ancestor-"+str(i), "Folder"+str(i), nodes[-1]["id"])
            path.append(value["name"])
            nodes.append(value)
        ancestor_id = nodes[ancestor_index]["id"]
        replies = [response(n) for n in nodes]
        if upload:
            replies.append(response({"uploadUrl": "https://upload.invalid/generated"}))
            client = self.client(replies, transfers=[
                response(file_item(ancestor_id, parent=nodes[-1]["id"]), status=201)])
            with self.assertRaises(GraphError):
                client.upload(io.BytesIO(b"abc"), "/".join(path), "report.txt")
        else:
            replies.append(response(folder(ancestor_id, "New", nodes[-1]["id"]), status=201))
            client = self.client(replies)
            with self.assertRaises(GraphError):
                client.create_folder("New", "/".join(path))
    return test


def invalid_utf8_scope_case(index):
    def test(self):
        anchor = "\ufffdanchor-" + str(index)
        root = folder(anchor)
        link = BASE + "/drives/drive/items/%FFanchor-" + str(index) + "/children?token=opaque"
        client = self.client([response(root), page(link=link), page()], root=root, folder_id=anchor)
        with self.assertRaises((ValueError, GraphError)):
            client.list_children()
        self.assertEqual(len(self.calls()), 3)
    return test


def bad_metadata_case(index):
    def test(self):
        kind = index % 8
        row = file_item(str(index), "record-"+str(index)+".txt")
        if kind == 0:
            row["size"] = -index-1
        elif kind == 1:
            row["size"] = False
        elif kind == 2:
            row["parentReference"]["id"] = "outside-"+str(index)
        elif kind == 3:
            row["parentReference"]["driveId"] = "other-"+str(index)
        elif kind == 4:
            row["file"] = None
        elif kind == 5:
            row["folder"] = {}
        elif kind == 6:
            row["deleted"] = {}
        else:
            row["error"] = {"code": "synthetic-"+str(index)}
        client = self.client([response(folder()), page(row)])
        with self.assertRaises(GraphError):
            client.list_children()
    return test


def bad_download_case(index):
    def test(self):
        size = index + 10
        mode = index % 4
        data = b"x"*(size-1 if mode == 0 else size+1 if mode == 1 else size)
        reply = binary(data=data, status=206 if mode == 2 else 200,
                       headers={"Content-Length": str(size+1)} if mode == 3 else None)
        client = self.client([response(folder()), response(file_item(size=size)), reply])
        with tempfile.TemporaryDirectory() as local:
            with self.assertRaises(GraphError):
                client.download_file("report.txt", local)
            self.assertEqual(list(Path(local).iterdir()), [])
        reply.close.assert_called_once()
    return test


_rng = random.Random(20261119)
_alphabet = "abcXYZ09éßΩ🧪%+#&'()=-_ "
for _index in range(120):
    _name = "fixture-" + "".join(_rng.choice(_alphabet) for _ in range(24)) + ".bin"
    setattr(OneDriveGenerated, "test_valid_name_%03d" % _index, valid_name_case(_name))
_bad_prefixes = ("../", "/", "a//", "a/../", "a/./", "bad\\", "C:", "bad\x00", "bad\n", "bad?", "bad|", "bad*")
for _index in range(120):
    _path = _bad_prefixes[_index % len(_bad_prefixes)] + str(_index) + ".bin"
    setattr(OneDriveGenerated, "test_invalid_path_%03d" % _index, bad_path_case(_path))
for _index in range(60):
    _paths = (
        "/drives/other-%d/items/anchor/children" % _index,
        "/drives/drive/items/other-%d/children" % _index,
        "/drives/drive/root/search(q='record-%d')" % _index,
    )
    _target = BASE + _paths[_index % 3] + "?$skiptoken=unchanged%2B%2F%3D"
    setattr(OneDriveGenerated, "test_foreign_continuation_%03d" % _index, foreign_case(_target))
for _index in range(30):
    _depth = 2 + _index % 5
    _ancestor = 1 + _index % (_depth-1)
    setattr(OneDriveGenerated, "test_ancestor_confirmation_%03d" % _index,
            ancestor_case(_depth, _ancestor, _index % 2 == 0))
for _index in range(10):
    setattr(OneDriveGenerated, "test_invalid_utf8_continuation_%03d" % _index, invalid_utf8_scope_case(_index))
for _index in range(40):
    setattr(OneDriveGenerated, "test_bad_metadata_%03d" % _index, bad_metadata_case(_index))
for _index in range(20):
    setattr(OneDriveGenerated, "test_incomplete_download_%03d" % _index, bad_download_case(_index))

GENERATED_CASE_COUNT = sum(name.startswith("test_") for name in vars(OneDriveGenerated))
assert GENERATED_CASE_COUNT == 400
