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
"""Offline regressions for OneDrive binding, queries, and file transfers."""
from copy import deepcopy
import hashlib
import io
import json
import os
from pathlib import Path
import tempfile
from types import SimpleNamespace as NS
import unittest
from unittest.mock import Mock, patch
from urllib.parse import urlsplit
import zipfile

from m365.graph import Drive, GraphError
from m365.graph._graph import GraphAPI
from test_sharepoint import Offline, response
from test_transfer_hardening import binary

BASE = "https://graph.microsoft.com/v1.0"
DRIVE = "drive"
ROOT = "anchor"
CHUNK = 327680


def folder(identifier=ROOT, name="Archive", parent=None, **extra):
    result = dict(id=identifier, name=name, folder={})
    if parent is not None:
        result["parentReference"] = {"id": parent, "driveId": DRIVE}
    result.update(extra)
    return result


def file_item(identifier="f", name="report.txt", parent=ROOT, size=3, **extra):
    result = dict(id=identifier, name=name, size=size, file={},
                  parentReference={"id": parent, "driveId": DRIVE})
    result.update(extra)
    return result


def page(*items, link=None):
    result = {"value": list(items)}
    if link is not None:
        result["@odata.nextLink"] = link
    return response(result)


class DriveFixture(Offline):
    def client(self, replies=(), transfers=(), *, root=None, token_engine=None, **options):
        self.created_apis = []
        self.provider = token_engine or NS(access_token="SYNTHETIC-TOKEN")
        if root is None:
            root = folder()
        def factory(parent):
            api = GraphAPI(parent)
            api.session.request = Mock(side_effect=[response(root), *replies])
            api.transfer_session.request = Mock(side_effect=list(transfers))
            api.session.close = Mock(wraps=api.session.close)
            api.transfer_session.close = Mock(wraps=api.transfer_session.close)
            self.created_apis.append(api)
            return api
        self.factory = Mock(side_effect=factory)
        args = dict(drive_id=DRIVE, folder_id=ROOT, max_retries=0)
        args.update(options)
        with patch("m365.graph._storage.GraphAPI", self.factory):
            client = Drive(self.provider, **args)
        self.api = self.created_apis[-1]
        self.addCleanup(client.close)
        return client

    def calls(self):
        return self.api.session.request.call_args_list

    def paths(self):
        return [urlsplit(call.args[1]).path for call in self.calls()]


class OneDriveBasics(DriveFixture):
    def test_exact_folder_binding_has_no_site_or_drive_root_discovery(self):
        client = self.client()
        self.assertEqual(client.drive_id, DRIVE)
        self.assertEqual(client.folder_id, ROOT)
        self.assertEqual(self.paths(), ["/v1.0/drives/drive/items/anchor"])

    def test_explicit_full_drive_root_binds_its_resolved_id(self):
        client = self.client(root=folder("resolved-root", "OneDrive"), folder_id=None)
        self.assertEqual(client.folder_id, "resolved-root")
        self.assertEqual(self.paths(), ["/v1.0/drives/drive/root"])

    def test_get_root_metadata_is_copied_and_cannot_retarget(self):
        client = self.client([response(folder()), response(folder())])
        value = client.get_item()
        value["id"] = "other"
        self.assertEqual(client.get_item()["id"], ROOT)
        with self.assertRaises(AttributeError):
            client.drive_id = "other"
        with self.assertRaises(AttributeError):
            client.folder_id = "other"

    def test_list_root_files_and_folders(self):
        rows = [file_item(), folder("sub", "Projects", ROOT)]
        client = self.client([response(folder()), page(*rows)])
        self.assertEqual(client.list_children(), rows)
        self.assertEqual(self.paths()[-1], "/v1.0/drives/drive/items/anchor/children")

    def test_relative_nested_path_is_resolved_under_each_parent_id(self):
        sub = folder("sub", "Projects", ROOT)
        item = file_item(parent="sub")
        client = self.client([response(folder()), response(sub), response(item)])
        self.assertEqual(client.get_item("Projects/report.txt"), item)
        self.assertEqual(self.paths()[-2:], ["/v1.0/drives/drive/items/anchor:/Projects",
                                            "/v1.0/drives/drive/items/sub:/report.txt"])

    def test_listing_subfolder_stays_at_that_folder(self):
        sub = folder("sub", "Projects", ROOT)
        item = file_item(parent="sub")
        client = self.client([response(folder()), response(sub), page(item)])
        self.assertEqual(client.list_children("Projects"), [item])
        self.assertEqual(self.paths()[-1], "/v1.0/drives/drive/items/sub/children")

    def test_empty_folder_returns_empty_list(self):
        client = self.client([response(folder()), page()])
        self.assertEqual(client.list_children(), [])

    def test_opaque_pagination_is_not_rewritten(self):
        link = BASE + "/drives/drive/items/anchor/children?$skiptoken=A%2BB+z%26&x=%252F"
        a, b = file_item("a", "a.txt"), file_item("b", "b.txt")
        client = self.client([response(folder()), page(a, link=link), page(b)])
        self.assertEqual(client.list_children(), [a, b])
        self.assertEqual(self.calls()[-1].args[1], link)

    def test_filename_search_is_bounded_local_name_search_not_indexed(self):
        a = file_item("a", "REPORT.txt")
        sub = folder("sub", "Projects", ROOT)
        b = file_item("b", "report-2.txt", "sub")
        client = self.client([response(folder()), page(a, sub), page(b)])
        self.assertEqual(client.find_files("report"), [a, b])
        self.assertFalse(any("search(" in p for p in self.paths()))

    def test_exact_and_partial_search_are_distinct(self):
        a, b = file_item("a", "Report.txt"), file_item("b", "Report.txt.bak")
        client = self.client([response(folder()), page(a, b)])
        self.assertEqual(client.find_files("report.txt", exact_match=True), [a])

    def test_nonrecursive_search_does_not_fetch_children(self):
        a, sub = file_item(), folder("sub", "Projects", ROOT)
        client = self.client([response(folder()), page(a, sub)])
        self.assertEqual(client.find_files("report", recursive=False), [a])
        self.assertEqual(len(self.calls()), 3)

    def test_folder_search_retains_folder_role(self):
        sub = folder("sub", "Projects", ROOT)
        client = self.client([response(folder()), page(file_item(), sub)])
        self.assertEqual(client.find_folders("projects", recursive=False), [sub])

    def test_zero_limit_does_not_issue_an_operation_request(self):
        client = self.client()
        self.assertEqual(client.find_files("report", limit=0), [])
        self.assertEqual(client.find_folders("projects", limit=0), [])
        self.assertEqual(len(self.calls()), 1)

    def test_explicit_search_limit_is_an_intentional_subset(self):
        a, b = file_item("a", "report-a.txt"), file_item("b", "report-b.txt")
        link = BASE + "/drives/drive/items/anchor/children?page=2"
        client = self.client([response(folder()), page(a, b, link=link)])
        self.assertEqual(client.find_files("report", limit=1), [a])
        self.assertEqual(len(self.calls()), 3)

    def test_nested_folder_creation_is_fail_on_conflict(self):
        sub = folder("sub", "Projects", ROOT)
        created = folder("new", "New", "sub")
        client = self.client([response(folder()), response(sub), response(created, status=201)])
        self.assertEqual(client.create_folder("New", "Projects"), created)
        call = self.calls()[-1]
        self.assertEqual(call.args[0], "POST")
        self.assertEqual(call.kwargs["json"], {"name": "New", "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
        self.assertTrue(call.args[1].endswith("/items/sub/children"))

    def test_create_folder_conflict_propagates_without_retry_or_merge(self):
        client = self.client([response(folder()), response({"error": {}}, status=409)])
        with self.assertRaises(GraphError) as error:
            client.create_folder("Existing")
        self.assertEqual(error.exception.status_code, 409)
        self.assertEqual(len(self.calls()), 3)

    def test_create_result_cannot_be_parent_or_another_kind(self):
        for item in (folder(ROOT, "New", ROOT), file_item(name="New"), folder("new", "Other", ROOT)):
            with self.subTest(item=item):
                client = self.client([response(folder()), response(item, status=201)])
                with self.assertRaises(GraphError):
                    client.create_folder("New")

    def test_close_is_idempotent_and_context_manager_closes(self):
        client = self.client()
        with client as selected:
            self.assertIs(selected, client)
        client.close()
        self.api.session.close.assert_called_once()
        self.api.transfer_session.close.assert_called_once()
        with self.assertRaises(ValueError):
            client.get_item()
        with self.assertRaises(ValueError):
            client.__enter__()

    def test_constructor_failure_closes_both_sessions(self):
        with self.assertRaises(GraphError):
            self.client(root=folder("wrong"))
        for api in self.created_apis:
            api.session.close.assert_called_once()
            api.transfer_session.close.assert_called_once()

    def test_provider_is_used_without_browser_or_secret_reconfiguration(self):
        class Provider:
            def __init__(self):
                self.calls = 0
            @property
            def access_token(self):
                self.calls += 1
                return "TOKEN-" + str(self.calls)
        provider = Provider()
        client = self.client([response(folder())], token_engine=provider)
        client.get_item()
        self.assertEqual(provider.calls, 2)
        self.assertEqual(self.calls()[0].kwargs["headers"]["Authorization"], "Bearer TOKEN-1")
        self.assertEqual(self.calls()[1].kwargs["headers"]["Authorization"], "Bearer TOKEN-2")

    def test_child_drive_id_may_be_omitted_but_not_contradicted(self):
        row = file_item()
        row["parentReference"].pop("driveId")
        client = self.client([response(folder()), page(row)])
        self.assertEqual(client.list_children(), [row])


class OneDriveTransfers(DriveFixture):
    def session(self):
        return response({"uploadUrl": "https://upload.example.invalid/session?capability=SYNTHETIC"})

    def test_small_upload_uses_resolved_folder_id_not_drive_root(self):
        committed = file_item(size=3)
        source = io.BytesIO(b"abc")
        source.seek(2)
        client = self.client([response(folder()), self.session()],
                             transfers=[response(committed, status=201)])
        self.assertEqual(client.upload(source, new_file_name="report.txt"), committed)
        self.assertEqual(source.tell(), 2)
        self.assertTrue(self.calls()[-1].args[1].endswith("/items/anchor:/report.txt:/createUploadSession"))
        self.assertEqual(self.calls()[-1].kwargs["json"]["item"]["@microsoft.graph.conflictBehavior"], "fail")
        transfer = self.api.transfer_session.request.call_args
        self.assertEqual(transfer.kwargs["data"], b"abc")
        self.assertFalse(any(k.lower() == "authorization" for k in transfer.kwargs["headers"]))

    def test_nested_upload_uses_direct_child_folder_id(self):
        sub = folder("sub", "Projects", ROOT)
        committed = file_item(parent="sub")
        client = self.client([response(folder()), response(sub), self.session()],
                             transfers=[response(committed, status=201)])
        client.upload(io.BytesIO(b"abc"), "Projects", "report.txt")
        self.assertIn("/items/sub:/report.txt:/", self.calls()[-1].args[1])

    def test_chunked_upload_validates_ranges_and_completion(self):
        payload = b"x" * CHUNK + b"tail"
        committed = file_item(size=len(payload))
        client = self.client([response(folder()), self.session()], transfers=[
            response({"nextExpectedRanges": [str(CHUNK)+"-"]}, status=202),
            response(committed, status=201)])
        self.assertEqual(client.upload(io.BytesIO(payload), new_file_name="report.txt", chunk_size=CHUNK), committed)
        calls = self.api.transfer_session.request.call_args_list
        self.assertEqual(b"".join(c.kwargs["data"] for c in calls), payload)
        self.assertEqual(calls[0].kwargs["headers"]["Content-Range"], "bytes 0-327679/327684")
        self.assertEqual(calls[1].kwargs["headers"]["Content-Range"], "bytes 327680-327683/327684")

    def test_explicit_replace_is_not_the_default(self):
        committed = file_item()
        client = self.client([response(folder()), self.session()], transfers=[response(committed, status=200)])
        client.upload(io.BytesIO(b"abc"), new_file_name="report.txt", conflict_behavior="replace")
        self.assertEqual(self.calls()[-1].kwargs["json"]["item"]["@microsoft.graph.conflictBehavior"], "replace")

    def test_rename_conflict_behavior_requires_valid_confirmed_metadata(self):
        committed = file_item(name="report (1).txt")
        client = self.client([response(folder()), self.session()], transfers=[response(committed, status=201)])
        result = client.upload(io.BytesIO(b"abc"), new_file_name="report.txt", conflict_behavior="rename")
        self.assertEqual(result["name"], "report (1).txt")

    def test_real_local_file_upload_uses_file_basename(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory)/"report.txt"
            source.write_bytes(b"abc")
            client = self.client([response(folder()), self.session()], transfers=[response(file_item(), status=201)])
            self.assertEqual(client.upload(source)["id"], "f")

    def test_upload_completion_requires_the_exact_parent_and_drive(self):
        for parent in ({}, {"id": "outside", "driveId": DRIVE}, {"id": ROOT, "driveId": "other"},
                       {"id": ROOT}, None):
            with self.subTest(parent=parent):
                committed = file_item(parentReference=parent)
                client = self.client([response(folder()), self.session()], transfers=[response(committed, status=201)])
                with self.assertRaises(GraphError):
                    client.upload(io.BytesIO(b"abc"), new_file_name="report.txt")
                self.assertEqual(self.api.transfer_session.request.call_count, 1)

    def test_partial_or_progress_final_response_never_claims_success(self):
        for status in (202, 204, 206, 299):
            with self.subTest(status=status):
                client = self.client([response(folder()), self.session()], transfers=[response(file_item(), status=status)])
                with self.assertRaises(GraphError):
                    client.upload(io.BytesIO(b"abc"), new_file_name="report.txt")

    def test_session_creation_requires_complete_200_response(self):
        for status in (201, 202, 206):
            with self.subTest(status=status):
                client = self.client([response(folder()), response({"uploadUrl": "https://upload.example.invalid/x"}, status=status)])
                with self.assertRaises(GraphError):
                    client.upload(io.BytesIO(b"abc"), new_file_name="report.txt")
                self.api.transfer_session.request.assert_not_called()

    def test_empty_upload_never_creates_a_session(self):
        client = self.client([response(folder())])
        with self.assertRaises(ValueError):
            client.upload(io.BytesIO(b""), new_file_name="empty.txt")
        self.assertFalse(any(c.args[0] != "GET" for c in self.calls()))
        self.api.transfer_session.request.assert_not_called()

    def test_binary_download_matches_original_sha256(self):
        payload = bytes(range(256)) * 100 + "Unicode content: café 🧪".encode()
        item = file_item(size=len(payload))
        reply = binary(data=payload, headers={"Content-Length": str(len(payload))})
        client = self.client([response(folder()), response(item), reply])
        with tempfile.TemporaryDirectory() as directory:
            result = client.download_file("report.txt", directory)
            self.assertEqual(hashlib.sha256(result.read_bytes()).digest(), hashlib.sha256(payload).digest())
            self.assertFalse(list(Path(directory).glob(".msftgraph-*")))
        reply.close.assert_called_once()

    def test_download_missing_path_is_an_error_not_an_empty_file(self):
        client = self.client([response(folder()), response({}, status=404)])
        with tempfile.TemporaryDirectory() as directory, self.assertRaises(GraphError):
            client.download_file("missing.txt", directory)
        self.assertEqual(len(self.calls()), 3)

    def test_download_cannot_overwrite_existing_local_file(self):
        client = self.client([response(folder()), response(file_item())])
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)/"report.txt"
            target.write_bytes(b"KEEP")
            with self.assertRaises(FileExistsError):
                client.download_file("report.txt", directory)
            self.assertEqual(target.read_bytes(), b"KEEP")
            self.assertEqual(len(self.calls()), 3)

    def test_partial_download_is_not_published(self):
        for status in (202, 204, 206):
            with self.subTest(status=status):
                reply = binary(status=status)
                client = self.client([response(folder()), response(file_item()), reply])
                with tempfile.TemporaryDirectory() as directory:
                    with self.assertRaises(GraphError):
                        client.download_file("report.txt", directory)
                    self.assertFalse(list(Path(directory).iterdir()))
                reply.close.assert_called_once()

    def test_truncated_oversized_and_interrupted_downloads_leave_no_file(self):
        cases = [binary(data=b"ab"), binary(data=b"abcd")]
        reset = binary()
        reset.iter_content = Mock(side_effect=__import__("requests").ConnectionError("SYNTHETIC PRIVATE URL"))
        cases.append(reset)
        for reply in cases:
            client = self.client([response(folder()), response(file_item()), reply])
            with tempfile.TemporaryDirectory() as directory:
                with self.assertRaises(GraphError) as error:
                    client.download_file("report.txt", directory)
                self.assertNotIn("PRIVATE", str(error.exception))
                self.assertFalse(list(Path(directory).iterdir()))
            reply.close.assert_called_once()

    def test_download_folder_zip_preserves_empty_dirs_bytes_and_originals(self):
        payload = b"abc"
        sub = folder("empty", "Empty", ROOT)
        client = self.client([response(folder()), page(file_item(), sub), page(), binary(data=payload)])
        with tempfile.TemporaryDirectory() as directory:
            archive = client.download_folder("", directory, compress=True, max_workers=1)
            root = Path(directory)/"Archive"
            self.assertEqual((root/"report.txt").read_bytes(), payload)
            self.assertTrue((root/"Empty").is_dir())
            with zipfile.ZipFile(archive) as z:
                self.assertIsNone(z.testzip())
                self.assertEqual(set(z.namelist()), {"report.txt", "Empty/"})
                self.assertEqual(z.read("report.txt"), payload)

    def test_completely_empty_folder_zip_has_a_directory_entry(self):
        client = self.client([response(folder()), page()])
        with tempfile.TemporaryDirectory() as directory:
            archive = client.download_folder("", directory, compress=True)
            with zipfile.ZipFile(archive) as z:
                self.assertEqual(z.namelist(), ["Archive/"])
            self.assertTrue((Path(directory)/"Archive").is_dir())

    def test_folder_plan_failure_creates_no_local_output(self):
        sub = folder("sub", "Sub", ROOT)
        client = self.client([response(folder()), page(sub), response({}, status=403)])
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(GraphError):
                client.download_folder("", directory)
            self.assertEqual(list(Path(directory).iterdir()), [])

    def test_folder_download_rejects_existing_root_or_zip_without_overwrite(self):
        for archive_exists in (False, True):
            client = self.client([response(folder()), page()])
            with tempfile.TemporaryDirectory() as directory:
                target = Path(directory)/("Archive.zip" if archive_exists else "Archive")
                if archive_exists:
                    target.write_bytes(b"KEEP")
                else:
                    target.mkdir()
                with self.assertRaises(FileExistsError):
                    client.download_folder("", directory, compress=archive_exists)
                if archive_exists:
                    self.assertEqual(target.read_bytes(), b"KEEP")

    def test_source_position_is_restored_after_upload_failure(self):
        source = io.BytesIO(b"abc")
        source.seek(2)
        client = self.client([response(folder()), self.session()], transfers=[response({}, status=403)])
        with self.assertRaises(GraphError):
            client.upload(source, new_file_name="report.txt")
        self.assertEqual(source.tell(), 2)
