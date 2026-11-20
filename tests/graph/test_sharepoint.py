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
"""Offline contract regressions; no credentials, tenants, or network required."""
import ast
import io
import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import requests

from m365.graph import Site, GraphError, _syslib as syslib, _dateplumb as dateplumb
from m365.graph._graph import GraphAPI
from m365.graph._site_resolution import drive_path
from m365.graph.operations.doc_sync import Upload
from m365.graph.operations.doc_transfer import Download, _child, _compress
from m365.graph.operations.doc_dispatch import Move
from m365.graph.operations.doc_management import File as FileSearch, Folder as FolderSearch, FileFolder
from m365.graph._site_resolution import library_relative


def response(body=None, status=200, headers=None, content=None):
    r = requests.Response()
    r.status_code = status
    r.headers.update({"Content-Type": "application/json"})
    r.headers.update(headers or {})
    r._content = json.dumps(body).encode() if content is None else content
    r.close = Mock()
    return r


class Offline(unittest.TestCase):
    def setUp(self):
        # Any unmocked HTTP call is a hard test failure, not an accidental live request.
        self.network = patch("requests.sessions.Session.request", side_effect=AssertionError("Live network forbidden"))
        self.network.start()
        self.addCleanup(self.network.stop)

    def api(self, token="test-token"):
        parent = SimpleNamespace(graph_url="https://graph.microsoft.com/v1.0",
            request_timeout=(10, 60), max_retries=2, get_token=lambda: token)
        api = GraphAPI(parent)
        self.addCleanup(api.close)
        api.session.request = Mock()
        api.transfer_session.request = Mock()
        return api

    def site(self, replies, **kwargs):
        session = requests.Session()
        session.request = Mock(side_effect=replies)
        with patch("m365.graph._graph.requests.Session", return_value=session):
            site = Site(SimpleNamespace(access_token="org-specific-token"), **kwargs)
        self.addCleanup(site.api.close)
        return site, session.request


class TransportTests(Offline):
    def test_preserves_arbitrary_json(self):
        api = self.api()
        api.session.request.return_value = response({"uploadUrl": "https://uploads.example/session?secret=1"})
        body = {"item": {"@microsoft.graph.conflictBehavior": "fail"}}
        api.graph_call(resource_url="drives/d/root:/f:/createUploadSession", method="POST", json=body)
        self.assertEqual(api.session.request.call_args.kwargs["json"], body)
        self.assertEqual(api.session.request.call_args.kwargs["timeout"], (10, 60))
        self.assertFalse(api.session.request.call_args.kwargs["allow_redirects"])

    def test_transfer_does_not_get_bearer_or_cookies(self):
        api = self.api()
        api.transfer_session.request.return_value = response({"id": "i"}, status=201)
        result = api.graph_call(resource_url="https://uploads.example/u?capability=private", method="PUT",
            data=b"abc", headers={"authorization": "bad", "Cookie": "bad"}, preauthenticated=True)
        self.assertEqual(result, {"id": "i"})
        headers = api.transfer_session.request.call_args.kwargs["headers"]
        self.assertFalse(any(k.lower() in ("authorization", "cookie") for k in headers))
        api.session.request.assert_not_called()

    def test_external_graph_url_rejected_before_token_use(self):
        api = self.api()
        for url in ("https://evil.example/u", "//evil.example/u", "http://graph.microsoft.com/v1.0/me",
                    "https://graph.microsoft.com/v1.0/../beta/me", "https://graph.microsoft.com.evil/v1.0/me"):
            with self.subTest(url=url), self.assertRaises(ValueError):
                api.graph_call(resource_url=url)
        api.session.request.assert_not_called()

    def test_204_returns_empty_object(self):
        api = self.api()
        api.session.request.return_value = response(status=204, content=b"")
        self.assertEqual(api.graph_call(resource_url="me"), {})

    def test_bytes_upload_returns_json_metadata(self):
        api = self.api()
        api.session.request.return_value = response({"id": "i", "size": 3})
        self.assertEqual(api.graph_call(resource_url="drives/d/items/i/content", method="PUT", data=b"abc"),
                         {"id": "i", "size": 3})

    @patch("m365.graph._graph.time.sleep")
    def test_retry_after_and_exhaustion(self, sleep):
        api = self.api()
        api.session.request.side_effect = [response(status=429, headers={"Retry-After": "2"}), response({"id": "ok"})]
        self.assertEqual(api.graph_call(resource_url="me"), {"id": "ok"})
        sleep.assert_called_once_with(2)
        api.session.request.side_effect = None
        api.session.request.return_value = response(status=503)
        with self.assertRaises(GraphError) as ctx:
            api.graph_call(resource_url="me")
        self.assertEqual(ctx.exception.status_code, 503)

    def test_http_date_retry_after(self):
        self.assertEqual(GraphAPI._delay(response(headers={"Retry-After": "Mon, 01 Jan 2001 00:00:00 GMT"}), 0), 0)
        self.assertEqual(GraphAPI._delay(response(headers={"Retry-After": "invalid"}), 2), 4)

    @patch("m365.graph._graph.time.sleep")
    def test_long_retry_after_does_not_retry_prematurely(self, sleep):
        api = self.api()
        api.session.request.return_value = response(status=429, headers={"Retry-After": "120"})
        with self.assertRaises(GraphError) as ctx:
            api.graph_call(resource_url="me")
        self.assertEqual(ctx.exception.code, "retry_after_exceeds_budget")
        self.assertEqual(api.session.request.call_count, 1)
        sleep.assert_not_called()

    def test_legacy_data_dict_is_json_not_form_data(self):
        api = self.api()
        api.session.request.return_value = response({"id": "f"})
        api.graph_call(resource_url="drives/d/items/f", method="PATCH", data={"name": "new"})
        self.assertEqual(api.session.request.call_args.kwargs["json"], {"name": "new"})
        self.assertIsNone(api.session.request.call_args.kwargs["data"])

    def test_malformed_json_response_is_redacted(self):
        api = self.api()
        api.session.request.return_value = response(content=b"PRIVATE payload")
        with self.assertRaises(GraphError) as ctx:
            api.graph_call(resource_url="me")
        self.assertEqual(ctx.exception.code, "invalid_json_response")
        self.assertNotIn("PRIVATE", str(ctx.exception))

    def test_no_blind_post_retry(self):
        api = self.api()
        api.session.request.side_effect = requests.Timeout("SECRET URL and TOKEN")
        with self.assertRaises(GraphError) as ctx:
            api.graph_call(resource_url="drives/d/root/children", method="POST", json={"name": "f"})
        self.assertNotIn("SECRET", str(ctx.exception))
        self.assertEqual(api.session.request.call_count, 1)

    def test_redacted_http_errors(self):
        api = self.api()
        api.transfer_session.request.return_value = response({"error": "SECRET TOKEN"}, status=403)
        with self.assertRaises(GraphError) as ctx:
            api.graph_call(resource_url="https://uploads.example/u?secret=PRIVATE", preauthenticated=True)
        self.assertNotIn("PRIVATE", str(ctx.exception))
        self.assertNotIn("SECRET", str(ctx.exception))

    def test_graph_content_redirect_strips_authorization(self):
        api = self.api()
        api.session.request.return_value = response(status=302, headers={"Location": "https://cdn.example/download?secret=1"})
        api.transfer_session.request.return_value = response(content=b"abc", headers={"Content-Type": "application/octet-stream"})
        r = api.graph_call(resource_url="drives/d/items/i/content", stream=True)
        self.assertEqual(r.content, b"abc")
        self.assertNotIn("Authorization", api.transfer_session.request.call_args.kwargs["headers"])

    def test_paginate_and_cycle_detection(self):
        api = self.api()
        api.session.request.side_effect = [response({"value": [1], "@odata.nextLink": "https://graph.microsoft.com/v1.0/list?page=2"}), response({"value": [2]})]
        self.assertEqual(list(api.iter_collection("list")), [1, 2])
        api.session.request.side_effect = None
        api.session.request.return_value = response({"value": [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/list"})
        with self.assertRaises(GraphError):
            list(api.iter_collection("list"))

    def test_batch_partial_failure_is_not_success(self):
        api = self.api()
        api.session.request.return_value = response({"responses": [{"id": "1", "status": 429}]})
        with self.assertRaises(GraphError):
            api.graph_call(batch_requests=[{"id": "1", "method": "GET", "url": "/me"}])
        with self.assertRaises(ValueError):
            api.graph_call(batch_requests=[], batch_size=21)

    def test_batch_success_and_incomplete_response(self):
        api = self.api()
        api.session.request.return_value = response({"responses": [{"id": "1", "status": 200, "body": {"id": "x"}}]})
        result = api.graph_call(batch_requests=[{"id": "1", "method": "GET", "url": "me"}])
        self.assertEqual(result["1"]["body"], {"id": "x"})
        api.session.request.return_value = response({"responses": []})
        with self.assertRaises(GraphError):
            api.graph_call(batch_requests=[{"id": "1", "method": "GET", "url": "me"}])


class SiteTests(Offline):
    def test_site_url_and_default_drive(self):
        opaque = "tenant.example,uuid-one,uuid-two"
        site, request = self.site([response({"id": opaque, "webUrl": "https://tenant.example/teams/Operations"}),
                                  response({"id": "opaque-library", "name": "Library"})],
                                  site_name="https://tenant.example/teams/Operations")
        self.assertEqual(site.site_id, opaque)
        self.assertEqual(site.drive_id, "opaque-library")
        self.assertIn("sites/tenant.example:/teams/Operations", request.call_args_list[0].args[1])
        self.assertIn("sites/tenant.example,uuid-one,uuid-two/drive", request.call_args_list[1].args[1])

    def test_named_drive_on_later_page(self):
        site, request = self.site([response({"id": "s", "webUrl": "https://tenant.example/sites/A"}),
            response({"value": [{"id": "default", "name": "Documents"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/sites/s/drives?page=2"}),
            response({"value": [{"id": "selected", "name": "Other Library"}]})], site_id="s", drive_name="Other Library")
        self.assertEqual(site.drive_id, "selected")
        self.assertEqual(request.call_count, 3)

    def test_explicit_drive_must_belong_to_site(self):
        with self.assertRaises(ValueError):
            self.site([response({"id": "s", "webUrl": "https://tenant.example/sites/A"}), response({"value": [{"id": "another", "name": "X"}]})], site_id="s", drive_id="wrong")

    def test_short_name_uses_provider_root(self):
        site, request = self.site([response({"webUrl": "https://another.example/"}),
            response({"id": "s", "webUrl": "https://another.example/sites/Operations"}), response({"id": "d", "name": "Files"})], site_name="Operations")
        self.assertIn("sites/another.example:/sites/Operations", request.call_args_list[1].args[1])

    def test_retarget_failure_rolls_back_and_raises(self):
        site, request = self.site([response({"id": "s", "webUrl": "https://tenant.example/sites/A"}), response({"id": "d", "name": "Files"})], site_id="s")
        request.side_effect = [response(status=404)]
        with self.assertRaises(GraphError):
            site.site_name = "https://tenant.example/sites/Missing"
        self.assertEqual((site.site_id, site.drive_id, site.drive_name), ("s", "d", "Files"))

    def test_organizations_keep_separate_tokens_and_bindings(self):
        a, b = self.api("org-a"), self.api("org-b")
        a.session.request.return_value = b.session.request.return_value = response({})
        a.graph_call(resource_url="sites/a")
        b.graph_call(resource_url="sites/b")
        self.assertEqual(a.session.request.call_args.kwargs["headers"]["Authorization"], "Bearer org-a")
        self.assertEqual(b.session.request.call_args.kwargs["headers"]["Authorization"], "Bearer org-b")
        self.assertIsNot(a.session, b.session)

    def test_path_encoding_not_library_name_assumptions(self):
        self.assertEqual(drive_path("Documents/a #100%.txt"), "Documents/a%20%23100%25.txt")
        self.assertEqual(drive_path("a%20b"), "a%2520b")
        for path in ("../a", "a/../b", "a\\b", "a//b"):
            with self.assertRaises(ValueError):
                drive_path(path)


class UploadTests(Offline):
    def client(self, replies):
        api = Mock()
        api.graph_call.side_effect = replies
        return SimpleNamespace(api=api, drive_id="chosen-drive", site_id="s")

    def test_small_upload_uses_selected_drive_and_conflict_fail(self):
        client = self.client([{"uploadUrl": "https://upload.example/session"},
            {"status_code": 201, "body": {"id": "i", "name": "100%.txt", "file": {}, "size": 3}}])
        stream = io.BytesIO(b"abc")
        stream.seek(2)
        result = Upload(client).upload(stream, "Documents/Quarter #1", "100%.txt")
        self.assertEqual(result["id"], "i")
        self.assertEqual(stream.tell(), 2)
        create, put = client.api.graph_call.call_args_list
        self.assertIn("drives/chosen-drive/root:/Documents/Quarter%20%231/100%25.txt", create.kwargs["resource_url"])
        self.assertEqual(create.kwargs["json"]["item"]["@microsoft.graph.conflictBehavior"], "fail")
        self.assertTrue(put.kwargs["preauthenticated"])
        self.assertEqual(put.kwargs["headers"]["Content-Range"], "bytes 0-2/3")

    def test_sequential_aligned_chunks(self):
        client = self.client([{"uploadUrl": "https://upload.example/s"},
            {"status_code": 202, "body": {"nextExpectedRanges": ["327680-"]}},
            {"status_code": 201, "body": {"id": "i", "name": "file", "file": {}, "size": 327681}}])
        Upload(client, chunk_size=327680, conflict_behavior="rename").upload(io.BytesIO(b"x" * 327681), "", "file")
        puts = client.api.graph_call.call_args_list[1:]
        self.assertEqual([r.kwargs["headers"]["Content-Range"] for r in puts], ["bytes 0-327679/327681", "bytes 327680-327680/327681"])

    def test_lost_chunk_ack_queries_server_before_resume(self):
        client = self.client([{"uploadUrl": "https://upload.example/s"}, GraphError(code="transport_error"),
            {"nextExpectedRanges": ["327680-"]},
            {"status_code": 201, "body": {"id": "i", "name": "file", "file": {}, "size": 327681}}])
        Upload(client, chunk_size=327680).upload(io.BytesIO(b"x" * 327681), "", "file")
        calls = client.api.graph_call.call_args_list
        self.assertNotIn("method", calls[2].kwargs)
        self.assertEqual(calls[3].kwargs["headers"]["Content-Range"], "bytes 327680-327680/327681")

    def test_lost_final_ack_is_not_silently_restarted(self):
        client = self.client([{"uploadUrl": "https://upload.example/s"}, GraphError(code="transport_error"), GraphError(404)])
        with self.assertRaises(GraphError) as ctx:
            Upload(client).upload(io.BytesIO(b"x"), "", "file")
        self.assertEqual(ctx.exception.code, "upload_completion_unknown")
        self.assertEqual(client.api.graph_call.call_count, 3)

    def test_failures_are_visible(self):
        client = self.client([GraphError(409)])
        with self.assertRaises(GraphError):
            Upload(client).upload(io.BytesIO(b"a"), "", "f")
        client = self.client([{"uploadUrl": "https://upload.example/s"},
            {"status_code": 201, "body": {"id": "i", "name": "f", "file": {}, "size": 99}}])
        with self.assertRaises(GraphError):
            Upload(client).upload(io.BytesIO(b"a"), "", "f")

    def test_validation_precedes_http(self):
        client = self.client([])
        with self.assertRaises(ValueError):
            Upload(client).upload(io.BytesIO(b""), "", "empty")
        with self.assertRaises(ValueError):
            Upload(client).upload(io.BytesIO(b"a"), "", "../escape")
        with tempfile.TemporaryDirectory() as root, self.assertRaises(ValueError):
            Upload(client).upload(root, "")
        with self.assertRaises(ValueError):
            Upload(client, conflict_behavior="unsafe")
        client.api.graph_call.assert_not_called()


class DownloadTests(Offline):
    def client(self):
        return SimpleNamespace(api=Mock(), drive_id="selected-library")

    @staticmethod
    def stream(data=b"abc"):
        r = Mock(status_code=200, headers={})
        r.iter_content.return_value = iter([data])
        return r

    def test_all_folder_pages_and_selected_drive(self):
        client = self.client()
        api = self.api()
        client.api = api
        api.session.request.side_effect = [response({"id": "folder", "name": "Tree", "folder": {}}),
            response({"value": [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/selected-library/items/folder/children?page=2"}),
            response({"value": [{"id": "file", "name": "f.txt", "file": {}, "size": 3}]})]
        r = self.stream()
        original = api.graph_call
        def request(**kwargs):
            return r if kwargs.get("stream") else original(**kwargs)
        with patch.object(api, "graph_call", side_effect=request):
            with tempfile.TemporaryDirectory() as root:
                result = Download(client).folder(root, {"id": "folder", "folder": {}})
                self.assertEqual((result[0] / "f.txt").read_bytes(), b"abc")
        self.assertEqual(api.session.request.call_count, 3)
        r.close.assert_called_once()

    def test_download_failure_propagates_no_zip_or_data_deletion(self):
        client = self.client()
        r = self.stream()
        r.iter_content.side_effect = requests.ConnectionError("failed stream")
        client.api.graph_call.return_value = r
        with tempfile.TemporaryDirectory() as root:
            folder = Path(root) / "downloads"
            folder.mkdir()
            keep = folder / "keep.txt"
            keep.write_text("keep")
            with self.assertRaises(GraphError):
                Download(client, compress=True).file(folder, {"id": "i", "name": "f", "file": {}})
            self.assertEqual(keep.read_text(), "keep")
            self.assertFalse((Path(root) / "downloads.zip").exists())
            self.assertEqual(list(folder.iterdir()), [keep])
        r.close.assert_called_once()

    def test_compression_preserves_originals_and_ignores_unrelated_files(self):
        client = self.client()
        client.api.graph_call.return_value = self.stream()
        with tempfile.TemporaryDirectory() as root:
            folder = Path(root) / "downloads"
            folder.mkdir()
            (folder / "unrelated").write_text("keep")
            archive = Download(client, compress=True).file(folder, {"id": "i", "name": "f", "size": 3, "file": {}})
            import zipfile
            with zipfile.ZipFile(archive) as z:
                self.assertEqual(z.namelist(), ["f"])
            self.assertEqual((folder / "f").read_bytes(), b"abc")
            self.assertTrue((folder / "unrelated").exists())

    def test_existing_destination_is_never_overwritten(self):
        client = self.client()
        with tempfile.TemporaryDirectory() as root:
            target = Path(root) / "f"
            target.write_text("original")
            with self.assertRaises(FileExistsError):
                Download(client).file(root, {"id": "i", "name": "f", "file": {}})
            self.assertEqual(target.read_text(), "original")
        client.api.graph_call.assert_not_called()

    def test_unsafe_names_and_metadata_fail_closed(self):
        with tempfile.TemporaryDirectory() as root:
            for name in ("../outside", "C:evil", "CON.txt", "f.", "a/b", "a\\b"):
                with self.subTest(name=name), self.assertRaises(ValueError):
                    _child(root, name)
        with self.assertRaises(ValueError):
            Download(self.client()).file("unused", {"value": []})

    def test_size_mismatch_removes_partial(self):
        client = self.client()
        client.api.graph_call.return_value = self.stream(b"ab")
        with tempfile.TemporaryDirectory() as root:
            with self.assertRaises(GraphError):
                Download(client).file(root, {"id": "i", "name": "f", "file": {}, "size": 3})
            self.assertEqual(list(Path(root).iterdir()), [])

    def test_flat_zip_collision_rejected(self):
        with tempfile.TemporaryDirectory() as root:
            p = Path(root)
            with self.assertRaises(ValueError):
                _compress(p, [p / "a/f", p / "b/f"], False)


class UtilityAndMoveTests(Offline):
    def test_utilities_are_internal_packages(self):
        self.assertEqual(syslib.__name__, "m365.graph._syslib")
        self.assertEqual(dateplumb.__name__, "m365.graph._dateplumb")
        self.assertEqual(syslib.Buffer.convert_size("4MB"), 4194304)
        self.assertIsNotNone(dateplumb.now())
        self.assertEqual(dateplumb.parse("2026-11-17").year, 2026)
        from m365.graph._syslib._utils import VALID_FILE_EXTENSIONS
        self.assertIn(".txt", VALID_FILE_EXTENSIONS)

    def test_no_active_global_utility_imports(self):
        root = Path(__file__).resolve().parents[2] / "src" / "m365" / "graph"
        self.assertTrue(root.is_dir())
        for file in root.rglob("*.py"):
            if any(part.startswith('.') or part == '__pycache__' for part in file.relative_to(root).parts):
                continue  # Historical/deprecated files are not part of the distribution.
            for node in ast.walk(ast.parse(file.read_text(encoding="utf-8-sig"))):
                if isinstance(node, ast.Import):
                    self.assertFalse(any(n.name.split(".")[0] in ("syslib", "dateplumb") for n in node.names), str(file))
                if isinstance(node, ast.ImportFrom) and node.level == 0:
                    self.assertNotIn((node.module or "").split(".")[0], ("syslib", "dateplumb"), str(file))

    def test_move_returns_result_and_propagates_failure(self):
        client = SimpleNamespace(api=Mock(), drive_id="d", search_by=SimpleNamespace(filefoldername=Mock(return_value=[
            {"id": "f", "name": "file.txt", "file": {}}, {"id": "folder", "folder": {}}])))
        client.api.graph_call.return_value = {"id": "f", "name": "file.txt", "parentReference": {"id": "folder"}}
        self.assertEqual(Move(client).move("file.txt", "folder", None)["id"], "f")
        client.api.graph_call.side_effect = GraphError(403)
        with self.assertRaises(GraphError):
            Move(client).move("file.txt", "folder", None)


class SearchAndBoundaryTests(Offline):
    def client(self):
        return SimpleNamespace(api=Mock(), drive_id="selected", drive_web_url="https://tenant.example/teams/A/Library", search_mode="indexed")

    def test_file_search_escapes_odata_literals_and_stays_in_drive(self):
        client = self.client()
        client.api.iter_collection.return_value = [{"id": "f", "name": "O'Brien #1.txt", "file": {}},
                                                   {"id": "other", "name": "different.txt", "file": {}}]
        self.assertEqual(FileSearch(client).find("O'Brien #1.txt")[0]["id"], "f")
        self.assertEqual(client.api.iter_collection.call_args.args[0],
                         "drives/selected/root/search(q='O%27%27Brien%20%231.txt')")

    def test_library_relative_path_does_not_strip_documents(self):
        client = self.client()
        client.api.graph_call.return_value = {"status_code": 200, "body": {"id": "f", "file": {}, "name": "a #1.txt"}}
        FileSearch(client).find("Documents/a #1.txt")
        self.assertEqual(client.api.graph_call.call_args.kwargs["resource_url"], "drives/selected/root:/Documents/a%20%231.txt")

    def test_absolute_path_requires_selected_library(self):
        client = self.client()
        self.assertEqual(library_relative(client, "https://tenant.example/teams/A/Library/a%20b"), "a b")
        for url in ("https://tenant.example/teams/A/LibraryOther/a", "https://another.example/teams/A/Library/a"):
            with self.assertRaises(ValueError):
                library_relative(client, url)

    def test_search_404_is_empty_but_permissions_are_not_hidden(self):
        client = self.client()
        client.api.graph_call.side_effect = GraphError(404)
        self.assertEqual(FileSearch(client).find("missing/file"), [])
        client.api.graph_call.side_effect = GraphError(403)
        with self.assertRaises(GraphError):
            FileSearch(client).find("missing/file")

    def test_folder_children_use_collection_pagination(self):
        client = self.client()
        client.api.iter_collection.side_effect = [[{"id": "folder", "name": "Target", "folder": {}}], [{"id": "child"}]]
        result = FolderSearch(client).find("Target", include_children=True)
        self.assertEqual(result[0]["children"], [{"id": "child"}])
        self.assertEqual(client.api.iter_collection.call_args.args[0], "drives/selected/items/folder/children")

    def test_ambiguous_move_search_never_selects_arbitrary_result(self):
        client = self.client()
        client.api.iter_collection.side_effect = [[{"id": "1", "name": "f", "file": {}}, {"id": "2", "name": "f", "file": {}}],
            [{"id": "dir", "name": "folder", "folder": {}}]]
        with self.assertRaises(ValueError):
            FileFolder(client).find("f", "folder")

    def test_transfer_session_suppresses_implicit_netrc_auth(self):
        api = self.api()
        with patch("requests.sessions.get_netrc_auth", return_value=("username", "secret")) as netrc:
            prepared = api.transfer_session.prepare_request(requests.Request("PUT", "https://uploads.example/u"))
        self.assertNotIn("Authorization", prepared.headers)
        netrc.assert_not_called()

    def test_off_graph_nextlink_never_receives_token(self):
        api = self.api()
        api.session.request.return_value = response({"value": [], "@odata.nextLink": "https://evil.example/steal"})
        with self.assertRaises(ValueError):
            list(api.iter_collection("collection"))
        self.assertEqual(api.session.request.call_count, 1)

    def test_stream_download_uses_selected_drive(self):
        client = self.client()
        client.api.graph_call.return_value = DownloadTests.stream()
        with tempfile.TemporaryDirectory() as root:
            Download(client).file(root, {"id": "i", "name": "file", "file": {}})
        self.assertEqual(client.api.graph_call.call_args.kwargs["resource_url"], "drives/selected/items/i/content")

    def test_client_options_reject_invalid_timeout(self):
        with self.assertRaises(ValueError):
            Site(SimpleNamespace(access_token="unused"), site_id="s", request_timeout=float("inf"))

    def test_context_manager_closes_both_sessions(self):
        site, _ = self.site([response({"id": "s", "webUrl": "https://tenant.example/sites/A"}), response({"id": "d", "name": "Files"})], site_id="s")
        with patch.object(site.api, "close") as close:
            with site as entered:
                self.assertIs(entered, site)
            close.assert_called_once()

    def test_move_response_mismatch_is_not_success(self):
        client = SimpleNamespace(api=Mock(), drive_id="d", search_by=SimpleNamespace(filefoldername=Mock(return_value=[
            {"id": "f", "name": "file.txt", "file": {}}, {"id": "folder", "folder": {}}])))
        client.api.graph_call.return_value = {"id": "unexpected", "name": "file.txt"}
        with self.assertRaises(RuntimeError):
            Move(client).move("file.txt", "folder", None)


if __name__ == "__main__":
    unittest.main()
