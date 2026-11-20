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
"""Regression tests for transfer status, scoped metadata, and mailbox cursors."""

import gzip
import hashlib
import io
from pathlib import Path
import tempfile
from types import SimpleNamespace
from unittest.mock import Mock

import requests

from m365.graph import GraphError
from m365.graph.operations.doc_management import File as FileSearch
from m365.graph.operations.doc_sync import Upload
from m365.graph.operations.doc_transfer import Download, _save_response
from m365.graph.operations.mail_management import Delta
from test_sharepoint import Offline, response


BASE = "https://graph.microsoft.com/v1.0"
CHUNK = 327680


def file_item(name="contract.txt", size=3, **extra):
    return {"id": "file-id", "name": name, "file": {}, "size": size, **extra}


def binary(status=200, data=b"abc", headers=None):
    reply = response({}, status=status, content=data, headers=headers)
    reply._content_consumed = True
    return reply


class FullFileDownloadTests(Offline):
    def client(self, *replies):
        api = self.api()
        api.session.request.side_effect = replies
        return SimpleNamespace(api=api, drive_id="drive-id")

    def test_partial_and_no_content_are_not_full_files(self):
        for status, size in ((206, None), (206, 3), (204, 0), (202, 3)):
            with self.subTest(status=status, size=size), tempfile.TemporaryDirectory() as folder:
                reply = binary(status, headers={"Content-Range": "bytes 0-2/9"} if status == 206 else None)
                client = self.client(reply)
                with self.assertRaises(GraphError) as caught:
                    Download(client).file(folder, file_item(size=size))
                self.assertEqual((caught.exception.status_code, caught.exception.code),
                                 (status, "unexpected_download_status"))
                reply.close.assert_called_once()
                self.assertEqual(list(Path(folder).iterdir()), [])

    def test_redirect_to_partial_response_fails_and_closes_both_responses(self):
        hop = binary(302, headers={"Location": "https://download.example.invalid/file"})
        partial = binary(206, headers={"Content-Range": "bytes 0-2/9"})
        client = self.client(hop)
        client.api.transfer_session.request.return_value = partial
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(GraphError) as caught:
                Download(client).file(folder, file_item(size=None))
            self.assertEqual(caught.exception.status_code, 206)
            self.assertEqual(list(Path(folder).iterdir()), [])
        hop.close.assert_called_once()
        partial.close.assert_called_once()

    def test_raw_stream_can_deliberately_request_a_range(self):
        hop = binary(302, headers={"Location": "https://download.example.invalid/file"})
        partial = binary(206, headers={"Content-Range": "bytes 0-2/9"})
        client = self.client(hop)
        client.api.transfer_session.request.return_value = partial
        result = client.api.graph_call(resource_url="drives/drive-id/items/file-id/content",
                                       stream=True, headers={"Range": "bytes=0-2"})
        try:
            self.assertEqual(result.status_code, 206)
            self.assertEqual(result.content, b"abc")
            sent = client.api.transfer_session.request.call_args.kwargs["headers"]
            self.assertEqual(sent["Range"], "bytes=0-2")
            self.assertNotIn("Authorization", sent)
        finally:
            result.close()

    def test_unexpected_content_range_on_200_is_rejected(self):
        client = self.client(binary(headers={"content-range": "bytes 0-2/9"}))
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(GraphError) as caught:
                Download(client).file(folder, file_item())
            self.assertEqual(caught.exception.code, "unexpected_content_range")
            self.assertEqual(list(Path(folder).iterdir()), [])

    def test_length_header_prevents_publishing_without_item_size(self):
        for length, code in (("9", "download_size_mismatch"), ("2", "download_size_mismatch"),
                             ("3, 9", "invalid_content_length"), ("+3", "invalid_content_length"),
                             ("3, ", "invalid_content_length")):
            with self.subTest(length=length), tempfile.TemporaryDirectory() as folder:
                client = self.client(binary(headers={"content-length": length}))
                with self.assertRaises(GraphError) as caught:
                    Download(client).file(folder, file_item(size=None))
                self.assertEqual(caught.exception.code, code)
                self.assertEqual(list(Path(folder).iterdir()), [])

    def test_wire_and_item_length_conflict_stops_before_read(self):
        reply = Mock(status_code=200, headers={"Content-Length": "99"})
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(GraphError):
                _save_response(reply, Path(folder) / "file.txt", 3)
            reply.iter_content.assert_not_called()
            reply.close.assert_called_once()
            self.assertEqual(list(Path(folder).iterdir()), [])

    def test_identity_and_compressed_lengths_preserve_verified_bytes(self):
        content = b"fixture bytes\n" * 80
        for headers in ({"Content-Length": str(len(content))},
                        {"Content-Length": f"{len(content)}, {len(content)}"},
                        {"Content-Length": str(len(gzip.compress(content))), "Content-Encoding": "gzip"},
                        {"Transfer-Encoding": "chunked"}, {}):
            with self.subTest(headers=headers), tempfile.TemporaryDirectory() as folder:
                client = self.client(binary(data=content, headers=headers))
                path, = Download(client).file(folder, file_item(size=len(content)))
                self.assertEqual(hashlib.sha256(path.read_bytes()).digest(), hashlib.sha256(content).digest())

    def test_conflicting_length_and_wire_framing_is_rejected(self):
        client = self.client(binary(headers={"Transfer-Encoding": "chunked", "Content-Length": "3"}))
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(GraphError) as caught:
                Download(client).file(folder, file_item())
            self.assertEqual(caught.exception.code, "invalid_content_length")
            self.assertEqual(list(Path(folder).iterdir()), [])

    def test_stream_failure_preserves_neighbor_and_leaves_no_staging_file(self):
        def chunks():
            yield b"abc"
            raise requests.exceptions.ChunkedEncodingError("synthetic broken stream")

        reply = Mock(status_code=200, headers={})
        reply.iter_content.return_value = chunks()
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            keep = root / "existing.txt"
            keep.write_bytes(b"keep")
            with self.assertRaises(GraphError) as caught:
                _save_response(reply, root / "incoming.txt", 9)
            self.assertEqual(caught.exception.code, "download_stream_failed")
            reply.close.assert_called_once()
            self.assertEqual(list(root.iterdir()), [keep])
            self.assertEqual(keep.read_bytes(), b"keep")


class UploadCommitTests(Offline):
    def upload(self, *replies, size=3):
        api = self.api()
        api.session.request.return_value = response({"uploadUrl": "https://upload.example.invalid/fixture"})
        api.transfer_session.request.side_effect = replies
        data = io.BytesIO(b"x" * size)
        return api, lambda: Upload(SimpleNamespace(api=api, drive_id="drive-id"), chunk_size=CHUNK).upload(
            data, "scratch", "contract.txt")

    def test_202_and_206_cannot_commit_even_with_matching_item_metadata(self):
        for status in (202, 204, 206):
            with self.subTest(status=status):
                api, send = self.upload(response(file_item(), status=status))
                with self.assertRaises(GraphError) as caught:
                    send()
                self.assertEqual((caught.exception.status_code, caught.exception.code),
                                 (status, "upload_not_committed"))
                self.assertEqual(api.session.request.call_count, 1)
                self.assertEqual(api.transfer_session.request.call_count, 1)

    def test_progress_requires_202_and_cannot_claim_an_item(self):
        for status, body in (
            (201, {"nextExpectedRanges": [f"{CHUNK}-"]}),
            (202, {"id": "item", "nextExpectedRanges": [f"{CHUNK}-"]}),
            (202, {"file": {}, "nextExpectedRanges": [f"{CHUNK}-"]}),
            (202, {"nextExpectedRanges": ["0-"]}),
        ):
            with self.subTest(status=status, body=body):
                api, send = self.upload(response(body, status=status), size=CHUNK + 1)
                with self.assertRaises(GraphError) as caught:
                    send()
                self.assertEqual(caught.exception.code, "unexpected_upload_acknowledgement")
                self.assertEqual(api.transfer_session.request.call_count, 1)

    def test_completion_with_progress_or_error_body_fails_closed(self):
        for body in ({"nextExpectedRanges": ["0-"]}, {"error": {}},
                     file_item(nextExpectedRanges=["0-"])):
            with self.subTest(body=body):
                api, send = self.upload(response(body, status=201))
                with self.assertRaises(GraphError):
                    send()
                self.assertEqual(api.transfer_session.request.call_count, 1)

    def test_valid_create_and_replace_200_201_return_the_same_item_shape(self):
        for status in (200, 201):
            with self.subTest(status=status):
                api, send = self.upload(
                    response({"nextExpectedRanges": [f"{CHUNK}-"]}, status=202),
                    response(file_item(size=CHUNK + 1), status=status), size=CHUNK + 1)
                self.assertEqual(send(), file_item(size=CHUNK + 1))
                ranges = [c.kwargs["headers"]["Content-Range"]
                          for c in api.transfer_session.request.call_args_list]
                self.assertEqual(ranges, [f"bytes 0-{CHUNK - 1}/{CHUNK + 1}",
                                          f"bytes {CHUNK}-{CHUNK}/{CHUNK + 1}"])


class DriveMetadataTests(Offline):
    def client(self, *pages, search_mode="enumerate"):
        api = self.api()
        api.session.request.side_effect = [response(page) for page in pages]
        return SimpleNamespace(api=api, drive_id="drive-id", search_mode=search_mode)

    def test_folder_metadata_must_match_request_before_creating_any_path(self):
        folder = {"id": "wanted", "name": "Folder", "folder": {}}
        for extra in ({"id": "wrong"}, {"parentReference": {"driveId": "wrong"}},
                      {"folder": None}, {"file": {}}, {"remoteItem": {}}, {"deleted": {}}):
            with self.subTest(extra=extra), tempfile.TemporaryDirectory() as root:
                client = self.client(dict(folder, **extra))
                with self.assertRaises((GraphError, ValueError)):
                    Download(client).folder(root, folder)
                self.assertEqual(client.api.session.request.call_count, 1)
                self.assertEqual(list(Path(root).iterdir()), [])

    def test_direct_lookup_rejects_contradictory_name_or_parent_path(self):
        for row in (file_item("wrong.txt"), file_item(parentReference={
                "driveId": "drive-id", "path": "/drives/drive-id/root:/Different"})):
            with self.subTest(row=row):
                client = self.client(row)
                with self.assertRaises(GraphError):
                    FileSearch(client).find("Folder/contract.txt")

    def test_direct_lookup_accepts_encoded_parent_and_case_insensitive_unicode_name(self):
        row = file_item("CAFÉ + REPORT.TXT", parentReference={
            "driveId": "drive-id", "path": "/drives/drive-id/root:/Project%20%231"})
        client = self.client(row)
        self.assertEqual(FileSearch(client).find("Project #1/café + report.txt"), [row])

    def test_indexed_lookup_rejects_invalid_name_in_both_modes(self):
        for name in (None, True, "", "  ", "folder/report"):
            for exact in (True, False):
                with self.subTest(name=name, exact=exact):
                    client = self.client({"value": [file_item(name)]}, search_mode="indexed")
                    with self.assertRaises(GraphError) as caught:
                        FileSearch(client).find("report", exact_match=exact)
                    self.assertEqual(caught.exception.code, "invalid_search_item")

    def test_continuation_cannot_change_concrete_parent_or_collection_kind(self):
        for tail in ("items/different/children", "root/children", "root/search(q='file')"):
            with self.subTest(tail=tail):
                client = self.client({"value": [file_item()],
                                      "@odata.nextLink": BASE + "/drives/drive-id/" + tail + "?page=2"})
                with self.assertRaises(GraphError) as caught:
                    next(client.api.iter_collection("drives/drive-id/items/folder/children"))
                self.assertEqual(caught.exception.code, "collection_foreign_parent")
                self.assertEqual(client.api.session.request.call_count, 1)

    def test_child_cannot_claim_other_parent_and_no_partial_page_yields(self):
        client = self.client({"value": [file_item(), file_item("other.txt", parentReference={"id": "wrong"})]})
        with self.assertRaises(GraphError) as caught:
            next(client.api.iter_collection("drives/drive-id/items/folder/children"))
        self.assertEqual(caught.exception.code, "collection_foreign_parent")

    def test_root_alias_canonicalizes_once_without_rewriting_opaque_token(self):
        alias = BASE + "/drives/drive-id/items/canonical/children?$skiptoken=a%2Bb%2F%3D"
        client = self.client(
            {"value": [file_item(parentReference={"id": "canonical"})], "@odata.nextLink": alias},
            {"value": []})
        self.assertEqual(len(list(client.api.iter_collection("drives/drive-id/root/children"))), 1)
        self.assertEqual(client.api.session.request.call_args.args[1], alias)

    def test_root_alias_cannot_canonicalize_to_contradictory_id(self):
        alias = BASE + "/drives/drive-id/items/wrong/children?$skiptoken=opaque"
        client = self.client({"value": [file_item(parentReference={"id": "known"})], "@odata.nextLink": alias})
        with self.assertRaises(GraphError):
            next(client.api.iter_collection("drives/drive-id/root/children"))
        self.assertEqual(client.api.session.request.call_count, 1)

    def test_name_traversal_rejects_continuation_to_different_parent(self):
        alias = BASE + "/drives/drive-id/items/other/children?$skiptoken=opaque"
        client = self.client({"value": [file_item()], "@odata.nextLink": alias})
        client.search_root = "wanted"
        with self.assertRaises(GraphError) as caught:
            FileSearch(client).find("contract.txt")
        self.assertEqual(caught.exception.code, "search_foreign_parent")
        self.assertEqual(client.api.session.request.call_count, 1)


class DownloadWrapperTests(Offline):
    def test_mixed_input_rejects_error_or_malformed_entry_before_any_fetch(self):
        for wrapper in (
            {"value": [file_item(), {"error": {"message": "SYNTHETIC_SECRET"}}]},
            {"responses": [{"status": 200, "body": file_item()}, {"status": 403}]},
            {"responses": [{"status": 200, "body": file_item()}, {"status": "200", "body": {}}]},
            {"value": [file_item(), {"value": "wrong"}]},
        ):
            with self.subTest(wrapper=wrapper), tempfile.TemporaryDirectory() as folder:
                api = Mock()
                client = SimpleNamespace(api=api, drive_id="drive-id")
                root = Path(folder) / "not-created"
                with self.assertRaises(GraphError) as caught:
                    Download(client).file(root, wrapper)
                self.assertNotIn("SYNTHETIC_SECRET", str(caught.exception))
                api.graph_call.assert_not_called()
                self.assertFalse(root.exists())

    def test_valid_batch_wrapper_and_mixed_facets_remain_supported(self):
        c = SimpleNamespace(api=Mock(), drive_id="drive-id")
        c.api.graph_call.return_value = binary()
        wrapper = {"responses": [{"status": 200, "body": {
            "value": [{"id": "folder", "name": "Folder", "folder": {}}, file_item()]}}]}
        with tempfile.TemporaryDirectory() as folder:
            item, = Download(c).file(folder, wrapper)
            self.assertEqual(item.read_bytes(), b"abc")
            self.assertEqual(c.api.graph_call.call_count, 1)


class DeltaCursorTests(Offline):
    def test_message_and_folder_delta_require_exactly_one_valid_cursor(self):
        for method in ("messages", "folders"):
            for keys, code in (({}, "invalid_delta_cursor"),
                              ({"@odata.nextLink": "next", "@odata.deltaLink": "delta"}, "invalid_delta_cursor"),
                              ({"@odata.deltaLink": "  "}, "invalid_mail_continuation"),
                              ({"@odata.nextLink": "opaque\n"}, "invalid_mail_continuation")):
                with self.subTest(method=method, keys=keys):
                    request = Mock(return_value={"status_code": 200, "body": {"value": [{"id": "a"}], **keys}})
                    client = SimpleNamespace(api=SimpleNamespace(graph_call=request), base_resource="me")
                    with self.assertRaises(GraphError) as caught:
                        getattr(Delta(client), method)()
                    self.assertEqual(caught.exception.code, code)
                    self.assertEqual(request.call_count, 1)

    def test_empty_delta_page_retains_the_opaque_sync_cursor(self):
        cursor = BASE + "/me/mailFolders/delta?$deltatoken=opaque%2B%2F%3D"
        request = Mock(return_value={"status_code": 200, "body": {"value": [], "@odata.deltaLink": cursor}})
        client = SimpleNamespace(api=SimpleNamespace(graph_call=request), base_resource="me")
        self.assertEqual(Delta(client).folders(),
                         {"value": [], "next_link": None, "delta_link": cursor})
