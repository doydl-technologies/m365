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
"""Offline regressions for malformed responses, transfer safety and resource bounds."""
import io
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import Mock, patch

from m365.graph import GraphError
from m365.graph._drive_walk import walk_drive
from m365.graph._graph import checked_drive_item
from m365.graph.operations.doc_management import File as SearchFile
from m365.graph.operations.doc_sync import Upload
from m365.graph.operations.doc_transfer import _child, _save_response
from m365.graph.operations.mail_transfer import AttachmentUpload
from test_sharepoint import Offline, response


class Collections330(Offline):
    def test_present_invalid_link_rejected_before_current_page_yields(self):
        for value in (None, "", " ", False, 0, [], {}, "https://graph.microsoft.com/v1.0/me\n"):
            with self.subTest(value=value):
                a = self.api()
                a.session.request.return_value = response({"value": [{"id": "a"}], "@odata.nextLink": value})
                with self.assertRaises(GraphError) as ctx:
                    next(a.iter_collection("drives/d/root/children"))
                self.assertEqual(ctx.exception.code, "invalid_collection_continuation")

    def test_absent_link_is_valid_end(self):
        a = self.api(); a.session.request.return_value = response({"value": [{"id": "a"}]})
        self.assertEqual(list(a.iter_collection("drives/d/root/children")), [{"id": "a"}])

    def test_foreign_link_is_not_requested(self):
        a = self.api(); a.session.request.return_value = response({"value": [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/other/root/children"})
        with self.assertRaises(GraphError) as ctx: list(a.iter_collection("drives/d/root/children"))
        self.assertEqual(ctx.exception.code, "collection_foreign_drive")
        self.assertEqual(a.session.request.call_count, 1)

    def test_same_drive_opaque_query_is_unchanged(self):
        a = self.api(); link = "https://graph.microsoft.com/v1.0/drives/d/items/root-id/children?$skiptoken=a%2Bb%3D&custom=X%252F"
        a.session.request.side_effect = [response({"value": [], "@odata.nextLink": link}), response({"value": [{"id": "a"}]})]
        self.assertEqual(len(list(a.iter_collection("drives/d/root/children"))), 1)
        self.assertEqual(a.session.request.call_args.args[1], link)

    def test_foreign_parent_rejected_before_yield(self):
        a = self.api(); a.session.request.return_value = response({"value": [{"id": "a", "parentReference": {"driveId": "other"}}]})
        with self.assertRaises(GraphError): next(a.iter_collection("drives/d/root/children"))

    def test_malformed_parent_and_drive_path_fail_structurally(self):
        for parent in (None, [], "bad", {"path": None}, {"driveId": None}, {"path": "/drives/other/root:/A"}):
            with self.subTest(parent=parent), self.assertRaises(GraphError):
                checked_drive_item({"id": "a", "parentReference": parent}, "d")

    def test_name_lookup_cannot_hide_invalid_continuation(self):
        a = self.api(); a.session.request.return_value = response({"value": [{"id": "a", "name": "x", "file": {}}], "@odata.nextLink": False})
        c = NS(api=a, drive_id="d")
        with self.assertRaises(GraphError): SearchFile(c).find("x")

    def test_name_walk_does_not_follow_foreign_drive(self):
        a = self.api(); a.session.request.return_value = response({"value": [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/other/root/children"})
        with self.assertRaises(GraphError): list(walk_drive(NS(api=a, drive_id="d")))
        self.assertEqual(a.session.request.call_count, 1)

    def test_direct_path_foreign_drive_and_shortcut_rejected(self):
        for extra in ({"parentReference": {"driveId": "other"}}, {"remoteItem": {"id": "external"}}):
            with self.subTest(extra=extra):
                a = self.api(); a.session.request.return_value = response(dict(id="a", name="x", file={}, **extra))
                with self.assertRaises(GraphError): SearchFile(NS(api=a, drive_id="d")).find("Folder/x")

    def test_indexed_lookup_rejects_explicit_foreign_parent(self):
        a = self.api(); a.session.request.return_value = response({"value": [{"id": "a", "name": "x", "file": {}, "parentReference": {"driveId": "other"}}]})
        with self.assertRaises(GraphError): SearchFile(NS(api=a, drive_id="d", search_mode="indexed")).find("x")


class Batch330(Offline):
    def batch(self, a): return a.graph_call(batch_requests=[{"id": "1", "method": "GET", "url": "/me"}])

    def test_invalid_statuses_raise_redacted_graph_errors(self):
        for value in (None, False, True, "200", 0, -1, 199, 600, [], {}):
            with self.subTest(value=value):
                a = self.api(); a.session.request.return_value = response({"responses": [{"id": "1", "status": value}]})
                with self.assertRaises(GraphError) as ctx: self.batch(a)
                self.assertEqual(ctx.exception.code, "invalid_batch_status")

    def test_missing_status_is_not_key_error(self):
        a = self.api(); a.session.request.return_value = response({"responses": [{"id": "1"}]})
        with self.assertRaises(GraphError) as ctx: self.batch(a)
        self.assertEqual(ctx.exception.code, "invalid_batch_status")

    def test_non_list_and_non_object_subresponses_fail_closed(self):
        for value in (None, {}, "bad", [None], [1]):
            with self.subTest(value=value):
                a = self.api(); a.session.request.return_value = response({"responses": value})
                with self.assertRaises(GraphError): self.batch(a)

    def test_redirect_is_not_success_or_implicitly_followed(self):
        a = self.api(); a.session.request.return_value = response({"responses": [{"id": "1", "status": 302, "headers": {"Location": "https://private.example/secret"}}]})
        with self.assertRaises(GraphError) as ctx: self.batch(a)
        self.assertEqual(ctx.exception.code, "batch_subrequest_redirect")
        self.assertNotIn("private", str(ctx.exception)); self.assertEqual(a.session.request.call_count, 1)

    def test_all_success_statuses_retained(self):
        for status in (200, 201, 202, 204, 206, 299):
            with self.subTest(status=status):
                a = self.api(); a.session.request.return_value = response({"responses": [{"id": "1", "status": status}]})
                self.assertEqual(self.batch(a)["1"]["status"], status)


class Upload330(Offline):
    def item(self, **change):
        item = {"id": "f", "name": "report.txt", "size": 3, "file": {}, "parentReference": {"driveId": "d", "path": "/drives/d/root:/Target"}}
        item.update(change); return item

    def complete(self, item, conflict="fail", folder="Target"):
        a = self.api(); a.session.request.return_value = response({"uploadUrl": "https://upload.example/session"})
        a.transfer_session.request.return_value = response(item, status=201)
        stream = io.BytesIO(b"abc"); stream.seek(2)
        try: return Upload(NS(api=a, drive_id="d"), conflict_behavior=conflict).upload(stream, folder, "report.txt")
        finally: self.assertEqual(stream.tell(), 2)

    def test_valid_completion_with_destination(self): self.assertEqual(self.complete(self.item())["id"], "f")

    def test_valid_completion_without_parent_does_not_invent_it(self):
        item = self.item(); item.pop("parentReference")
        self.assertNotIn("parentReference", self.complete(item))

    def test_wrong_drive_name_id_or_facet_rejected(self):
        for change in ({"id": {}}, {"id": ""}, {"id": True}, {"name": "wrong.txt"}, {"name": None},
                       {"file": None}, {"file": False}, {"folder": {}}, {"remoteItem": {}}, {"deleted": {}},
                       {"parentReference": {"driveId": "other"}}, {"parentReference": None}):
            with self.subTest(change=change), self.assertRaises(GraphError): self.complete(self.item(**change))

    def test_missing_required_completion_fields_rejected(self):
        for key in ("name", "file", "size"):
            item = self.item(); item.pop(key)
            with self.subTest(key=key), self.assertRaises(GraphError): self.complete(item)

    def test_boolean_size_is_not_one_byte(self):
        a = self.api(); a.session.request.return_value = response({"uploadUrl": "https://upload.example/s"})
        a.transfer_session.request.return_value = response(self.item(size=True), status=201)
        with self.assertRaises(GraphError): Upload(NS(api=a, drive_id="d")).upload(io.BytesIO(b"x"), "Target", "report.txt")

    def test_wrong_parent_path_rejected(self):
        for path in ("/drives/d/root:/Other", "/drives/other/root:/Target", "unproven-path"):
            with self.subTest(path=path), self.assertRaises(GraphError): self.complete(self.item(parentReference={"path": path}))

    def test_encoded_parent_path_and_case_match(self):
        item = self.item(name="REPORT.TXT", parentReference={"driveId": "d", "path": "/drives/d/root:/Folder%20%231"})
        self.assertEqual(self.complete(item, folder="Folder #1")["id"], "f")

    def test_root_parent_path_supported(self):
        self.assertEqual(self.complete(self.item(parentReference={"path": "/drives/d/root:"}), folder="")["id"], "f")

    def test_rename_policy_allows_server_rename(self):
        self.assertEqual(self.complete(self.item(name="report (1).txt"), conflict="rename")["name"], "report (1).txt")

    def test_rename_policy_still_rejects_unsafe_name_and_foreign_drive(self):
        for change in ({"name": "../escape"}, {"parentReference": {"driveId": "other"}}):
            with self.subTest(change=change), self.assertRaises(GraphError): self.complete(self.item(**change), conflict="rename")

    def test_incomplete_or_error_commit_never_replayed(self):
        a = self.api(); a.session.request.return_value = response({"uploadUrl": "https://upload.example/s"})
        a.transfer_session.request.return_value = response(self.item(name="wrong.txt"), status=201)
        with self.assertRaises(GraphError): Upload(NS(api=a, drive_id="d")).upload(io.BytesIO(b"abc"), "Target", "report.txt")
        self.assertEqual(a.transfer_session.request.call_count, 1); self.assertEqual(a.session.request.call_count, 1)

    def test_invalid_range_end_syntax_and_bounds(self):
        for ranges in (["327680-garbage"], ["327680-0"], ["0-655360"], ["327680"], ["1-"], ["0-", "3-"], [None], [False], ["0-" + "9" * 100]):
            with self.subTest(ranges=ranges), self.assertRaises(GraphError): Upload._offset({"nextExpectedRanges": ranges}, 655360)

    def test_valid_open_closed_and_multiple_missing_ranges(self):
        for ranges in (["327680-"], ["327680-655359"], ["327680-655359", "0-327679"]):
            with self.subTest(ranges=ranges):
                self.assertEqual(Upload._offset({"nextExpectedRanges": ranges}, 655360), min(int(s.split("-")[0]) for s in ranges))

    def test_malformed_ack_rejected_through_real_transport(self):
        a = self.api(); a.session.request.return_value = response({"uploadUrl": "https://upload.example/s"})
        a.transfer_session.request.return_value = response({"nextExpectedRanges": ["327680-bad"]}, status=202)
        with self.assertRaises(GraphError): Upload(NS(api=a, drive_id="d"), chunk_size=327680).upload(io.BytesIO(b"x" * 327681), "", "f")
        self.assertEqual(a.transfer_session.request.call_count, 1)


class LocalFiles330(Offline):
    def test_superscript_reserved_names_rejected_pre_io(self):
        for name in ("COM\u00b9", "com\u00b2.txt", "CoM\u00b3.bin", "LPT\u00b9", "lpt\u00b2.txt", "LPT\u00b3"):
            with self.subTest(name=name), self.assertRaises(ValueError): _child(Path(tempfile.gettempdir()), name)

    def test_unrelated_unicode_names_remain_valid(self):
        for name in ("report-\u00b9.txt", "company\u00e9.pdf", "COM10.txt", "lpt-report.txt"):
            with tempfile.TemporaryDirectory() as td: self.assertEqual(_child(td, name).name, name)

    def stream(self, chunks):
        r = Mock(status_code=200, headers={}); r.iter_content.return_value = iter(chunks); return r

    def test_oversize_aborts_at_first_excess_chunk(self):
        with tempfile.TemporaryDirectory() as td:
            seen = []
            def chunks():
                for chunk in (b"x" * 65536, b"y" * 65536): seen.append(len(chunk)); yield chunk
            r = Mock(status_code=200, headers={}); r.iter_content.return_value = chunks()
            with self.assertRaises(GraphError): _save_response(r, Path(td) / "f", 3)
            self.assertEqual(seen, [65536]); self.assertEqual(list(Path(td).iterdir()), [])
            r.close.assert_called_once()

    def test_invalid_expected_size_closes_before_consuming(self):
        for size in (True, -1, 1.0, "3", [], {}):
            with self.subTest(size=size), tempfile.TemporaryDirectory() as td:
                r = self.stream([b"x"])
                with self.assertRaises(GraphError): _save_response(r, Path(td) / "f", size)
                r.iter_content.assert_not_called(); r.close.assert_called_once()

    def test_zero_and_unspecified_size_still_supported(self):
        for size, content in ((0, b""), (None, b"abc")):
            with tempfile.TemporaryDirectory() as td:
                path = Path(td) / "f"; _save_response(self.stream([content]), path, size)
                self.assertEqual(path.read_bytes(), content)

    def test_close_failure_cannot_leak_staging_file(self):
        with tempfile.TemporaryDirectory() as td:
            r = self.stream([b"x"]); r.close.side_effect = RuntimeError("synthetic close failure")
            with self.assertRaises(RuntimeError): _save_response(r, Path(td) / "f", 3)
            self.assertEqual(list(Path(td).iterdir()), [])


class Outlook330(Offline):
    def test_error_body_on_final_201_rejected(self):
        a = self.api(); a.transfer_session.request.return_value = response({"error": {"message": "SYNTHETIC_PRIVATE"}}, status=201)
        with self.assertRaises(GraphError) as ctx: AttachmentUpload(NS(api=a))._upload_large("https://upload.example/s", b"abc")
        self.assertNotIn("PRIVATE", str(ctx.exception)); self.assertEqual(a.transfer_session.request.call_count, 1)

    def test_error_body_on_progress_200_rejected(self):
        a = self.api(); a.transfer_session.request.return_value = response({"error": {}, "nextExpectedRanges": ["327680"]})
        with self.assertRaises(GraphError): AttachmentUpload(NS(api=a))._upload_large("https://upload.example/s", b"x" * 327681)

    def test_empty_201_remains_valid(self):
        a = self.api(); a.transfer_session.request.return_value = response({}, status=201)
        self.assertTrue(AttachmentUpload(NS(api=a))._upload_large("https://upload.example/s", b"abc")["completed"])


if __name__ == "__main__": unittest.main()
