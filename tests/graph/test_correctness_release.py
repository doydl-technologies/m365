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
"""Offline regressions for search, mailbox, path and date behavior."""
import io
import tempfile
import unittest
from datetime import datetime, date, time, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import requests

from m365.graph import Site, Mail, GraphError, _dateplumb as dateplumb
from m365.graph._graph import GraphAPI
from m365.graph.operations.doc_management import File, Folder
from m365.graph.operations.doc_dispatch import Move
from m365.graph._preprocessing.transforms import re_format_folderpath
from test_sharepoint import Offline, response


class SelectedSearchTests(Offline):
    def client(self, pages, **options):
        client = SimpleNamespace(drive_id="selected", graph_url="https://graph.microsoft.com/v1.0",
            search_mode="enumerate", search_root=None, search_max_requests=20,
            search_max_items=100, search_max_depth=5)
        for k, v in options.items(): setattr(client, k, v)
        client.api = self.api()
        client.api.session.request.side_effect = [response(p) for p in pages]
        return client

    def test_default_does_not_use_unsupported_index(self):
        c = self.client([{"value": [{"id": "f", "name": "O'Brien é.txt", "file": {}}]}])
        self.assertEqual(File(c).find("O'Brien é.txt")[0]["id"], "f")
        url = c.api.session.request.call_args.args[1]
        self.assertIn("/root/children?", url)
        self.assertNotIn("search(q=", url)

    def test_subfolder_scope_and_real_pagination(self):
        c = self.client([
            {"value": [{"id": "dir", "name": "Inner", "folder": {}}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/selected/items/scope/children?page=2"},
            {"value": [{"id": "1", "name": "REPORT.txt", "file": {}}]},
            {"value": [{"id": "2", "name": "report.txt", "file": {}}]}], search_root="scope")
        self.assertEqual({r["id"] for r in File(c).find("report.txt")}, {"1", "2"})
        self.assertIn("/items/scope/children", c.api.session.request.call_args_list[0].args[1])

    def test_contains_match_and_sorting(self):
        rows = [{"id": "1", "name": "Report-a", "file": {}, "createdDateTime": "2026-01-01T00:00:00Z"},
                {"id": "2", "name": "Report-b", "file": {}, "createdDateTime": "2026-02-01T00:00:00Z"}]
        c = self.client([{"value": rows}])
        self.assertEqual(File(c).find("REPORT", exact_match=False, sort_order="recent", num_files=1)[0]["id"], "2")

    def test_item_limit_never_returns_partial_match(self):
        c = self.client([{"value": [{"id": "1", "name": "f", "file": {}}, {"id": "2", "name": "f", "file": {}}]}], search_max_items=1)
        with self.assertRaises(GraphError) as e: File(c).find("f")
        self.assertEqual(e.exception.code, "search_item_limit")

    def test_request_budget(self):
        c = self.client([{"value": [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/selected/root/children?page=2"}], search_max_requests=1)
        with self.assertRaises(GraphError) as e: File(c).find("f")
        self.assertEqual(e.exception.code, "search_request_limit")

    def test_depth_budget(self):
        c = self.client([{"value": [{"id": "dir", "name": "a", "folder": {}}]}, {"value": [{"id": "deep", "name": "b", "folder": {}}]}], search_max_depth=1)
        with self.assertRaises(GraphError) as e: File(c).find("f")
        self.assertEqual(e.exception.code, "search_depth_limit")

    def test_shortcuts_are_not_traversed_or_controlling_matches(self):
        c = self.client([{"value": [{"id": "r", "name": "f", "file": {}, "folder": {}, "remoteItem": {"id": "outside"}}]}])
        self.assertEqual(File(c).find("f"), [])
        self.assertEqual(c.api.session.request.call_count, 1)

    def test_cycle_and_foreign_nextlink(self):
        for link in ("https://graph.microsoft.com/v1.0/drives/selected/root/children?$top=200", "https://evil.example/steal"):
            c = self.client([{"value": [], "@odata.nextLink": link}])
            with self.subTest(link=link), self.assertRaises((ValueError, GraphError)): File(c).find("f")
            self.assertEqual(c.api.session.request.call_count, 1)

    def test_foreign_drive_and_duplicate_items_fail_closed(self):
        for rows in ([{"id": "f", "name": "f", "file": {}, "parentReference": {"driveId": "other"}}],
                     [{"id": "f", "name": "f", "file": {}}, {"id": "f", "name": "f", "file": {}}]):
            with self.subTest(rows=rows), self.assertRaises(GraphError): File(self.client([{"value": rows}])).find("f")

    def test_permission_errors_are_not_hidden_as_no_matches(self):
        c = self.client([])
        c.api.session.request.side_effect = None
        c.api.session.request.return_value = response(status=403)
        with self.assertRaises(GraphError): File(c).find("f")

    def test_configuration_validation_before_network(self):
        for options in ({"search_mode": "auto-broaden"}, {"search_max_requests": 0}, {"search_max_items": True}):
            with self.subTest(options=options), self.assertRaises(ValueError):
                Site(SimpleNamespace(access_token="fake"), site_id="s", **options)


class PathAndMoveTests(Offline):
    def site_helpers(self):
        parent = SimpleNamespace(drive_name="Documents", drive_web_url="https://tenant.example/teams/A/Shared%20Documents", web_url="https://tenant.example/teams/A")
        parent.url = Site.URL(parent)
        parent.path = Site.PATH(parent)
        return parent

    def test_remove_only_one_library_prefix(self):
        s = self.site_helpers()
        self.assertEqual(s.path.drop_drivename("Documents/Reports/Documents/a.txt"), "Reports/Documents/a.txt")
        self.assertEqual(s.path.drop_drivename("Reports/Documents/a.txt"), "Reports/Documents/a.txt")
        self.assertEqual(s.path.drop_drivename("Shared%20Documents/Reports/a.txt"), "Reports/a.txt")

    def test_extension_strip_preserves_url_components(self):
        s = self.site_helpers()
        source = "https://tenant.example/teams/A/Shared%20Documents/a.txt?download=1#section"
        self.assertEqual(s.path.strip_domain_extension(source), source.replace("a.txt", "a"))
        self.assertEqual(s.path.strip_domain_extension("report.txt"), "report")
        self.assertEqual(s.path.strip_domain_extension(".hidden"), ".hidden")

    def test_relative_path_custom_site_and_library(self):
        s = self.site_helpers()
        self.assertEqual(s.path.relative_path(s.drive_web_url + "/Reports/a.txt", drop_drive=True), "Reports/a.txt")
        self.assertEqual(s.path.relative_path(s.drive_web_url + "/Reports/a.txt"), "Shared Documents/Reports/a.txt")
        with self.assertRaises(ValueError): s.path.relative_path("https://other.example/a.txt", drop_drive=True)

    def test_formatter_converts_and_does_not_mutate_input(self):
        s = self.site_helpers()
        original = [{"webUrl": s.drive_web_url + "/Reports"}]
        self.assertEqual(re_format_folderpath(original, s), [{"webUrl": "Reports"}])
        self.assertEqual(original[0]["webUrl"], s.drive_web_url + "/Reports")

    def test_formatter_propagates_invalid_url(self):
        with self.assertRaises(ValueError): re_format_folderpath({"webUrl": "https://evil.example/a"}, self.site_helpers())

    def move_client(self):
        return SimpleNamespace(drive_id="d", api=Mock(), search_by=SimpleNamespace(filefoldername=Mock(return_value=[
            {"id": "file", "name": "a.txt", "file": {}}, {"id": "folder", "folder": {}}])))

    def test_move_default_uses_resolved_basename(self):
        client = self.move_client()
        client.api.graph_call.return_value = {"id": "file", "name": "a.txt", "parentReference": {"id": "folder"}}
        self.assertEqual(Move(client).move("Reports/a.txt", "Archive/Target", None)["name"], "a.txt")
        self.assertEqual(client.api.graph_call.call_args.kwargs["json"]["name"], "a.txt")

    def test_move_rejects_path_in_destination_name(self):
        client = self.move_client()
        with self.assertRaises(ValueError): Move(client).move("Reports/a.txt", "Archive/Target", "../bad.txt")
        client.api.graph_call.assert_not_called()

    def test_contains_time_all_documented_types(self):
        for value, expected in [(time(1), True), (time(), False), (time(1, tzinfo=timezone.utc), True),
                                (datetime(2026, 1, 1), False), (datetime(2026, 1, 1, 1), True), (date(2026, 1, 1), False)]:
            with self.subTest(value=value): self.assertIs(dateplumb.contains_time(value), expected)


class MailSafetyTests(Offline):
    def mailbox(self):
        with patch.object(Mail, "_check_mailbox_access", return_value=None):
            mb = Mail(SimpleNamespace(access_token="synthetic-token"), mailbox="test@example.test")
        self.addCleanup(mb.api.close)
        mb.api.session.request = Mock()
        mb.api.transfer_session.request = Mock()
        return mb

    def test_graph_and_prefer_headers_and_payload(self):
        mb = self.mailbox()
        mb.api.session.request.return_value = response({"id": "m"})
        headers = {"Prefer": 'outlook.body-content-type="text"'}
        mb.api.graph_call(resource_url="users/test@example.test/messages", method="POST", json={"subject": "synthetic"}, headers=headers)
        sent = mb.api.session.request.call_args.kwargs
        self.assertEqual(sent["json"], {"subject": "synthetic"})
        self.assertIn('IdType="ImmutableId"', sent["headers"]["Prefer"])
        self.assertIn('outlook.body-content-type="text"', sent["headers"]["Prefer"])
        self.assertNotIn("IdType", headers["Prefer"])
        self.assertEqual(sent["headers"]["Authorization"], "Bearer synthetic-token")
        self.assertFalse(sent["allow_redirects"])

    def test_external_url_rejected_and_explicit_capability_has_no_bearer(self):
        mb = self.mailbox()
        with self.assertRaises(ValueError): mb.api.graph_call(resource_url="https://upload.example.test/u", method="PUT", data=b"abc")
        mb.api.transfer_session.request.return_value = response({"id": "a"})
        mb.api.graph_call(resource_url="https://upload.example.test/u", method="PUT", data=b"abc", headers={"Authorization": "bad", "Cookie": "bad"}, preauthenticated=True)
        sent = mb.api.transfer_session.request.call_args.kwargs
        self.assertFalse(any(k.lower() in ("authorization", "cookie") for k in sent["headers"]))
        mb.api.session.request.assert_not_called()

    def test_legacy_json_errors_are_redacted_streams_raise(self):
        mb = self.mailbox()
        mb.api.session.request.return_value = response({"error": "PRIVATE URL/TOKEN"}, status=403)
        result = mb.api.graph_call(resource_url="me/messages")
        self.assertIn("error", result)
        self.assertNotIn("PRIVATE", result["error"])
        with self.assertRaises(GraphError): mb.api.graph_call(resource_url="me/messages", stream=True)

    def test_batch_keeps_immutable_id_preferences(self):
        mb = self.mailbox()
        mb.api.session.request.return_value = response({"responses": [{"id": "1", "status": 200, "body": {}}]})
        mb.api.graph_call(batch_requests=[{"id": "1", "method": "GET", "url": "/me/messages/m"}])
        sent = mb.api.session.request.call_args.kwargs["json"]["requests"][0]
        self.assertIn('IdType="ImmutableId"', sent["headers"]["Prefer"])

    def attachment(self, mb, name, chunks=(b"abc",), size=None):
        metadata = {"id": "a", "name": name}
        if size is not None: metadata["size"] = size
        stream = Mock(status_code=200, headers={}); stream.iter_content.return_value = iter(chunks)
        mb.api.graph_call = Mock(side_effect=[{"status_code": 200, "body": metadata}, stream])
        return stream

    def test_attachment_success_staged_and_closed(self):
        mb = self.mailbox(); stream = self.attachment(mb, "a.txt", size=3)
        with tempfile.TemporaryDirectory() as root:
            result = mb.operations.AttachmentDownloader().download("m", root, attachment_id="a")
            self.assertEqual(Path(result).read_bytes(), b"abc")
            self.assertEqual(len(list(Path(root).iterdir())), 1)
        stream.close.assert_called_once()

    def test_attachment_unsafe_names_are_rejected_before_stream(self):
        for name in ("../escape", "a/b", "a\\b", "C:escape", "CON.txt", "a.", "a "):
            mb = self.mailbox(); self.attachment(mb, name)
            with tempfile.TemporaryDirectory() as root, self.subTest(name=name):
                with self.assertRaises(ValueError): mb.operations.AttachmentDownloader().download("m", root, attachment_id="a")
                self.assertEqual(list(Path(root).iterdir()), [])
                self.assertEqual(mb.api.graph_call.call_count, 1)

    def test_attachment_existing_file_never_overwritten(self):
        mb = self.mailbox(); self.attachment(mb, "a.txt")
        with tempfile.TemporaryDirectory() as root:
            existing = Path(root) / "a.txt"; existing.write_bytes(b"KEEP")
            with self.assertRaises(FileExistsError): mb.operations.AttachmentDownloader().download("m", root, attachment_id="a")
            self.assertEqual(existing.read_bytes(), b"KEEP")
            self.assertEqual(mb.api.graph_call.call_count, 1)

    def test_attachment_stream_failure_cleans_partial(self):
        mb = self.mailbox(); stream = self.attachment(mb, "a.txt")
        stream.iter_content.side_effect = requests.ConnectionError("PRIVATE URL")
        with tempfile.TemporaryDirectory() as root:
            with self.assertRaises(GraphError) as e: mb.operations.AttachmentDownloader().download("m", root, attachment_id="a")
            self.assertNotIn("PRIVATE", str(e.exception))
            self.assertEqual(list(Path(root).iterdir()), [])
        stream.close.assert_called_once()

    def test_attachment_size_mismatch_cleans_partial(self):
        mb = self.mailbox(); stream = self.attachment(mb, "a.txt", size=100)
        with tempfile.TemporaryDirectory() as root:
            with self.assertRaises(GraphError): mb.operations.AttachmentDownloader().download("m", root, attachment_id="a")
            self.assertEqual(list(Path(root).iterdir()), [])
        stream.close.assert_called_once()

    def test_large_attachment_explicitly_uses_preauthenticated_transport(self):
        mb = self.mailbox()
        size, chunk = 3 * 1024 * 1024, 327680
        acknowledgements = [{"status_code": 200, "body": {"nextExpectedRanges": [str(end)]}}
                            for end in range(chunk, size, chunk)]
        mb.api.graph_call = Mock(side_effect=[{"uploadUrl": "https://upload.example.test/u"}] +
                                acknowledgements + [{"status_code": 201, "body": {}}])
        mb.operations.AttachmentUploader().upload("m", b"x" * (3 * 1024 * 1024), file_name="a.bin")
        calls = mb.api.graph_call.call_args_list
        self.assertFalse(calls[0].kwargs.get("preauthenticated", False))
        self.assertTrue(all(c.kwargs["preauthenticated"] for c in calls[1:]))


if __name__ == "__main__": unittest.main()
