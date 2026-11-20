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
"""Offline SharePoint pagination, path and archive regressions."""
import io
import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock
from urllib.parse import parse_qs, urlsplit, quote

from m365.graph import GraphError
from m365.graph._local_paths import native_path, public_path, windows_path
from m365.graph._site_resolution import library_relative
from m365.graph.operations.doc_management import File
from m365.graph.operations.doc_sync import Upload
from m365.graph.operations.doc_transfer import Download, _compress, _save_response
from test_sharepoint import Offline, response


GUID = "f0f54d45-7187-490e-9e71-d573a180ee66"


class Pagination310(Offline):
    def scan(self, pages, query):
        api = self.api()
        api.session.request.side_effect = [response(p) for p in pages]
        result = list(api.iter_collection("drives/selected/items/folder/children" + query))
        return api, result

    def test_drive_children_adds_unique_name_tie_break(self):
        api, rows = self.scan([{"value": [{"id": "a"}]}],
                             "?$orderby=lastModifiedDateTime%20desc&$top=2")
        query = parse_qs(urlsplit(api.session.request.call_args.args[1]).query)
        self.assertEqual(query["$orderby"], ["lastModifiedDateTime desc,name asc"])
        self.assertEqual(query["$top"], ["2"])
        self.assertEqual(rows, [{"id": "a"}])

    def test_name_order_is_not_appended_twice(self):
        api, _ = self.scan([{"value": []}], "?$orderby=size%20desc,name%20desc")
        self.assertEqual(parse_qs(urlsplit(api.session.request.call_args.args[1]).query)["$orderby"],
                         ["size desc,name desc"])

    def test_root_path_and_site_drive_children_get_tie_break(self):
        for endpoint in ("drives/d/root/children", "drives/d/root:/A/B:/children",
                         "sites/s/drive/root/children"):
            with self.subTest(endpoint=endpoint):
                api = self.api()
                api.session.request.return_value = response({"value": []})
                list(api.iter_collection(endpoint + "?$orderby=size%20desc"))
                query = parse_qs(urlsplit(api.session.request.call_args.args[1]).query)
                self.assertEqual(query["$orderby"], ["size desc,name asc"])

    def test_nextlink_is_not_rewritten_or_reconstructed(self):
        nextlink = "https://graph.microsoft.com/v1.0/drives/selected/items/folder/children?$skiptoken=opaque%2B%2F%3D&$top=2"
        api, rows = self.scan([{"value": [{"id": "a"}], "@odata.nextLink": nextlink},
                              {"value": [{"id": "b"}]}], "?$orderby=size%20desc")
        self.assertEqual(api.session.request.call_args_list[1].args[1], nextlink)
        self.assertEqual([r["id"] for r in rows], ["a", "b"])

    def test_explicit_continuation_is_already_opaque(self):
        api = self.api()
        url = "https://graph.microsoft.com/v1.0/drives/d/root/children?$skiptoken=a%2Bb&$orderby=size%20desc"
        api.session.request.return_value = response({"value": []})
        list(api.iter_collection(url))
        self.assertEqual(api.session.request.call_args.args[1], url)

    def test_duplicate_id_across_pages_raises_not_silent_dedup(self):
        with self.assertRaises(GraphError) as caught:
            self.scan([{"value": [{"id": "a"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/selected/items/folder/children?page=2"},
                       {"value": [{"id": "a"}, {"id": "b"}]}], "")
        self.assertEqual(caught.exception.code, "collection_duplicate_item")

    def test_duplicate_id_within_page_raises(self):
        with self.assertRaises(GraphError) as caught:
            self.scan([{"value": [{"id": "a"}, {"id": "a"}]}], "")
        self.assertEqual(caught.exception.code, "collection_duplicate_item")

    def test_select_retains_fields_and_adds_id_for_completeness_guard(self):
        api, _ = self.scan([{"value": [{"id": "a", "name": "f"}]}], "?$select=name&$top=1")
        query = parse_qs(urlsplit(api.session.request.call_args.args[1]).query)
        self.assertEqual(query["$select"], ["name,id"])

    def test_missing_identity_is_explicit_failure(self):
        with self.assertRaises(GraphError) as caught:
            self.scan([{"value": [{"name": "f"}]}], "")
        self.assertEqual(caught.exception.code, "invalid_collection_item")

    def test_delta_can_repeat_ids_without_child_scan_semantics(self):
        api = self.api()
        api.session.request.return_value = response({"value": [{"id": "a"}, {"id": "a", "deleted": {}}]})
        self.assertEqual(len(list(api.iter_collection("drives/d/root/delta"))), 2)

    def test_non_drive_collection_is_not_given_drive_order(self):
        api = self.api()
        api.session.request.return_value = response({"value": [1, 2]})
        list(api.iter_collection("users/u/messages?$orderby=receivedDateTime%20desc"))
        self.assertNotIn("name", api.session.request.call_args.args[1])

    def test_unsupported_query_http_error_is_not_hidden_or_retried_broadly(self):
        api = self.api()
        api.session.request.return_value = response(status=400)
        with self.assertRaises(GraphError):
            list(api.iter_collection("drives/d/root/children?$orderby=createdDateTime%20desc"))
        self.assertEqual(api.session.request.call_count, 1)

    def test_duplicate_orderby_is_rejected_before_http(self):
        api = self.api()
        with self.assertRaises(ValueError):
            list(api.iter_collection("drives/d/root/children?$orderby=name&$orderby=size"))
        api.session.request.assert_not_called()


class DocumentURLs310(Offline):
    def client(self, pages):
        parent = SimpleNamespace(graph_url="https://graph.microsoft.com/v1.0",
            request_timeout=(1, 2), max_retries=0, get_token=lambda: "fake")
        api = self.api()
        api.session.request.side_effect = [response(p) for p in pages]
        return SimpleNamespace(api=api, drive_id="selected",
            web_url="https://tenant.example/teams/Operations",
            drive_web_url="https://tenant.example/teams/Operations/Contracts",
            search_max_requests=10, search_max_items=100, search_max_depth=4)

    def row(self, id_="doc", name="O'Brien #100%.csv", guid=GUID):
        return {"id": id_, "name": name, "file": {}, "parentReference": {"driveId": "selected"},
                "sharepointIds": {"listItemUniqueId": guid}}

    def url(self, guid=GUID, name="O'Brien #100%.csv"):
        return "https://tenant.example/teams/Operations/_layouts/15/Doc.aspx?sourcedoc=%7B" + guid + "%7D&file=" + quote(name, safe="") + "&action=default&mobileredirect=true"

    def test_browser_document_guid_resolves_nested_unencoded_path(self):
        c = self.client([{"value": [{"id": "dir", "name": "A #100%", "folder": {}}]},
                         {"value": [self.row()]}])
        self.assertEqual(library_relative(c, self.url()), "A #100%/O'Brien #100%.csv")
        self.assertIn("sharepointIds", c.api.session.request.call_args.args[1])

    def test_high_level_filename_url_lookup_returns_selected_item(self):
        c = self.client([{"value": [self.row()]}, self.row()])
        self.assertEqual(File(c).find(self.url())[0]["id"], "doc")
        self.assertIn("/root:/O%27Brien%20%23100%25.csv", c.api.session.request.call_args.args[1])

    def test_clean_library_url_still_works_without_traversal(self):
        c = self.client([])
        self.assertEqual(library_relative(c, c.drive_web_url + "/x%20%23%25.txt"), "x #%.txt")
        c.api.session.request.assert_not_called()

    def test_sourcedoc_identity_not_filename_guess(self):
        c = self.client([{"value": [self.row(guid="de1743d3-34c7-4c67-b7b6-17fb3bdb56cd")]}])
        with self.assertRaises(ValueError):
            library_relative(c, self.url())
        self.assertEqual(c.api.session.request.call_count, 1)

    def test_duplicate_guid_in_distinct_items_is_ambiguous(self):
        c = self.client([{"value": [self.row("one"), self.row("two")]}])
        with self.assertRaises(ValueError):
            library_relative(c, self.url())

    def test_display_name_contradiction_rejected(self):
        c = self.client([{"value": [self.row(name="different.csv")]}])
        with self.assertRaises(ValueError):
            library_relative(c, self.url())

    def test_unknown_browser_identity_does_not_fallback_to_index_or_other_drive(self):
        c = self.client([{"value": []}])
        with self.assertRaises(ValueError):
            library_relative(c, self.url())
        self.assertEqual(c.api.session.request.call_count, 1)
        self.assertNotIn("search(", c.api.session.request.call_args.args[1])

    def test_cross_site_host_http_and_malformed_query_rejected_without_http(self):
        variants = [self.url().replace("tenant.example", "other.example"),
                    self.url().replace("/teams/Operations/", "/teams/Other/"),
                    self.url().replace("https://", "http://"),
                    self.url(guid="not-a-guid"), self.url() + "&sourcedoc=" + GUID,
                    self.url() + "&redirect=https://evil.example",
                    self.url().replace("tenant.example", "user@tenant.example")]
        for url in variants:
            with self.subTest(url=url):
                c = self.client([])
                with self.assertRaises(ValueError):
                    library_relative(c, url)
                c.api.session.request.assert_not_called()

    def test_browser_lookup_respects_budgets_and_no_partial_unique_success(self):
        c = self.client([{"value": [self.row(), {"id": "dir", "name": "Inner", "folder": {}}]}])
        c.search_max_requests = 1
        with self.assertRaises(GraphError) as caught:
            library_relative(c, self.url())
        self.assertEqual(caught.exception.code, "search_request_limit")

    def test_shortcuts_and_foreign_drive_cannot_supply_document_identity(self):
        row = self.row()
        row["remoteItem"] = {"id": "outside"}
        c = self.client([{"value": [row]}])
        with self.assertRaises(ValueError):
            library_relative(c, self.url())
        row = self.row()
        row["parentReference"]["driveId"] = "outside"
        c = self.client([{"value": [row]}])
        with self.assertRaises(GraphError):
            library_relative(c, self.url())


class LocalPathsAndZIP310(Offline):
    def test_drive_unc_and_existing_extended_path(self):
        for source, wanted in [(r"C:\files\report.txt", r"\\?\C:\files\report.txt"),
                               (r"\\host\share\a.txt", r"\\?\UNC\host\share\a.txt"),
                               (r"\\?\C:\files\a.txt", r"\\?\C:\files\a.txt")]:
            with self.subTest(source=source):
                self.assertEqual(windows_path(source), wanted)
        for unsafe in (r"\\.\PhysicalDrive0", r"\\?\GLOBALROOT\Device\Disk"):
            with self.assertRaises(ValueError):
                windows_path(unsafe)

    @unittest.skipUnless(os.name == "nt", "Windows UTF-16 path length")
    def test_astral_unicode_paths_keep_needed_extended_prefix(self):
        normal = "C:\\data\\" + "\U0001f600" * 70 + "\\" + "\U0001f600" * 70 + ".txt"
        self.assertLess(len(normal), 248)
        self.assertGreater(len(normal.encode("utf-16-le")) // 2, 248)
        self.assertTrue(str(public_path(native_path(normal))).startswith("\\\\?\\"))

    @unittest.skipUnless(os.name == "nt", "Real Windows long-path filesystem case")
    def test_long_download_and_source_upload_work_without_host_setting_changes(self):
        temporary = tempfile.TemporaryDirectory()
        temporary.name = str(native_path(temporary.name))
        with temporary as root:
            deep = native_path(Path(root) / ("a" * 100) / ("b" * 100))
            deep.mkdir(parents=True)
            name = "long-" + "c" * 100 + ".txt"
            client = SimpleNamespace(api=Mock(), drive_id="d")
            stream = Mock(status_code=200, headers={}); stream.iter_content.return_value = iter([b"abc"])
            client.api.graph_call.return_value = stream
            path = Download(client).file(deep, {"id": "f", "name": name, "file": {}, "size": 3})[0]
            self.assertTrue(str(path).startswith("\\\\?\\"))
            self.assertEqual(path.read_bytes(), b"abc")
            client.api.graph_call.side_effect = [{"uploadUrl": "https://upload.example/u"},
                {"status_code": 201, "body": {"id": "i", "name": "copy.txt", "file": {}, "size": 3}}]
            Upload(client).upload(path, "", "copy.txt")
            self.assertEqual(client.api.graph_call.call_args.kwargs["data"], b"abc")
            archive = _compress(deep, [path], True)
            with zipfile.ZipFile(archive) as z:
                self.assertEqual(z.read(name), b"abc")

    def test_empty_and_nested_directories_retained_without_unrelated_content(self):
        with tempfile.TemporaryDirectory() as parent:
            root = Path(parent) / "Tree"; root.mkdir()
            file = root / "keep.txt"; file.write_bytes(b"keep")
            empty = root / "Nested" / "Empty"; empty.mkdir(parents=True)
            (root / "unrelated").mkdir()
            archive = _compress(root, [file], True, directories=[empty])
            with zipfile.ZipFile(archive) as z:
                self.assertEqual(set(z.namelist()), {"keep.txt", "Nested/Empty/"})
                self.assertEqual(z.read("keep.txt"), b"keep")
                self.assertTrue(z.getinfo("Nested/Empty/").is_dir())
            self.assertEqual(file.read_bytes(), b"keep")

    def test_wholly_empty_folder_archive_has_directory_entry(self):
        with tempfile.TemporaryDirectory() as parent:
            root = Path(parent) / "Empty"; root.mkdir()
            archive = _compress(root, [], True, directories=[root])
            with zipfile.ZipFile(archive) as z:
                self.assertEqual(z.namelist(), ["Empty/"])

    def test_flat_empty_folder_name_collision_rejected(self):
        with tempfile.TemporaryDirectory() as parent:
            root = Path(parent)
            one, two = root / "a" / "Empty", root / "b" / "Empty"
            one.mkdir(parents=True); two.mkdir(parents=True)
            with self.assertRaises(ValueError):
                _compress(root, [], False, directories=[one, two])

    def test_archive_no_overwrite_preserves_existing_archive_and_files(self):
        with tempfile.TemporaryDirectory() as parent:
            root = Path(parent) / "Tree"; root.mkdir()
            keep = root / "file"; keep.write_bytes(b"keep")
            archive = Path(str(root) + ".zip"); archive.write_bytes(b"original")
            with self.assertRaises(FileExistsError):
                _compress(root, [keep], True)
            self.assertEqual(archive.read_bytes(), b"original")
            self.assertEqual(keep.read_bytes(), b"keep")
            self.assertFalse(list(Path(parent).glob(".msftgraph-*")))

    def test_actual_folder_download_tracks_empty_directory_entries(self):
        client = SimpleNamespace(api=Mock(), drive_id="d")
        client.api.graph_call.return_value = {"id": "root", "name": "Tree", "folder": {}}
        client.api.iter_collection.side_effect = [[{"id": "empty", "name": "Empty", "folder": {}}], []]
        with tempfile.TemporaryDirectory() as parent:
            archive = Download(client, compress=True, preserve_dir_structure=True).folder(parent, {"id": "root", "folder": {}})[0]
            with zipfile.ZipFile(archive) as z:
                self.assertEqual(z.namelist(), ["Empty/"])
            self.assertTrue((Path(parent) / "Tree" / "Empty").is_dir())

    @unittest.skipUnless(os.name == "nt", "Windows returned-tree namespace")
    def test_short_folder_root_keeps_prefix_when_descendants_are_long(self):
        temporary = tempfile.TemporaryDirectory()
        temporary.name = str(native_path(temporary.name))
        with temporary as parent:
            client = SimpleNamespace(api=Mock(), drive_id="d")
            stream = Mock(status_code=200, headers={}); stream.iter_content.return_value = iter([b"x"])
            client.api.graph_call.side_effect = [{"id": "root", "name": "Tree", "folder": {}}, stream]
            inner = "inner-" + "a" * 100
            filename = "file-" + "b" * 120
            client.api.iter_collection.side_effect = [[{"id": "inner", "name": inner, "folder": {}}],
                [{"id": "f", "name": filename, "file": {}, "size": 1}]]
            root = Download(client).folder(parent, {"id": "root", "folder": {}})[0]
            self.assertTrue(str(root).startswith("\\\\?\\"))
            self.assertEqual((root / inner / filename).read_bytes(), b"x")

    def test_bad_destination_still_closes_stream(self):
        stream = Mock(status_code=200, headers={})
        with self.assertRaises((TypeError, ValueError)):
            _save_response(stream, None, 0)
        stream.close.assert_called_once()


class NumericValidation310(Offline):
    def test_boolean_and_noninteger_file_limits_rejected_before_lookup(self):
        client = SimpleNamespace(api=Mock(), drive_id="d")
        for value in (True, False, 0, -1, 1.5, "2"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                File(client).find("f", num_files=value)
        client.api.graph_call.assert_not_called()

    def test_workers_validate_before_remote_or_local_operations(self):
        client = SimpleNamespace(api=Mock(), drive_id="d")
        for value in (True, False, 0, -1, 1.5, "2"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                Download(client, max_workers=value)
        client.api.graph_call.assert_not_called()

    def test_chunk_sizes_reject_bool_and_float_before_http(self):
        client = SimpleNamespace(api=Mock(), drive_id="d")
        for value in (True, False, 327680.0, 0, -327680):
            with self.subTest(value=value), self.assertRaises(ValueError):
                Upload(client, chunk_size=value)
        client.api.graph_call.assert_not_called()

    def test_batch_sizes_reject_bool_and_float_before_http(self):
        api = self.api()
        for value in (True, False, 1.5, 0, 21, "2"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                api.graph_call(batch_requests=[], batch_size=value)
        api.session.request.assert_not_called()


if __name__ == "__main__":
    unittest.main()
