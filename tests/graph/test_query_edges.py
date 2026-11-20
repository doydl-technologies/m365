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
"""Offline regressions for query flags and remote timestamp error boundaries."""
from types import SimpleNamespace as NS
from unittest.mock import patch
import unittest

from m365.graph import Site, GraphError
from m365.graph.operations.doc_management import File, Folder
from test_sharepoint import Offline, response
from test_storage_parity import StorageFixture
from test_onedrive import ROOT, page

INVALID_FLAGS = ("false", "true", "", 0, 1, None, [], {}, b"false", object())
INVALID_DATES = (
    "0001-01-01T00:00:00+14:00", "9999-12-31T23:59:59-14:00",
    "not-a-timestamp", "2026-02-30T00:00:00Z", "2026-01-01T00:00:00+24:00",
    "2026-01-01T00:00:00+02:60", "PRIVATE invalid date", 42, True, ["2026-01-01"], {"year":2026},
)

def row(identifier="f", name="report.txt", kind="file", **fields):
    return {"id":identifier, "name":name, kind:{}, **fields}

class LegacyQueryEdges(Offline):
    def client(self, bodies=(), mode="indexed"):
        c = NS(api=self.api(), drive_id="d", search_mode=mode, search_root=None,
               search_max_requests=20, search_max_items=100, search_max_depth=5)
        c.api.session.request.side_effect = [response(b) for b in bodies]
        c.search_by = Site.SearchBy(c)
        return c

    def test_invalid_filename_flag_rejected_before_http_in_every_lookup_mode(self):
        for mode in ("indexed","enumerate","path"):
            for value in INVALID_FLAGS:
                with self.subTest(mode=mode,value=repr(value)):
                    c = self.client(mode="indexed" if mode=="path" else mode)
                    term = "Parent/report.txt" if mode=="path" else "report.txt"
                    with self.assertRaisesRegex(ValueError,"exact_match"):
                        c.search_by.filename(term,exact_match=value)
                    c.api.session.request.assert_not_called()

    def test_invalid_folder_flags_rejected_before_http_in_every_lookup_mode(self):
        for mode in ("indexed","enumerate","path"):
            for key in ("exact_match","include_children"):
                for value in INVALID_FLAGS:
                    with self.subTest(mode=mode,key=key,value=repr(value)):
                        c = self.client(mode="indexed" if mode=="path" else mode)
                        term = "Parent/Target" if mode=="path" else "Target"
                        with self.assertRaisesRegex(ValueError,key):
                            c.search_by.foldername(term,**{key:value})
                        c.api.session.request.assert_not_called()

    def test_valid_filename_exact_and_partial_controls_keep_their_meaning(self):
        values = [row("one"),row("two","report-other.txt")]
        for mode in ("indexed","enumerate"):
            for exact, expected in ((True,[values[0]]),(False,values)):
                with self.subTest(mode=mode,exact=exact):
                    c = self.client([{"value":values}],mode)
                    self.assertEqual(c.search_by.filename("report" if not exact else "report.txt",exact_match=exact),expected)

    def test_include_children_false_avoids_the_child_request(self):
        value = row("target","Target","folder")
        c = self.client([{"value":[value]}])
        self.assertEqual(c.search_by.foldername("Target",include_children=False),[value])
        self.assertEqual(c.api.session.request.call_count,1)

    def test_include_children_true_retains_expansion(self):
        target, child = row("target","Target","folder"), row(parentReference={"id":"target","driveId":"d"})
        c = self.client([{"value":[target]},{"value":[child]}])
        result = c.search_by.foldername("Target",include_children=True)
        self.assertEqual(result[0]["children"],[child])
        self.assertEqual(c.api.session.request.call_count,2)

    def test_false_exact_folder_flag_still_means_partial_matching(self):
        value = row("target","Target-Archive","folder")
        c = self.client([{"value":[value]}])
        self.assertEqual(c.search_by.foldername("Target",exact_match=False),[value])

    def test_bad_sorted_dates_raise_controlled_error_for_all_lookup_modes(self):
        for mode in ("indexed","enumerate","path"):
            for order in ("recent","oldest"):
                for value in INVALID_DATES:
                    with self.subTest(mode=mode,order=order,value=value):
                        item = row(createdDateTime=value)
                        c = self.client([item] if mode=="path" else [{"value":[item]}],
                                        "indexed" if mode=="path" else mode)
                        with self.assertRaises(GraphError) as error:
                            c.search_by.filename("Parent/report.txt" if mode=="path" else "report.txt",sort_order=order)
                        self.assertEqual(error.exception.code,"invalid_created_datetime")
                        self.assertIsNone(error.exception.__cause__)
                        self.assertTrue(error.exception.__suppress_context__)
                        self.assertNotIn("PRIVATE",str(error.exception))

    def test_missing_dates_keep_the_existing_missing_reason(self):
        for fields in ({},{"createdDateTime":None},{"createdDateTime":""}):
            with self.subTest(fields=fields):
                c = self.client([{"value":[row(**fields)]}])
                with self.assertRaises(GraphError) as error:
                    c.search_by.filename("report.txt",sort_order="recent")
                self.assertEqual(error.exception.code,"missing_created_datetime")

    def test_valid_extreme_utc_dates_are_not_rejected(self):
        values = [row("min",createdDateTime="0001-01-01T00:00:00Z"),
                  row("max",createdDateTime="9999-12-31T23:59:59Z")]
        for order, expected in (("oldest",values),("recent",values[::-1])):
            with self.subTest(order=order):
                c = self.client([{"value":values}])
                self.assertEqual(c.search_by.filename("report.txt",sort_order=order,num_files=2),expected)

    def test_valid_offset_dates_sort_by_utc_instant(self):
        values = [row("one",createdDateTime="2026-01-01T01:00:00+02:00"),
                  row("two",createdDateTime="2025-12-31T23:30:00Z"),
                  row("three",createdDateTime="2026-01-01T00:00:00-02:00")]
        c = self.client([{"value":values[::-1]}])
        self.assertEqual(c.search_by.filename("report.txt",sort_order="oldest",num_files=3),values)

    def test_legacy_sort_preserves_implicit_single_result(self):
        values = [row("old",createdDateTime="2026-01-01T00:00:00Z"),
                  row("new",createdDateTime="2026-02-01T00:00:00Z")]
        c = self.client([{"value":values}])
        self.assertEqual(c.search_by.filename("report.txt",sort_order="recent"),values[1:])

    def test_unsorted_lookup_does_not_require_a_parseable_timestamp(self):
        item = row(createdDateTime="unusable")
        c = self.client([{"value":[item]}])
        self.assertEqual(c.search_by.filename("report.txt"),[item])

    def test_later_bad_date_cannot_be_hidden_by_num_files(self):
        values = [row("good",createdDateTime="2026-01-01T00:00:00Z"),
                  row("bad",createdDateTime="not-a-timestamp")]
        c = self.client([{"value":values}])
        with self.assertRaises(GraphError):
            c.search_by.filename("report.txt",sort_order="recent",num_files=1)

    def test_unrelated_parser_runtime_error_is_not_swallowed(self):
        c = self.client([{"value":[row(createdDateTime="2026-01-01T00:00:00Z")]}])
        with patch("m365.graph.operations.doc_management.parse_graph_datetime",
                   side_effect=RuntimeError("SYNTHETIC parser defect")):
            with self.assertRaisesRegex(RuntimeError,"SYNTHETIC"):
                c.search_by.filename("report.txt",sort_order="recent")

class ModernQueryEdges(StorageFixture):
    def add_item(self, kind, name, date):
        identifier = self.service.add(name,data=b"x" if kind=="file" else None)
        self.service.nodes[identifier]["createdDateTime"] = date
        return identifier

    def test_invalid_legacy_flags_on_site_do_not_start_an_operation(self):
        c = self.client("site")
        before = list(self.service.records)
        for value in INVALID_FLAGS:
            for method, keyword in ((c.search_by.filename,"exact_match"),
                                    (c.search_by.foldername,"exact_match"),
                                    (c.search_by.foldername,"include_children")):
                with self.subTest(keyword=keyword,value=repr(value)),self.assertRaises(ValueError):
                    method("report",**{keyword:value})
        self.assertEqual(self.service.records,before)
        self.assertEqual(self.mutations(),[])

    def test_valid_extreme_dates_sort_correctly_for_both_item_kinds_and_facades(self):
        for item_kind in ("file","folder"):
            old = self.add_item(item_kind,"edge-"+item_kind+"-old","0001-01-01T00:00:00Z")
            new = self.add_item(item_kind,"edge-"+item_kind+"-new","9999-12-31T23:59:59Z")
            for facade in ("site","drive"):
                c = self.client(facade)
                method = c.find_files if item_kind=="file" else c.find_folders
                for order, expected in (("oldest",[old,new]),("recent",[new,old])):
                    with self.subTest(facade=facade,item_kind=item_kind,order=order):
                        self.assertEqual([r["id"] for r in method("edge-"+item_kind,sort_order=order)],expected)
        self.assertEqual(self.mutations(),[])

    def test_unsorted_and_name_sorted_queries_ignore_unused_date_metadata(self):
        i = self.add_item("file","report.txt","0001-01-01T00:00:00+14:00")
        for facade in ("site","drive"):
            c = self.client(facade)
            for order in (None,"name","name_desc"):
                with self.subTest(facade=facade,order=order):
                    self.assertEqual(c.find_files("report",sort_order=order)[0]["id"],i)

    def test_late_paginated_bad_date_never_yields_sorted_partial_success(self):
        from urllib.parse import urlsplit
        good = self.add_item("file","report-good.txt","2026-01-01T00:00:00Z")
        bad = self.add_item("file","report-bad.txt","9999-12-31T23:59:59-14:00")
        original = self.service.graph
        link = "https://graph.microsoft.com/v1.0/drives/drive/items/anchor/children?$skiptoken=late"
        def pages(m,u,**kw):
            if u == link:
                return page(self.service.metadata(bad))
            if urlsplit(u).path.endswith("/items/anchor/children"):
                return page(self.service.metadata(good),link=link)
            return original(m,u,**kw)
        with patch.object(self.service,"graph",side_effect=pages):
            for facade in ("site","drive"):
                with self.subTest(facade=facade),self.assertRaises(GraphError) as error:
                    self.client(facade).find_files("report",sort_order="recent",limit=1)
                self.assertEqual(error.exception.code,"invalid_created_datetime")
        self.assertEqual(self.mutations(),[])

    def test_valid_offset_and_fractional_dates_still_sort_correctly(self):
        values = [("a","2026-01-01T01:00:00+02:00"),
                  ("b","2025-12-31T23:30:00.1234567Z"),
                  ("c","2026-01-01T00:00:00-02:00")]
        ids = [self.add_item("file",name+".txt",date) for name,date in values]
        for facade in ("site","drive"):
            with self.subTest(facade=facade):
                self.assertEqual([r["id"] for r in self.client(facade).find_files(".txt",sort_order="oldest")],ids)

    def test_runtime_parser_defect_is_not_swallowed_by_modern_search(self):
        self.add_item("file","report.txt","2026-01-01T00:00:00Z")
        for facade in ("site","drive"):
            with self.subTest(facade=facade),patch("m365.graph._core_dates.parse_graph_datetime",
                                                 side_effect=RuntimeError("SYNTHETIC parser defect")):
                with self.assertRaisesRegex(RuntimeError,"SYNTHETIC"):
                    self.client(facade).find_files("report",sort_order="recent")

def malformed_date_case(facade, item_kind, order):
    def test(self):
        i = self.add_item(item_kind,"report","invalid")
        c = self.client(facade)
        method = c.find_files if item_kind=="file" else c.find_folders
        for value in INVALID_DATES:
            self.service.nodes[i]["createdDateTime"] = value
            with self.subTest(value=value),self.assertRaises(GraphError) as error:
                method("report",sort_order=order)
            self.assertEqual(error.exception.code,"invalid_created_datetime")
            self.assertIsNone(error.exception.__cause__)
            self.assertTrue(error.exception.__suppress_context__)
            self.assertNotIn("PRIVATE",str(error.exception))
        self.assertEqual(self.mutations(),[])
    return test

for facade in ("site","drive"):
    for item_kind in ("file","folder"):
        for order in ("recent","oldest"):
            setattr(ModernQueryEdges,f"test_{facade}_{item_kind}_{order}_malformed_dates",
                    malformed_date_case(facade,item_kind,order))

if __name__ == "__main__":
    unittest.main()
