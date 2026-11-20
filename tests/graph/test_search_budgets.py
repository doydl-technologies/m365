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
"""Offline regressions for bounded legacy search and composite operations."""
from types import SimpleNamespace as NS
from unittest.mock import patch
from urllib.parse import quote
import unittest

from m365.graph import GraphError
from m365.graph._graph import _CollectionBudget
from m365.graph.operations.doc_management import File, Folder, FileFolder
from test_sharepoint import Offline, response

BASE = "https://graph.microsoft.com/v1.0"
GUID = "de1743d3-34c7-4c67-b7b6-17fb3bdb56cd"

def row(identifier="f", name="report", kind="file", parent=None, **fields):
    result = {"id": identifier, "name": name, kind: {}}
    if parent is not None:
        result["parentReference"] = {"id": parent, "driveId": "d"}
    result.update(fields)
    return result

def index_link(query="report", token="next"):
    return BASE + "/drives/d/root/search(q='" + quote(query.replace("'", "''"), safe="") + "')?$skiptoken=" + token

def children_link(parent="target", token="next"):
    return BASE + "/drives/d/items/" + parent + "/children?$skiptoken=" + token

class SearchBudgetTests(Offline):
    def client(self, pages=(), **options):
        c = NS(api=self.api(), drive_id="d", search_mode="indexed",
               search_root=None, search_max_requests=20, search_max_items=100,
               search_max_depth=5, web_url="https://tenant.example/sites/Test",
               drive_web_url="https://tenant.example/sites/Test/Documents")
        for key, value in options.items():
            setattr(c, key, value)
        c.api.session.request.side_effect = [p if hasattr(p, "status_code") else response(p) for p in pages]
        return c

    def fails(self, c, call, code, requests):
        with self.assertRaises(GraphError) as error:
            call()
        self.assertEqual(error.exception.code, code)
        self.assertEqual(c.api.session.request.call_count, requests)
        self.assertTrue(all(p.args[0] == "GET" for p in c.api.session.request.call_args_list))

    def test_original_indexed_item_overflow(self):
        c = self.client([{"value": [row(str(n), "report-"+str(n)) for n in range(12)]}], search_max_items=3)
        self.fails(c, lambda: File(c).find("report", exact_match=False), "search_item_limit", 1)

    def test_original_indexed_page_overflow(self):
        c = self.client([{"value": [row(str(n))], "@odata.nextLink": index_link(token=str(n))}
                         for n in range(10)], search_max_requests=3)
        self.fails(c, lambda: File(c).find("report"), "search_request_limit", 3)

    def test_distinct_empty_pages_still_consume_requests(self):
        c = self.client([{"value": [], "@odata.nextLink": index_link(token=str(n))} for n in range(10)],
                        search_max_requests=3)
        self.fails(c, lambda: File(c).find("report"), "search_request_limit", 3)

    def test_nonmatching_rows_consume_item_budget(self):
        c = self.client([{"value": [row(str(n), "unrelated-"+str(n)) for n in range(8)]}], search_max_items=3)
        self.fails(c, lambda: File(c).find("report"), "search_item_limit", 1)

    def test_remote_shortcuts_consume_item_budget(self):
        c = self.client([{"value": [row(str(n), remoteItem={"id":"external"}) for n in range(8)]}],
                        search_max_items=3)
        self.fails(c, lambda: File(c).find("report"), "search_item_limit", 1)

    def test_explicit_result_limit_does_not_hide_incomplete_scan(self):
        c = self.client([{"value": [row()], "@odata.nextLink": index_link()}], search_max_requests=1)
        self.fails(c, lambda: File(c).find("report", num_files=1), "search_request_limit", 1)

    def test_sorted_result_limit_does_not_hide_incomplete_scan(self):
        c = self.client([{"value": [row(createdDateTime="2026-01-01T00:00:00Z")],
                          "@odata.nextLink": index_link()}], search_max_requests=1)
        self.fails(c, lambda: File(c).find("report", sort_order="recent", num_files=1), "search_request_limit", 1)

    def test_exact_item_and_request_boundary_succeeds(self):
        values = [row("one"), row("two")]
        c = self.client([{"value": values[:1], "@odata.nextLink": index_link()},
                         {"value": values[1:]}], search_max_requests=2, search_max_items=2)
        self.assertEqual(File(c).find("report"), values)
        self.assertEqual(c.api.session.request.call_count, 2)

    def test_oversized_page_yields_no_first_item(self):
        c = self.client([{"value": [row("one"), row("two")]}])
        budget = _CollectionBudget(10, 1, prefix="search")
        self.fails(c, lambda: next(c.api.iter_collection("drives/d/root/search(q='report')", _budget=budget)),
                   "search_item_limit", 1)

    def test_over_budget_page_response_is_closed(self):
        reply = response({"value": [row("one"), row("two")]})
        c = self.client([reply], search_max_items=1)
        self.fails(c, lambda: File(c).find("report"), "search_item_limit", 1)
        reply.close.assert_called_once()

    def test_operation_budget_resets_for_each_independent_call(self):
        c = self.client([{"value": [row()]}, {"value": [row()]}], search_max_requests=1, search_max_items=1)
        self.assertEqual(len(File(c).find("report")), 1)
        self.assertEqual(len(File(c).find("report")), 1)

    def test_failure_does_not_poison_a_later_operation(self):
        c = self.client([{"value": [row()], "@odata.nextLink": index_link()},
                         {"value": [row()]}], search_max_requests=1)
        self.fails(c, lambda: File(c).find("report"), "search_request_limit", 1)
        self.assertEqual(len(File(c).find("report")), 1)

    def test_opaque_continuation_is_forwarded_unchanged(self):
        link = index_link(token="a%2BB+z%26&custom=%252F")
        c = self.client([{"value": [], "@odata.nextLink": link}, {"value": [row()]}],
                        search_max_requests=2, search_max_items=1)
        self.assertEqual(File(c).find("report"), [row()])
        self.assertEqual(c.api.session.request.call_args.args[1], link)

    def test_foreign_continuation_still_fails_without_following(self):
        c = self.client([{"value": [row()], "@odata.nextLink": BASE+"/drives/other/root/search(q='report')"}])
        self.fails(c, lambda: File(c).find("report"), "collection_context_changed", 1)

    def test_permission_error_is_not_hidden(self):
        c = self.client([response({}, status=403)], search_max_requests=1)
        self.fails(c, lambda: File(c).find("report"), "request_failed", 1)

    def test_transport_retries_are_not_extra_logical_pages(self):
        c = self.client([response({}, status=429, headers={"Retry-After":"0"}), {"value":[row()]}],
                        search_max_requests=1, search_max_items=1)
        with patch("m365.graph._graph.time.sleep") as sleep:
            self.assertEqual(File(c).find("report"), [row()])
        self.assertEqual(c.api.session.request.call_count, 2)
        sleep.assert_called_once_with(0)

    def test_low_level_collections_without_budget_remain_compatible(self):
        c = self.client([{"value": [row("one"), row("two")]}], search_max_items=1)
        self.assertEqual(len(list(c.api.iter_collection("drives/d/root/search(q='report')"))), 2)

    def test_default_item_ceiling_is_finite(self):
        c = self.client([{"value": [row(str(n)) for n in range(5001)]}])
        del c.search_max_items
        self.fails(c, lambda: File(c).find("report"), "search_item_limit", 1)

    def test_default_page_ceiling_bounds_empty_continuations(self):
        c = self.client([{"value": [], "@odata.nextLink": index_link(token=str(n))} for n in range(201)])
        del c.search_max_requests
        self.fails(c, lambda: File(c).find("report"), "search_request_limit", 200)

    def test_folder_expansion_shares_item_budget(self):
        target = row("target", "Target", "folder")
        c = self.client([{"value": [target]}, {"value": [row("one", parent="target"), row("two", parent="target")]}],
                        search_max_items=2)
        self.fails(c, lambda: Folder(c).find("Target", include_children=True), "search_item_limit", 2)

    def test_multiple_folder_expansions_share_request_budget(self):
        c = self.client([{"value": [row("first","Target","folder"),row("second","Target","folder")]},
                         {"value": []}], search_max_requests=2)
        self.fails(c, lambda: Folder(c).find("Target", include_children=True), "search_request_limit", 2)

    def test_child_empty_continuations_share_lookup_budget(self):
        c = self.client([{"value": [row("target","Target","folder")]},
                         {"value": [], "@odata.nextLink": children_link()}], search_max_requests=2)
        self.fails(c, lambda: Folder(c).find("Target", include_children=True), "search_request_limit", 2)

    def test_expanded_folder_succeeds_at_exact_shared_boundary(self):
        child = row("child", parent="target")
        c = self.client([{"value": [row("target","Target","folder")]}, {"value": [child]}],
                        search_max_requests=2, search_max_items=2)
        result = Folder(c).find("Target", include_children=True)
        self.assertEqual(result[0]["children"], [child])

    def test_explicit_folder_path_read_is_counted_before_expansion(self):
        c = self.client([row("target","Target","folder")], search_max_requests=1)
        self.fails(c, lambda: Folder(c).find("Nested/Target", include_children=True), "search_request_limit", 1)

    def test_enumeration_and_expansion_share_request_budget(self):
        target = row("target", "Target", "folder")
        c = self.client([{"value": [target]}, {"value": []}], search_mode="enumerate", search_max_requests=2)
        self.fails(c, lambda: Folder(c).find("Target", include_children=True), "search_request_limit", 2)

    def test_enumeration_and_expansion_share_processed_row_budget(self):
        c = self.client([{"value": [row("target","Target","folder")]},
                         {"value": [row("child",parent="target")]},
                         {"value": [row("child",parent="target")]}],
                        search_mode="enumerate", search_max_items=2)
        self.fails(c, lambda: Folder(c).find("Target", include_children=True), "search_item_limit", 3)

    def test_combined_lookup_shares_request_budget(self):
        c = self.client([{"value": [row()]}], search_max_requests=1)
        self.fails(c, lambda: FileFolder(c).find("report", "Target"), "search_request_limit", 1)

    def test_combined_lookup_shares_item_budget(self):
        c = self.client([{"value": [row()]}, {"value": [row("target","Target","folder")]}], search_max_items=1)
        self.fails(c, lambda: FileFolder(c).find("report", "Target"), "search_item_limit", 2)

    def test_combined_lookup_succeeds_at_shared_boundary(self):
        f, d = row(), row("target","Target","folder")
        c = self.client([{"value": [f]}, {"value": [d]}], search_max_requests=2, search_max_items=2)
        self.assertEqual(FileFolder(c).find("report","Target"), [f,d])

    def test_combined_direct_path_lookups_share_requests(self):
        c = self.client([row(name="report")], search_max_requests=1)
        self.fails(c, lambda: FileFolder(c).find("Nested/report","Nested/Target"), "search_request_limit", 1)

    def test_combined_enumerations_share_requests(self):
        c = self.client([{"value": [row()]}], search_mode="enumerate", search_max_requests=1)
        self.fails(c, lambda: FileFolder(c).find("report","Target"), "search_request_limit", 1)

    def test_browser_guid_scan_and_final_path_read_share_requests(self):
        f = row(sharepointIds={"listItemUniqueId":GUID})
        c = self.client([{"value": [f]}], search_max_requests=1)
        url = c.web_url + "/_layouts/15/Doc.aspx?sourcedoc="+GUID+"&file=report"
        self.fails(c, lambda: File(c).find(url), "search_request_limit", 1)

    def test_browser_guid_scan_and_final_path_read_share_rows(self):
        f = row(sharepointIds={"listItemUniqueId":GUID})
        c = self.client([{"value": [f]}, row()], search_max_items=1)
        url = c.web_url + "/_layouts/15/Doc.aspx?sourcedoc="+GUID+"&file=report"
        self.fails(c, lambda: File(c).find(url), "search_item_limit", 2)

    def test_browser_guid_lookup_succeeds_at_shared_boundary(self):
        f = row(sharepointIds={"listItemUniqueId":GUID})
        c = self.client([{"value": [f]}, row()], search_max_requests=2, search_max_items=2)
        url = c.web_url + "/_layouts/15/Doc.aspx?sourcedoc="+GUID+"&file=report"
        self.assertEqual(File(c).find(url), [row()])

    def test_budget_validation_precedes_http(self):
        for field in ("search_max_requests", "search_max_items"):
            for value in (0, -1, True, 1.0, "1", None):
                with self.subTest(field=field,value=value):
                    c = self.client([], **{field:value})
                    with self.assertRaises(ValueError):
                        File(c).find("report")
                    c.api.session.request.assert_not_called()

if __name__ == "__main__":
    unittest.main()
