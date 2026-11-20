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
"""Offline query invariants, with transport-level fault injection and no live data."""
import unittest
from types import SimpleNamespace as NS
from unittest.mock import Mock
from urllib.parse import quote

from m365.graph._graph import CollectionScope, GraphError, read_response
from m365.graph._mail_common import checked_mail_page
from m365.graph.operations.doc_management import File as FileSearch
from m365.graph.operations.mail_management import Attachment, Delta, Folder, Message
from test_mail_correctness_320 import MailClientFixture
from test_sharepoint import response


BASE = "https://graph.microsoft.com/v1.0"


def drive_item(name="report.txt", identifier="f", **fields):
    return dict(id=identifier, name=name, file={}, **fields)


class FileQueries(MailClientFixture):
    def files(self, replies, mode="indexed"):
        c = NS(api=self.api(), drive_id="d", search_mode=mode, search_root=None,
               search_max_depth=5, search_max_items=50, search_max_requests=20)
        c.api.session.request.side_effect = replies
        return c

    def test_partial_index_search_applies_casefolded_filename_predicate(self):
        rows = [drive_item("unrelated.txt", "a"), drive_item("REPORT-é.txt", "b")]
        c = self.files([response({"value": rows})])
        self.assertEqual(FileSearch(c).find("report", exact_match=False), [rows[1]])

    def test_partial_matching_agrees_between_index_and_enumeration(self):
        rows = [drive_item("Straße.txt", "a"), drive_item("other.txt", "b")]
        for mode in ("indexed", "enumerate"):
            with self.subTest(mode=mode):
                c = self.files([response({"value": rows})], mode)
                self.assertEqual(FileSearch(c).find("STRASSE", exact_match=False), [rows[0]])

    def test_index_search_does_not_stop_after_a_nonmatching_page(self):
        link = BASE + "/drives/d/root/search(q='report')?$skiptoken=a%2Bb"
        hit = drive_item()
        c = self.files([response({"value": [drive_item("other", "b")], "@odata.nextLink": link}),
                        response({"value": [hit]})])
        self.assertEqual(FileSearch(c).find("report", exact_match=False), [hit])
        self.assertEqual(c.api.session.request.call_args.args[1], link)

    def test_duplicate_index_ids_fail_even_if_metadata_is_identical(self):
        row = drive_item()
        for other in (row, drive_item(parentReference={"id": "other"})):
            with self.subTest(other=other):
                c = self.files([response({"value": [row, other]})])
                with self.assertRaises(GraphError) as error:
                    FileSearch(c).find("report.txt")
                self.assertEqual(error.exception.code, "search_duplicate_item")

    def test_duplicate_index_ids_across_pages_are_not_last_write_wins(self):
        link = BASE + "/drives/d/root/search(q='report.txt')?page=2"
        c = self.files([response({"value": [drive_item()], "@odata.nextLink": link}),
                        response({"value": [drive_item(parentReference={"id": "B"})]})])
        with self.assertRaises(GraphError):
            FileSearch(c).find("report.txt")
        self.assertEqual(c.api.session.request.call_count, 2)

    def test_index_rejects_foreign_drive_collection_or_search_expression(self):
        for path in ("drives/e/root/search(q='report.txt')", "drives/d/root/children",
                     "drives/d/root/search(q='different.txt')", "users/other/messages"):
            with self.subTest(path=path):
                c = self.files([response({"value": [drive_item()], "@odata.nextLink": BASE + "/" + path})])
                with self.assertRaises(GraphError):
                    FileSearch(c).find("report.txt")
                self.assertEqual(c.api.session.request.call_count, 1)

    def test_opaque_index_continuation_is_forwarded_unchanged(self):
        link = BASE + "/drives/d/root/search(q='report.txt')?$skiptoken=a%2BB+z%26&custom=%252F"
        c = self.files([response({"value": [], "@odata.nextLink": link}),
                        response({"value": [drive_item()]})])
        self.assertEqual(len(FileSearch(c).find("report.txt")), 1)
        self.assertEqual(c.api.session.request.call_args.args[1], link)

    def test_same_path_with_equivalent_segment_encoding_is_allowed(self):
        link = BASE + "/drives/%64/root/search(q='report.txt')?$skiptoken=opaque"
        c = self.files([response({"value": [], "@odata.nextLink": link}), response({"value": []})])
        self.assertEqual(FileSearch(c).find("report.txt"), [])
        self.assertEqual(c.api.session.request.call_args.args[1], link)

    def test_encoded_slash_is_not_a_path_separator_for_scope(self):
        scope = CollectionScope(BASE + "/users/a%2Fb/messages")
        with self.assertRaises(GraphError):
            scope.check_page_url(BASE + "/users/a/b/messages")
        scope.check_page_url(BASE + "/users/a%2fb/messages?token=untouched")

    def test_bad_continuation_never_yields_a_current_page_match(self):
        c = self.files([response({"value": [drive_item()], "@odata.nextLink": BASE + "/drives/e/root/children"})])
        with self.assertRaises(GraphError):
            next(c.api.iter_collection("drives/d/root/search(q='report.txt')"))

    def test_deleted_removed_and_error_markers_are_not_search_results(self):
        for mode in ("indexed", "enumerate", "path"):
            for marker in ("deleted", "@removed", "error"):
                with self.subTest(mode=mode, marker=marker):
                    row = drive_item(**{marker: {}})
                    page = row if mode == "path" else {"value": [row]}
                    c = self.files([response(page)], mode)
                    with self.assertRaises(GraphError):
                        FileSearch(c).find("Folder/report.txt" if mode == "path" else "report.txt")

    def test_error_envelope_with_value_is_not_a_valid_page(self):
        for mode in ("indexed", "enumerate"):
            with self.subTest(mode=mode):
                c = self.files([response({"value": [drive_item()], "error": {"code": "PRIVATE"}})], mode)
                with self.assertRaises(GraphError) as error:
                    FileSearch(c).find("report.txt")
                self.assertNotIn("PRIVATE", str(error.exception))

    def test_collection_and_path_reads_require_200_and_close_responses(self):
        for mode in ("indexed", "enumerate", "path"):
            for status in (201, 202, 204, 205, 206, 299):
                with self.subTest(mode=mode, status=status):
                    reply = response(drive_item() if mode == "path" else {"value": [drive_item()]}, status=status)
                    c = self.files([reply], mode)
                    with self.assertRaises(GraphError) as error:
                        FileSearch(c).find("Folder/report.txt" if mode == "path" else "report.txt")
                    self.assertEqual(error.exception.status_code, status)
                    reply.close.assert_called_once()

    def test_later_page_status_error_cannot_return_a_partial_search(self):
        link = BASE + "/drives/d/root/search(q='report.txt')?page=2"
        c = self.files([response({"value": [drive_item()], "@odata.nextLink": link}),
                        response({"value": []}, status=206)])
        with self.assertRaises(GraphError):
            FileSearch(c).find("report.txt")

    def test_raw_upload_status_contract_is_unchanged(self):
        for status in (201, 202, 204):
            with self.subTest(status=status):
                api = self.api()
                api.session.request.return_value = response({"nextExpectedRanges": ["1"]}, status=status)
                result = api.graph_call(resource_url="drives/d/root/createUploadSession", method="POST")
                self.assertEqual(result, {} if status == 204 else {"nextExpectedRanges": ["1"]})


class MailQueries(MailClientFixture):
    def test_mailbox_folder_and_collection_continuations_are_bound(self):
        endpoints = ("users/u/messages", "users/u/mailFolders/A/messages",
                     "users/u/messages/m/attachments", "users/u/mailFolders")
        for endpoint in endpoints:
            for wrong in ("users/other/messages", "users/u/mailFolders/B/messages",
                          "users/u/messages/other/attachments", "users/u/mailFolders/B/childFolders"):
                if endpoint == wrong:
                    continue
                with self.subTest(endpoint=endpoint, wrong=wrong):
                    c = self.client([response({"value": [], "@odata.nextLink": BASE + "/" + wrong})])
                    from m365.graph._mail_common import mail_pages
                    with self.assertRaises(GraphError):
                        list(mail_pages(c.api.graph_call, endpoint))
                    self.assertEqual(c.api.session.request.call_count, 1)

    def test_mail_continuation_checked_even_when_first_page_meets_limit(self):
        c = self.client([response({"value": [{"id": "a"}], "@odata.nextLink": BASE + "/users/other/messages"})])
        with self.assertRaises(GraphError):
            Message(c).find(top=1)
        self.assertEqual(c.api.session.request.call_count, 1)

    def test_mail_opaque_continuation_preserves_bytes_and_immutable_headers(self):
        link = BASE + "/users/u/messages?$skiptoken=a%2Bb+z%26"
        c = self.client([response({"value": [{"id": "a"}], "@odata.nextLink": link}),
                         response({"value": [{"id": "b"}]})])
        self.assertEqual(Message(c).find(top=2), [{"id": "a"}, {"id": "b"}])
        self.assertEqual(c.api.session.request.call_args.args[1], link)
        for call in c.api.session.request.call_args_list:
            self.assertIn('IdType="ImmutableId"', call.kwargs["headers"]["Prefer"])

    def test_malformed_message_objects_raise_redacted_errors(self):
        for row in ({"subject": "invoice"}, {"id": ""}, {"id": False},
                    {"id": "a", "error": {"private": "PRIVATE"}},
                    {"id": "a", "@removed": {}}, {"id": "a", "deleted": {}}):
            with self.subTest(row=row):
                c = self.client([response({"value": [row]})])
                with self.assertRaises(GraphError) as error:
                    Message(c).find()
                self.assertNotIn("PRIVATE", str(error.exception))

    def test_wrong_address_object_types_are_rejected_with_and_without_query(self):
        for query in (None, "invoice"):
            for field, value in (("from", "bad"), ("from", []), ("sender", [1]),
                                 ("from", {"emailAddress": "bad"}),
                                 ("sender", {"emailAddress": 7}),
                                 ("from", {"emailAddress": {"address": []}}),
                                 ("sender", {"emailAddress": {"name": {}}})):
                with self.subTest(query=query, field=field, value=value):
                    c = self.client([response({"value": [{"id": "a", "subject": "invoice", field: value}]})])
                    with self.assertRaises(GraphError) as error:
                        Message(c).find(query=query)
                    self.assertEqual(error.exception.code, "invalid_mail_address")

    def test_nullable_addresses_and_text_remain_valid(self):
        row = {"id": "a", "subject": "invoice", "bodyPreview": None, "from": None,
               "sender": {"emailAddress": {"address": None, "name": None}}}
        c = self.client([response({"value": [row]})])
        self.assertEqual(Message(c).find(query="invoice"), [row])

    def test_custom_projection_may_intentionally_omit_id(self):
        c = self.client([response({"value": [{"subject": "invoice"}]})])
        self.assertEqual(Message(c).find(select="subject"), [{"subject": "invoice"}])
        self.assertIn("$select=subject", c.api.session.request.call_args.args[1])

    def test_invalid_mail_folder_names_do_not_crash_fallback(self):
        for name in (None, [], {}, False, 4, "", " "):
            with self.subTest(name=name):
                c = self.client([response({"value": [{"id": "a", "displayName": name}]})])
                with self.assertRaises(GraphError) as error:
                    Folder(c).find("inbox", recursive=False)
                self.assertEqual(error.exception.code, "invalid_mail_folder")

    def test_every_mail_read_requires_completed_status(self):
        for action in (lambda c: Message(c).find(), lambda c: Attachment(c).list("m"),
                       lambda c: Folder(c).find(recursive=False), lambda c: Delta(c).messages(),
                       lambda c: Delta(c).folders(), lambda c: Message(c).get("m"),
                       lambda c: Attachment(c).get("m", "a"), lambda c: Folder(c).resolve("id"),
                       lambda c: Message(c).get("m", mime=True)):
            for status in (201, 202, 204, 206):
                with self.subTest(action=action, status=status):
                    reply = response({"value": [], "id": "a", "@odata.deltaLink": "cursor"}, status=status)
                    c = self.client([reply])
                    with self.assertRaises(GraphError) as error:
                        action(c)
                    self.assertEqual(error.exception.status_code, status)
                    reply.close.assert_called_once()

    def test_delta_retains_repeated_events_and_tombstones(self):
        values = [{"id": "a"}, {"id": "a", "@removed": {"reason": "deleted"}}]
        c = self.client([response({"value": values, "@odata.deltaLink": "cursor"})])
        self.assertEqual(Delta(c).messages()["value"], values)

    def test_zero_message_folder_or_attachment_limit_makes_no_request(self):
        for action in (lambda c: Message(c).find(top=0), lambda c: Message(c).find(top=0, max_items=3),
                       lambda c: Message(c).find(max_items=0), lambda c: Attachment(c).list("m", max_items=0),
                       lambda c: Folder(c).find(top=0)):
            with self.subTest(action=action):
                c = self.client()
                self.assertEqual(action(c), [])
                c.api.session.request.assert_not_called()

    def test_invalid_counts_are_rejected_before_http(self):
        for value in (True, False, -1, 1.8, "3", [], {}, float("inf")):
            for action in (lambda c, v: Message(c).find(top=v), lambda c, v: Message(c).find(max_items=v),
                           lambda c, v: Folder(c).find(top=v), lambda c, v: Attachment(c).list("m", max_items=v),
                           lambda c, v: Delta(c).messages(top=v), lambda c, v: Delta(c).folders(top=v)):
                with self.subTest(value=value, action=action):
                    c = self.client()
                    with self.assertRaises(ValueError):
                        action(c, value)
                    c.api.session.request.assert_not_called()

    def test_delta_zero_page_size_is_not_an_empty_initial_sync(self):
        for method in ("messages", "folders"):
            c = self.client()
            with self.subTest(method=method), self.assertRaises(ValueError):
                getattr(Delta(c), method)(top=0)
            c.api.session.request.assert_not_called()

    def test_explicit_bad_cursor_never_starts_an_initial_sync(self):
        for value in ("", " ", "\n", False, 0, [], {}, "cursor\nsecret"):
            for method in ("messages", "folders"):
                with self.subTest(value=value, method=method):
                    c = self.client()
                    with self.assertRaises(ValueError):
                        getattr(Delta(c), method)(delta_link=value)
                    c.api.session.request.assert_not_called()

    def test_wrong_type_unread_is_not_silently_ignored(self):
        for value in ("false", 0, 1, [], {}):
            c = self.client()
            with self.subTest(value=value), self.assertRaises(ValueError):
                Message(c).find(unread=value)
            c.api.session.request.assert_not_called()

    def test_boolean_unread_filters_remain_distinct(self):
        for value, expected in ((True, "false"), (False, "true"), (None, None)):
            c = self.client([response({"value": []})])
            Message(c).find(unread=value)
            url = c.api.session.request.call_args.args[1]
            self.assertEqual("$filter=isRead eq " + expected in url if expected else "$filter" not in url, True)

    def test_delta_410_still_exposes_resync_not_payload(self):
        c = self.client([response({"error": "PRIVATE"}, status=410)])
        with self.assertRaises(GraphError) as error:
            Delta(c).messages()
        self.assertEqual((error.exception.status_code, error.exception.code), (410, "delta_resync_required"))
        self.assertNotIn("PRIVATE", str(error.exception))

    def test_read_metadata_protocol_rejects_missing_or_invalid_status(self):
        for envelope in ({}, {"status_code": True, "body": {}}, {"status_code": "200", "body": {}},
                         {"status_code": 200}, {"value": []}):
            with self.subTest(envelope=envelope), self.assertRaises(GraphError):
                read_response(Mock(return_value=envelope), resource_url="anything")

    def test_malformed_later_page_fails_even_after_matching_message(self):
        link = BASE + "/users/u/messages?token=2"
        c = self.client([response({"value": [{"id": "a"}], "@odata.nextLink": link}),
                         response({"value": [{"id": "b", "sender": []}]})])
        with self.assertRaises(GraphError):
            Message(c).find(top=2)
        self.assertEqual(c.api.session.request.call_count, 2)

    def test_invalid_folder_selectors_never_widen_or_coerce_a_query(self):
        for selector in (False, 0, 1, [], {}, "", " ", {"id": ["bad"]}, {"id": 7}, {"displayName": {}}):
            for action in (lambda c, v: Message(c).find(folder=v), lambda c, v: Delta(c).messages(folder=v)):
                with self.subTest(selector=selector, action=action):
                    c = self.client()
                    with self.assertRaises(ValueError):
                        action(c, selector)
                    c.api.session.request.assert_not_called()

    def test_valid_folder_selectors_keep_the_selected_scope(self):
        for selector in ("A/+#", {"id": "A/+#"}, {"displayName": "A/+#"}):
            c = self.client([response({"value": []})])
            Message(c).find(folder=selector)
            self.assertIn("/mailFolders/A%2F%2B%23/messages", c.api.session.request.call_args.args[1])

    def test_normal_mail_duplicates_raise_instead_of_truncating_coverage(self):
        row = {"id": "a"}
        link = BASE + "/users/u/messages?token=2"
        for replies in ([response({"value": [row, row]})],
                        [response({"value": [row], "@odata.nextLink": link}), response({"value": [row]})]):
            c = self.client(replies)
            with self.assertRaises(GraphError) as error:
                Message(c).find(top=3)
            self.assertEqual(error.exception.code, "mailbox_duplicate_item")

    def test_one_shot_projection_is_not_consumed_twice(self):
        c = self.client([response({"value": [{"subject": "invoice"}]})])
        self.assertEqual(Message(c).find(select=iter(["subject"])), [{"subject": "invoice"}])
        self.assertIn("$select=subject&", c.api.session.request.call_args.args[1])
