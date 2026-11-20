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
"""Offline mailbox read, attachment and helper regressions."""
import copy
import io
import json
import unittest
from datetime import datetime, timezone, timedelta, tzinfo
from types import SimpleNamespace
from unittest.mock import Mock
from urllib.parse import parse_qs, urlsplit

from m365.graph import GraphError
from m365.graph._mail_graph import MailboxAPI
from m365.graph._mail_common import mail_id
from m365.graph.operations.mail_management import Folder, Message, Attachment, Delta
from m365.graph.operations.mail_transfer import AttachmentUpload, AttachmentDownload
from m365.graph.operations.mail_compose import Draft
from m365.graph.operations.mail_dispatch import Move as MailMove, MessageState, Delete
from m365.graph.operations.doc_dispatch import Move
from m365.graph._preprocessing import lists, transforms
from test_sharepoint import Offline, response


class MailClientFixture(Offline):
    def client(self, replies=()):
        c = SimpleNamespace(graph_url="https://graph.microsoft.com/v1.0", get_token=lambda: "fake",
            request_timeout=(1, 2), max_retries=0, use_immutable_ids=True, base_resource="users/u")
        c.api = MailboxAPI(c)
        self.addCleanup(c.api.close)
        c.api.session.request = Mock(side_effect=list(replies))
        c.api.transfer_session.request = Mock()
        return c


class MailReads320(MailClientFixture):
    def test_message_page_error_cannot_return_partial_success(self):
        c = self.client([response({"value": [{"id": "a"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/u/messages?page=2"}),
                         response({"error": "PRIVATE BODY"}, status=503)])
        with self.assertRaises(GraphError) as caught: Message(c).find(top=10)
        self.assertEqual(caught.exception.status_code, 503)
        self.assertNotIn("PRIVATE", str(caught.exception))

    def test_attachment_page_error_cannot_return_partial_success(self):
        c = self.client([response({"value": [{"id": "a"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/u/messages/m/attachments?page=2"}),
                         response({}, status=403)])
        with self.assertRaises(GraphError) as caught: Attachment(c).list("m")
        self.assertEqual(caught.exception.status_code, 403)

    def test_nonpaginated_message_error_is_not_empty_success(self):
        c = self.client([response({}, status=403)])
        with self.assertRaises(GraphError): Message(c).find(auto_paginate=False)

    def test_messages_and_attachments_reject_page_cycles(self):
        for kind in (Message, Attachment):
            with self.subTest(kind=kind.__name__):
                path = "messages" if kind is Message else "messages/m/attachments"
                link = "https://graph.microsoft.com/v1.0/users/u/" + path + "?page=2"
                c = self.client([response({"value": [], "@odata.nextLink": link}),
                                 response({"value": [], "@odata.nextLink": link})])
                with self.assertRaises(GraphError) as caught:
                    Message(c).find() if kind is Message else Attachment(c).list("m")
                self.assertEqual(caught.exception.code, "pagination_cycle")
                self.assertEqual(c.api.session.request.call_count, 2)

    def test_default_safety_budget_raises_instead_of_returning_partial(self):
        c = self.client([response({"value": [{"id": str(i)}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/u/messages?page=" + str(i)})
                         for i in range(201)])
        with self.assertRaises(GraphError) as caught: Message(c).find(max_items=300)
        self.assertEqual(caught.exception.code, "mailbox_page_limit")
        self.assertEqual(c.api.session.request.call_count, 200)

    def test_explicit_page_limit_is_an_intentional_subset(self):
        c = self.client([response({"value": [{"id": "a"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/u/messages?page=2"})])
        self.assertEqual(Message(c).find(page_limit=1), [{"id": "a"}])
        self.assertEqual(c.api.session.request.call_count, 1)

    def test_bad_page_envelopes_fail_not_empty_results(self):
        for page in ({}, {"value": {}}, {"value": [None]}, {"value": None}):
            with self.subTest(page=page):
                c = self.client([response(page)])
                with self.assertRaises(GraphError): Attachment(c).list("m")

    def test_malformed_nextlink_is_rejected(self):
        c = self.client([response({"value": [], "@odata.nextLink": 42})])
        with self.assertRaises(GraphError) as caught: Message(c).find()
        self.assertEqual(caught.exception.code, "invalid_mail_continuation")

    def test_delta_410_requires_resync_not_empty_success(self):
        for method in ("messages", "folders"):
            with self.subTest(method=method):
                c = self.client([response({"error": "PRIVATE"}, status=410)])
                with self.assertRaises(GraphError) as caught: getattr(Delta(c), method)()
                self.assertEqual((caught.exception.status_code, caught.exception.code), (410, "delta_resync_required"))
                self.assertNotIn("PRIVATE", str(caught.exception))

    def test_delta_other_errors_keep_status(self):
        c = self.client([response({}, status=403)])
        with self.assertRaises(GraphError) as caught: Delta(c).messages()
        self.assertEqual(caught.exception.status_code, 403)

    def test_delta_preserves_repeated_events_and_opaque_links(self):
        values = [{"id": "a"}, {"id": "a", "@removed": {"reason": "deleted"}}]
        for cursor_key, next_link, delta_link in (
            ("@odata.nextLink", "opaque-next", None),
            ("@odata.deltaLink", None, "opaque-delta"),
        ):
            with self.subTest(cursor=cursor_key):
                c = self.client([response({"value": values, cursor_key: next_link or delta_link})])
                result = Delta(c).messages(delta_link="opaque-input")
                self.assertEqual(result, {"value": values, "next_link": next_link, "delta_link": delta_link})
                self.assertTrue(c.api.session.request.call_args.args[1].endswith("/opaque-input"))

    def test_zero_limit_does_not_fetch(self):
        c = self.client([])
        self.assertEqual(Message(c).find(max_items=0), [])
        self.assertEqual(Attachment(c).list("m", max_items=0), [])
        c.api.session.request.assert_not_called()

    def test_hidden_children_flag_survives_attached_children(self):
        c = self.client([response({"value": [{"id": "f", "displayName": "Folder"}]}), response({"value": []})])
        Folder(c).find(include_hidden=True, include_children=True, recursive=False)
        for call in c.api.session.request.call_args_list:
            self.assertIn("includeHiddenFolders=true", call.args[1])

    def test_folder_name_ambiguity_not_first_match(self):
        c = self.client([response({}, status=404),
            response({"value": [{"id": "a", "displayName": "Archive"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/users/u/mailFolders?page=2"}),
            response({"value": [{"id": "b", "displayName": "Archive"}]}),
            response({"value": []}), response({"value": []})])
        with self.assertRaisesRegex(ValueError, "ambiguous"): Folder(c).resolve("Archive")

    def test_folder_permission_error_does_not_start_name_search(self):
        c = self.client([response({}, status=403)])
        with self.assertRaises(GraphError): Folder(c).resolve("Archive")
        self.assertEqual(c.api.session.request.call_count, 1)

    def test_explicit_folder_id_does_not_need_lookup(self):
        c = self.client([])
        folder = {"id": "f/+#"}
        self.assertIs(Folder(c).resolve(folder), folder)
        c.api.session.request.assert_not_called()

    def test_folder_tree_cycle_is_bounded(self):
        c = self.client([response({"value": [{"id": "a", "displayName": "A"}]}),
                         response({"value": [{"id": "a", "displayName": "A"}]})])
        with self.assertRaises(GraphError) as caught: Folder(c).find()
        self.assertEqual(caught.exception.code, "mailbox_folder_cycle")
        self.assertEqual(c.api.session.request.call_count, 2)

    def test_string_select_normalizes_single_field(self):
        c = self.client([response({"value": []})])
        Message(c).find(select="subject")
        self.assertEqual(parse_qs(urlsplit(c.api.session.request.call_args.args[1]).query)["$select"], ["subject"])

    def test_select_deduplicates_and_validates_before_http(self):
        c = self.client([response({"value": []})])
        Message(c).find(select=["id", "subject", "id"])
        self.assertEqual(parse_qs(urlsplit(c.api.session.request.call_args.args[1]).query)["$select"], ["id,subject"])
        before = c.api.session.request.call_count
        with self.assertRaises(ValueError): Message(c).find(select=[1])
        self.assertEqual(c.api.session.request.call_count, before)

    def test_reserved_message_and_attachment_ids_are_escaped(self):
        c = self.client([response({"id": "m"}), response({"id": "a"})])
        Message(c).get("opaque/id+#?%")
        Attachment(c).get("m/+#", "a/+#")
        self.assertIn("opaque%2Fid%2B%23%3F%25", c.api.session.request.call_args_list[0].args[1])
        self.assertIn("m%2F%2B%23/attachments/a%2F%2B%23", c.api.session.request.call_args_list[1].args[1])

    def test_reply_dispatch_and_folder_delta_ids_are_escaped(self):
        c = self.client([response({}), response({}), response({}),
                         response({"value": [], "@odata.deltaLink": "opaque-delta"})])
        c.tools = SimpleNamespace(recipients=lambda x: [])
        Draft(c).create_reply("m/+#", comment="test")
        Delete(c).delete("m/+#")
        MailMove(c).move("m/+#", {"id": "destination/raw"})
        Delta(c).messages(folder="folder/+#", select="id")
        for call in c.api.session.request.call_args_list[:3]:
            self.assertIn("m%2F%2B%23", call.args[1])
        self.assertEqual(c.api.session.request.call_args_list[2].kwargs["json"]["destinationId"], "destination/raw")
        self.assertIn("folder%2F%2B%23", c.api.session.request.call_args_list[3].args[1])

    def test_ambiguous_attachment_name_is_not_first_match(self):
        c = self.client([response({"value": [{"id": "a", "name": "same.txt"}, {"id": "b", "name": "same.txt"}]})])
        with self.assertRaisesRegex(ValueError, "ambiguous"):
            AttachmentDownload(c)._resolve_attachment("m", attachment_name="same.txt")

    def test_raw_mailbox_error_contract_retains_redacted_status(self):
        c = self.client([response({"error": "PRIVATE"}, status=503)])
        result = c.api.graph_call(resource_url="users/u/messages")
        self.assertEqual(result["status_code"], 503)
        self.assertIn("error", result)
        self.assertNotIn("PRIVATE", json.dumps(result))


class AttachmentProtocol320(MailClientFixture):
    def test_intermediate_ack_and_empty_201_commit(self):
        c = self.client()
        c.api.transfer_session.request.side_effect = [
            response({"nextExpectedRanges": ["327680"]}),
            response(status=201, content=b"")]
        result = AttachmentUpload(c)._upload_large("https://outlook.example/session", b"x"*327681)
        self.assertEqual(result, {"completed": True, "status_code": 201, "size": 327681})
        for call in c.api.transfer_session.request.call_args_list:
            self.assertNotIn("Authorization", call.kwargs["headers"])

    def test_incorrect_intermediate_range_is_not_ignored(self):
        c = self.client()
        c.api.transfer_session.request.return_value = response({"nextExpectedRanges": ["0"]})
        with self.assertRaises(GraphError) as caught:
            AttachmentUpload(c)._upload_large("https://outlook.example/session", b"x"*327681)
        self.assertEqual(caught.exception.code, "unexpected_attachment_acknowledgement")
        self.assertEqual(c.api.transfer_session.request.call_count, 1)

    def test_early_201_does_not_commit_remaining_bytes(self):
        c = self.client()
        c.api.transfer_session.request.return_value = response(status=201, content=b"")
        with self.assertRaises(GraphError):
            AttachmentUpload(c)._upload_large("https://outlook.example/session", b"x"*327681)

    def test_final_progress_response_is_not_completion(self):
        c = self.client()
        c.api.transfer_session.request.return_value = response({"nextExpectedRanges": ["0"]})
        with self.assertRaises(GraphError) as caught:
            AttachmentUpload(c)._upload_large("https://outlook.example/session", b"x")
        self.assertEqual(caught.exception.code, "attachment_not_committed")

    def test_201_with_incomplete_body_is_not_completion(self):
        c = self.client()
        c.api.transfer_session.request.return_value = response({"nextExpectedRanges": ["0"]}, status=201)
        with self.assertRaises(GraphError):
            AttachmentUpload(c)._upload_large("https://outlook.example/session", b"x")

    def test_commit_metadata_size_must_match(self):
        c = self.client()
        c.api.transfer_session.request.return_value = response({"id": "a", "size": 5}, status=201)
        with self.assertRaises(GraphError) as caught:
            AttachmentUpload(c)._upload_large("https://outlook.example/session", b"x")
        self.assertEqual(caught.exception.code, "attachment_size_mismatch")

    def test_lost_put_is_not_blindly_replayed(self):
        import requests
        c = self.client()
        c.api.transfer_session.request.side_effect = requests.ConnectionError("simulated")
        with self.assertRaises(GraphError):
            AttachmentUpload(c)._upload_large("https://outlook.example/session", b"x")
        self.assertEqual(c.api.transfer_session.request.call_count, 1)

    def test_invalid_upload_session_fails_explicitly(self):
        c = self.client([response({})])
        with self.assertRaises(GraphError) as caught: AttachmentUpload(c).upload("m", b"x"*(3*1024*1024))
        self.assertEqual(caught.exception.code, "invalid_attachment_session")
        c.api.transfer_session.request.assert_not_called()

    def test_response_metadata_omits_location_and_closes_response(self):
        c = self.client()
        r = response(status=201, content=b"", headers={"Location": "https://outlook.example/PRIVATE"})
        c.api.transfer_session.request.return_value = r
        result = c.api.graph_call(resource_url="https://outlook.example/session", preauthenticated=True,
                                 method="PUT", data=b"x", response_metadata=True)
        self.assertEqual(result, {"status_code": 201, "body": {}})
        self.assertNotIn("PRIVATE", json.dumps(result))
        r.close.assert_called_once()

    def test_metadata_cannot_be_combined_with_stream_or_batch(self):
        c = self.client()
        with self.assertRaises(ValueError): c.api.graph_call(resource_url="x", stream=True, response_metadata=True)
        with self.assertRaises(ValueError): c.api.graph_call(batch_requests=[], response_metadata=True)
        c.api.session.request.assert_not_called()


class MoveAndUtilities320(Offline):
    def move_client(self, result):
        return SimpleNamespace(drive_id="d", api=SimpleNamespace(graph_call=Mock(return_value=result)),
            search_by=SimpleNamespace(filefoldername=Mock(return_value=[
                {"id": "f", "name": "file.txt", "file": {}}, {"id": "target", "folder": {}}])))

    def test_move_requires_correct_parent_and_drive(self):
        for result in ({"id": "f", "name": "file.txt"},
                       {"id": "f", "name": "file.txt", "parentReference": {"id": "wrong"}},
                       {"id": "f", "name": "file.txt", "parentReference": {"id": "target", "driveId": "other"}}):
            with self.subTest(result=result):
                c = self.move_client(result)
                with self.assertRaises(RuntimeError): Move(c).move("file.txt", "Target", None)
                self.assertEqual(c.api.graph_call.call_count, 1)

    def test_move_valid_parent_still_succeeds(self):
        c = self.move_client({"id": "f", "name": "file.txt", "parentReference": {"id": "target", "driveId": "d"}})
        self.assertEqual(Move(c).move("file.txt", "Target", None)["parentReference"]["id"], "target")

    def test_sets_empty_single_ambiguous_and_ordered_compatibility(self):
        self.assertIsNone(lists.ensure_single_item(set()))
        self.assertEqual(lists.ensure_single_item({"only"}), "only")
        self.assertEqual(lists.ensure_single_item(frozenset({"only"})), "only")
        with self.assertRaises(ValueError): lists.ensure_single_item({"one", "two"})
        self.assertEqual(lists.ensure_single_item([2, 1]), 2)
        self.assertEqual(lists.ensure_single_item((2, 1)), 2)

    def test_extrema_preserves_offsets_and_original_fields(self):
        rows = [{"id": "later", "createdDateTime": "2026-11-17T10:00:00-05:00"},
                {"id": "earlier", "createdDateTime": "2026-11-17T14:30:00Z"}]
        original = copy.deepcopy(rows)
        self.assertEqual(transforms.date_key_extrema(rows, "createdDateTime")[0], original[0])
        self.assertEqual(transforms.date_key_extrema(rows, "createdDateTime", maxima=False)[0], original[1])
        self.assertEqual(rows, original)

    def test_extrema_naive_policy_not_host_timezone(self):
        rows = [{"id": "naive", "when": "2026-11-17 10:00:00"},
                {"id": "aware", "when": "2026-11-17T12:00:00Z"}]
        self.assertEqual(transforms.date_key_extrema(rows, "when")[0]["id"], "aware")
        self.assertEqual(transforms.date_key_extrema(rows, "when", naive_timezone=timezone(timedelta(hours=-5)))[0]["id"], "naive")

    def test_extrema_custom_formats_invalid_values_and_stable_ties(self):
        rows = [{"id": "a", "when": "17/11/2026 10:00 -0500"}, {"id": "b", "when": "17/11/2026 15:00 +0000"},
                {"id": "bad", "when": "not-a-date"}]
        result = transforms.date_key_extrema(rows, "when", num_items=2, from_format="%d/%m/%Y %H:%M %z")
        self.assertEqual([r["id"] for r in result], ["a", "b"])
        self.assertEqual(result[0]["when"], rows[0]["when"])

    def test_extrema_counts_validate_and_zero_is_empty(self):
        for value in (True, -1, 1.5):
            with self.assertRaises(ValueError): transforms.date_key_extrema([], "when", num_items=value)
        self.assertEqual(transforms.date_key_extrema([], "when", num_items=0), [])

    def test_naive_policy_cannot_fall_back_to_the_host_timezone(self):
        class MissingOffset(tzinfo):
            def utcoffset(self, dt): return None
        with self.assertRaises(ValueError):
            transforms.date_key_extrema([], "when", naive_timezone=MissingOffset())

    def test_mail_id_validation(self):
        for value in (None, "", 1, "a\nb"):
            with self.assertRaises(ValueError): mail_id(value)
        self.assertEqual(mail_id("opaque/id+#?%"), "opaque%2Fid%2B%23%3F%25")


if __name__ == "__main__":
    unittest.main()
