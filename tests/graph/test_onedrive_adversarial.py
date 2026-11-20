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
"""Boundary attacks, auth-mode controls and stress cases for the OneDrive facade."""
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
import io
import os
from pathlib import Path
import tempfile
from types import SimpleNamespace as NS
from unittest.mock import Mock, patch
from urllib.parse import quote

import requests
from azauthlib import AuthenticationError, InteractionRequiredError

from m365.graph import Drive, GraphError
from test_onedrive import BASE, CHUNK, DRIVE, ROOT, DriveFixture, file_item, folder, page
from test_sharepoint import response
from test_transfer_hardening import binary


class OneDriveBoundaries(DriveFixture):
    def test_non_provider_objects_fail_clearly_without_http(self):
        for provider in (False, 0, 42, "raw-token", {}, [], object()):
            with self.subTest(provider_type=type(provider).__name__):
                with patch("m365.graph._storage.GraphAPI") as factory:
                    from m365.graph._graph import GraphAPI
                    created = []
                    def make(parent):
                        api = GraphAPI(parent)
                        api.session.request = Mock(side_effect=AssertionError("Unexpected HTTP"))
                        created.append(api)
                        return api
                    factory.side_effect = make
                    with self.assertRaises(ValueError):
                        Drive(provider, DRIVE, folder_id=ROOT)
                    for api in created:
                        api.session.request.assert_not_called()

    def test_exact_255_utf16_name_boundary_remains_valid(self):
        for name in ("a"*255, "😀"*127+"a"):
            with self.subTest(name_length=len(name)):
                client = self.client([response(folder()), response(file_item(name=name))])
                self.assertEqual(client.get_item(name)["name"], name)

    def test_opaque_query_bytes_are_not_decoded_with_continuation_path(self):
        link = BASE + "/drives/drive/items/anchor/children?$skiptoken=A%FF%2B%3D"
        client = self.client([response(folder()), page(link=link), page(file_item())])
        self.assertEqual(len(client.list_children()), 1)
        self.assertEqual(self.calls()[-1].args[1], link)

    def test_invalid_configuration_fails_before_transport_creation(self):
        variants = [
            {"drive_id": None}, {"drive_id": ""}, {"drive_id": " "}, {"drive_id": True},
            {"drive_id": "\ud800"}, {"folder_id": ""}, {"folder_id": False},
            {"folder_id": "\x00"}, {"folder_id": []}, {"folder_id": {}},
            {"max_requests": 0}, {"max_requests": True}, {"max_items": -1},
            {"max_items": False}, {"max_depth": -1}, {"max_depth": True},
            {"max_retries": -1}, {"max_retries": 11}, {"max_retries": False},
            {"request_timeout": None}, {"request_timeout": 0}, {"request_timeout": float("nan")},
            {"request_timeout": True}, {"request_timeout": (1,)}, {"request_timeout": (1, 2, 3)},
            {"request_timeout": (1, 0)}, {"request_timeout": (1, float("inf"))},
            {"graph_url": "http://graph.microsoft.com/v1.0"},
            {"graph_url": "https://user:secret@graph.microsoft.com/v1.0"},
            {"graph_url": "https://graph.microsoft.com/v1.0?secret=x"},
            {"graph_url": "https://graph.microsoft.com/v1.0#fragment"},
            {"graph_url": "https://graph.microsoft.com:bad/v1.0"},
            {"graph_url": "https://graph.microsoft.com"},
            {"graph_url": None},
        ]
        for values in variants:
            options = dict(drive_id=DRIVE, folder_id=ROOT)
            options.update(values)
            with self.subTest(values=values), patch("m365.graph._storage.GraphAPI") as factory:
                with self.assertRaises((ValueError, TypeError)):
                    Drive(NS(access_token="unused"), **options)
                factory.assert_not_called()

    def test_missing_provider_fails_before_any_request(self):
        with patch("m365.graph._storage.GraphAPI") as factory, self.assertRaises(ValueError):
            Drive(None, DRIVE, folder_id=ROOT)
        factory.assert_not_called()

    def test_invalid_provider_token_is_not_sent_or_leaked(self):
        for token in (None, False, 42, "", " ", "SYNTHETIC-SECRET\ninjected", "secret\x7f"):
            with self.subTest(token=token):
                with self.assertRaises(GraphError) as error:
                    self.client(token_engine=NS(access_token=token))
                self.assertNotIn("SYNTHETIC-SECRET", str(error.exception))
                self.created_apis[-1].session.request.assert_not_called()

    def test_folder_permission_denial_never_falls_back_to_a_broader_root(self):
        client = self.client()
        self.api.session.request.side_effect = [response({}, status=403)]
        with self.assertRaises(GraphError) as error:
            client.list_children()
        self.assertEqual(error.exception.status_code, 403)
        self.assertEqual(len(self.calls()), 2)
        self.assertFalse(any(p.endswith("/root") or "/sites/" in p for p in self.paths()))

    def test_constructor_rejects_wrong_kind_shortcut_or_foreign_drive(self):
        variants = [file_item(ROOT), folder("other"), folder(remoteItem={}),
                    folder(parentReference={"driveId": "other"}),
                    folder(file={}), {"id": ROOT, "folder": {}},
                    folder(deleted={}), folder(error={}), folder(name=None)]
        for value in variants:
            with self.subTest(value=value), self.assertRaises(GraphError):
                self.client(root=value)

    def test_paths_cannot_escape_or_silently_widen_scope(self):
        client = self.client()
        bad = ["../outside", "a/../b", ".", "..", "a/.", "a//b", "/outside",
               "a/", "\\outside", "C:\\outside", "https://other.example/x", "a\x00b",
               "a\nb", "a\x7fb", "\ud800", [], {}, False, None, "x"*256, "😀"*128]
        for value in bad:
            actions = [
                lambda v: client.get_item(v), lambda v: client.list_children(v),
                lambda v: client.find_files("x", folder_path=v),
                lambda v: client.find_folders("x", folder_path=v),
                lambda v: client.upload(io.BytesIO(b"x"), v, "x.txt"),
                lambda v: client.download_file(v, "unused-local-directory"),
                lambda v: client.download_folder(v, "unused-local-directory"),
                lambda v: client.create_folder("New", v),
            ]
            for action in actions:
                with self.subTest(path=repr(value), action=action), self.assertRaises((ValueError, TypeError)):
                    action(value)
                self.assertEqual(len(self.calls()), 1)

    def test_encoded_looking_name_is_literal_not_decoded_into_traversal(self):
        sub = folder("sub", "%2E%2E", ROOT)
        client = self.client([response(folder()), response(sub)])
        self.assertEqual(client.get_item("%2E%2E")["id"], "sub")
        self.assertTrue(self.calls()[-1].args[1].endswith("/anchor:/%252E%252E"))

    def test_unicode_quotes_spaces_percent_and_hash_names_encode_once(self):
        for name in ("O'Brien #1.txt", "100%.txt", "café 🧪.txt", "a+b & c.txt", ".gitignore"):
            with self.subTest(name=name):
                client = self.client([response(folder()), response(file_item(name=name))])
                self.assertEqual(client.get_item(name)["name"], name)
                self.assertTrue(self.calls()[-1].args[1].endswith("anchor:/"+quote(name, safe="")))

    def test_opaque_drive_and_anchor_ids_are_single_encoded_segments(self):
        identifier = "opaque/id+#?%"
        root = folder(identifier)
        client = self.client(root=root, folder_id=identifier, drive_id="drive/id+#?%")
        self.assertTrue(self.calls()[0].args[1].endswith(
            "/drives/drive%2Fid%2B%23%3F%25/items/opaque%2Fid%2B%23%3F%25"))

    def test_current_metadata_cannot_redirect_the_bound_root(self):
        client = self.client([response(folder("outside"))])
        with self.assertRaises(GraphError) as error:
            client.list_children()
        self.assertEqual(error.exception.code, "onedrive_root_id_mismatch")
        self.assertEqual(len(self.calls()), 2)

    def test_path_lookup_rejects_wrong_name_parent_drive_and_cycles(self):
        for row in (file_item(name="other.txt"), file_item(parent="outside"),
                    file_item(parentReference={}), file_item(parentReference={"id": ROOT, "driveId": "other"}),
                    file_item(identifier=ROOT), file_item(remoteItem={}), file_item(error={}),
                    file_item(deleted={}), file_item(folder={})):
            with self.subTest(row=row):
                client = self.client([response(folder()), response(row)])
                with self.assertRaises(GraphError):
                    client.get_item("report.txt")

    def test_parent_or_name_malformed_metadata_never_becomes_success(self):
        for row in (file_item(parentReference=None), file_item(parentReference=[]),
                    file_item(name=None), file_item(name="../escape"),
                    file_item(identifier="\x00"), file_item(identifier=[]),
                    file_item(size=True), file_item(size=-1), file_item(size="3")):
            with self.subTest(row=row):
                client = self.client([response(folder()), page(row)])
                with self.assertRaises(GraphError):
                    client.list_children()

    def test_child_marker_facets_and_parent_context_are_checked_for_every_item(self):
        for row in (file_item(remoteItem={}), file_item(deleted={}), file_item(error={}),
                    file_item(parent="outside"), file_item(identifier=ROOT),
                    file_item(file=None), file_item(folder={}), {"id": "f", "name": "no-facet"}):
            with self.subTest(row=row):
                client = self.client([response(folder()), page(file_item("good", "good.txt"), row)])
                with self.assertRaises(GraphError):
                    client.list_children()

    def test_duplicate_ids_and_casefold_name_collisions_are_not_silently_deduplicated(self):
        for rows in ([file_item(), file_item()],
                     [file_item("a", "REPORT.txt"), file_item("b", "report.TXT")]):
            client = self.client([response(folder()), page(*rows)])
            with self.assertRaises(GraphError):
                client.list_children()

    def test_same_drive_different_folder_continuation_is_not_fetched(self):
        for link in (BASE + "/drives/drive/items/outside/children?token=x",
                     BASE + "/drives/other/items/anchor/children?token=x",
                     BASE + "/drives/drive/root/children?token=x",
                     BASE + "/users/other/messages?token=x", "https://evil.invalid/anything"):
            with self.subTest(link=link):
                client = self.client([response(folder()), page(file_item(), link=link)])
                with self.assertRaises((GraphError, ValueError)):
                    client.list_children()
                self.assertEqual(len(self.calls()), 3)

    def test_bad_nextlink_values_are_not_treated_as_end_of_data(self):
        for value in (None, "", " ", False, 0, [], {}, BASE+"/drives/drive/items/anchor/children\n"):
            client = self.client([response(folder()), response({"value": [], "@odata.nextLink": value})])
            with self.subTest(value=value), self.assertRaises(GraphError):
                client.list_children()

    def test_paging_cycles_and_later_http_failures_cannot_return_partial_lists(self):
        link = BASE + "/drives/drive/items/anchor/children?page=2"
        for last in (page(link=link), response({}, status=503)):
            client = self.client([response(folder()), page(file_item(), link=link), last])
            with self.assertRaises(GraphError):
                client.list_children()
            self.assertEqual(len(self.calls()), 4)

    def test_get_and_list_reject_incomplete_success_statuses(self):
        for status in (201, 202, 204, 205, 206, 299):
            with self.subTest(status=status):
                client = self.client([response(folder(), status=status)])
                with self.assertRaises(GraphError):
                    client.get_item()
                client = self.client([response(folder()), response({"value": []}, status=status)])
                with self.assertRaises(GraphError):
                    client.list_children()

    def test_malformed_collection_envelopes_raise_redacted_errors(self):
        for body in (None, [], {}, {"value": None}, {"value": {}}, {"value": [None]},
                     {"value": [], "error": {"message": "SYNTHETIC-SECRET"}}):
            client = self.client([response(folder()), response(body)])
            with self.subTest(body=body), self.assertRaises(GraphError) as error:
                client.list_children()
            self.assertNotIn("SYNTHETIC-SECRET", str(error.exception))

    def test_request_budget_stops_before_next_metadata_request(self):
        client = self.client([response(folder())], max_requests=1)
        with self.assertRaises(GraphError) as error:
            client.list_children()
        self.assertEqual(error.exception.code, "onedrive_request_limit")
        self.assertEqual(len(self.calls()), 2)

    def test_item_budget_is_shared_across_pages_and_tree(self):
        client = self.client([response(folder()), page(file_item("a", "a"), file_item("b", "b"))], max_items=1)
        with self.assertRaises(GraphError) as error:
            client.find_files("a")
        self.assertEqual(error.exception.code, "onedrive_item_limit")

    def test_depth_limit_and_ancestor_cycles_fail_before_download(self):
        sub = folder("sub", "Sub", ROOT)
        for replies in ([response(folder()), page(sub)],):
            client = self.client(replies, max_depth=0)
            with tempfile.TemporaryDirectory() as directory:
                with self.assertRaises(GraphError) as error:
                    client.download_folder("", directory)
                self.assertEqual(error.exception.code, "onedrive_depth_limit")
                self.assertEqual(list(Path(directory).iterdir()), [])
        ancestor = folder(ROOT, "Ancestor", "sub")
        client = self.client([response(folder()), page(sub), page(ancestor)])
        with self.assertRaises(GraphError):
            client.find_files("report")

    def test_non_folder_paths_are_not_traversed_as_folders(self):
        for action in (lambda c: c.list_children("report.txt"), lambda c: c.get_item("report.txt/child"),
                       lambda c: c.upload(io.BytesIO(b"x"), "report.txt", "x"),
                       lambda c: c.find_files("x", folder_path="report.txt"),
                       lambda c: c.create_folder("New", "report.txt")):
            client = self.client([response(folder()), response(file_item())])
            with self.assertRaises(ValueError):
                action(client)

    def test_invalid_search_controls_do_not_make_requests(self):
        client = self.client()
        for values in ({"term": ""}, {"term": None}, {"term": "x\n"}, {"term": "\ud800"},
                       {"term": "x", "limit": -1}, {"term": "x", "limit": True},
                       {"term": "x", "limit": 1.5}, {"term": "x", "exact_match": "false"},
                       {"term": "x", "recursive": 1}):
            with self.subTest(values=values), self.assertRaises(ValueError):
                client.find_files(**values)
            self.assertEqual(len(self.calls()), 1)

    def test_invalid_remote_names_do_not_create_folders_or_uploads(self):
        client = self.client()
        for name in ("", ".", "..", "a/b", "a\\b", "CON.txt", "NUL", "LPT¹.txt", "x.", "x ", "a\x00"):
            with self.subTest(name=name):
                with self.assertRaises(ValueError):
                    client.create_folder(name)
                with self.assertRaises(ValueError):
                    client.upload(io.BytesIO(b"x"), new_file_name=name)
                self.assertEqual(len(self.calls()), 1)

    def test_invalid_upload_protocol_controls_fail_before_reads(self):
        client = self.client()
        for values in ({"chunk_size": True}, {"chunk_size": 1}, {"chunk_size": 60*1024*1024},
                       {"conflict_behavior": "overwrite"}, {"conflict_behavior": None}):
            with self.subTest(values=values), self.assertRaises(ValueError):
                client.upload(io.BytesIO(b"abc"), new_file_name="x", **values)
            self.assertEqual(len(self.calls()), 1)

    def test_invalid_download_controls_fail_without_file_work(self):
        client = self.client()
        for values in ({"max_workers": True}, {"max_workers": 0}, {"max_workers": 1.5},
                       {"compress": "false"}):
            with self.subTest(values=values), self.assertRaises(ValueError):
                client.download_folder("", "unused-local-directory", **values)
            self.assertEqual(len(self.calls()), 1)


class OneDriveHardTransfers(DriveFixture):
    def test_renamed_upload_confirmation_must_remain_a_safe_readable_item(self):
        for row in (file_item(name="CON"), file_item(name="report. "),
                    file_item(name="x"*256), file_item(name="report\x7f.txt"),
                    file_item(identifier="bad\nidentity"), file_item(identifier="bad\ud800identity")):
            client = self.client([response(folder()), response({"uploadUrl": "https://upload.invalid/x"})],
                                 transfers=[response(row, status=201)])
            with self.subTest(row=row), self.assertRaises(GraphError):
                client.upload(io.BytesIO(b"abc"), new_file_name="report.txt", conflict_behavior="rename")

    def test_invalid_session_urls_never_receive_graph_bearer(self):
        for url in ("http://upload.invalid/x", "https://user:secret@upload.invalid/x",
                    "//upload.invalid/x", "https://upload.invalid/x#fragment"):
            client = self.client([response(folder()), response({"uploadUrl": url})])
            with self.subTest(url=url), self.assertRaises(ValueError):
                client.upload(io.BytesIO(b"abc"), new_file_name="report.txt")
            self.api.transfer_session.request.assert_not_called()

    def test_chunk_acknowledgements_cannot_skip_bytes_or_claim_early_commit(self):
        for body, status in (({"nextExpectedRanges": ["0-"]}, 202),
                             ({"nextExpectedRanges": [str(CHUNK*2)+"-"]}, 202),
                             (file_item(size=CHUNK+1), 201),
                             ({"id": "f", "nextExpectedRanges": [str(CHUNK)+"-"]}, 202)):
            client = self.client([response(folder()), response({"uploadUrl": "https://upload.invalid/x"})],
                                 transfers=[response(body, status=status)])
            with self.subTest(body=body, status=status), self.assertRaises(GraphError):
                client.upload(io.BytesIO(b"x"*(CHUNK+1)), new_file_name="report.txt", chunk_size=CHUNK)
            self.assertEqual(self.api.transfer_session.request.call_count, 1)

    def test_unknown_upload_commit_is_not_replayed_as_a_success(self):
        client = self.client([response(folder()), response({"uploadUrl": "https://upload.invalid/x"})],
                             transfers=[requests.Timeout("SYNTHETIC-SECRET"), response({}, status=404)])
        with self.assertRaises(GraphError) as error:
            client.upload(io.BytesIO(b"abc"), new_file_name="report.txt")
        self.assertEqual(error.exception.code, "upload_completion_unknown")
        self.assertNotIn("SYNTHETIC-SECRET", str(error.exception))
        self.assertEqual([c.args[0] for c in self.api.transfer_session.request.call_args_list], ["PUT", "GET"])

    def test_commit_cannot_reuse_the_parent_or_anchor_identity(self):
        client = self.client([response(folder()), response({"uploadUrl": "https://upload.invalid/x"})],
                             transfers=[response(file_item(identifier=ROOT), status=201)])
        with self.assertRaises(GraphError):
            client.upload(io.BytesIO(b"abc"), new_file_name="report.txt")

    def test_new_nested_folder_cannot_reuse_the_anchor_id(self):
        sub = folder("sub", "Sub", ROOT)
        client = self.client([response(folder()), response(sub), response(folder(ROOT, "New", "sub"), status=201)])
        with self.assertRaises(GraphError):
            client.create_folder("New", "Sub")

    def test_create_folder_conflicting_confirmation_is_not_success(self):
        for body, status in ((folder("new", "Wrong", ROOT), 201), (folder("new", "New", "outside"), 201),
                             (file_item(name="New"), 201), (folder("new", "New", ROOT), 200),
                             (folder(ROOT, "New", ROOT), 201), (folder("new", "New", ROOT), 202)):
            client = self.client([response(folder()), response(body, status=status)])
            with self.subTest(body=body, status=status), self.assertRaises(GraphError):
                client.create_folder("New")
            self.assertEqual([c.args[0] for c in self.calls()], ["GET", "GET", "POST"])

    def test_redirected_download_strips_bearer_and_rejects_partial_content(self):
        hop = response({}, status=302, headers={"Location": "https://download.invalid/x?secret=SYNTHETIC"})
        partial = binary(status=206)
        client = self.client([response(folder()), response(file_item()), hop], transfers=[partial])
        with tempfile.TemporaryDirectory() as directory, self.assertRaises(GraphError):
            client.download_file("report.txt", directory)
        self.assertFalse(any(k.lower() in ("authorization", "cookie") for k in
                             self.api.transfer_session.request.call_args.kwargs["headers"]))
        hop.close.assert_called_once()
        partial.close.assert_called_once()

    def test_stream_interruption_leaves_no_partial_file_and_preserves_neighbors(self):
        broken = binary()
        def chunks(**kwargs):
            yield b"a"
            raise requests.ConnectionError("SYNTHETIC-SECRET")
        broken.iter_content = chunks
        client = self.client([response(folder()), response(file_item()), broken])
        with tempfile.TemporaryDirectory() as directory:
            neighbor = Path(directory)/"neighbor.txt"
            neighbor.write_bytes(b"KEEP")
            with self.assertRaises(GraphError) as error:
                client.download_file("report.txt", directory)
            self.assertNotIn("SYNTHETIC-SECRET", str(error.exception))
            self.assertEqual(neighbor.read_bytes(), b"KEEP")
            self.assertEqual([p.name for p in Path(directory).iterdir()], ["neighbor.txt"])
        broken.close.assert_called_once()

    def test_concurrent_local_downloads_have_one_winner_and_no_corrupt_output(self):
        client = self.client()
        def serve(method, url, **kwargs):
            if url.endswith("/items/anchor"):
                return response(folder())
            if url.endswith("/items/anchor:/report.txt"):
                return response(file_item())
            if url.endswith("/items/f/content"):
                return binary(data=b"abc")
            raise AssertionError("Unexpected endpoint")
        self.api.session.request.side_effect = serve
        with tempfile.TemporaryDirectory() as directory:
            def run(_):
                try:
                    return client.download_file("report.txt", directory)
                except FileExistsError:
                    return None
            with ThreadPoolExecutor(max_workers=16) as pool:
                values = list(pool.map(run, range(32)))
            self.assertEqual(sum(v is not None for v in values), 1)
            self.assertEqual((Path(directory)/"report.txt").read_bytes(), b"abc")
            self.assertEqual([p.name for p in Path(directory).iterdir()], ["report.txt"])

    def test_512_pages_stay_ordered_and_opaque(self):
        links = [BASE + "/drives/drive/items/anchor/children?$skiptoken=page%2B"+str(i) for i in range(1, 512)]
        replies = [response(folder())]
        replies.extend(page(file_item(str(i), str(i)+".txt"), link=links[i]) if i < 511
                       else page(file_item(str(i), str(i)+".txt")) for i in range(512))
        client = self.client(replies, max_requests=520)
        rows = client.list_children()
        self.assertEqual([r["id"] for r in rows], [str(i) for i in range(512)])
        self.assertEqual([c.args[1] for c in self.calls()[3:]], links)

    def test_large_page_hits_exact_item_limit_and_rejects_one_more(self):
        rows = [file_item(str(i), str(i)+".txt") for i in range(5000)]
        client = self.client([response(folder()), page(*rows)])
        self.assertEqual(len(client.list_children()), 5000)
        client = self.client([response(folder()), page(*rows, file_item("extra", "extra.txt"))])
        with self.assertRaises(GraphError) as error:
            client.list_children()
        self.assertEqual(error.exception.code, "onedrive_item_limit")


class OneDriveProfiles(DriveFixture):
    def test_explicit_custom_cloud_scopes_and_endpoint_are_preserved(self):
        client, auth, _, authentication = self.from_profile(
            "client_certificate", graph_url="https://graph.microsoft.us/v1.0",
            scopes=["https://graph.microsoft.us/.default"])
        self.assertEqual(authentication.call_args.kwargs["scopes"], ["https://graph.microsoft.us/.default"])
        self.assertTrue(client._api.session.request.call_args.args[1].startswith("https://graph.microsoft.us/v1.0/"))

    def from_profile(self, method, *, interactive=False, silent_error=None, **kwargs):
        self.client()
        auth = Mock()
        auth.access_token = "SYNTHETIC-PROFILE-TOKEN"
        if silent_error is not None:
            auth.Silent.side_effect = silent_error
        with patch.dict(os.environ, {}, clear=True), \
                patch("azauthlib.credentials.load_public_configuration", return_value={"auth_method": method}) as config, \
                patch("azauthlib.Authentication", return_value=auth) as authentication, \
                patch("m365.graph._storage.GraphAPI", self.factory):
            client = Drive.from_profile("Saved Connection", DRIVE, folder_id=ROOT,
                                           allow_interactive=interactive, **kwargs)
        self.addCleanup(client.close)
        return client, auth, config, authentication

    def test_certificate_and_secret_profiles_use_application_auth(self):
        for method in ("client_certificate", "client_credentials"):
            client, auth, config, authentication = self.from_profile(method)
            auth.Build.Default.assert_called_once_with()
            auth.ClientCredentials.assert_called_once_with()
            auth.Silent.assert_not_called()
            auth.Interactive.assert_not_called()
            config.assert_called_once_with(profile="Saved Connection")
            self.assertEqual(client.folder_id, ROOT)
            self.assertEqual(authentication.call_args.kwargs["scopes"], ["https://graph.microsoft.com/.default"])

    def test_cached_user_profiles_do_not_open_ui(self):
        for method in ("interactive", "device_code"):
            client, auth, _, _ = self.from_profile(method)
            auth.Silent.assert_called_once_with()
            auth.ClientCredentials.assert_not_called()
            auth.Interactive.assert_not_called()
            auth.DeviceCodeFlow.assert_not_called()

    def test_user_interaction_requires_explicit_opt_in(self):
        for method in ("interactive", "device_code"):
            with self.subTest(method=method), self.assertRaises(InteractionRequiredError):
                self.from_profile(method, silent_error=InteractionRequiredError("Sign-in needed"))

    def test_explicit_opt_in_uses_the_saved_user_flow_only(self):
        for method in ("interactive", "device_code"):
            _, auth, _, _ = self.from_profile(method, interactive=True,
                                              silent_error=InteractionRequiredError("Sign-in needed"))
            (auth.Interactive if method == "interactive" else auth.DeviceCodeFlow).assert_called_once_with()
            (auth.DeviceCodeFlow if method == "interactive" else auth.Interactive).assert_not_called()

    def test_other_auth_failures_do_not_trigger_interactive_fallback(self):
        with self.assertRaises(AuthenticationError):
            self.from_profile("interactive", interactive=True, silent_error=AuthenticationError("Blocked"))

    def test_invalid_profile_arguments_fail_before_loading_or_authenticating(self):
        variants = [
            {"profile": ""}, {"profile": False}, {"drive_id": ""},
            {"folder_id": ""}, {"allow_interactive": "false"}, {"max_requests": 0},
            {"unexpected_option": True}, {"scopes": []}, {"scopes": [None]}, {"scopes": " "},
            {"graph_url": "https://graph.microsoft.us/v1.0"},
        ]
        for options in variants:
            values = dict(profile="Saved Connection", drive_id=DRIVE, folder_id=ROOT)
            values.update(options)
            with self.subTest(options=options), patch.dict(os.environ, {}, clear=True), \
                    patch("azauthlib.credentials.load_public_configuration") as config, \
                    patch("azauthlib.Authentication") as authentication:
                with self.assertRaises((ValueError, TypeError)):
                    Drive.from_profile(**values)
                config.assert_not_called()
                authentication.assert_not_called()

    def test_environment_credentials_cannot_silently_override_a_profile(self):
        for key in ("CLIENT_ID", "TENANT_ID", "CLIENT_SECRET", "AZURE_AUTHORITY"):
            with self.subTest(key=key), patch.dict(os.environ, {key: ""}, clear=True), \
                    patch("azauthlib.credentials.load_public_configuration") as config, \
                    patch("azauthlib.Authentication") as authentication:
                with self.assertRaises(ValueError):
                    Drive.from_profile("Saved Connection", DRIVE, folder_id=ROOT)
                config.assert_not_called()
                authentication.assert_not_called()
