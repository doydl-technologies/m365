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
"""Stateful storage parity and adversarial tests; all HTTP is synthetic."""
from copy import deepcopy
import hashlib
import inspect
import io
import os
from pathlib import Path
import tempfile
from types import SimpleNamespace as NS
import unittest
from unittest.mock import Mock, patch
from urllib.parse import parse_qs, unquote, urlsplit
import zipfile

import requests
from azauthlib import AuthenticationError, InteractionRequiredError
from m365.graph import Site, Drive, GraphError
from m365.graph._storage import _StorageClient
from test_onedrive import DRIVE, ROOT, folder, file_item, page, response
from test_onedrive_model import MemoryDrive
from test_sharepoint import Offline


SITE_URL = "https://fixture.invalid/sites/Test"
METHODS = ("get_item", "list_children", "find_files", "find_folders", "find_file_folder",
           "upload", "download_file", "download_folder", "create_folder", "move", "rename")


class StorageService(MemoryDrive):
    def __init__(self):
        super().__init__()
        self.patch_override = None
        self.before_patch = None
        self.denied_root = False
        self.nodes[ROOT]["eTag"] = '"root"'

    def add(self, name, parent=ROOT, data=None):
        identifier = super().add(name, parent, data)
        self.nodes[identifier]["eTag"] = '"v1"'
        return identifier

    def graph(self, method, url, **kwargs):
        path = urlsplit(url).path
        if method == "GET" and path in ("/v1.0/sites/site", "/v1.0/sites/fixture.invalid:/sites/Test"):
            self.records.append((method, url))
            return response({"id": "site", "webUrl": SITE_URL})
        if method == "GET" and path in ("/v1.0/sites/site/drive", "/v1.0/sites/site/drives"):
            self.records.append((method, url))
            data = {"id": DRIVE, "name": "Documents", "webUrl": SITE_URL+"/Documents"}
            return page(data, {"id": "other", "name": "Other"}) if path.endswith("/drives") else response(data)
        if method == "GET" and path == "/v1.0/drives/drive/root":
            self.records.append((method, url))
            return response({}, status=403) if self.denied_root else response(self.metadata(ROOT))
        if method == "GET" and path in ("/v1.0/drives/other/root", "/v1.0/drives/other/items/other-root"):
            self.records.append((method, url))
            return response({"id": "other-root", "name": "Other", "folder": {}})
        if method == "GET" and path == "/v1.0/drives/other/items/other-root/children":
            self.records.append((method, url))
            return page()
        if method == "PATCH":
            self.records.append((method, url))
            assert path.startswith("/v1.0/drives/drive/items/")
            assert parse_qs(urlsplit(url).query) == {"@microsoft.graph.conflictBehavior": ["fail"]}
            identifier = unquote(path.rsplit("/", 1)[1], errors="strict")
            if self.before_patch:
                self.before_patch(self, identifier, kwargs)
            node = self.nodes[identifier]
            if kwargs["headers"].get("If-Match") != node["eTag"]:
                return response({}, status=412)
            body = kwargs["json"]
            parent, name = body["parentReference"]["id"], body["name"]
            if self.lookup(parent, name) not in (None, identifier):
                return response({}, status=409)
            node["parentReference"] = {"id": parent, "driveId": DRIVE}
            node["name"], node["eTag"] = name, '"v2"'
            result = response(self.metadata(identifier))
            return self.patch_override(result) if self.patch_override else result
        return super().graph(method, url, **kwargs)


class StorageFixture(Offline):
    def setUp(self):
        super().setUp()
        self.service = StorageService()
        self.sessions = []
        self.transport = patch("m365.graph._graph.requests.Session", side_effect=self.session)
        self.transport.start()
        self.addCleanup(self.transport.stop)

    def session(self):
        session = requests.sessions.Session()
        def request(method, url, **kwargs):
            handler = self.service.transfer if urlsplit(url).hostname == "upload.invalid" else self.service.graph
            return handler(method, url, **kwargs)
        session.request = Mock(side_effect=request)
        session.close = Mock(wraps=session.close)
        self.sessions.append(session)
        return session

    def client(self, kind, **options):
        provider = NS(access_token="SYNTHETIC-STORAGE-TOKEN")
        if kind == "drive":
            client = Drive(provider, DRIVE, folder_id=ROOT, **options)
        else:
            mapped = {{"max_requests": "search_max_requests", "max_items": "search_max_items",
                       "max_depth": "search_max_depth"}.get(k, k): v for k, v in options.items()}
            client = Site(provider, site_name=SITE_URL, **mapped)
        self.addCleanup(client.close)
        return client

    def mutations(self):
        return [(method, url) for method, url in self.service.records if method != "GET"]


class StorageArchitecture(StorageFixture):
    def test_shared_engine_and_matching_modern_signatures(self):
        self.assertTrue(issubclass(Drive, _StorageClient))
        for name in METHODS:
            self.assertIs(getattr(Drive, name), getattr(_StorageClient, name))
            self.assertEqual(inspect.signature(getattr(Site, name)), inspect.signature(getattr(Drive, name)))

    def test_site_context_does_not_create_storage_sessions_until_used(self):
        client = self.client("site")
        self.assertEqual(len(self.sessions), 2)
        client.list_children()
        self.assertEqual(len(self.sessions), 4)
        self.assertFalse(any(s.close.called for s in self.sessions[:2]))
        self.assertTrue(all(s.close.called for s in self.sessions[2:]))

    def test_snapshot_is_owned_and_not_retargeted_by_site_rebind(self):
        item = self.service.add("old.txt", data=b"OLD")
        site = self.client("site")
        with site.as_drive() as view:
            site.update_drive_name("Other")
            self.assertEqual(site.drive_id, "other")
            self.assertEqual(view.drive_id, DRIVE)
            self.assertEqual(view.get_item("old.txt")["id"], item)
            self.assertEqual(site.list_children(), [])
        self.assertTrue(all(s.close.called for s in self.sessions[2:]))

    def test_independent_view_survives_site_close(self):
        site = self.client("site")
        with site.as_drive() as view:
            site.close()
            self.assertEqual(view.get_item()["id"], ROOT)
        before = len(self.service.records)
        with self.assertRaises(ValueError):
            site.list_children()
        self.assertEqual(len(self.service.records), before)

    def test_scoped_site_never_requests_drive_root(self):
        sub = self.service.add("Restricted")
        self.service.add("allowed.txt", sub, b"YES")
        self.service.add("outside.txt", data=b"NO")
        self.service.denied_root = True
        site = self.client("site", search_root=sub)
        self.assertEqual([r["name"] for r in site.find_files(".txt")], ["allowed.txt"])
        self.assertFalse(any(urlsplit(u).path.endswith("/root") for _, u in self.service.records))
        with self.assertRaises(ValueError):
            site.as_drive("outside")

    def test_modern_site_search_does_not_inherit_unsafe_index_scope(self):
        site = self.client("site", search_mode="indexed")
        self.service.add("safe.txt", data=b"YES")
        self.assertEqual(len(site.find_files("safe")), 1)
        self.assertFalse(any("search(q=" in u for _, u in self.service.records))

    def test_failed_bound_view_closes_only_its_own_sessions(self):
        site = self.client("site")
        self.service.denied_root = True
        with self.assertRaises(GraphError):
            site.as_drive()
        self.assertFalse(any(s.close.called for s in self.sessions[:2]))
        self.assertTrue(all(s.close.called for s in self.sessions[2:]))

    def test_legacy_namespace_entry_points_remain_present(self):
        site = self.client("site")
        self.assertTrue(callable(site.operations.Uploader))
        self.assertTrue(callable(site.operations.Downloader))
        self.assertTrue(callable(site.operations.Mover))
        self.assertTrue(callable(site.search_by.filename))
        self.assertTrue(callable(site.search_by.foldername))
        self.assertTrue(callable(site.search_by.filefoldername))
        self.assertTrue(set(METHODS) <= set(dir(site)))


class StorageWorkflows(StorageFixture):
    def test_upload_search_move_rename_download_for_both_facades(self):
        for kind in ("site", "drive"):
            with self.subTest(kind=kind), self.client(kind) as client:
                name = kind+"-Incoming"
                client.create_folder(name)
                payload = bytes(range(256))*3 + "Résumé 🧪".encode()
                item = client.upload(io.BytesIO(payload), new_file_name=kind+" #1.bin")
                found = client.find_files(kind+" #1", recursive=False)
                self.assertEqual([r["id"] for r in found], [item["id"]])
                pair = client.find_file_folder(kind+" #1.bin", name)
                self.assertEqual(pair[0]["id"], item["id"])
                moved = client.move(kind+" #1.bin", name)
                renamed = client.rename(name+"/"+moved["name"], kind+" renamed.bin")
                self.assertEqual(renamed["id"], item["id"])
                with tempfile.TemporaryDirectory() as local:
                    downloaded = client.download_file(name+"/"+renamed["name"], local)
                    self.assertEqual(hashlib.sha256(downloaded.read_bytes()).digest(), hashlib.sha256(payload).digest())
                self.assertEqual(client.list_children(name)[0]["id"], item["id"])

    def test_folder_zip_and_no_clobber_are_shared(self):
        for kind in ("site", "drive"):
            with self.subTest(kind=kind), self.client(kind) as client:
                client.create_folder(kind)
                client.create_folder("Empty", kind)
                client.upload(io.BytesIO(b"UNCHANGED"), kind, "file.bin")
                with tempfile.TemporaryDirectory() as local:
                    archive = client.download_folder(kind, local, compress=True, max_workers=1)
                    with zipfile.ZipFile(archive) as z:
                        self.assertEqual(z.read("file.bin"), b"UNCHANGED")
                        self.assertIn("Empty/", z.namelist())
                    with self.assertRaises(FileExistsError):
                        client.download_folder(kind, local, compress=True)

    def test_scope_escape_never_reaches_a_mutation(self):
        for kind in ("site", "drive"):
            with self.subTest(kind=kind), self.client(kind) as client:
                for path in ("../outside", "/root/file", "https://evil.invalid/x", "a//b", "a/../b"):
                    with self.subTest(path=path), self.assertRaises(ValueError):
                        client.move(path, "")
                    with self.assertRaises(ValueError):
                        client.rename(path, "safe")
                self.assertEqual(self.mutations(), [])

    def test_no_implicit_overwrite_or_suffix_inference(self):
        for kind in ("site", "drive"):
            with self.subTest(kind=kind), self.client(kind) as client:
                self.service.add(kind+".txt", data=b"ORIGINAL")
                result = client.rename(kind+".txt", kind)
                self.assertEqual(result["name"], kind)
                self.assertEqual(self.service.nodes[result["id"]]["_data"], b"ORIGINAL")


class StorageSearch(StorageFixture):
    def test_sorted_limit_scans_all_matching_metadata(self):
        first = self.service.add("Report-z", data=b"a")
        last = self.service.add("Report-a", data=b"b")
        self.service.nodes[first]["createdDateTime"] = "2025-01-01T00:00:00Z"
        self.service.nodes[last]["createdDateTime"] = "2026-01-01T00:00:00Z"
        for kind in ("site", "drive"):
            client = self.client(kind)
            self.assertEqual([r["id"] for r in client.find_files("Report", sort_order="recent", limit=1)], [last])
            self.assertEqual([r["id"] for r in client.find_files("Report", sort_order="oldest")], [first, last])
            self.assertEqual([r["id"] for r in client.find_files("Report", sort_order="name")], [last, first])
            self.assertEqual([r["id"] for r in client.find_files("Report", sort_order="name_desc")], [first, last])

    def test_sorted_missing_or_invalid_datetime_is_not_silently_ranked(self):
        item = self.service.add("a.txt", data=b"x")
        for date in (None, "", "not-a-date", 42):
            if date is None:
                self.service.nodes[item].pop("createdDateTime", None)
            else:
                self.service.nodes[item]["createdDateTime"] = date
            for kind in ("site", "drive"):
                with self.subTest(kind=kind, date=date), self.assertRaises(GraphError):
                    self.client(kind).find_files(".txt", sort_order="recent")

    def test_sort_and_combined_controls_are_strict(self):
        for kind in ("site", "drive"):
            client = self.client(kind)
            for sort in ("bogus", True, 7, []):
                with self.subTest(kind=kind, sort=sort), self.assertRaises(ValueError):
                    client.find_files("x", sort_order=sort)
            for term in ("", None, "bad\nterm"):
                with self.assertRaises(ValueError):
                    client.find_file_folder(term, "folder")
                with self.assertRaises(ValueError):
                    client.find_file_folder("file", term)
            self.assertEqual(client.find_files("x", limit=0), [])
            self.assertEqual(self.mutations(), [])

    def test_combined_search_requires_unique_full_population(self):
        sub = self.service.add("Sub")
        self.service.add("target")
        self.service.add("source.txt", data=b"x")
        self.service.add("source.txt", sub, b"y")
        for kind in ("site", "drive"):
            with self.subTest(kind=kind), self.assertRaises(ValueError):
                self.client(kind).find_file_folder("source.txt", "target")

    def test_combined_search_shares_one_operation_budget(self):
        self.service.add("target")
        self.service.add("source.txt", data=b"x")
        for kind in ("site", "drive"):
            with self.subTest(kind=kind), self.assertRaises(GraphError):
                self.client(kind, max_requests=2).find_file_folder("source.txt", "target")

    def test_nested_search_cannot_accept_the_anchor_as_a_child(self):
        sub = self.service.add("Sub")
        self.service.nodes[ROOT]["parentReference"] = {"id": sub, "driveId": DRIVE}
        for kind in ("site", "drive"):
            with self.subTest(kind=kind), self.assertRaises(GraphError):
                self.client(kind).find_folders("Archive", folder_path="Sub")


    def test_known_nested_ancestor_cannot_reappear_as_a_child(self):
        parent = self.service.add("Parent")
        child = self.service.add("Child", parent)
        original = self.service.graph
        def forged(method, url, **kwargs):
            if method == "GET" and urlsplit(url).path.endswith("/items/"+child+"/children"):
                return page(folder(parent, "Parent", child))
            return original(method, url, **kwargs)
        with patch.object(self.service, "graph", side_effect=forged):
            for kind in ("site", "drive"):
                client = self.client(kind)
                with self.subTest(kind=kind, operation="list"), self.assertRaises(GraphError):
                    client.list_children("Parent/Child")
                with self.subTest(kind=kind, operation="limited-search"), self.assertRaises(GraphError):
                    client.find_folders("Parent", folder_path="Parent/Child", limit=1)

    def test_sorted_search_never_uses_a_partial_early_page(self):
        first = self.service.add("a.txt", data=b"x")
        original = self.service.graph
        link = "https://graph.microsoft.com/v1.0/drives/drive/items/anchor/children?$skiptoken=opaque"
        def pages(method, url, **kwargs):
            if url == link:
                return response({}, status=503)
            if method == "GET" and urlsplit(url).path.endswith("/items/anchor/children"):
                return page(self.service.metadata(first), link=link)
            return original(method, url, **kwargs)
        with patch.object(self.service, "graph", side_effect=pages):
            for kind in ("site", "drive"):
                with self.subTest(kind=kind), self.assertRaises(GraphError):
                    self.client(kind, max_retries=0).find_files(".txt", sort_order="name", limit=1)


class StorageMutationGuards(StorageFixture):
    def item(self, kind="file"):
        return self.service.add("source.bin", data=b"ORIGINAL") if kind == "file" else self.service.add("Source")

    def test_destination_conflict_preserves_both_items(self):
        for kind in ("site", "drive"):
            source = self.service.add(kind, data=b"SOURCE")
            destination = self.service.add(kind+"-Dest")
            existing = self.service.add(kind, destination, b"KEEP")
            with self.assertRaises(FileExistsError):
                self.client(kind).move(kind, kind+"-Dest")
            self.assertEqual(self.service.nodes[source]["parentReference"]["id"], ROOT)
            self.assertEqual(self.service.nodes[existing]["_data"], b"KEEP")
        self.assertEqual(self.mutations(), [])

    def test_root_move_rename_and_descendant_destination_rejected(self):
        top = self.service.add("Top")
        self.service.add("Inner", top)
        for kind in ("site", "drive"):
            client = self.client(kind)
            for action in (lambda: client.move("", "Top"), lambda: client.rename("", "Changed"),
                           lambda: client.move("Top", "Top"), lambda: client.move("Top", "Top/Inner")):
                with self.assertRaises(ValueError):
                    action()
        self.assertEqual(self.mutations(), [])

    def test_missing_etag_blocks_mutation_but_explicit_etag_can_be_used(self):
        source = self.item()
        self.service.nodes[source].pop("eTag")
        for kind in ("site", "drive"):
            with self.assertRaises(GraphError):
                self.client(kind).rename("source.bin", "changed.bin")
        self.assertEqual(self.mutations(), [])
        self.service.nodes[source]["eTag"] = '"current"'
        result = self.client("drive").rename("source.bin", "changed.bin", if_match='"current"')
        self.assertEqual(result["id"], source)

    def test_invalid_etags_do_not_reach_patch(self):
        self.item()
        for kind in ("site", "drive"):
            client = self.client(kind)
            for value in ("", "*", "bad\r\nHeader:value", "x y", True, 42, [], "é", "\ud800"):
                with self.subTest(kind=kind, value=value), self.assertRaises(ValueError):
                    client.rename("source.bin", "changed.bin", if_match=value)
        self.assertEqual(self.mutations(), [])

    def test_stale_explicit_etag_keeps_source_unchanged(self):
        source = self.item()
        for kind in ("site", "drive"):
            with self.assertRaises(GraphError) as error:
                self.client(kind).rename("source.bin", "changed.bin", if_match='"stale"')
            self.assertEqual(error.exception.status_code, 412)
            self.assertEqual(self.service.nodes[source]["name"], "source.bin")

    def test_concurrent_source_change_is_not_replayed(self):
        source = self.item()
        self.service.before_patch = lambda service, identifier, _: service.nodes[identifier].update(eTag='"newer"')
        with self.assertRaises(GraphError) as error:
            self.client("drive").rename("source.bin", "changed.bin")
        self.assertEqual(error.exception.status_code, 412)
        self.assertEqual(len(self.mutations()), 1)
        self.assertEqual(self.service.nodes[source]["name"], "source.bin")

    def test_concurrent_destination_conflict_is_not_overwritten(self):
        source = self.item()
        def collide(service, identifier, options):
            service.add("changed.bin", data=b"WINNER")
        self.service.before_patch = collide
        with self.assertRaises(GraphError) as error:
            self.client("site").rename("source.bin", "changed.bin")
        self.assertEqual(error.exception.status_code, 409)
        self.assertEqual(len(self.mutations()), 1)
        self.assertEqual(self.service.nodes[source]["name"], "source.bin")
        self.assertEqual(self.service.nodes[self.service.lookup(ROOT, "changed.bin")]["_data"], b"WINNER")

    def test_ambiguous_commit_never_replayed_or_automatically_reversed(self):
        source = self.item()
        def lost(_):
            raise requests.ConnectionError("Synthetic lost acknowledgement")
        self.service.patch_override = lost
        with self.assertRaises(GraphError):
            self.client("site").rename("source.bin", "changed.bin")
        self.assertEqual(len(self.mutations()), 1)
        self.assertEqual(self.service.nodes[source]["name"], "changed.bin")

    def test_noop_is_read_only_but_stale_explicit_etag_fails(self):
        source = self.item()
        for kind in ("site", "drive"):
            client = self.client(kind)
            self.assertEqual(client.move("source.bin")["id"], source)
            with self.assertRaises(GraphError) as error:
                client.rename("source.bin", "source.bin", if_match='"stale"')
            self.assertEqual(error.exception.status_code, 412)
        self.assertEqual(self.mutations(), [])

    def test_case_only_rename_preserves_identity(self):
        source = self.item()
        value = self.client("drive").rename("source.bin", "SOURCE.bin")
        self.assertEqual(value["id"], source)
        self.assertEqual(self.service.nodes[source]["name"], "SOURCE.bin")

    def test_folder_move_keeps_children_and_cannot_break_depth_limit(self):
        source = self.item("folder")
        child = self.service.add("inside.bin", source, b"KEEP")
        self.service.add("Dest")
        for kind in ("site", "drive"):
            with self.assertRaises(GraphError):
                self.client(kind, max_depth=2).move("Source", "Dest")
        self.assertEqual(self.mutations(), [])
        result = self.client("drive").move("Source", "Dest")
        self.assertEqual(result["id"], source)
        self.assertEqual(self.service.nodes[child]["parentReference"]["id"], source)

    def test_destination_parent_id_not_root_alias_is_sent(self):
        self.item()
        self.service.add("Dest")
        client = self.client("drive")
        moved = client.move("source.bin", "Dest")
        back = client.move("Dest/source.bin", "")
        self.assertEqual(moved["id"], back["id"])
        self.assertEqual(back["parentReference"]["id"], ROOT)

    def test_invalid_confirmation_is_not_success_even_after_commit(self):
        for status in (201, 202, 204, 206):
            for kind in ("site", "drive"):
                name = kind+str(status)+".bin"
                self.service.add(name, data=b"ORIGINAL")
                self.service.patch_override = lambda r, s=status: response(r.json(), status=s)
                with self.subTest(kind=kind, status=status), self.assertRaises(GraphError):
                    self.client(kind).rename(name, "renamed-"+name)

    def test_contradictory_commit_metadata_is_not_accepted(self):
        changes = (
            lambda row: row.update(id="unrelated"),
            lambda row: row.update(name="wrong"),
            lambda row: row["parentReference"].update(id="outside"),
            lambda row: row["parentReference"].update(driveId="other"),
            lambda row: row["parentReference"].pop("driveId"),
            lambda row: row.update(size=999),
            lambda row: row.update(folder={}),
            lambda row: row.update(remoteItem={}),
            lambda row: row.update(error={}),
        )
        for index, change in enumerate(changes):
            name = str(index)+".bin"
            self.service.add(name, data=b"ORIGINAL")
            def corrupt(reply, change=change):
                row = reply.json()
                change(row)
                return response(row)
            self.service.patch_override = corrupt
            with self.subTest(index=index), self.assertRaises(GraphError):
                self.client("drive").rename(name, "renamed-"+name)
        self.assertEqual(len(self.mutations()), len(changes))

    def test_unknown_source_and_wrong_destination_kind_do_not_mutate(self):
        self.item()
        for kind in ("site", "drive"):
            client = self.client(kind)
            with self.assertRaises(GraphError):
                client.move("missing", "")
            with self.assertRaises(ValueError):
                client.move("source.bin", "source.bin")
        self.assertEqual(self.mutations(), [])


class StorageProfiles(StorageFixture):
    def profile_client(self, kind, method, *, allow=False, error=None, **kwargs):
        auth = Mock()
        auth.access_token = "SYNTHETIC-AUTH-TOKEN"
        if error:
            auth.Silent.side_effect = error
        with patch.dict(os.environ, {}, clear=True), \
                patch("azauthlib.Authentication", return_value=auth) as factory, \
                patch("azauthlib.credentials.load_public_configuration",
                      return_value={"auth_method": method}):
            client = (Drive.from_profile("Saved", DRIVE, folder_id=ROOT, allow_interactive=allow, **kwargs)
                      if kind == "drive" else Site.from_profile("Saved", SITE_URL, allow_interactive=allow, **kwargs))
        self.addCleanup(client.close)
        return client, auth, factory

    def test_profile_application_flows_are_identical(self):
        for kind in ("site", "drive"):
            for method in ("client_credentials", "client_certificate"):
                with self.subTest(kind=kind, method=method):
                    client, auth, factory = self.profile_client(kind, method)
                    auth.ClientCredentials.assert_called_once_with()
                    auth.Silent.assert_not_called()
                    auth.Interactive.assert_not_called()
                    self.assertEqual(client.get_item()["id"], ROOT)

    def test_profile_silent_delegated_flows_never_open_ui(self):
        for kind in ("site", "drive"):
            for method in ("interactive", "device_code"):
                client, auth, _ = self.profile_client(kind, method)
                auth.Silent.assert_called_once_with()
                auth.Interactive.assert_not_called()
                auth.DeviceCodeFlow.assert_not_called()

    def test_profile_interaction_requires_opt_in_for_both_facades(self):
        for kind in ("site", "drive"):
            for method in ("interactive", "device_code"):
                before = len(self.service.records)
                with self.assertRaises(InteractionRequiredError):
                    self.profile_client(kind, method, error=InteractionRequiredError("Needed"))
                self.assertEqual(len(self.service.records), before)
                client, auth, _ = self.profile_client(kind, method, allow=True,
                                                       error=InteractionRequiredError("Needed"))
                selected = auth.Interactive if method == "interactive" else auth.DeviceCodeFlow
                selected.assert_called_once_with()

    def test_other_profile_errors_never_fall_back_to_ui(self):
        for kind in ("site", "drive"):
            with self.assertRaises(AuthenticationError):
                self.profile_client(kind, "interactive", allow=True, error=AuthenticationError("Blocked"))

    def test_unknown_site_options_fail_before_authentication(self):
        with patch("azauthlib.Authentication") as factory:
            for options in ({"nonsense": True}, {"search_max_requests": 0},
                            {"graph_url": "http://unsafe.invalid/v1.0"}, {"search_max_depth": 0},
                            {"search_mode": "bad"}, {"search_root": ""}):
                with self.assertRaises((TypeError, ValueError)):
                    Site.from_profile("Saved", SITE_URL, **options)
            factory.assert_not_called()

    def test_profile_environment_overrides_block_both_facades(self):
        with patch.dict(os.environ, {"CLIENT_SECRET": "SYNTHETIC"}, clear=True), \
                patch("azauthlib.Authentication") as factory:
            with self.assertRaises(ValueError):
                Site.from_profile("Saved", SITE_URL)
            with self.assertRaises(ValueError):
                Drive.from_profile("Saved", DRIVE)
            factory.assert_not_called()

    def test_legacy_factories_are_available_for_both_without_interactive_fallback(self):
        auth = Mock()
        auth.access_token = "SYNTHETIC-AUTH"
        for kind, cls, resource in (("site", Site, SITE_URL), ("drive", Drive, DRIVE)):
            for name, options in (
                ("from_default_env", {}),
                ("from_env_file", {"env_path": "synthetic.env"}),
                ("from_os_env", {}),
                ("from_entry", {"client_id": "client", "tenant_id": "tenant"}),
            ):
                auth.reset_mock()
                with self.subTest(kind=kind, factory=name), \
                        patch("azauthlib.Authentication", return_value=auth), \
                        patch("m365.graph.sharepoint._AzAuth", return_value=auth):
                    client = getattr(cls, name)(resource, **options)
                    self.addCleanup(client.close)
                    auth.Silent.assert_called_once_with(scopes=None)
                    auth.Interactive.assert_not_called()

    def test_invalid_drive_builder_controls_fail_before_authentication(self):
        with patch("azauthlib.Authentication") as factory:
            with self.assertRaises(ValueError):
                Drive.from_entry(DRIVE, client_id="client", tenant_id="tenant", max_requests=0)
            factory.assert_not_called()


class GeneratedStorageMoves(StorageFixture):
    pass


def valid_move_case(index, kind):
    def test(self):
        client = self.client(kind)
        old = "source-"+str(index)+".bin"
        new = "Café # % O'Brien 🧪 "+str(index)+".bin"
        source = self.service.add(old, data=bytes([index]))
        actual = client.rename(old, new)
        self.assertEqual(actual["id"], source)
        self.assertEqual(self.service.nodes[source]["_data"], bytes([index]))
        self.assertEqual(client.get_item(new)["id"], source)
    return test


def invalid_move_case(index, kind):
    def test(self):
        client = self.client(kind)
        self.item = self.service.add("source.bin", data=b"KEEP")
        paths = ("../x", "/x", "x//y", "x/../y", "https://outside.invalid/x",
                 "x\\y", "x/./y", "x\x00y", "x?y", "x:y")
        with self.assertRaises(ValueError):
            client.move(paths[index % len(paths)], "")
        self.assertEqual(self.mutations(), [])
        self.assertEqual(self.service.nodes[self.item]["_data"], b"KEEP")
    return test


for _index in range(40):
    for _kind in ("site", "drive"):
        setattr(GeneratedStorageMoves, "test_valid_%s_%02d" % (_kind, _index), valid_move_case(_index, _kind))
        setattr(GeneratedStorageMoves, "test_invalid_%s_%02d" % (_kind, _index), invalid_move_case(_index, _kind))
