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

"""Required auth dependency and real-provider integration without tenant access."""
from contextlib import ExitStack
from email import message_from_bytes
import os
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from azauthlib import Authentication, InteractionRequiredError
from msal import SerializableTokenCache

from m365.graph import Mail, Site
from test_sharepoint import response


class AuthProviderTests(unittest.TestCase):
    def setUp(self):
        stack = ExitStack()
        self.addCleanup(stack.close)
        workdir = Path(stack.enter_context(TemporaryDirectory()))
        self.config = {"client_id": "synthetic-app", "tenant_id": "synthetic-tenant",
                       "token_path": str(workdir / "test-token.json"), "scopes": ["User.Read"]}
        self.now = 1_800_000_000
        stack.enter_context(patch("azauthlib.authentication.time",
                                  SimpleNamespace(time=lambda: self.now)))
        self.settings = stack.enter_context(patch(
            "azauthlib.authentication.get_settings",
            return_value=SimpleNamespace(env_path=workdir / "unused.env",
                                         token_path=workdir / "unused-cache.json")))
        self.cache = stack.enter_context(patch("azauthlib.authentication.load_token_cache",
                                              return_value=SerializableTokenCache()))
        self.msal = stack.enter_context(patch("azauthlib.authentication.PublicClientApplication"))
        self.app = self.msal.return_value
        self.app.get_accounts.return_value = [{"username": "fixture@example.invalid"}]
        self.app.acquire_token_silent.side_effect = [
            {"access_token": "initial-fixture-token", "expires_in": 3600},
            {"access_token": "renewed-fixture-token", "expires_in": 3600},
        ]
        self.graph = stack.enter_context(patch("requests.sessions.Session.request",
                                              side_effect=AssertionError("Unexpected network call")))

    def test_sharepoint_factory_and_next_request_use_real_provider_refresh(self):
        self.graph.side_effect = [
            response({"id": "fixture-site", "webUrl": "https://fixture.invalid/sites/Test"}),
            response({"id": "fixture-drive", "name": "Documents"}),
            response({"id": "fixture-root"}),
        ]
        with Site.from_entry("https://fixture.invalid/sites/Test", **self.config) as site:
            self.assertIsInstance(site.token_engine, Authentication)
            self.assertEqual(site.drive_id, "fixture-drive")
            self.now += 3601
            item = site.api.graph_call(resource_url="drives/fixture-drive/root")
        self.assertEqual(item, {"id": "fixture-root"})
        headers = [call.kwargs["headers"]["Authorization"] for call in self.graph.call_args_list]
        self.assertEqual(headers, ["Bearer initial-fixture-token", "Bearer initial-fixture-token",
                                   "Bearer renewed-fixture-token"])
        self.assertEqual(self.app.acquire_token_silent.call_count, 2)
        self.cache.assert_called_once_with(self.config["token_path"], allow_plaintext=False)

    def test_mailbox_factory_authenticates_with_real_provider(self):
        self.graph.side_effect = [response({"value": []}), response({"id": "fixture-user"})]
        mailbox = Mail.from_entry(**self.config)
        self.assertIsInstance(mailbox.token_engine, Authentication)
        self.assertEqual(mailbox.mailbox_profile, {"id": "fixture-user"})
        self.assertEqual(self.graph.call_count, 2)
        for call in self.graph.call_args_list:
            self.assertEqual(call.kwargs["headers"]["Authorization"], "Bearer initial-fixture-token")
        self.app.acquire_token_silent.assert_called_once()

    def test_missing_cached_sign_in_never_launches_interactive_or_requests_graph(self):
        self.app.get_accounts.return_value = []
        for factory, args in ((Site.from_entry, ("https://fixture.invalid/sites/Test",)),
                              (Mail.from_entry, ())):
            with self.subTest(factory=factory), self.assertRaises(InteractionRequiredError):
                factory(*args, **self.config)
        self.app.acquire_token_interactive.assert_not_called()
        self.graph.assert_not_called()

    def test_custom_provider_does_not_construct_azauthlib(self):
        self.graph.side_effect = [
            response({"id": "fixture-site", "webUrl": "https://fixture.invalid/sites/Test"}),
            response({"id": "fixture-drive", "name": "Documents"}),
            response({"value": []}), response({"id": "fixture-user"}),
        ]
        supplied = SimpleNamespace(access_token="injected-fixture-token")
        with Site(token_engine=supplied, site_name="https://fixture.invalid/sites/Test") as site:
            self.assertIs(site.token_engine, supplied)
        self.assertIs(Mail(token_engine=supplied).token_engine, supplied)
        self.settings.assert_not_called()
        self.msal.assert_not_called()
        for call in self.graph.call_args_list:
            self.assertEqual(call.kwargs["headers"]["Authorization"], "Bearer injected-fixture-token")


class DependencyMetadataTests(unittest.TestCase):
    @unittest.skipUnless(os.environ.get("M365_DISTRIBUTION_DIR"), "Built wheel not supplied")
    def test_azauthlib_is_unconditional_and_dates_remain_optional(self):
        directory = Path(os.environ["M365_DISTRIBUTION_DIR"])
        wheel, = directory.glob("m365-*.whl")
        with ZipFile(wheel) as archive:
            metadata, = [name for name in archive.namelist() if name.endswith(".dist-info/METADATA")]
            headers = message_from_bytes(archive.read(metadata))
        self.assertEqual(headers["Requires-Python"], ">=3.10")
        requires = headers.get_all("Requires-Dist")
        auth, = [value for value in requires if value.startswith("azauthlib")]
        self.assertNotIn(";", auth)
        bounds = auth.removeprefix("azauthlib").strip(" ()").split(",")
        self.assertEqual({bound.strip() for bound in bounds}, {">=6.0.6", "<7"})
        self.assertTrue(any(value.startswith("dately") and 'extra == "dates"' in value
                            for value in requires))
        self.assertIn("auth", headers.get_all("Provides-Extra"))
