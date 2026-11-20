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
"""Offline regression coverage for SharePoint path discovery and S3 bucket paths."""

from types import SimpleNamespace
import unittest

from m365.graph import Site, _syslib as syslib
from m365.graph._syslib._paradigm import (
    __CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_S3_SPECIFIC__ as S3_RULES,
)


_RESERVED_SUFFIXES = ("-s3alias", "--ol-s3", ".mrap", "--x-s3", "--table-s3")
_RESERVED_PREFIXES = ("xn--", "sthree-", "amzn-s3-demo-")


class SharePointPathDiscoveryTests(unittest.TestCase):
    def test_dir_exposes_the_intended_path_helpers(self):
        parent = SimpleNamespace(url=object())
        path = Site.PATH(parent)
        expected = (
            "drop_drivename",
            "strip_domain_extension",
            "is_sharepoint_folderpath",
            "is_sharepoint_filepath",
            "relative_path",
            "file_ext",
            "join_file_ext",
            "is_path_struct",
        )
        self.assertEqual(dir(path), sorted(expected))
        for name in expected:
            with self.subTest(name=name):
                self.assertTrue(callable(getattr(path, name)))
        self.assertIs(path.parent, parent)
        self.assertIs(path.url, parent.url)
        self.assertEqual(path.remove_leading_slash("/Folder/File.txt"), "Folder/File.txt")


class S3BucketPathTests(unittest.TestCase):
    def test_reserved_suffixes_are_rejected_by_match_and_fullmatch(self):
        for suffix in _RESERVED_SUFFIXES:
            name = "example" + suffix
            with self.subTest(name=name):
                self.assertIsNone(S3_RULES.match(name))
                self.assertIsNone(S3_RULES.fullmatch(name))

    def test_reserved_prefixes_are_still_rejected(self):
        for prefix in _RESERVED_PREFIXES + ("sthree-configurator",):
            with self.subTest(prefix=prefix):
                self.assertIsNone(S3_RULES.match(prefix + "example"))

    def test_s3_rule_matches_the_whole_nonempty_name(self):
        for name in ("", "example\n", "example\n--x-s3"):
            with self.subTest(name=name):
                self.assertIsNone(S3_RULES.match(name))

    def test_suffix_literals_are_only_reserved_at_the_end(self):
        for name in (
            "example", "example.s3.archive", "example-mrap", "exampleamrap",
            "example-111122223333-us-west-2-an",
            *("example" + suffix + "-archive" for suffix in _RESERVED_SUFFIXES),
        ):
            with self.subTest(name=name):
                self.assertEqual(S3_RULES.match(name).group(), name)

    def test_public_path_type_rejects_reserved_s3_buckets(self):
        for provider in ("S3", "AWS", "Amazon"):
            for suffix in _RESERVED_SUFFIXES:
                with self.subTest(provider=provider, suffix=suffix):
                    result = syslib.path_type(
                        "example" + suffix + "/folder/report.pdf",
                        cloud=True,
                        provider=provider,
                    )
                    self.assertEqual(result, {"folder": False, "file": False})

    def test_public_path_type_accepts_ordinary_s3_paths_and_object_names(self):
        for path, expected in (
            ("example-bucket/folder/", {"folder": True, "file": False}),
            ("example-bucket/report.mrap", {"folder": False, "file": True}),
            ("example-bucket/folder--x-s3/report.pdf", {"folder": False, "file": True}),
            ("example--table-s3-archive/report.pdf", {"folder": False, "file": True}),
        ):
            with self.subTest(path=path):
                self.assertEqual(syslib.path_type(path, cloud=True, provider="S3"), expected)

    def test_s3_reservation_does_not_change_other_providers(self):
        for provider in ("Wasabi", "Backblaze", "SharePoint"):
            for suffix in _RESERVED_SUFFIXES:
                with self.subTest(provider=provider, suffix=suffix):
                    result = syslib.path_type(
                        "example" + suffix + "/folder/report.pdf",
                        cloud=True,
                        provider=provider,
                    )
                    self.assertEqual(result, {"folder": False, "file": True})


if __name__ == "__main__":
    unittest.main()
