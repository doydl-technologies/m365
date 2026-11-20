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
"""Core ISO handling must never require the optional advanced backend."""
import unittest
import ast
from pathlib import Path
from datetime import datetime, timezone, timedelta
from types import SimpleNamespace
from unittest.mock import Mock, patch

from m365.graph import _dateplumb as dateplumb
from m365.graph._core_dates import parse_graph_datetime
from m365.graph._dateplumb import _advanced
from m365.graph.operations.doc_management import File


class CoreDates(unittest.TestCase):
    def setUp(self):
        self.guard = patch.object(_advanced, "_backend", side_effect=AssertionError("Core tried to load dately"))
        self.loader = self.guard.start()
        self.addCleanup(self.guard.stop)

    def test_core_has_no_advanced_backend_import(self):
        import m365.graph
        file = Path(m365.graph.__file__).resolve().parent / "_core_dates.py"
        for node in ast.walk(ast.parse(file.read_text(encoding="utf-8-sig"))):
            if isinstance(node, ast.Import):
                self.assertFalse(any(n.name.split('.')[0] == 'dately' for n in node.names))
            if isinstance(node, ast.ImportFrom):
                self.assertNotEqual((node.module or '').split('.')[0], 'dately')

    def test_iso_timestamps_preserve_z_and_numeric_offsets(self):
        utc = parse_graph_datetime("2026-11-17T01:02:03Z")
        self.assertEqual(utc.tzinfo, timezone.utc)
        other = parse_graph_datetime("2026-11-17T01:02:03-0530")
        self.assertEqual(other.utcoffset(), timedelta(hours=-5, minutes=-30))
        self.assertEqual(parse_graph_datetime("2026-11-17T01:02:03").tzinfo, timezone.utc)
        self.loader.assert_not_called()

    def test_graph_fraction_precision_python38_compatible(self):
        for fraction, microsecond in (("1", 100000), ("123", 123000), ("123456", 123456), ("1234567", 123456)):
            self.assertEqual(parse_graph_datetime("2026-11-17T01:02:03." + fraction + "Z").microsecond, microsecond)

    def test_invalid_iso_and_offset_values_do_not_fall_back(self):
        for value in ("2026-02-29", "2026-11-17T25:00:00Z", "2026-11-17T12:00:00+1260", "tomorrow", "2026", "2026-11"):
            with self.subTest(value=value), self.assertRaises(ValueError): parse_graph_datetime(value)
        self.loader.assert_not_called()

    def test_public_basic_parse_without_advanced_backend(self):
        self.assertEqual(dateplumb.parse("2026-11-17"), datetime(2026, 11, 17))
        self.assertEqual(dateplumb.quick_parse("2026-11-17 12:34"), datetime(2026, 11, 17, 12, 34))
        self.assertEqual(dateplumb.parse("2026-11-17T12:34:56Z", include_timezone=True).tzinfo, timezone.utc)
        self.assertIsNone(dateplumb.parse("2026-11-17T12:34:56Z").tzinfo)  # Legacy public flag.
        self.assertEqual(dateplumb.parse("2026-11-17T12:34:56+02:00").utcoffset(), timedelta(hours=2))
        self.assertEqual(len(dateplumb.parse(["2026-11-17T12:34:56.1234567Z", "2026-11-18"])), 2)
        with self.assertRaises(ValueError): dateplumb.parse("unsupported phrase")
        with self.assertRaises(ValueError): dateplumb.parse(None)
        self.loader.assert_not_called()

    def test_legacy_time_substitution_and_time_detection(self):
        self.assertEqual(dateplumb.sub.time("2026-11-17 12:34", hours=13), "2026-11-17 13:34")
        self.assertTrue(dateplumb.contains_time(datetime(2026, 11, 17, 1)))

    def test_graph_sorting_does_not_load_advanced_dates(self):
        client = SimpleNamespace(drive_id="d", search_mode="indexed", api=Mock())
        client.api.iter_collection.return_value = [
            {"id": "1", "name": "f", "file": {}, "createdDateTime": "2026-11-17T10:00:00-05:00"},
            {"id": "2", "name": "f", "file": {}, "createdDateTime": "2026-11-17T14:00:00Z"}]
        self.assertEqual(File(client).find("f", sort_order="recent")[0]["id"], "1")
        self.loader.assert_not_called()


class OptionalBoundary(unittest.TestCase):
    def setUp(self):
        _advanced._backend.cache_clear()
        self.addCleanup(_advanced._backend.cache_clear)

    def test_all_six_missing_dependency_errors_are_actionable(self):
        calls = [lambda: dateplumb.pro_parse("2026-11-17"), lambda: dateplumb.conversion("2026-11-17"),
                 lambda: dateplumb.detect("2026-11-17"), lambda: dateplumb.fragment("2026-11-17", "year"),
                 lambda: dateplumb.sequence("2026-11-17", "2026-11-18"), lambda: dateplumb.sub.date("2026-11-17", year=2027)]
        with patch.object(_advanced.importlib, "import_module", side_effect=ModuleNotFoundError("missing", name="dately")):
            for operation in calls:
                with self.subTest(operation=operation), self.assertRaisesRegex(ImportError, r"m365\[dates\]"):
                    operation()

    def test_dependency_failure_does_not_mutate_dictionary(self):
        value = {"when": "2026-11-17"}
        with patch.object(_advanced.importlib, "import_module", side_effect=ModuleNotFoundError("missing", name="dately")):
            with self.assertRaises(ImportError): dateplumb.conversion(value, dict_keys="when", dict_inplace=True)
        self.assertEqual(value, {"when": "2026-11-17"})

    def test_supported_version_and_api_are_checked(self):
        for version in ("4.7.7", "4.8.0", "5.0.0", "dev", None):
            with self.subTest(version=version), patch.object(_advanced.importlib, "import_module", return_value=SimpleNamespace(__version__=version)):
                with self.assertRaisesRegex(ImportError, "Unsupported"): _advanced._backend()
        with patch.object(_advanced.importlib, "import_module", return_value=SimpleNamespace(__version__="4.7.8")):
            with self.assertRaisesRegex(ImportError, "Incomplete"): _advanced._backend()

    def test_transitive_dependency_error_retains_cause(self):
        with patch.object(_advanced.importlib, "import_module", side_effect=ModuleNotFoundError("missing numbr", name="numbr")):
            with self.assertRaises(ImportError) as error: _advanced._backend()
        self.assertIsInstance(error.exception.__cause__, ModuleNotFoundError)


if __name__ == "__main__": unittest.main()
