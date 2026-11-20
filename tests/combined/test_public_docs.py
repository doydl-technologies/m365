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
"""Keep internal helpers out of application-facing documentation examples."""
from pathlib import Path
import re
import unittest


class PublicDocumentationTests(unittest.TestCase):
    def documents(self):
        root = Path(__file__).resolve().parents[2]
        for name in ("README.md", "STORAGE.md"):
            path = root / name
            # Source distributions intentionally omit README files.
            if path.name == "README.md" and not path.exists():
                continue
            yield name, path.read_text(encoding="utf-8")

    def test_repository_uses_one_project_readme_and_one_storage_guide(self):
        root = Path(__file__).resolve().parents[2]
        for retired in ("src/m365/graph/README.md", "GRAPH.md", "ONEDRIVE.md"):
            with self.subTest(retired=retired):
                self.assertFalse((root / retired).exists())
        self.assertTrue((root / "STORAGE.md").is_file())
        readme = root / "README.md"
        if readme.exists():
            self.assertIn("](STORAGE.md)", readme.read_text(encoding="utf-8"))
        for name, text in self.documents():
            with self.subTest(document=name):
                self.assertNotRegex(text, r"(?i)\b(?:GRAPH|ONEDRIVE)\.md\b")

    def test_documents_define_internal_api_boundary(self):
        for name, text in self.documents():
            with self.subTest(document=name):
                for helper in ("_dateplumb", "_preprocessing", "_syslib"):
                    self.assertIn("`" + helper + "`", text)
                self.assertIn("internal", text)
                self.assertIn("must not be used by applications", text)

    def test_examples_do_not_import_internal_modules_or_names(self):
        internal = {"dateplumb", "preprocessing", "syslib", "_dateplumb", "_preprocessing", "_syslib"}
        for name, text in self.documents():
            for line in text.splitlines():
                match = re.match(r"\s*(?:from|import)\s+(m365(?:\.[\w]+)*)\b(.*)", line)
                if match:
                    with self.subTest(document=name, line=line):
                        parts = set(match[1].split("."))
                        self.assertFalse(parts & internal)
                        self.assertFalse(any(part.startswith("_") for part in parts))
                        imported = re.findall(r"\b\w+\b", match[2])
                        self.assertFalse(set(imported) & internal)
                        self.assertFalse(any(part.startswith("_") for part in imported))

    def test_readmes_do_not_advertise_helper_capabilities(self):
        for name, text in self.documents():
            with self.subTest(document=name):
                self.assertNotIn("| 📅 Dates |", text)
                self.assertNotIn("| 🧹 Data processing |", text)
                self.assertNotIn("| 💾 Local systems |", text)
