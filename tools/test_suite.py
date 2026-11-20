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
"""Isolated offline suites for source and installed distributions."""
import argparse
import json
import os
from pathlib import Path
import socket
import sys
import unittest
from unittest.mock import patch


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("suite", choices=["graph", "combined"])
    parser.add_argument("--installed", action="store_true")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--dist", type=Path)
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[1]
    if not args.installed:
        sys.path.insert(0, str(root / "src"))
    args.output.mkdir(parents=True, exist_ok=True)
    if args.dist:
        os.environ["M365_DISTRIBUTION_DIR"] = str(args.dist.resolve())
        os.environ["M365_DIST"] = str(args.dist.resolve())
    attempted = []
    def blocked(*a, **kw):
        attempted.append(True)
        raise AssertionError("Unexpected network request")
    with patch.object(socket.socket, "connect", side_effect=blocked), \
            patch.object(socket, "create_connection", side_effect=blocked):
        suite = unittest.defaultTestLoader.discover(str(root / "tests" / args.suite))
        with (args.output / "tests.log").open("w", encoding="utf-8") as log:
            result = unittest.TextTestRunner(stream=log, verbosity=2).run(suite)
    import m365
    summary = dict(suite=args.suite, installed=args.installed, package=m365.__file__,
                   python=sys.version, tests=result.testsRun,
                   passed=result.testsRun-len(result.skipped)-len(result.failures)-len(result.errors),
                   failed=len(result.failures), errors=len(result.errors),
                   skipped=len(result.skipped), network_attempts=len(attempted))
    (args.output / "results.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(json.dumps(summary, indent=2))
    return 0 if result.wasSuccessful() and not attempted else 1


if __name__ == "__main__":
    sys.exit(main())
