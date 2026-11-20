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
"""Verify logo attribution sources and, optionally, actual built distributions.

Set M365_DISTRIBUTION_DIR to a fresh build output directory to exercise
the wheel and source-distribution checks as part of release validation.
"""
import os
from pathlib import Path
import tarfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[2]
ASSETS = 'assets/third-party/'
FILES = ('microsoft-sharepoint.svg', 'SHAREPOINT-NOTICE.txt',
         'microsoft-onedrive.svg', 'ONEDRIVE-NOTICE.txt',
         'microsoft-outlook.svg', 'OUTLOOK-NOTICE.txt')


class ThirdPartyAssetsTests(unittest.TestCase):
    def test_logos_and_notices_present(self):
        for name in FILES:
            with self.subTest(name=name):
                self.assertTrue((ROOT / ASSETS / name).is_file())

    def test_notices_identify_asset_and_owner(self):
        for notice, logo in (('SHAREPOINT-NOTICE.txt', 'microsoft-sharepoint.svg'),
                             ('ONEDRIVE-NOTICE.txt', 'microsoft-onedrive.svg'),
                             ('OUTLOOK-NOTICE.txt', 'microsoft-outlook.svg')):
            with self.subTest(notice=notice):
                text = (ROOT / ASSETS / notice).read_text(encoding='utf-8')
                self.assertIn(logo, text)
                self.assertIn('Microsoft Corporation', text)
                self.assertIn('does not claim ownership', text)
                self.assertIn('does not grant permission', text)

    @unittest.skipUnless(os.environ.get('M365_DISTRIBUTION_DIR'),
                         'Set M365_DISTRIBUTION_DIR to check built artifacts')
    def test_root_assets_in_sdist_and_notices_in_wheel_are_unchanged(self):
        directory = Path(os.environ['M365_DISTRIBUTION_DIR'])
        wheels, sdists = list(directory.glob('m365-*.whl')), list(directory.glob('m365-*.tar.gz'))
        self.assertEqual(len(wheels), 1)
        self.assertEqual(len(sdists), 1)
        with zipfile.ZipFile(wheels[0]) as wheel, tarfile.open(sdists[0], 'r:gz') as sdist:
            for name in FILES:
                with self.subTest(name=name):
                    expected = (ROOT / ASSETS / name).read_bytes()
                    self.assertNotIn('m365/graph/' + ASSETS + name, wheel.namelist())
                    if name.endswith('-NOTICE.txt'):
                        matches = [p for p in wheel.namelist()
                                   if p.endswith('.dist-info/licenses/' + ASSETS + name)]
                        self.assertEqual(len(matches), 1)
                        self.assertEqual(wheel.read(matches[0]), expected)
                    matches = [m for m in sdist.getmembers() if m.name.endswith('/' + ASSETS + name)]
                    self.assertEqual(len(matches), 1)
                    with sdist.extractfile(matches[0]) as stream:
                        self.assertEqual(stream.read(), expected)

    def test_all_root_branding_assets_and_notices_are_in_sdist(self):
        directory = os.environ.get('M365_DISTRIBUTION_DIR')
        if not directory:
            self.skipTest('Build artifacts required')
        sdist, = Path(directory).glob('m365-*.tar.gz')
        with tarfile.open(sdist, 'r:gz') as archive:
            for path in (ROOT / 'assets').rglob('*'):
                if path.is_file():
                    relative = path.relative_to(ROOT).as_posix()
                    matches = [m for m in archive.getmembers() if m.name.endswith('/' + relative)]
                    self.assertEqual(len(matches), 1, relative)
                    with archive.extractfile(matches[0]) as member:
                        self.assertEqual(member.read(), path.read_bytes(), relative)


if __name__ == '__main__':
    unittest.main()
