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
"""Local Windows long-path support without changing host/tenant configuration."""
import ntpath
import os
from pathlib import Path


def windows_path(value):
    """Use the Win32 extended namespace for ordinary absolute drive/UNC paths."""
    raw = os.fspath(value)
    if raw.startswith("\\\\.\\"):
        raise ValueError("Device paths are not supported transfer destinations")
    if raw.startswith("\\\\?\\"):
        tail = raw[4:]
        if not (tail.startswith("UNC\\") or (len(tail) >= 3 and tail[1:3] == ":\\")):
            raise ValueError("Only filesystem extended paths are supported")
        return raw
    absolute = ntpath.abspath(raw)
    if absolute.startswith("\\\\"):
        return "\\\\?\\UNC\\" + absolute[2:]
    return "\\\\?\\" + absolute


def native_path(value):
    return Path(windows_path(value) if os.name == "nt" else value)


def public_path(value, *, keep_extended=False):
    """Keep short Windows paths familiar; retain the prefix when it is needed."""
    p = Path(value)
    raw = str(p)
    if os.name == "nt" and raw.startswith("\\\\?\\") and not keep_extended:
        normal = "\\\\" + raw[8:] if raw.startswith("\\\\?\\UNC\\") else raw[4:]
        # Win32 path limits count UTF-16 code units, not Python code points.
        if len(normal.encode("utf-16-le")) // 2 < 248:
            return Path(normal)
    return p
