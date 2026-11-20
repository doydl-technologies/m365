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
"""Recursively locate a filename starting in this module's directory.

Return formats follow absolute_path and relative_path. The first match follows
unsorted directory enumeration. Directory links follow normal filesystem
semantics; this helper is not a sandbox or a containment check.
"""
import os
import re

def file(target_file, absolute_path=True, relative_path=False):
    """
    Find the first matching filename beneath this module's directory tree.
    """
    try: root_path = os.path.dirname(os.path.abspath(__file__))
    except NameError: root_path = os.path.dirname(os.path.abspath("__file__"))
    if not root_path: return "Root directory not found."
    def find(root_path,target_file):
        if os.path.isdir(root_path):
            for entry in os.listdir(root_path):
                full = os.path.join(root_path,entry)
                if os.path.isfile(full) and entry==target_file:
                    return os.path.abspath(full) if absolute_path else (os.path.relpath(full,start='.') if relative_path else full)
                elif os.path.isdir(full):
                    r = find(full,target_file)
                    if r: return r
        return None
    return find(root_path,target_file)

__all__ = ['file']
def __dir__(): return __all__
