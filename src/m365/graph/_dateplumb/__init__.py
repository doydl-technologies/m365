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
import sys
from .base import dtparse

__all__ = ['parse', 'now', 'unix_timestamp', 'nowCT', 'subtract_months', 'days_in_year', 'weeks_in_year', 'ET', 'make_timezone_aware', 'conversion', 'detect', 'fragment', 'sequence', 'sub', 'quick_parse', 'pro_parse']

__methods__ = ['parse', 'now', 'unix_timestamp', 'nowCT', 'subtract_months', 'days_in_year', 'weeks_in_year', 'make_timezone_aware', 'conversion', 'detect', 'fragment','quick_parse', 'pro_parse', 'sequence', 'contains_time']
__inner_classes__ = ['ET', 'sub']
__base_methods__ = __methods__ + __inner_classes__


for method_name in __base_methods__:
    if hasattr(dtparse, method_name):
        attr = getattr(dtparse, method_name)
        globals()[method_name] = attr

for instance_name in __inner_classes__:
    instance = getattr(dtparse, instance_name, None)
    if instance is not None:
        globals()[instance_name] = instance
    else:
        print(f"Warning: '{instance_name}' not found in 'dtparse'. Check if it is initialized properly.")

del sys, __methods__, __inner_classes__,  __base_methods__, method_name, instance_name, instance, base, dtparse, attr
