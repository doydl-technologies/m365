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
"""Expose shared filesystem helpers through the host pathdir instance.

The facade binds path inspection, extension handling, search, directory lookup,
stream operations and serialization helpers from syslib.base. base_path is
resolved dynamically from the shared host rather than copied at import time.
Filesystem and cloud-path patterns describe syntax, not remote resource access.
"""
import sys

from .base import host as __syslib

__all__ = [
    'base_path','chdir','cleanup_cache','create_data_directory','files','full_path',
    'invalidate_cache','path_type','search','directories','is_path_struct','join',
    'kind','get_module_dir','resetdir','Text','File','Dirs','Save','Buffer','Stream',
    'Load','delete'
]

_exposed_methods = [
    'chdir','cleanup_cache','create_data_directory','files','full_path','invalidate_cache',
    'path_type','search','directories','is_path_struct','join','kind','get_module_dir','resetdir'
]

# On Windows, also expose roaming_path and create_shortcut if available.
if sys.platform.startswith("win") and hasattr(__syslib, 'roaming_path') and hasattr(__syslib, 'create_shortcut'):
    _exposed_methods.append('roaming_path')
    _exposed_methods.append('create_shortcut')    

for method_name in _exposed_methods:
    if hasattr(__syslib, method_name):
        globals()[method_name] = getattr(__syslib, method_name)

_inner_instances = ['Text', 'File', 'Dirs', 'Save', 'Buffer', 'Stream', 'Load']
for instance_name in _inner_instances:
    instance = getattr(__syslib, instance_name, None)
    if instance is not None:
        globals()[instance_name] = instance

def __getattr__(name):
    """Resolve base_path dynamically from the shared host state."""
    if name == 'base_path':
        return __syslib.base_path
    raise AttributeError(f"module {__name__} has no attribute {name}")

del sys, _exposed_methods, _inner_instances, method_name, instance_name, instance, base
del _loc, _paradigm
