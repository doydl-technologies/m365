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
"""Compiled structural patterns for local and cloud-provider paths.

These expressions classify supported path shapes. They are not complete
operating-system filename validators, existence checks or authorization checks.
Provider bucket rules and structural rules are separate.
"""
import re
# Optional drive prefix; accepts either separator and requires an extension.
# Structural match only: this does not enforce reserved names or UNC rules.
__WINDOWS_FILE_PATH_REGEX__ = re.compile(r"^(?:[a-zA-Z]:)?[\\/]?[\w\s.-]*(?:[\\/][\w\s.-]+)*\.[a-zA-Z0-9]+$")
# File-shaped path with an extension; both slash styles are accepted.
# The name is retained for compatibility, not strict POSIX validation.
__UNIX_FILE_PATH_REGEX__ = re.compile(r'^[\\/]?[\w\s.-]*(?:[\\/][\w\s.-]+)*\.[a-zA-Z0-9]+$')
# Drive-absolute directory shape with word/space/hyphen components.
# Trailing separators and dotted components are excluded.
__WINDOWS_DIRECTORY_PATH_REGEX__ = re.compile(r"^[a-zA-Z]:[\\/][\w\s-]+(?:[\\/][\w\s-]+)*$")
# ASCII alphanumeric/space/underscore/hyphen directory components.
# Outer slashes are optional; dot-relative and dotted components are excluded.
__UNIX_DIRECTORY_PATH_REGEX__ = re.compile(r'^\/?([A-Za-z0-9 _-]+\/)*[A-Za-z0-9 _-]+\/?$')
# Drive-absolute, backslash-separated path shape, without a UNC prefix.
# Other component characters are not checked against Windows filename rules.
__WINDOWS_PATH_STRUCTURE_REGEX__ = re.compile(r'^[A-Za-z]:\\(?:[^\\]+(?:\\|$))*$')
# Absolute slash-separated path, or dot/dot-dot-relative form.
# Components only exclude slashes; the absolute-form branch also matches empty input.
__UNIX_PATH_STRUCTURE_REGEX__ = re.compile(r'^(\/[^\/]+)*\/?$|^\.{1,2}(\/[^\/]+)*\/?$')
# Separator-based path shape rejecting drive-backslash prefixes per segment.
# This is not a provider-specific path validator.
__CLOUDPROVIDER_TRADITIONAL_PATH_STRUCTURE_REGEX__ = re.compile(r'^[\\/]?((?![A-Za-z]:\\)[^\\/]+[\\/])+[^\\/]*[\\/]?$')
# Slash-separated alphanumeric/space/underscore/hyphen folder components.
# Excludes drive prefixes without enforcing provider-specific naming constraints.
__CLOUDPROVIDER_TRADITIONAL_FOLDER_PATH_REGEX__ = re.compile(r'^\/?((?![A-Za-z]:)[A-Za-z0-9 _-]+\/)*[A-Za-z0-9 _-]+\/?$')
# File-shaped path with a 1-10 character alphanumeric extension.
# Rejects drive prefixes and accepts either slash style.
__CLOUDPROVIDER_TRADITIONAL_FILE_PATH_REGEX__ = re.compile(r'^(?![A-Za-z]:)[\\/]?[\w\s.-]*(?:[\\/][\w\s.-]+)*\.[a-zA-Z0-9]{1,10}$')
# Capture the first slash-delimited component, allowing an optional leading slash.
__CLOUDPROVIDER_BUCKET_NAME_EXISTENCE_REGEX__ = re.compile(r'^/?([^/]+)')
# Case-insensitive provider aliases. Anchors apply to individual alternatives;
# this expression is not an exact-name validator for every alias.
__CLOUDPROVIDER_SELECTION_REGEX__ = re.compile(r'^(?P<aws>amazon|s3|aws)|(?P<wasabi>wasabi)|(?P<backblaze>backblaze|bb|backblazeb2|backblaze b2|b2|bb2|bbb2)$', re.IGNORECASE)
# Bucket shape: 3-63 lowercase alphanumeric/dot/hyphen characters,
# alphanumeric ends, no consecutive dots and no IPv4-shaped value.
__CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_CASESENSITIVE__ = re.compile(r'^(?!.*\.\.)(?!.*-$)(?!^-)(?!^\d+\.\d+\.\d+\.\d+$)([a-z0-9][a-z0-9.-]{1,61}[a-z0-9])$')
# The same bucket-shape constraints, with uppercase letters also allowed.
__CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_IGNORECASE__ = re.compile(r'^(?!.*\.\.)(?!.*-$)(?!^-)(?!^\d+\.\d+\.\d+\.\d+$)([a-zA-Z0-9][a-zA-Z0-9.-]{1,61}[a-zA-Z0-9])$')

# Check reserved forms against the entire name before consuming it. The caller
# checks the general bucket shape separately. The -an suffix depends on the
# bucket namespace and is not rejected by this namespace-agnostic path check.
__CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_S3_SPECIFIC__ = re.compile(
    r'\A(?!xn--|sthree-|amzn-s3-demo-)(?!.*(?:-s3alias|--ol-s3|\.mrap|--x-s3|--table-s3)\Z).+\Z'
)


__all__ = [
    "__WINDOWS_FILE_PATH_REGEX__",
    "__UNIX_FILE_PATH_REGEX__",
    "__WINDOWS_DIRECTORY_PATH_REGEX__",
    "__UNIX_DIRECTORY_PATH_REGEX__",
    "__WINDOWS_PATH_STRUCTURE_REGEX__",
    "__UNIX_PATH_STRUCTURE_REGEX__",
    "__CLOUDPROVIDER_TRADITIONAL_PATH_STRUCTURE_REGEX__",
    "__CLOUDPROVIDER_TRADITIONAL_FILE_PATH_REGEX__",
    "__CLOUDPROVIDER_TRADITIONAL_FOLDER_PATH_REGEX__",
    "__CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_CASESENSITIVE__",
    "__CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_IGNORECASE__",
    "__CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_S3_SPECIFIC__",
    "__CLOUDPROVIDER_BUCKET_NAME_EXISTENCE_REGEX__",
    "__CLOUDPROVIDER_SELECTION_REGEX__",
]
def __dir__(): return __all__
