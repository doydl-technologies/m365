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
"""URL-preserving compatibility helpers; never apply filesystem normalization."""
import posixpath
from urllib.parse import urlsplit, urlunsplit, unquote

from ._site_resolution import library_relative


def strip_file_extension(value):
    if not isinstance(value, str):
        raise TypeError("Paths must be strings")
    parsed = urlsplit(value)
    path, _ = posixpath.splitext(parsed.path)
    return urlunsplit((parsed.scheme, parsed.netloc, path, parsed.query, parsed.fragment))


def drop_library_prefix(client, value):
    if not isinstance(value, str):
        raise TypeError("Paths must be strings")
    if urlsplit(value).scheme or value.startswith("//"):
        return library_relative(client, value)
    value = unquote(value)
    first, separator, remainder = value.lstrip("/").partition("/")
    aliases = {client.drive_name} if client.drive_name else set()
    if client.drive_web_url:
        aliases.add(unquote(urlsplit(client.drive_web_url).path.rstrip("/").rsplit("/", 1)[-1]))
    # Strip ONE leading library component only; preserve real nested folders.
    return remainder if separator and first in aliases else value


def relative_path(client, value, site_name=None, remove_leading_slash=True,
                  remove_domain_extension=False, drop_drive=False):
    if not isinstance(value, str):
        raise TypeError("Paths must be strings")
    parsed = urlsplit(value)
    if parsed.scheme or parsed.netloc:
        if drop_drive:
            result = library_relative(client, value)
        else:
            base = urlsplit(client.web_url)
            path, base_path = unquote(parsed.path), unquote(base.path).rstrip("/")
            if (parsed.scheme != "https" or parsed.netloc.lower() != base.netloc.lower()
                    or not (path == base_path or path.startswith(base_path + "/"))):
                raise ValueError("URL is outside the bound SharePoint site")
            result = path[len(base_path):].lstrip("/")
    else:
        result = drop_library_prefix(client, value) if drop_drive else unquote(value)
    if remove_domain_extension:
        result = strip_file_extension(result)
    return result.lstrip("/") if remove_leading_slash else "/" + result.lstrip("/")
