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
"""Unix platform directory lookup with environment-based path defaults."""

import os
from pathlib import Path as PathlibPath
import sys


class _unix_dir:
    def __init__(self):
        self._home = self._get_home_directory()
        self._downloads = self._get_downloads()
        if sys.platform.startswith("darwin"):
            self._applications = self._get_applications()
        else:
            self._applications = None
        self._system = self._get_system_directory()

    def _get_home_directory(self):
        return PathlibPath.home()

    def _get_downloads(self):
        downloads = self._home / "Downloads"
        if downloads.exists() and downloads.is_dir():
            return downloads
        else:
            raise FileNotFoundError("The Downloads directory does not exist.")

    if sys.platform.startswith("darwin"):
        def _get_applications(self):
            """Retrieve the applications directory (macOS only)."""
            applications = PathlibPath("/Applications")
            if applications.exists() and applications.is_dir():
                return applications
            return None

    def _get_system_directory(self):
        """Retrieve the system directory (e.g., /usr or /System)."""
        if os.name == "posix":
            system_dir = PathlibPath("/System") if PathlibPath("/System").exists() else PathlibPath("/usr")
            return system_dir
        return None

    @property
    def Downloads(self):
        return self._downloads           
           
    @property
    def System(self):
        return self._system

    @property
    def UserBin(self):
        return self._validate_directory("/usr/bin")

    @property
    def LocalBin(self):
        return self._validate_directory("/usr/local/bin")

    @property
    def Opt(self):
        return self._validate_directory("/opt")

    @property
    def LocalUserBin(self):
        return self._validate_directory("~/.local/bin", resolve_user=True)

    @property
    def LocalShare(self):
        return self._validate_directory("~/.local/share", resolve_user=True)


    if sys.platform.startswith("darwin"):
        @property
        def Applications(self):
            return self._applications
           
    if sys.platform.startswith("linux"):
        @property
        def Snap(self):
            return self._validate_directory("/snap")

        @property
        def ShareApplications(self):
            return self._validate_directory("/usr/share/applications")
           
        @property
        def FlatpakSystem(self):
            return self._validate_directory("/var/lib/flatpak")       
           
        @property
        def FlatpakUser(self):
            return self._validate_directory("~/.local/share/flatpak", resolve_user=True)
           

    def _validate_directory(self, path, resolve_user=False):
        """Validate that the directory exists and return its path."""
        if resolve_user:
            path = PathlibPath(path).expanduser()
        else:
            path = PathlibPath(path)

        if path.exists() and path.is_dir():
            return path
        else:
            raise FileNotFoundError(f"The directory '{path}' does not exist or is not a directory.")

    def Home(self, directory=None):
        """Return the current user's home directory or an existing path relative to its parent."""
        if directory:
            homeDir = PathlibPath.home()
            directory = PathlibPath(directory)
            fpath = PathlibPath(homeDir).parent
            dirpath = fpath / directory
            if dirpath.exists() and dirpath.is_dir():
                return str(dirpath)
            else:
                raise FileNotFoundError(f"The specified directory '{directory}' does not exist or is not a directory.")
        else:
            return PathlibPath.home()  

    def __dir__(self):
        """List directory helpers available on the current platform."""
        if sys.platform.startswith("darwin"):
            return ['Downloads', 'System', 'Home', 'Applications', 'UserBin', 'LocalBin', 'Opt', 'LocalUserBin', 'LocalShare']
        elif sys.platform.startswith("linux"):
            return ['Downloads', 'System', 'Home', 'UserBin', 'LocalBin', 'Opt', 'LocalUserBin', 'LocalShare', 'Snap', 'ShareApplications', 'FlatpakSystem', 'FlatpakUser']


__all__ = ['_unix_dir']
def __dir__(): return __all__
