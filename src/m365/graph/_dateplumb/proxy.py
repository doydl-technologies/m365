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

"""Lazy module or attribute access through a cached proxy.

The target is imported on first use. Calls and attribute operations delegate
to the resolved target. First access is not synchronized between threads.
"""
import importlib
from typing import Any, TypeVar, Generic, Optional


__all__ = ['Proxy']





_T = TypeVar("_T")

class Proxy(Generic[_T]):
    """Defer importing module_name and resolving an optional attr_name.

    With attr_name=None the target is the module itself. Calls and attribute reads
    load and cache the target. Import errors propagate; missing attributes raise
    AttributeError. First access is not synchronized.
    """
    def __init__(self, module_name: str, attr_name: Optional[str] = None) -> None:    
        """Store the module name and optional attribute without importing the target."""
        self._module_name = module_name
        self._attr_name = attr_name
        self._wrapped: Any = None

    def _load(self) -> _T:
        """Import and cache the module or requested attribute on first use."""
        if self._wrapped is None:
            module = importlib.import_module(self._module_name)
            if self._attr_name is not None:
                self._wrapped = getattr(module, self._attr_name)
            else:
                self._wrapped = module
        return self._wrapped  # type: ignore        

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        """Call the resolved target with the supplied arguments."""
        return self._load()(*args, **kwargs)

    def __getattr__(self, name: str) -> Any:
        """Read an attribute from the resolved target."""
        return getattr(self._load(), name)

    def __repr__(self):
        """
        Return a string representation of the Proxy instance.

        Returns:
        -----------
            str :
                A descriptive string indicating the target module and attribute,
                and whether the target has been loaded.

        Notes:
        -----------
            - Helps with debugging and logging by displaying the loading state.
            - The representation changes after the attribute is loaded.
        """
        if self._wrapped is not None:
            return f"<Proxy for {self._module_name}.{self._attr_name} (loaded: {self._wrapped!r})>"
        return f"<Proxy for {self._module_name}.{self._attr_name} (not loaded)>"

def __dir__():
    return __all__
