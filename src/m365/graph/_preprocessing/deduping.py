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

import json

def find_identical_items(data, keys_to_compare=None):
    """
    Finds and returns identical items in a list of dictionaries,
    including nested structures, based on specific keys.

    Args:
        data (list): A list of dictionaries or nested dictionaries.
        keys_to_compare (list): Optional list of keys to consider for comparison.
                                If None, compares all keys.

    Returns:
        list: A list of tuples, each containing identical items.
    """

    def serialize_dict(d):
        """Serialize a dictionary based on specific keys to a JSON string."""
        if keys_to_compare is not None:
            filtered_dict = {key: d[key] for key in keys_to_compare if key in d}
            return json.dumps(filtered_dict, sort_keys=True)
        else:
            return json.dumps(d, sort_keys=True)

    seen = {}
    duplicates = []

    def flatten_and_identify(item):
        """Recursively flatten any nested lists of dictionaries and identify duplicates."""
        if isinstance(item, dict):
            serialized_item = serialize_dict(item)
            if serialized_item in seen:
                seen[serialized_item].append(item)
            else:
                seen[serialized_item] = [item]
        elif isinstance(item, list):
            for sub_item in item:
                flatten_and_identify(sub_item)

    for element in data:
        flatten_and_identify(element)

    for items in seen.values():
        if len(items) > 1:
            duplicates.append(tuple(items))

    return duplicates
   
def find_identical_dicts(dicts, ignore_keys=None):
    """
    Finds identical dictionaries in a list of dictionaries, ignoring specified keys.

    Args:
        dicts (list): A list of dictionaries to check for identical entries.
        ignore_keys (list): List of keys to ignore during comparison.

    Returns:
        list: A list of tuples, where each tuple contains the indices of identical dictionaries.
    """
    identical_pairs = []
    ignore_keys = ignore_keys or ["@microsoft.graph.downloadUrl"]

    simplified_dicts = []
    for d in dicts:
        simplified_dict = {k: v for k, v in d.items() if k not in ignore_keys}
        simplified_dicts.append(simplified_dict)

    for i in range(len(simplified_dicts)):
        for j in range(i + 1, len(simplified_dicts)):
            if simplified_dicts[i] == simplified_dicts[j]:
                identical_pairs.append((i, j))

    return identical_pairs   

def remove_identical_dicts(dicts, ignore_keys=None):
    """Keep the first dictionary in each group of duplicates, ignoring selected keys."""
    ignore_keys = ignore_keys or ["@microsoft.graph.downloadUrl"]
    unique_dicts = []
    seen = []

    for index, d in enumerate(dicts):
        simplified_dict = {k: v for k, v in d.items() if k not in ignore_keys}
        
        if not any(simplified_dict == seen_dict for seen_dict in seen):
            unique_dicts.append(d)
            seen.append(simplified_dict)

    return unique_dicts



__all__ = [
    "find_identical_items",
    "find_identical_dicts",
    "remove_identical_dicts",
]
def __dir__(): return __all__
