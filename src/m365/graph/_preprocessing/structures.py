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

def count_folders(data):
    """
    Count the number of folders in a nested structure.
    
    Parameters:
        data (list): List of dictionaries representing folders and their children.
        
    Returns:
        int: The total number of folders in the structure.
    """
    folder_count = 0

    def count_item(item):
        nonlocal folder_count
        # If the item has 'children', it is considered a folder
        if "children" in item and isinstance(item["children"], list):
            folder_count += 1
            for child in item["children"]:
                count_item(child)

    for top_level_item in data:
        count_item(top_level_item)

    return folder_count
   
def count_items(data):
    folder_count = 0
    file_count = 0

    def count_item(item):
        nonlocal folder_count, file_count
        if 'folder' in item and isinstance(item['folder'], list):
            folder_count += 1
            for x in item['folder']:
                count_item(x)
        if 'file' in item:
            file_count += 1

    for top_level_item in data:
        count_item(top_level_item)

    return folder_count, file_count
   
   
def flatten_separate(data):
    """
    Flatten the folder structure so that the top-level folder is the first element
    and all subsequent child items follow sequentially, excluding specified keys.
    
    Parameters:
        data (list): List of dictionaries representing folders and their children.
        
    Returns:
        list: A flattened list where the top-level folder is the first element
              and all children follow in order, with specific keys excluded.
    """
    flattened = []
    excluded_keys = {'eTag', 'cTag', 'shared', 'children'}

    def flatten_item(item):
        """
        Recursively flatten a folder or file and its children, excluding specific keys.
        """
        item_copy = {k: v for k, v in item.items() if k not in excluded_keys}
        flattened.append(item_copy)
        
        if "children" in item and isinstance(item["children"], list):
            for child in item["children"]:
                flatten_item(child)

    # Only the first item is treated as the root.
    for top_level_item in data:
        flatten_item(top_level_item)

    return flattened

def flatten_into_child(data):
    """
    Flatten the folder structure so that only child items appear in the list,
    each including data from its respective parent folder.
    Specific keys ('eTag', 'cTag', 'shared') are excluded from the data.
    
    Parameters:
        data (list): List of dictionaries representing folders and their children.
        
    Returns:
        list: A flattened list where only children are listed, each including its parent folder's data.
    """
    flattened = []
    excluded_keys = {'eTag', 'cTag', 'shared', 'children'}

    def flatten_item(item, parent_data=None):
        """
        Recursively flatten a folder or file and its children, embedding parent data into each child
        and excluding specific keys.
        
        Args:
            item (dict): Current item to flatten.
            parent_data (dict): Data of the parent item, passed down to children.
        """
        item_copy = {k: v for k, v in item.items() if k not in excluded_keys}

        if parent_data:
            item_copy['parent_info'] = {k: v for k, v in parent_data.items() if k not in excluded_keys}
            # Emit only descendants, not top-level folders.
            flattened.append(item_copy)

        if "children" in item and isinstance(item["children"], list):
            parent_data_copy = {k: v for k, v in item_copy.items()}
            for child in item["children"]:
                flatten_item(child, parent_data_copy)

    # Traverse every top-level folder.
    for top_level_item in data:
        flatten_item(top_level_item)

    return flattened


__all__ = [
    "count_folders",
    "count_items",
    "flatten_separate",
    "flatten_into_child",
]
def __dir__(): return __all__
