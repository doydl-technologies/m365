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

import math
from collections.abc import Iterable
from copy import deepcopy

import numpy as np

def deduplicate(*lists):
    """Deduplicate values while preserving their first occurrence."""
    def _flat(l):
        """Flatten nested lists."""
        stack = list(reversed(l))
        flattened = []
        while stack:
            item = stack.pop()
            if isinstance(item, list):
                stack.extend(reversed(item))
            else:
                flattened.append(item)
        return flattened
    
    result = []
    for items in lists:
        seen = set()
        unique_list = []
        for item in items:
            if item not in seen:
                unique_list.append(item)
                seen.add(item)
        result.append(unique_list)
    return tuple(result) if len(result) > 1 else _flat(result)

def remove_item_from_list(item_list, item_to_remove=None, drop_NA=False, inplace=False, verbose=False):
    """
    Removes an item or missing values from the provided list.

    Args:
    item_list (list): The list from which to remove the item.
    item_to_remove (any): The item to be removed from the list.
    drop_NA (bool): If True, removes missing values from the list.
    inplace (bool): If True, modifies the original list. If False, returns a new list.
    verbose (bool): If True, prints messages about the operation results.

    Returns:
    list: The original or updated list after attempting to remove the item or missing values.
    """
    def is_missing(value):
        return (isinstance(value, float) and math.isnan(value)) or value is None or value is pd.NaT or (isinstance(value, float) and np.isnan(value))
    
    if drop_NA:
        if inplace:
            i = 0
            while i < len(item_list):
                if is_missing(item_list[i]):
                    item_list.pop(i)
                else:
                    i += 1
            if verbose:
                print("Missing values have been removed.")
            return item_list
        else:
            new_list = [item for item in item_list if not is_missing(item)]
            if verbose:
                print("Missing values have been removed from the new list.")
            return new_list

    if item_to_remove is not None:
        if inplace:
            try:
                item_list.remove(item_to_remove)
                if verbose:
                    print(f"Item '{item_to_remove}' has been removed successfully.")
            except ValueError:
                if verbose:
                    print(f"Item '{item_to_remove}' not found in the list.")
            return item_list
        else:
            new_list = [item for item in item_list if item != item_to_remove]
            if verbose:
                removed_count = len(item_list) - len(new_list)
                if removed_count:
                    print(f"Item '{item_to_remove}' removed {removed_count} times.")
                else:
                    print(f"Item '{item_to_remove}' not found in the list.")
            return new_list

    if verbose:
        print("No operation performed.")
    return item_list

def compare(list1, list2, operation):
    """
    Compare two lists based on the specified operation.

    :param list1: First list to compare.
    :param list2: Second list to compare.
    :param operation: Type of comparison ('diff' for differences, 'alike' for similarities).
    :return: List of elements based on the specified operation.
    """
    if operation == 'diff':
        return list(set(list1) - set(list2))
    elif operation == 'alike':
        return list(set(list1) & set(list2))
    else:
        raise ValueError("Invalid operation. Choose 'diff' or 'alike'.")

def flatten(l, exhaustive=False, recurse=False):
    """
    Flatten a nested data structure (e.g., lists, tuples, sets, and dictionaries)
    into a single, flat list.

    Parameters
    ----------
    container : iterable
        The input data structure to be flattened. It can be a list, tuple, set,
        dictionary, or a combination of these types.
    
    exhaustive : bool, optional (default=False)
        If True, both keys and values from dictionaries will be included in the
        flattened result. If False, only the values from dictionaries will be included.
    
    recurse : bool, optional (default=False)
        If True, the function uses a recursive approach to flatten the data structure.
        If False, it uses an iterative approach with a stack.

    Returns
    -------
    list
        A flat list containing all elements from the input data structure.

    Notes
    -----
    - Strings and bytes are treated as atomic elements and will not be flattened
      into individual characters.
    - The function handles nested lists, tuples, sets, and dictionaries, including
      combinations of these types.
    - If `exhaustive=True` and `recurse=True`, the function will deeply flatten all
      elements, including both keys and values from dictionaries.
    """	
    def flatten_recursive(item):
        """Recursively flatten the structure."""
        if isinstance(item, (list, tuple, set)):
            for subitem in item:
                yield from flatten_recursive(subitem)
        elif isinstance(item, dict):
            if exhaustive:
                for key, value in item.items():
                    yield from flatten_recursive(key)
                    yield from flatten_recursive(value)
            else:
                for value in item.values():
                    yield from flatten_recursive(value)
        else:
            yield item

    def flatten_iterative(item):
        """Iteratively flatten the structure using a stack."""
        stack = [item]
        while stack:
            current = stack.pop()
            if isinstance(current, (list, tuple, set)):
                stack.extend(reversed(current))  # Reverse to maintain the order
            elif isinstance(current, dict):
                if exhaustive:
                    for pair in reversed(list(current.items())):
                        stack.extend(pair)
                else:
                    stack.extend(reversed(list(current.values())))
            else:
                yield current

    if recurse:
        return list(flatten_recursive(l))
    else:
        return list(flatten_iterative(l))
       
def force_n(l, count=1):
    """
    Ensures the input is returned as a list with up to `count` items.
    
    If the input is a non-empty list or iterable, it returns a list containing
    up to the first `count` items of the iterable.
    If the input is a scalar value or an empty list/iterable, it returns a list
    with the input as its sole element or an empty list if the input is empty.

    Parameters:
    - l: The input to be processed, which can be a list, tuple, set,
      string, bytes, or scalar value.
    - count (int): The number of items to include in the returned list. Defaults to 1.

    Returns:
    - list: A list containing up to `count` items from the input or the input
      itself if it's not iterable.
    """
    count = int(str(count))
    if count == 0:
        count = 1
    if l is None: 
        return []
    if isinstance(l, (str, bytes)):  # Treat strings and bytes as atomic values
        return [l]
    if isinstance(l, Iterable):
        input_list = list(l)
        count = min(count, len(input_list))
        return input_list[:count]
    return [l]
   
def unlist_single(l):
    """Return the first item of a nonempty list, None for an empty list, or the input otherwise."""
    if isinstance(l, list):
        return l[0] if l else None
    return l

def ensure_single_item(collection):
    """
    Ensures that the input value is returned as a single item.
    """
    if isinstance(collection, (set, frozenset)):
        if len(collection) > 1:
            raise ValueError("Unordered collection is ambiguous; supply one item or an ordered sequence")
        return next(iter(collection), None)
    if isinstance(collection, (list, tuple)) and collection:
        return collection[0]
    elif isinstance(collection, (list, tuple)) and not collection:
        return None
    else:
        return collection   
   

__all__ = [
    "deduplicate",
    "remove_item_from_list",
    "compare",
    "flatten",
    "force_n",
    "unlist_single",
    "ensure_single_item",
]
def __dir__(): return __all__
