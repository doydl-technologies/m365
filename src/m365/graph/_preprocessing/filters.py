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

import re

__any_regex__ = r"\S+"


def filter_dicts_by_key_value(data_structure, key, value, exact=True, search_nested=False, dynamic_search=False):
    """
    Filters dictionaries by key and value from a nested data structure,
    with options to dynamically switch to nested search and control the application of dynamic search.
    
    1. List of Dictionaries
    2. List of Dictionaries Containing Nested Lists of Dictionaries
    3. List of Dictionaries with Nested Dictionaries
    4. Dictionary of Lists Containing Dictionaries
    5. List of Lists of Dictionaries
    6. Dictionary with Nested Lists and Dictionaries
    7. Dictionary with Nested Dictionaries and Lists
    8. Complex Multi-level Nesting    
    
    """
    def _root_key_exists(data, key):
        """
        Traverses nested lists until dictionaries are found, then checks for a key
        at the root level of each dictionary found at the point of discovery in the nesting.
        """
        if isinstance(data, dict):
            return key in data
        elif isinstance(data, list):
            for item in data:
                if isinstance(item, (list, dict)):
                    if _root_key_exists(item, key):
                        return True
        return False
    
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
       
    def _recurse_items(items, root):
        """
        Recursively searches through the data structure to find dictionaries
        that match the specified key and value.
        """
        if isinstance(items, dict):
            if ((exact and items.get(key) == value) or (not exact and value in str(items.get(key)))) and (root not in matched_roots):
                matched_roots.append(root)
            if search_nested:
                for v in items.values():
                    if isinstance(v, (dict, list)):
                        _recurse_items(v, root)
        elif isinstance(items, list):
            for item in items:
                _recurse_items(item, root)

    # Only switch to nested matching when root-level keys are absent.
    if dynamic_search and not _root_key_exists(data_structure, key):
        search_nested = True

    data_structure = data_structure if isinstance(data_structure, list) else [data_structure]
    value = value[0] if isinstance(value, list) else value
    matched_roots = []

    if isinstance(data_structure, list):
        for item in data_structure:
            _recurse_items(item, item)
    else:
        _recurse_items(data_structure, data_structure)

    return _flat(matched_roots) if matched_roots else matched_roots

def find_items_by_key(data, itemType):
    """
    Recursively filter a list of dictionaries (or nested lists of dictionaries)
    to return only the items that contain the specified key (itemType).
    
    Args:
    - data (list or dict): A list of dictionaries or a list containing nested lists/dictionaries.
    - itemType (str): The key to filter for (e.g., 'file').
    
    Returns:
    - list: A list of dictionaries that contain the specified key.
    """
    def recursive_filter(items):
        """Filter matching items in a nested list."""
        result = []
        if isinstance(items, list):
            for item in items:
                if isinstance(item, dict):
                    if itemType in item:
                        result.append(item)
                    else:
                        result.extend(recursive_filter(item.get('items', [])))
                elif isinstance(item, list):
                    result.extend(recursive_filter(item))
        elif isinstance(items, dict):
            for key, value in items.items():
                if isinstance(value, list):
                    result.extend(recursive_filter(value))
                elif isinstance(value, dict):
                    result.extend(recursive_filter([value]))
        return result
    return recursive_filter(data)
   
def xfilter(d, filter_key, filter_value, regex=False):
    """
    Filters a complex data structure, recursively checking dictionaries and lists for the presence
    of a specified key-value pair or a key with a value matching a regex pattern.

    Parameters:
        d (dict or list): The data structure to be filtered, which can be a dictionary, a list, or a nested combination.
        filter_key (str): The key to search for within the data structure.
        filter_value (Any): The value associated with `filter_key` that the function looks for to include an element in the filtered result.
        regex (bool): If True, treat `filter_value` as a regex pattern to match against the values.

    Returns:
        dict, list, or None: 
            - Returns a filtered dictionary or list where only elements containing the specified key-value pair are included.
            - If the filter criteria is not met anywhere in the data structure, returns `None`.
    """
    if regex:
        pattern = re.compile(filter_value)
    
    def recursive_filter(data):
        if isinstance(data, dict):
            filtered_dict = {}
            for key, value in data.items():
                if isinstance(value, (dict, list)):
                    filtered_value = recursive_filter(value)
                    if filtered_value:
                        filtered_dict[key] = filtered_value
                elif key == filter_key:
                    if (regex and pattern.search(str(value))) or (not regex and value == filter_value):
                        return data
            return filtered_dict if filtered_dict else None
        
        elif isinstance(data, list):
            filtered_list = []
            for item in data:
                filtered_item = recursive_filter(item)
                if filtered_item:
                    filtered_list.append(filtered_item)
            return filtered_list if filtered_list else None
        return None
    return recursive_filter(d)[0].get(filter_key) if isinstance(recursive_filter(d), list) and recursive_filter(d) else recursive_filter(d).get(filter_key) if recursive_filter(d) and not isinstance(recursive_filter(d), list) else recursive_filter(d)

def find_first_instance(data):
    """Return the first dictionary found in nested lists, or None."""
    if isinstance(data, dict):
        return data 
    elif isinstance(data, list):
        for item in data:
            result = find_first_instance(item)
            if result is not None:
                return [result] if not isinstance(result, list) else result
    return None 

def root_key_exists(data, key):
    """
    Traverses nested lists until dictionaries are found, then checks for a key
    at the root level of each dictionary found at the point of discovery in the nesting.

    Parameters:
    ----------
    data : any
        The nested data structure to be searched, which can be deeply nested with lists and dictionaries.

    key : str
        The key to search for at the root level of the found dictionaries.

    Returns:
    -------
    bool
        True if the key exists at the root level of any discovered dictionary; False otherwise.
    """
    if isinstance(data, dict):
        return key in data

    elif isinstance(data, list):
        for item in data:
            if isinstance(item, (list, dict)):
                if root_key_exists(item, key):
                    return True
    return False

def filter_deep(data_structure, key, value, exact=True, search_nested=False, dynamic_search=False, use_regex=False, depth="outer"):
    """
    Filters dictionaries by a specified key and value within a complex, potentially nested data structure.
    This function allows for a flexible approach to search, supporting exact and partial matches, regular expression evaluation, and control over the search depth within nested structures.
    
    Parameters:
        data_structure (list or dict): The data structure to be searched. This can be a single dictionary,
            a list of dictionaries, or a more complex nested structure containing lists and dictionaries at various levels.
        key (str): The dictionary key to search for.
        value (str or pattern): The value to search for, which can be a string for direct comparison or a pattern if `use_regex` is True.
        exact (bool, optional): If True, the function performs an exact match against the value. If False, it checks if the value is a substring of the dictionary value. Defaults to True.
        search_nested (bool, optional): If True, enables the search within nested dictionaries and lists. If False, the search is limited to the first level of the data structure. Defaults to False.
        dynamic_search (bool, optional): If True, the function dynamically adjusts to perform a nested search if the key is not found at the top level of the data structure. Defaults to False.
        use_regex (bool, optional): If True, treats `value` as a regular expression pattern, allowing for flexible matching based on regex rules. Defaults to False.
        depth (str, optional): Controls the depth of the return value for found items. "outer" returns the top-most dictionary in the nested structure that contains the key, while "inner" returns the direct dictionary that contains the key. 
        This parameter is ignored if `search_nested` is False. Defaults to "outer".

    Returns:
        list: A list of dictionaries that match the given criteria. If no matches are found, returns an empty list.

    Examples:
        # Example data structure
        data = [
            {'name': 'Alice', 'age': 30, 'contacts': [{'email': 'alice@example.com'}, {'email': 'alice@work.com'}]},
            {'name': 'Bob', 'age': 25, 'contacts': [{'email': 'bob@example.com'}]},
            {'name': 'Charlie', 'nested': {'email': 'charlie@example.com'}}
        ]

        # Basic search
        print(filter_dicts_by_key_value(data, 'email', 'alice@example.com', exact=True))

        # Nested search
        print(filter_dicts_by_key_value(data, 'email', 'example.com', exact=False, search_nested=True, depth='inner'))

        # Using regular expressions
        print(filter_dicts_by_key_value(data, 'email', r'\\bexample\\.com$', use_regex=True, search_nested=True))

    Note:
        - Regular expression searches do not check for exact matches but search for patterns within the string values.
        - When `search_nested` is False and `depth` is set to "inner", `depth` will have no effect because there is no nesting to consider.
    """
    def _root_key_exists(data, key):
        if isinstance(data, dict):
            return key in data
        elif isinstance(data, list):
            for item in data:
                if isinstance(item, (list, dict)):
                    if _root_key_exists(item, key):
                        return True
        return False

    def _flat(l):
        stack = list(reversed(l))
        flattened = []
        while stack:
            item = stack.pop()
            if isinstance(item, list):
                stack.extend(reversed(item))
            else:
                flattened.append(item)
        return flattened

    def _matches_criteria(item_value):
        if use_regex:
            return re.search(value, str(item_value)) is not None
        else:
            return (exact and item_value == value) or (not exact and value in str(item_value))

    def _recurse_items(items, root, parent):
        if isinstance(items, dict):
            if _matches_criteria(items.get(key)):
                if depth == "outer" or not search_nested:
                    if root not in matched_roots:
                        matched_roots.append(root)
                elif depth == "inner":
                    if items not in matched_roots:
                        matched_roots.append(items)
            if search_nested:
                for v in items.values():
                    if isinstance(v, (dict, list)):
                        _recurse_items(v, root, items)
        elif isinstance(items, list):
            for item in items:
                _recurse_items(item, root, parent)

    if dynamic_search and not _root_key_exists(data_structure, key):
        search_nested = True

    data_structure = data_structure if isinstance(data_structure, list) else [data_structure]
    value = value[0] if isinstance(value, list) and not use_regex else value
    matched_roots = []

    if isinstance(data_structure, list):
        for item in data_structure:
            _recurse_items(item, item, None)
    else:
        _recurse_items(data_structure, data_structure, None)
    return _flat(matched_roots) if matched_roots else matched_roots



__all__ = [
    "filter_dicts_by_key_value",
    "find_items_by_key",
    "xfilter",
    "find_first_instance",
    "root_key_exists",
    "filter_deep",
]
def __dir__(): return __all__
