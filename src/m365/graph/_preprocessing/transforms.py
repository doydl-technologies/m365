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
from copy import deepcopy
from datetime import datetime, timezone, tzinfo
from .._core_dates import try_iso

from .. import _dateplumb as dateplumb

def format_key_value(data_structure, keys, func, *args, **kwargs):
    """Apply a function to selected values in a copy of a nested structure."""
    def _format(obj, key, func, *args, **kwargs):
        if isinstance(obj, dict):
            for k, v in list(obj.items()):  # Convert to list for safe iteration
                if k == key:
                    try:
                        obj[k] = func(v, *args, **kwargs)
                    except Exception as e:
                        print(f"Error converting value {v} for key {k}: {e}")
                elif isinstance(v, (dict, list)):
                    obj[k] = _format(v, key, func, *args, **kwargs)
        elif isinstance(obj, list):
            for i in range(len(obj)):
                obj[i] = _format(obj[i], key, func, *args, **kwargs)
        return obj
    new_data_structure = deepcopy(data_structure)
    if not isinstance(keys, list):
        keys = [keys]
    for key in keys:
        _format(new_data_structure, key, func, *args, **kwargs)
    return new_data_structure

def modify_key_value(data_structure, func, *args, **kwargs):
    """Apply a function to item dictionaries in a copy of a nested structure.

    Only dictionaries with both parentReference and name are passed to the function.
    """
    def _format(obj, func, *args, **kwargs):
        if isinstance(obj, dict):
            try:
                if 'parentReference' in obj and 'name' in obj:
                    func(obj, *args, **kwargs)
            except Exception as e:
                print(f"Error during transformation: {e}")
            for k, v in obj.items():
                if isinstance(v, (dict, list)):
                    _format(v, func, *args, **kwargs)
        elif isinstance(obj, list):
            for i in range(len(obj)):
                _format(obj[i], func, *args, **kwargs)

    new_data_structure = deepcopy(data_structure)
    _format(new_data_structure, func, *args, **kwargs)
    return new_data_structure

def date_key_extrema(data, key, num_items=1, maxima=True, from_format='%Y-%m-%d %H:%M:%S', *, naive_timezone=timezone.utc):
    """
    Find the items with the most recent or oldest datetimes for a specified key in a list of dictionaries.
    
    Args:
        data (list): List of dictionaries containing data items.
        key (str): Key to search for the datetime value.
        num_items (int): Number of items to return (default is 1).
        maxima (bool): True to find the most recent items, False for the oldest.
        from_format (str): The datetime format used in the input data (default is '%Y-%m-%d %H:%M:%S').
    
    Returns:
        list: List of dictionaries with the most recent or oldest datetimes, up to num_items.
    """
    if isinstance(num_items, bool) or not isinstance(num_items, int) or num_items < 0:
        raise ValueError("num_items must be a nonnegative integer")
    if not isinstance(naive_timezone, tzinfo):
        raise ValueError("naive_timezone must be a timezone object")
    try:
        if datetime(2000, 1, 1, tzinfo=naive_timezone).utcoffset() is None:
            raise ValueError("naive_timezone must supply a UTC offset")
    except (NotImplementedError, TypeError):
        raise ValueError("naive_timezone must supply a UTC offset") from None
    new_data_struct = deepcopy(data)
    parsed_items = []

    for item in new_data_struct:
        datetime_value = item.get(key)
        if datetime_value:
            try:
                parsed_datetime = try_iso(datetime_value)
                if parsed_datetime is None:
                    parsed_datetime = datetime.strptime(datetime_value, from_format)
                if parsed_datetime.tzinfo is None or parsed_datetime.utcoffset() is None:
                    parsed_datetime = parsed_datetime.replace(tzinfo=naive_timezone)
                parsed_datetime = parsed_datetime.astimezone(timezone.utc)
                parsed_items.append((parsed_datetime, item))
            except (ValueError, TypeError):
                continue  # Skip items where the date cannot be parsed
    
    parsed_items.sort(key=lambda x: x[0], reverse=maxima)
    
    top_items = [item for _, item in parsed_items[:num_items]]
    
    return top_items

def re_format_folderpath(data, sharepoint_instance, key='webUrl'):
    """Convert URL fields through the bound library; propagate invalid-path errors."""
    def visit(value):
        if isinstance(value, dict):
            return {k: sharepoint_instance.url.normalize_sharepoint_path(v) if k == key
                    else visit(v) for k, v in value.items()}
        if isinstance(value, list):
            return [visit(v) for v in value]
        return value
    return visit(deepcopy(data))

def re_format_filepath(data):
    def reformat(item):
        """Build an item path from its parent reference and name when both are present."""
        try:
            if 'parentReference' in item and 'path' in item['parentReference'] and 'name' in item:
                pattern = re.compile(r'\b(b![A-Za-z0-9_-]+(?:/root:/?|/root)?)')
                path = item['parentReference']['path']
                file_name = item['name']
                
                match = pattern.search(path)
                if match:
                    new_path = path[match.end():] + "/" + file_name
                else:
                    new_path = path + "/" + file_name
                
                item['parentReference']['path'] = new_path

        except Exception as e:
            print(f"Error during transformation for item {item}: {e}")
        return item

    data_struct = deepcopy(data) 
    new_data_struct = modify_key_value(data_struct, reformat)
    return new_data_struct
   
  
__all__ = ["date_key_extrema", "format_key_value", "re_format_folderpath", "re_format_filepath"]
def __dir__(): return __all__
