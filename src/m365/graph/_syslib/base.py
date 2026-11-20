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
"""Filesystem, path-shape and local data helpers exposed through pathdir.

The implementation combines platform directory locators, extension rules, path
normalization, filesystem search and file/stream helpers. Cloud-provider path
checks are syntactic and do not contact remote services or verify access.
There is no filesystem locking or encryption layer. Pickle loading requires
trusted input because deserialization can execute code.
"""
import os
import stat
import sys
from pathlib import Path as PathlibPath
import re
import time
import pickle
import logging
import shutil
import math
import io 
from typing import Union
import inspect

if not sys.platform.startswith("win"):
    import pwd
    import grp
else:
    pwd = None
    grp = None   
    
from ._utils import VALID_FILE_EXTENSIONS
from ._appdata import UserDataDirectory
from ._paradigm import (
    __WINDOWS_FILE_PATH_REGEX__,
    __UNIX_FILE_PATH_REGEX__,
    __WINDOWS_DIRECTORY_PATH_REGEX__,
    __UNIX_DIRECTORY_PATH_REGEX__,
    __WINDOWS_PATH_STRUCTURE_REGEX__,
    __UNIX_PATH_STRUCTURE_REGEX__,
    __CLOUDPROVIDER_TRADITIONAL_PATH_STRUCTURE_REGEX__,
    __CLOUDPROVIDER_TRADITIONAL_FILE_PATH_REGEX__,
    __CLOUDPROVIDER_TRADITIONAL_FOLDER_PATH_REGEX__,
    __CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_CASESENSITIVE__,
    __CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_IGNORECASE__,
    __CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_S3_SPECIFIC__,
    __CLOUDPROVIDER_BUCKET_NAME_EXISTENCE_REGEX__,
    __CLOUDPROVIDER_SELECTION_REGEX__,
)
if sys.platform.startswith("win"):
    from ._windows import *
elif sys.platform.startswith("linux") or sys.platform.startswith("darwin"):
    from ._unix import *



logger = logging.getLogger(__name__)


class PackageDirectoryLocator:
    """
    A helper class for isolating the package directory in projects using a source-based layout.

    This class provides utility methods to process and traverse filesystem paths in order to determine
    the location of the package code within a project. It is designed for scenarios where the package code
    is maintained in its own subdirectory (as is common in 'src' layouts), distinct from other modules that
    may reside in the parent directory. The class features methods for:

      - Splitting paths into components with special handling for Windows drive letters.
      - Recursively searching upward from a given path to locate a marker file (such as 'setup.py'),
        which is assumed to reside at the project root.
      - Determining the package directory by combining the found project root with the expected
        package subdirectory (often sharing the same name as the project), thereby isolating the
        package code from other modules.

    Methods:
        _split_path_components(path, add_next_dir=False):
            Splits a given filesystem path into its components, ensuring that drive letters and
            their following directory are kept together on Windows systems.

        find_package_directory(start_path):
            Starting from the provided path, this method traverses upward to locate the project root
            (using the presence of a marker file like 'setup.py') and then determines the package
            directory by selecting the appropriate subfolder. If the start path is the project root,
            it attempts to return a candidate package subdirectory (if it exists), or falls back to the
            project root otherwise.
    """
    @classmethod
    def _combine_drive_components(cls, path_parts):
        """
        Combines the drive letter and the following directory component (for Windows paths), if applicable.
        """
        parts = path_parts.copy()
        if len(parts) > 1:
            parts[0] = os.path.join(parts[0], parts[1])
            parts.pop(1)
        return parts

    @classmethod
    def _locate_project_root(cls, directory, boundary):
        """
        Recursively searches for any of the marker files or directories in the current directory and moves upward until the
        specified boundary is reached. The presence of one of these markers is used as an indicator of the
        project root. The markers checked include:
            - 'setup.py'
            - 'pyproject.toml'
            - 'setup.cfg'
            - 'MANIFEST.in'
            - 'README.md'
            - '.git'
            - '.hg'

        Parameters:
            directory (str): The directory to start the search in.
            boundary (str): The boundary directory beyond which the search will not proceed.

        Returns:
            str or None: The directory containing one of the markers if found; otherwise, None.
        """
        marker_files = {'setup.py', 'pyproject.toml', 'setup.cfg', 'MANIFEST.in', 'README.md', '.git', '.hg'}

        for name in os.listdir(directory):
            full_path = os.path.join(directory, name)
            if name in marker_files and (os.path.isfile(full_path) or os.path.isdir(full_path)):
                return os.path.normpath(directory)

        parent_directory = os.path.dirname(directory)
        if parent_directory and os.path.abspath(parent_directory) >= os.path.abspath(boundary):
            return cls._locate_project_root(parent_directory, boundary)
        else:
            return None

    @classmethod
    def _split_path_components(cls, path, combine_drive=True):
        """
        Splits the provided path into components, ensuring that on Windows the drive letter and
        the following directory are kept together. If combine_drive is True, the first two
        components are combined appropriately.

        Parameters:
            path (str): The filesystem path to split.
            combine_drive (bool): Whether to combine the drive letter with the following directory component (useful on Windows).

        Returns:
            list: A list of path components.
        """
        path = os.path.normpath(path)
        parts = path.split(os.sep)
        if len(parts) > 1 and parts[0].endswith(':'):
            first_part = parts[0] + os.sep + parts[1]
            if combine_drive:
                parts_combined = [first_part] + parts[2:]
                return cls._combine_drive_components(parts_combined)
            return [first_part] + parts[2:]
        else:
            return parts

    @classmethod
    def find_package_directory(cls, start_path, combine_drive=True):
        """
        Traverses upward from the given start_path to locate the project root (determined by the presence
        of a marker file like 'setup.py') and then determines the package directory from that root.

        The package directory is typically defined as the first subdirectory under the project root that
        relates to the package code. If the start_path is already the project root, the method attempts
        to locate a candidate package subdirectory (assumed to share the project name). If such a subdirectory
        exists, it is returned; otherwise, the project root itself is returned.

        Parameters:
            start_path (str): The filesystem path from which to begin the upward search.
            combine_drive (bool): When True, adjusts path splitting on Windows by combining the drive letter
                with the following directory component. Defaults to True.

        Returns:
            str or None: The absolute path to the identified package directory if successful; otherwise, None.
        """
        parts = cls._split_path_components(start_path, combine_drive=combine_drive)
        boundary = parts[0]
        project_root = cls._locate_project_root(start_path, boundary)
        if project_root is None:
            return None

        rel_path = os.path.normpath(os.path.relpath(start_path, project_root))
        rel_parts = rel_path.split(os.sep)

        # If the relative path is just '.', it means start_path == project_root.
        # In that case, try to return a candidate subdirectory matching the project name.
        if rel_parts == ['.']:
            project_name = os.path.basename(project_root)
            candidate = os.path.join(project_root, project_name)
            if os.path.isdir(candidate):
                return candidate
            else:
                return project_root

        package_folder = rel_parts[0]
        return os.path.join(project_root, package_folder)


class pathdir:
    def __init__(self, base_path=None, cache_timeout=5,
                 ):
        self.DataDirectory = UserDataDirectory
        self.DataDirectory_instance = None
        self.base_path = base_path if base_path else self._root()
        self.valid_file_extensions = VALID_FILE_EXTENSIONS

        # Restore base_path after the temporary search scope.
        self._initial_base_path = self.base_path


        self.Dirs = self.DirsClass()
        self.Save = self.SaveClass(self)
        self.Load = self.LoadClass(self)
        self.File = self.FileClass(self)
        self.Text = self.TextClass(self)
        self.Buffer = self.BufferClass()
        self.Stream = self.StreamClass()

        self._deleted_files = []
        self._deleted_directories = []
        self.cache = {}
        self.cache_timeout = cache_timeout * 60

        self.kind = self.path_type

    class DirsClass:
        def __init__(self):
            if sys.platform.startswith("win"):
                self._windows_dir_instance = _windows_dir() 
                self.Appdata = _app_data_wrapper(_app_data_path())
                self.Downloads = self._windows_dir_instance.Downloads        
                self.ProgramFiles = self._windows_dir_instance.ProgramFiles
                self.ProgramFiles86 = self._windows_dir_instance.ProgramFiles86
                self.System32 = self._windows_dir_instance.System32    
                self.CDrive = self._windows_dir_instance.CDrive() 

            elif sys.platform.startswith("linux") or sys.platform.startswith("darwin"):
                self._unix_dir_instance = _unix_dir() 
                self.Downloads = self._unix_dir_instance.Downloads        
                self.System = self._unix_dir_instance.System
                self.UserBin = self._unix_dir_instance.UserBin
                self.LocalBin = self._unix_dir_instance.LocalBin
                self.Opt = self._unix_dir_instance.Opt                
                self.LocalUserBin = self._unix_dir_instance.LocalUserBin   
                self.LocalShare = self._unix_dir_instance.LocalShare   
                self.Home = self._unix_dir_instance.Home                   
                
                if sys.platform.startswith("linux"):                    
                    self.Snap = self._unix_dir_instance.Snap   
                    self.ShareApplications = self._unix_dir_instance.ShareApplications                
                    self.FlatpakSystem = self._unix_dir_instance.FlatpakSystem   
                    self.FlatpakUser = self._unix_dir_instance.FlatpakUser                       
                                                
                elif sys.platform.startswith("darwin"):
                    self.Applications = self._unix_dir_instance.Applications   
 
        def __getattr__(self, item):
            if sys.platform.startswith("win") and item in ["Local", "LocalLow", "Roaming"]:
                return None
            raise AttributeError(f"'{type(self).__name__}' object has no attribute '{item}'")

        def __dir__(self):
            if sys.platform.startswith("win"):
                return ['Appdata', 'Local', 'LocalLow', 'Roaming', 'CDrive', 'Downloads', 'ProgramFiles', 'ProgramFiles86']
            elif sys.platform.startswith("linux"):
                return ['Downloads', 'System', 'Home', 'UserBin', 'LocalBin', 'Opt', 'LocalUserBin', 'LocalShare', 'Snap', 'ShareApplications', 'FlatpakSystem', 'FlatpakUser']
            elif sys.platform.startswith("darwin"):
                return ['Downloads', 'System', 'Home', 'Applications', 'UserBin', 'LocalBin', 'Opt', 'LocalUserBin', 'LocalShare']
        
    def _root(self):
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            current_dir = os.path.dirname(os.path.abspath("__file__"))

        file = next(PathlibPath(current_dir).glob("base.*"), None)

        try:
            if file and file.is_file() and file.exists():
                return str(PathlibPath(current_dir))
            return current_dir  # Use the working directory if a caller file cannot be determined.
        except AttributeError:
            return None

    def cleanup_cache(self):
        """Remove cache entries that have exceeded the timeout."""
        current_time = time.time()
        keys_to_remove = [key for key, entry in self.cache.items() if current_time - entry['last_used'] > self.cache_timeout]
        for key in keys_to_remove:
            del self.cache[key]

    def invalidate_cache(self):
        """Clear all cached search results."""
        self.cache.clear()

    @staticmethod
    def _unlist(data):
        """
        If the input is a list and contains only one item, return the single item.
        """
        if isinstance(data, list) and len(data) == 1:
            return data[0]
        return data
       
    @staticmethod       
    def _generate_unique_filename(path, filename=None):
        """
        Generate a unique filename in the specified directory by appending a counter to the name if needed.

        This function checks if a file with the given name already exists in the specified directory. If it does,
        it appends a number in parentheses (starting from 1) to the base name of the file until a unique filename is found.

        Parameters:
            path (str): The directory path where the file will be located. If `filename` is None, `path` is expected to include the full path to the file.
            filename (str, optional): The base filename to check for uniqueness in the provided directory. If None, the filename is extracted from the last component of the `path`.
        """         
        if filename is None:
            filename = os.path.basename(path) 
            path = os.path.dirname(path)  
        base_name, extension = os.path.splitext(filename)
        counter = 1
        new_filename = filename
        full_path = os.path.join(path, new_filename)  
        while os.path.exists(full_path):
            new_filename = f"{base_name} ({counter}){extension}"
            full_path = os.path.join(path, new_filename) 
            counter += 1        
        return str(PathlibPath(full_path))   
       
    def _splitdrive_win(self, path):
        """
        A minimal recreation of os.path.splitdrive for Windows-like paths:
          - UNC paths: e.g. '\\\\server\\share'
          - Drive letter: 'C:' or 'D:' 
        Returns (drive, path_without_drive)
        If no drive or UNC found, returns ('', path) as-is.
        """
        # Unify slashes first to handle backslash or forward slash
        p = self.Text.posixsep(path)
        if len(p) >= 2 and p[0] == '/' and p[1] == '/':
            parts = p.split('/')
            if len(parts) >= 3:
                unc_prefix = '//'.join(parts[:2]) # Rebuild the UNC prefix
                if len(parts) >= 4:
                    drive = '//' + parts[2] + '/' + parts[3]
                    remainder = '/' + '/'.join(parts[4:])
                else:
                    drive = '//' + parts[2]
                    remainder = ''
                return (drive, remainder)
            else:
                return (p, '')
        if len(p) >= 2 and p[1] == ':':
            return (p[:2], p[2:])
        return ('', path)

    def _collapse_dotdot(self, components):
        """Collapse dot components while preserving unmatched leading parent components."""
        new_comps = []
        for c in components:
            if c == '' or c == '.':
                continue
            if c == '..':
                if new_comps and new_comps[-1] != '..':
                    new_comps.pop() if new_comps else None
                else:
                    new_comps.append('..')
            else:
                new_comps.append(c)
        return new_comps  
       
    def _cloud_bucket_name_validation(self, path, provider="SharePoint"):
        """Validate a cloud path and apply bucket-name rules for recognized providers."""
        provider_lower = provider.lower()    	
        if not __CLOUDPROVIDER_TRADITIONAL_PATH_STRUCTURE_REGEX__.match(path):
            return False
        provider_match = __CLOUDPROVIDER_SELECTION_REGEX__.match(provider_lower) # Apply bucket rules only for recognized providers.
        if not provider_match:
            return True 
        bucket_name_match = re.search(__CLOUDPROVIDER_BUCKET_NAME_EXISTENCE_REGEX__, path)
        if not bucket_name_match:
            return True 
        bucket_name = bucket_name_match.group(1)
        if provider_match.group("aws"):  # S3 requires both the bucket-shape and reserved-name checks.
            return bool(
                __CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_CASESENSITIVE__.match(bucket_name)
                and __CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_S3_SPECIFIC__.match(bucket_name))
        elif provider_match.group("wasabi"):  # Wasabi follows S3-like validation rules
            return bool(__CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_CASESENSITIVE__.match(bucket_name))
        elif provider_match.group("backblaze"):  # Backblaze allows uppercase letters in bucket names
            return bool(__CLOUDPROVIDER_BUCKETNAME_VALIDITY_REGEX_IGNORECASE__.match(bucket_name))
        return True       
       
       
    def create_data_directory(self, dirname, auto_remove=True):
        """
        Create and manage user data directory.
        """
        self.DataDirectory_instance = self.DataDirectory(dirname=dirname, auto_remove=auto_remove)
        return self.DataDirectory_instance

    def search(self, base_path=None, name=None, regex=False, is_directory=False, listdir=False,
               as_list=False, IGNORECASE=False, tempFiles=False, depth_limit=None, shallowest_only=False, file_only=False):
        """
        Search for files or directories within a base path.

        Parameters:
            base_path (str): The base path to start the search from.
            name (str): The name or regular expression pattern to match against files or directories.
            regex (bool): Whether `name` is a regular expression.
            is_directory (bool): Whether to search for directories instead of files.
            listdir (bool): If True (and searching for directories), list all files within the directory.
            as_list (bool): If True, return the output as a list; otherwise as a newline-joined string.
            IGNORECASE (bool): If True and regex is True, perform a case-insensitive match.
            tempFiles (bool): Whether to include temporary files (those starting with '~$').
            depth_limit (int): The maximum directory depth to search.
            shallowest_only (bool): If True, return only the shallowest matching result.
            file_only (bool): If True, and a result is a full file path, return only its basename.        
            
        Returns:
            str or list: The found paths as a newline-separated string or a list, based on as_list.
        """
        def entry_matches(entry):
            """
            Determines if the entry should be included based on the search criteria.
            """
            # Exclude temporary files if requested.
            if not (tempFiles or not entry.startswith('~$')):
                return False

            if regex:
                return re.search(search_pattern, entry) is not None
            else:
                return entry == name

        def process_directory(root, dirs, files):
            """
            Processes a directory from os.walk, applying a depth check and selecting matching entries.
            """
            current_depth = root.count(os.sep) - base_path.count(os.sep)
            if depth_limit is not None and current_depth > depth_limit:
                dirs[:] = []  # prune deeper directories
                return []

            target_list = dirs if is_directory else files
            matches = []
            for entry in target_list:
                entry_path = str(PathlibPath(root, entry).resolve())
                if entry_matches(entry):
                    if listdir and is_directory:
                        # Return the matching directory's file paths rather than the directory itself.
                        matches.extend(self.full_path(entry_path))
                    else:
                        matches.append(entry_path)
            return matches        
        
        self.cleanup_cache()

        if not base_path:        	
            base_path = self.base_path

        if name is None:
            return None

        if os.path.isdir(name):
            name = os.path.basename(name)

        # Keep file-only results separate in the cache.
        cache_key = (base_path, name, regex, is_directory, listdir, as_list,
                     IGNORECASE, tempFiles, depth_limit, file_only)
        if cache_key in self.cache:
            entry = self.cache[cache_key]
            entry['last_used'] = time.time()
            return entry['result']

        if regex and IGNORECASE:
            search_pattern = re.compile(name, re.IGNORECASE)
        else:
            search_pattern = name

        FileList = []
        for root, dirs, files in os.walk(base_path):
            FileList.extend(process_directory(root, dirs, files))

        # Sort so that shallower paths appear first.
        sorted_files = sorted(FileList, key=lambda x: x.count(os.sep))

        if file_only:
            sorted_files = [os.path.basename(x) if self.is_path_struct(x, auto=True) else x for x in sorted_files]

        result = sorted_files if as_list else '\n'.join(sorted_files)

        if shallowest_only and sorted_files:
            result = sorted_files[0] if not as_list else [sorted_files[0]]

        self.cache[cache_key] = {'result': result, 'last_used': time.time()}
        return result

    def full_path(self, directory, ignore_temp_files=True, as_list=True, file_extension=None):
        """
        Get the full paths of all files in a directory, with options to ignore temporary files,
        filter by file extension, and return a single path as a string or as a list.

        Parameters:
            directory (str): The directory for which to get the full file paths.
            ignore_temp_files (bool, optional): If True, temporary files starting with '~$' will be ignored.
            as_list (bool, optional): If False, and exactly one result is found, return the result as a string.
            file_extension (str, optional): If provided, only files with this extension will be included.
        """
        directory_path = PathlibPath(directory).resolve()
        files_list = [f for f in directory_path.iterdir() if f.is_file()]
        if ignore_temp_files:
            files_list = [f for f in files_list if not f.name.startswith('~$')]
        if file_extension:
            if not file_extension.startswith('.'):
                file_extension = '.' + file_extension
            files_list = [f for f in files_list if f.suffix == file_extension]
        full_paths = [str(f) for f in files_list]
        if not as_list and len(full_paths) == 1:
            return full_paths[0]
        return full_paths

    def files(self, base_path=None, name=None, regex=False, IGNORECASE=False, depth_limit=None):
        """
        Populates the list of files to be managed based on specific search criteria.

        This method searches the specified base path for files that match a given pattern, 
        either as a regular expression or as a direct string match. The method updates 
        the internal list `_deleted_files` with the paths of files that match the criteria.

        Parameters:
            name (str): The pattern or exact name of the files to search for. If `None`, no files are added to the list. This can be a regular expression if `regex` is True.
            regex (bool): If True, `name` is treated as a regular expression. Default is False.
            IGNORECASE (bool): If True and `regex` is True, the regular expression matching will be case-insensitive. Default is False.
            depth_limit (int): The maximum depth of directories to traverse. `None` means no limit.
        """    	
        if not base_path:
            base_path = self.base_path
        if name is None:
            return self
        for root, dirs, files in os.walk(base_path):
            current_depth = root.count(os.sep) - self.base_path.count(os.sep)
            if depth_limit is not None and current_depth > depth_limit:
                dirs[:] = []
                continue
            if regex:
                pattern_flags = re.IGNORECASE if IGNORECASE else 0
                search_pattern = re.compile(name, pattern_flags)
            else:
                search_pattern = name
            for file in files:
                file_path = str(PathlibPath(root, file).resolve())
                if (regex and re.search(search_pattern, file)) or (not regex and search_pattern == file):
                    self._deleted_files.append(file_path)
        return self

    def directories(self, base_path=None, name=None, regex=False, IGNORECASE=False, depth_limit=None):
        """
        Populates the list of directories to be managed based on specific search criteria.

        This method searches the specified base path for directories that match a given pattern, 
        either as a regular expression or as a direct string match. The method updates 
        the internal list `_deleted_directories` with the paths of directories that match the criteria.

        Parameters:
            name (str): The pattern or exact name of the directories to search for. If `None`, no directories are added to the list. This can be a regular expression if `regex` is True.
            regex (bool): If True, `name` is treated as a regular expression. Default is False.
            IGNORECASE (bool): If True and `regex` is True, the regular expression matching will be case-insensitive. Default is False.
            depth_limit (int): The maximum depth of directories to traverse. `None` means no limit.
        """
        if not base_path:
            base_path = self.base_path
        if name is None:
            return self
        for root, dirs, _ in os.walk(base_path):
            current_depth = root.count(os.sep) - self.base_path.count(os.sep)
            if depth_limit is not None and current_depth >= depth_limit:
                dirs[:] = []
                continue
            if regex:
                pattern_flags = re.IGNORECASE if IGNORECASE else 0
                search_pattern = re.compile(name, pattern_flags)
            else:
                search_pattern = name

            for directory in dirs:
                dir_path = str(PathlibPath(root, directory).resolve())
                if (regex and re.search(search_pattern, directory)) or (not regex and search_pattern == directory):
                    self._deleted_directories.append(dir_path)
        return self

    def delete(self):
        """
        Property that deletes all files in the list when accessed.
        """
        if not self._deleted_files:
            logger.info('No files to delete.')
            return False
        all_successful = True
        for file in self._deleted_files:
            if not os.path.isfile(file):
                logger.warning(f'File not found: {file}')
                all_successful = False
                continue
            try:
                os.remove(file)
                logger.info(f'Successfully deleted {file}')
            except PermissionError as e:
                logger.error(f'Permission denied: {file}. Reason: {e}')
                all_successful = False
            except OSError as e:
                logger.error(f'Failed to delete {file}. Reason: {e}')
                all_successful = False
        self._deleted_files = []
        return all_successful

    if sys.platform.startswith("win"):
        def roaming_path(self, directory=None):
            base = self.Dirs.Appdata.Roaming        
            if base:
                roamingPath = base
                if directory:
                    directory = PathlibPath(directory)
                    dirpath = roamingPath / directory
                    if dirpath.exists() and dirpath.is_dir():
                        return str(dirpath)
                    else:
                        raise FileNotFoundError(f"The directory '{directory}' does not exist or is not a directory.")
                else:
                    return roamingPath
            else:
                raise FileNotFoundError(f"The directory '{directory}' does not exist or is not a directory.")
               
    def path_type(self, path, posix=False, cloud=False, provider="SharePoint"):
        """
        Determines whether the given input string represents a directory-like or file-like path.

        This method uses pattern matching based on the platform to identify if the given path is directory-like or file-like.
        It returns a dictionary with boolean values for 'is_dir' and 'is_file' keys.

        Args:
            path (str): The input path string to evaluate.
            posix (bool): If True, treats the path as POSIX-compliant (e.g., Linux/Mac).
            cloud (bool): If True, applies cloud-specific path patterns.    
        """
        norm_pathstr = None
        dir_pattern = None
        file_pattern = None
        
        if cloud:
            norm_pathstr = self.Text.posixsep(path)
            if not self._cloud_bucket_name_validation(norm_pathstr, provider=provider):
                return {'folder': False, 'file': False} 
            dir_pattern = __CLOUDPROVIDER_TRADITIONAL_FOLDER_PATH_REGEX__
            file_pattern = __CLOUDPROVIDER_TRADITIONAL_FILE_PATH_REGEX__   
            
        elif posix:
            norm_pathstr = self.Text.posixsep(path) 
            dir_pattern = __UNIX_DIRECTORY_PATH_REGEX__
            file_pattern = __UNIX_FILE_PATH_REGEX__
        else:
            if sys.platform.startswith('win'):
                norm_pathstr = self.Text.systemsep(path) 
                dir_pattern = __WINDOWS_DIRECTORY_PATH_REGEX__
                file_pattern = __WINDOWS_FILE_PATH_REGEX__
            elif sys.platform.startswith('linux') or sys.platform.startswith('darwin'):
                norm_pathstr = self.Text.systemsep(path) 
                dir_pattern = __UNIX_DIRECTORY_PATH_REGEX__
                file_pattern = __UNIX_FILE_PATH_REGEX__
        is_dir = bool(dir_pattern.match(str(norm_pathstr)))
        is_file = bool(file_pattern.match(str(norm_pathstr)))        
        return {'folder': is_dir, 'file': is_file} if cloud else {'dir': is_dir, 'file': is_file}

    def is_path_struct(self, path, posix=False, cloud=False, auto=False):
        """
        Checks if a given path conforms to a valid directory or file path structure
        based on the specified or detected platform.

        Parameters:
            path (str): The path string to validate.
            posix (bool, optional): If True, explicitly checks against POSIX (Linux/macOS) path rules.
                                    If False, defaults to Windows rules unless overridden by `cloud` or `auto`.
            cloud (bool, optional): If True, checks the path against known cloud provider path rules (e.g., SharePoint, S3).
            auto (bool, optional): If True, automatically detects the platform (Windows, Linux, macOS)
                                   and applies the appropriate path structure validation.
                                   This overrides `posix` if both are True.

        Returns:
            bool: True if the path matches the expected structure for the chosen or detected platform, False otherwise.

        Notes:
            - If `cloud` is True, it takes precedence over `posix` and `auto`.
            - If `auto` is True, platform detection is performed to decide the validation method.
            - If neither `cloud`, `posix`, nor `auto` is True, the method defaults to the current platform's expected path format.
        """
        if cloud:           
            norm_pathstr = self.Text.posixsep(path)
            return bool(__CLOUDPROVIDER_TRADITIONAL_PATH_STRUCTURE_REGEX__.match(path))
           
        if auto:
            if sys.platform.startswith("win"):
                norm_pathstr = self.Text.windowsep(path) 
                return bool(__WINDOWS_PATH_STRUCTURE_REGEX__.match(path))                
            elif sys.platform.startswith("linux") or sys.platform.startswith("darwin"):
                norm_pathstr = self.Text.posixsep(path) 
                return bool(__UNIX_PATH_STRUCTURE_REGEX__.match(path))  
               
        elif posix:
            norm_pathstr = self.Text.posixsep(path) 
            return bool(__UNIX_PATH_STRUCTURE_REGEX__.match(path))
        else:
            norm_pathstr = self.Text.systemsep(path) 
            return bool(__WINDOWS_PATH_STRUCTURE_REGEX__.match(path))

    def chdir(self, new_base_dir, package_dir=False):
        """
        Sets a new base directory for the file path finder.
        
        Parameters:        
            package_dir (bool): If True, set base_path to the package directory using PackageDirectoryLocator.
        """
        if os.path.isdir(new_base_dir):
            if package_dir:
                self.base_path = PackageDirectoryLocator.find_package_directory(new_base_dir)
            else:
                self.base_path = new_base_dir            
            logger.info(f"Base directory updated to: {self.base_path}") 
            
    def resetdir(self):
        """
        Resets the base directory to the original path determined during initialization.
        """
        self.base_path = self._initial_base_path
        logger.info(f"Base directory reset to: {self.base_path}")            
            
    def join(self, base, *segments, expand_user=True, expand_vars=True, normalize=True, auto_platform=True, platform=None):
        """
        A robust path join function that handles:
        - Windows drives & UNC paths
        - user expansions (~)
        - environment variable expansions
        - cross-platform slash normalization using self.Text.posixsep/self.Text.windowsep
        - '.' and '..' collapsing (if normalize=True).

        :param base: The starting path string.
        :param segments: Additional path segments to join.
        :param expand_user: If True, expands '~' -> home dir.
        :param expand_vars: If True, expands env vars like %USERPROFILE% or $HOME.
        :param normalize: If True, collapses '.' and '..'.
        :param auto_platform: If True, detect current OS. Otherwise, use `platform`.
        :param platform: 'win' or 'posix' to force that style. If None, auto-detect.
        :return: A string with the joined path in the chosen style.
        """
        if platform not in (None, 'win', 'posix'):
            raise ValueError("platform must be one of None, 'win', or 'posix'")

        if auto_platform:
            current_is_win = sys.platform.startswith('win')
            path_style = 'win' if current_is_win else 'posix'
        else:
            path_style = platform or ('win' if os.name == 'nt' else 'posix')

        def expand(p):
            if expand_user:
                p = os.path.expanduser(p)
            if expand_vars:
                p = os.path.expandvars(p)
            return p

        base = expand(str(base))
        segments = [expand(str(s)) for s in segments]

        drive, base_path = self._splitdrive_win(base) if path_style == 'win' else ('', base)

        base_path = self.Text.posixsep(base_path)
        segments_posix = [self.Text.posixsep(s) for s in segments]

        base_comps = [comp for comp in base_path.split('/') if comp]
        leading_slash = base_path.startswith('/')

        # An absolute segment resets the accumulated path.
        for seg in segments_posix:
            if path_style == 'win':
                seg_drive, seg_path = self._splitdrive_win(seg)
                if seg_drive and seg_drive.lower() != drive.lower():
                    drive = seg_drive
                    base_comps = []
                    leading_slash = False
                if seg_path.startswith('/'):
                    base_comps = []
                    leading_slash = True
                seg_comps = [c for c in seg_path.split('/') if c]
            else:
                if seg.startswith('/'):
                    base_comps = []
                    leading_slash = True
                seg_comps = [c for c in seg.split('/') if c]
            base_comps.extend(seg_comps)

        if normalize:
            base_comps = self._collapse_dotdot(base_comps)

        sep_func = self.Text.windowsep if path_style == 'win' else self.Text.posixsep
        joined_comps = '/'.join(base_comps)
        
        if drive.startswith('//'): 
            if joined_comps:
                final_path = drive.rstrip('/') + '/' + joined_comps
            else:
                final_path = drive
        else:
            if drive:
                if leading_slash:
                    if joined_comps:
                        final_path = drive + '/' + joined_comps
                    else:
                        final_path = drive + '/'
                else:
                    final_path = drive + joined_comps
            else:
                if leading_slash:
                    final_path = '/' + joined_comps
                else:
                    final_path = joined_comps

        final_path = sep_func(final_path)
        return final_path

    def get_module_dir(self):
        """
        Returns the absolute directory path of the current Python file if possible.
        Falls back to the current working directory if no real file path exists
        (e.g., in an embedded or interactive environment where code is run from
        memory rather than a file).

        Returns:
            str: The absolute path to the directory, or the current working
                 directory if no file-based path can be determined.
        """
        try:
            caller_frame = inspect.stack()[1]
            caller_file = caller_frame.filename
            path_obj = PathlibPath(caller_file).resolve()
            
            # Treat synthetic source names such as <string> as unavailable.
            if path_obj.name.startswith("<") and path_obj.name.endswith(">"):
                raise OSError("Placeholder filename detected.")
            
            return str(path_obj.parent)
        except (IndexError, TypeError, OSError):
            # Fallback: use the current working directory.
            return str(PathlibPath.cwd().resolve())
           
    if sys.platform.startswith("win"):
        def create_shortcut(self, target_file, shortcut_path, shortcut_description=None, shortcut_iconLocation=None, ps_version=None):
            """
            Create a Windows shortcut using PowerShell.

            This method creates a Windows shortcut for a specified target file by utilizing the Qlnk class,
            which builds and executes a PowerShell script to generate the shortcut. The shortcut is created
            at the specified path and can include a custom description and icon. If no icon is provided, the shortcut
            will use the target file's default icon.

            Parameters:
                target_file (str): The absolute or relative path of the file for which the shortcut is to be created.
                shortcut_path (str): The absolute or relative path where the shortcut (.lnk) file will be created.
                    If the provided path does not include the ".lnk" extension, it will be appended automatically.
                shortcut_description (str, optional): A custom description for the shortcut. Defaults to "My Shortcut"
                    if not specified.
                shortcut_iconLocation (str, optional): The path to a custom icon file for the shortcut.
                    Supported file types include '.ico', '.exe', '.dll', and '.cpl'. If omitted, the target file's icon is used.
                ps_version (str, optional): Specifies the PowerShell executable to use. Valid options are:
                    - 'powershell.exe' (Windows PowerShell)
                    - 'powershell' (alias for Windows PowerShell)
                    - 'pwsh' (PowerShell Core, cross-platform)
                    Defaults to 'powershell.exe' if not provided.

            Returns:
                tuple: A tuple (success, message) where:
                    - success (bool): True if the shortcut was created successfully, False otherwise.
                    - message (str): A descriptive message indicating the outcome of the operation, including any error details.

            Raises:
                NotImplementedError: If the method is invoked on a non-Windows platform.
                FileNotFoundError: If the specified target_file does not exist.
                ValueError: If an invalid icon file type is provided or other parameter issues are encountered.
            """
            qlnk = Qlnk()
            success, msg = qlnk.create(target_file, shortcut_path, shortcut_description, shortcut_iconLocation, ps_version)
            return success, msg
           

    class TextClass:
        """Normalize path-related text and delimiters."""
        def __init__(self, parent):
            self.parent = parent

        def _trimlist(self, items):
            """
            Filter a list of items by removing empty strings and strings consisting only of punctuation.
            """        	
            import string as s
            return [item for item in items if item.strip() and not all(char in s.punctuation for char in item.strip())]

        def _stripext(self, s):
            """
            Remove invalid file extension characters from a string by normalizing spaces 
            and validating against a list of known valid extensions.
            """        	
            matches = re.finditer(r'\.([a-zA-Z0-9 ]+)', s) # Capture dot-prefixed suffix components, allowing spaces.
            processed_text = s
            last_index = 0
            for match in matches:
                extension_dirty = match.group(1)
                extension_clean = ''.join(extension_dirty.split())
                longest_valid_extension = ''
                for i in range(1, len(extension_clean) + 1):
                    potential_extension = '.' + extension_clean[:i]
                    if potential_extension in self.parent.valid_file_extensions:
                        longest_valid_extension = potential_extension
                if longest_valid_extension:
                    start = match.start(1) + last_index
                    end = start + len(extension_dirty) 
                    new_end = start + len(longest_valid_extension) - 1 
                    processed_text = processed_text[:start] + longest_valid_extension[1:] + processed_text[end:]
                    last_index = new_end - len(longest_valid_extension[1:])  
            return processed_text

        def posixsep(self, s):
            """ Replace all backslashes with forward slashes and merge multiple slashes into one."""
            s = str(s)
            replaced = s.replace('\\', '/')
            chars = ''
            previous_char = ''
            for i in replaced:
                if i == '/' and previous_char == '/':
                    continue
                else:
                    chars += i
                    previous_char = i
            return chars

        def windowsep(self, s):
            """Replace all forward slashes with backslashes and normalize multiple backslashes."""
            s = str(s)    
            replaced = s.replace('/', '\\')
            chars = ''
            previous_char = ''
            for i in replaced:
                if i == '\\' and previous_char == '\\': 
                    continue
                else:
                    chars += i
                    previous_char = i
            return chars

        def systemsep(self, s):
            """ Normalize path separators based on the current operating system."""
            if sys.platform.startswith("win"):
                return self.windowsep(s)            
            elif sys.platform.startswith("linux") or sys.platform.startswith("darwin"):
                return self.posixsep(s)

        def stripspace(self, s):
            """
            Remove leading, trailing, and excessive spaces from a string, 
            and eliminate spaces near slashes, backslashes, or dots for cleaner formatting.
            """        	
            trimmed_string = s.strip()
            no_space_near_punctuation = re.sub(r'(\w)\s+([\/\\\.])', r'\1\2', trimmed_string) # Remove spaces before path punctuation.
            no_space_near_punctuation = re.sub(r'([\/\\\.])\s+(\w)', r'\1\2', no_space_near_punctuation) # Remove spaces after path punctuation.
            cleaned_string = re.sub(r'\s+', ' ', no_space_near_punctuation) # Collapse remaining runs of whitespace.
            return cleaned_string

        def stripend(self, s):
            """
            Remove any trailing characters from the given string that are not letters, digits, 
            underscores, or hyphens.
            """     
            return re.sub(r'[^a-zA-Z0-9_-]+$', '', s)

        def normalize_drive(self, s, position=False):
            """
            Normalize the drive letter in a string by removing spaces around it, 
            ensuring compatibility with Windows file path strings. Optionally return positions of drive letters.
            """          	
            if sys.platform.startswith('win'):
                pattern = r'(\s*)([A-Z]:)(\s*)' # Regex to match spaces around a Windows-style drive letter. 
                matches = []
                for match in re.finditer(pattern, s):
                    left_space_group = match.group(1)
                    drive_letter = match.group(2)
                    right_space_group = match.group(3)
                    start = match.start()
                    end = match.end()
                    start_adjusted = start + len(left_space_group)
                    if right_space_group:
                        non_space_index = end
                        while non_space_index < len(s) and s[non_space_index].isspace():
                            non_space_index += 1
                        s = s[:start] + drive_letter + s[non_space_index:]
                    matches.append((drive_letter, str(start_adjusted)))
                s = s.rstrip()
                return s if not position else (matches, s)
            else:
                return s

        def splitext(self, s, clean=True):
            """
            Split a string into components, preserving dots as part of extensions. 
            Optionally clean the resulting parts by removing empty or punctuation-only strings.
            """        	
            parts = re.split(r'(\.)', s) # Keep dots as separate capture groups.
            new_parts = []
            i = 0
            while i < len(parts):
                if parts[i] == '.' and i + 1 < len(parts):
                    combined = parts[i] + parts[i + 1]
                    new_parts.append(combined)
                    i += 2
                else:
                    new_parts.append(parts[i])
                    i += 1
            return self._trimlist(new_parts) if clean else new_parts

        def stripextension(self, s, posix=False):
            """
            Strip invalid file extensions from a given string, 
            reconstructing a cleaned version of the string.
            """        	
            norm_pathstr = self.systemsep(s) if not posix else self.posixsep(s)           
            s = self.stripspace(norm_pathstr)
            s = self.normalize_drive(s) if not posix else s
            s = self.stripend(s)              
            parts = self.splitext(s)
            new_parts = []
            for p in parts:
                part = self._stripext(p)
                new_parts.append(part)
            return "".join(new_parts) 

        def balance_grouping_symbols(self, string):
            """
            Removes unmatched grouping symbols (parentheses '()', brackets '[]', 
            and braces '{}') from the input string, ensuring all remaining symbols are balanced.
            """   
            def _balance(s, delimiters="()"):
                open_sym, close_sym = delimiters[0], delimiters[1]
                stack = []
                remove_indices = set()        
                for index, char in enumerate(s):
                    if char == open_sym:
                        stack.append(index)
                    elif char == close_sym:
                        if stack:
                            stack.pop()
                        else:
                            remove_indices.add(index)
                remove_indices.update(stack)
                new_string = ''.join([char for idx, char in enumerate(s) if idx not in remove_indices])    
                return new_string
            _delimiters = ["()", "[]", "{}"]
            for _delimiter in _delimiters:
                string = _balance(string, delimiters=_delimiter)
            return string
           
        def __dir__(self):
            return ['posixsep', 'windowsep', 'systemsep', 'stripspace', 'stripend', 'normalize_drive', 'splitext', 'stripextension', 'balance_grouping_symbols']

           
    class FileClass:
        """
        A utility class for managing file paths and extensions, with built-in methods to normalize,
        validate, and manipulate file paths and extensions for cross-platform compatibility.

        This class provides an `Extension` class as a nested class to handle operations specifically
        related to file extensions, such as extracting, validating, and appending extensions to file paths.
        """
        def __init__(self, parent):
            self.parent = parent
            self.Extension = self.ExtensionClass(self)            
            self.Info = self.InfoClass()                

        def __dir__(self):
            return ['Extension', 'Info']             
          
        class ExtensionClass:
            def __init__(self, slot):
                self.slot = slot

            def _normalize_path(self, path, posix=False):
                """
                Normalize a file path to conform to OS-specific path standards, strip unnecessary spaces,
                normalize the drive letter, and remove any non-alphanumeric characters from the end.
                """                
                norm_pathstr = self.slot.parent.Text.systemsep(path) if not posix else self.slot.parent.Text.posixsep(path) 
                return self.slot.parent.Text.stripextension(norm_pathstr, posix=posix)

            def _normalize_path_prefix(self, path, normalize_path=False, posix=False, cloud=False):
                """
                Ensure that the path starts with a dot if it doesn't follow the OS-specific file path structure.
                """
                if normalize_path:
                    path = self._normalize_path(path, posix=posix or cloud)
                else:
                    path = (self.slot.parent.Text.posixsep(path) if posix or cloud else self.slot.parent.Text.systemsep(path))

                if cloud:
                    if not __CLOUDPROVIDER_TRADITIONAL_PATH_STRUCTURE_REGEX__.match(path):
                        path = '.' + path.lower() if not path.startswith('.') else path
                    return path

                if posix:
                    if not __UNIX_PATH_STRUCTURE_REGEX__.match(path):
                        path = '.' + path.lower() if not path.startswith('.') else path
                    return path

                if sys.platform.startswith('win'):
                    if not __WINDOWS_PATH_STRUCTURE_REGEX__.match(path):
                        path = '.' + path.lower() if not path.startswith('.') else path
                elif sys.platform.startswith('linux') or sys.platform.startswith('darwin'):
                    if not __UNIX_PATH_STRUCTURE_REGEX__.match(path):
                        path = '.' + path.lower() if not path.startswith('.') else path
                return path

            def _normalize_extension(self, path, validate_every_extension=True):
                """
                Normalize and validate file extensions in the given path. If `validate_every_extension` is True,
                ensure all extensions in the path are valid; otherwise, concatenate only valid extensions.
                """
                extensions = []
                parts = path.split('.')
                if len(parts) > 1:
                    for ext in parts[1:]:
                        formatted_ext = f".{ext.lower()}"
                        if formatted_ext in self.slot.parent.valid_file_extensions:
                            extensions.append(formatted_ext)
                        elif validate_every_extension:
                            return None
                return "".join(extensions) if extensions else None

            def get(self, path, posix=False, cloud=False):
                """
                Extract and validate the file extension from a given path, including the period,
                while handling multiple extensions, hidden files, input validation, and full file paths.
                """
                normalized_path = self._normalize_path(path, posix=posix or cloud)
                if not self.slot.parent.path_type(normalized_path, posix=posix or cloud, cloud=cloud)["file"]:
                    return None
                if not isinstance(normalized_path, str) or not normalized_path:
                    return None
                basename = os.path.basename(normalized_path)
                if basename.startswith('.') and basename.count('.') == 1:
                    return None
                if '.' in basename:
                    return self._normalize_extension(basename)
                return None

            def has(self, path, posix=False, cloud=False):
                """
                Check if an extension is valid and format it with a leading period if not present.
                """
                normalized_path = self._normalize_path(path, posix=posix or cloud)
                formatted_path = self._normalize_path_prefix(normalized_path, normalize_path=False, posix=posix or cloud, cloud=cloud)
                if '.' in formatted_path:
                    return self._normalize_extension(formatted_path)
                return None

            def exists(self, path, posix=False, cloud=False):
                """
                Check if a valid file extension exists in the given path.
                """
                return bool(self.get(path, posix=posix or cloud, cloud=cloud))

            def strip(self, path, posix=False, cloud=False):
                """
                Remove the file extension from a given path.
                """
                normalized_path = self._normalize_path(path, posix=posix or cloud)
                extension = self.get(normalized_path, posix=posix or cloud, cloud=cloud)
                return normalized_path.replace(extension, "") if extension else normalized_path

            def join(self, extension, path, posix=False, cloud=False):
                """Append the extension unless the filename already ends with it."""
                normalized_path = self._normalize_path(path, posix=posix)
                normalized_extension = self.has(extension, posix=posix or cloud, cloud=cloud)
                if normalized_extension and not normalized_path.endswith(normalized_extension):
                    return f"{normalized_path}{normalized_extension}"
                return normalized_path

            def __dir__(self):
                return [
                    'get', 'has', 'exists', 'strip', 'join',
                    ]

               
        class InfoClass:
            def __init__(self):
                pass      
                
            def _interpret_permissions(self, octal_perms):
                """
                Interprets the octal permission string into a readable format.

                This method converts an octal permission string, typically three digits, into a human-readable format
                describing the permissions for the owner, group, and others.

                Parameters:
                    octal_perms (str): A string representing file permissions in octal format. Must be a three-digit string, where each digit is between 0 and 7, inclusive.

                """            	
                octal_perms = str(int(float(octal_perms)))    
                if len(octal_perms) != 3 or not all(c in '01234567' for c in octal_perms):
                    return "Invalid input: Permissions must be a three-digit octal number."

                permissions = {
                    '0': "No permissions",
                    '1': "Execute",
                    '2': "Write",
                    '3': "Write and execute",
                    '4': "Read",
                    '5': "Read and execute",
                    '6': "Read and write",
                    '7': "Read, write, and execute"
                }                
                result = [
                    'Owner: ' + permissions[octal_perms[0]],
                    'Group: ' + permissions[octal_perms[1]],
                    'Others: ' + permissions[octal_perms[2]]
                ]
                return ', '.join(result)

            def get(self, path, follow_symlinks=True):
                """
                Retrieves file or directory information at a specified path.

                This method returns detailed information about the file or directory at the given path, such as size,
                permissions, timestamps, and more. It can also resolve symbolic links based on the 'follow_symlinks' parameter.

                Parameters:
                    path (str): The filesystem path of the file or directory for which to gather information.
                    follow_symlinks (bool, optional): Whether to follow symbolic links or not. Defaults to True.

                Returns:
                    dict or None: A dictionary containing various attributes of the file or directory if the path exists, otherwise None.
                """            	
                try:
                    if not os.path.exists(path):
                        logger.error(f"Error: Path '{path}' does not exist.")
                        return None
                    stat_info = os.stat(path) if follow_symlinks else os.lstat(path)
                    file_info = {
                        'absolute_path': os.path.abspath(path),
                        'file': stat.S_ISREG(stat_info.st_mode),
                        'dir': stat.S_ISDIR(stat_info.st_mode),
                        'link': stat.S_ISLNK(stat_info.st_mode),
                        'size': stat_info.st_size,
                        'permissions_oct': oct(stat_info.st_mode)[-3:],  # Last three octal digits of the mode.
                        'permissions': [self._interpret_permissions(oct(stat_info.st_mode)[-3:])],  # Human-readable permissions.
                        'last_accessed': time.ctime(stat_info.st_atime),
                        'last_modified': time.ctime(stat_info.st_mtime),
                        'creation_time': time.ctime(stat_info.st_ctime),  # Creation time on Windows; metadata change time on Unix.
                        'device': stat_info.st_dev,
                        'inode': stat_info.st_ino,
                        'hard_links': stat_info.st_nlink,
                        'readable': os.access(path, os.R_OK),
                        'writable': os.access(path, os.W_OK),
                        'executable': os.access(path, os.X_OK)
                    }
                    if file_info['link']:
                        try:
                            file_info['link_target'] = os.readlink(path)
                        except OSError as e:
                            logger.error(f"Error resolving symlink for '{path}': {e}")
                    if sys.platform.startswith("win") and pwd and grp:
                        file_info['owner_user'] = pwd.getpwuid(stat_info.st_uid).pw_name
                        file_info['owner_group'] = grp.getgrgid(stat_info.st_gid).gr_name
                    else:
                        file_info['owner_user'] = stat_info.st_uid if sys.platform.startswith("win") else "N/A"
                        file_info['owner_group'] = stat_info.st_gid if sys.platform.startswith("win") else "N/A"
                    return file_info
                except OSError as e:
                    logger.error(f"Error accessing path '{path}': {e}")
                    return None

            def __dir__(self):
                return ['get']

    class SaveClass:
        def __init__(self, parent):
            self.parent = parent

        def to_pkl(self, data, filename="data.pkl", path=None, create_dir=True):
            """
            Save a collection of data to a file using pickle.

            Parameters:
                data: Data to save (must be picklable).
                filename: Name of the file to save the data in. Defaults to 'data.pkl'.
                path: Directory path to save the file. Defaults to the instance's base path.
                create_dir: Whether to create the directory if it doesn't exist. Defaults to True.
            """
            filename = self.parent.File.Extension.join("pkl", filename)
            save_path = PathlibPath(path) if path else self.parent.base_path
            save_path = save_path.resolve()
            if create_dir:
                save_path.mkdir(parents=True, exist_ok=True)
            file_path = save_path / filename
            try:
                with file_path.open('wb') as file:
                    pickle.dump(data, file)
                logger.info(f"Data successfully saved to: {file_path}")
            except (OSError, pickle.PickleError) as e:
                logger.error(f"Failed to save data to {file_path}. Error: {e}")

        def __dir__(self):
            return ['to_pkl']   
           
    class LoadClass:
        def __init__(self, parent):
            self.parent = parent

        def from_pkl(self, filename="data.pkl", path=None):
            """Load pickled data from a local file. Only load trusted input."""
            filename = self.parent.File.Extension.join("pkl", filename)
            load_path = PathlibPath(path) if path else self.parent.base_path
            file_path = (load_path / filename).resolve()
            if not file_path.exists():
                logger.warning(f"No data file found at: {file_path}")
                return None
            try:
                with file_path.open('rb') as file:
                    data = pickle.load(file)
                logger.info(f"Data successfully loaded from: {file_path}")
                return data
            except (OSError, pickle.PickleError) as e:
                logger.error(f"Failed to load data from {file_path}. Error: {e}")
                return None

        def __dir__(self):
            return ['from_pkl']   
           
    class BufferClass:
        """Inspect local data sizes and calculate chunk sizes."""
        def __init__(self):
            pass      

        class _chunk_util:
            """A helper class for chunk-related calculations."""
            @staticmethod
            def is_multiple_of_320(number):
                return number % 320 == 0

            @staticmethod
            def next_multiple_of_320(number):
                return number if number % 320 == 0 else number + (320 - number % 320)

        def _ismemory(self, source):
            return any(isinstance(source, cls) for cls in (bytes, io.BytesIO))

        def chunk_calculation(self, file_source, chunk_property=None):
            """
            Dynamically calculates chunk sizes for file uploads based on the file size. The method determines preferred chunk size, 
            maximum chunk size, and chunk multiples to optimize file upload efficiency and compatibility with storage systems.

            Parameters:
                file_source (str, io.BytesIO, or bytes): The data source for which chunk sizes are calculated.
                chunk_property (str, optional): The key or synonym for the specific property to return.
                    Supported keys/synonyms:
                        - "preferred_chunk_size" or "preferred"
                        - "max_chunk_size" or "max"
                        - "chunk_multiple" or "multiple"
                    If None, returns the full dictionary of chunk properties.

            Returns:
                dict or int: The full dictionary of chunk properties if `chunk_property` is None.
                            
            """
            __chunk_util = self._chunk_util() 
            file_size_bytes = self.size(file_source)
            base_chunk_size = 5 * 1024 * 1024

            # Choose a chunk size, ceiling, and alignment for each file-size range.
            if file_size_bytes < 1 * 1024 * 1024:
                preferred_chunk_size = 64 * 1024
                max_chunk_size = 256 * 1024
                chunk_multiple = 32 * 1024

            elif file_size_bytes <= 10 * 1024 * 1024:
                preferred_chunk_size = 1 * 1024 * 1024
                max_chunk_size = 5 * 1024 * 1024
                chunk_multiple = 128 * 1024
                
            elif file_size_bytes <= 100 * 1024 * 1024:
                preferred_chunk_size = 5 * 1024 * 1024
                max_chunk_size = 10 * 1024 * 1024
                chunk_multiple = 256 * 1024
                
            elif file_size_bytes <= 1 * 1024 * 1024 * 1024:
                preferred_chunk_size = 10 * 1024 * 1024
                max_chunk_size = 60 * 1024 * 1024
                chunk_multiple = 512 * 1024
            
            else:
                preferred_chunk_size = 50 * 1024 * 1024
                max_chunk_size = 100 * 1024 * 1024
                chunk_multiple = 1 * 1024 * 1024
                
            chunk_multiple = __chunk_util.next_multiple_of_320(chunk_multiple)
            adjusted_preferred_chunk_size = preferred_chunk_size - (preferred_chunk_size % chunk_multiple) # Round the preferred size down to the required alignment.
            full_chunks = file_size_bytes // adjusted_preferred_chunk_size
            remaining_bytes = file_size_bytes % adjusted_preferred_chunk_size
            chunk_sizes = [adjusted_preferred_chunk_size] * full_chunks
            if remaining_bytes > 0: # Round a nonempty final chunk up, subject to the size ceiling.
                last_chunk_size = remaining_bytes
                if last_chunk_size % chunk_multiple != 0:
                    last_chunk_size += chunk_multiple - (last_chunk_size % chunk_multiple)
                last_chunk_size = min(last_chunk_size, max_chunk_size)
                chunk_sizes.append(last_chunk_size)

            chunk_data = {
                "preferred_chunk_size": adjusted_preferred_chunk_size,
                "max_chunk_size": max_chunk_size,
                "chunk_multiple": chunk_multiple,
            }

            synonym_map = {
                "preferred_chunk_size": "preferred_chunk_size",
                "preferred": "preferred_chunk_size",
                "max_chunk_size": "max_chunk_size",
                "max": "max_chunk_size",
                "chunk_multiple": "chunk_multiple",
                "multiple": "chunk_multiple",
            }
            
            if chunk_property:
                resolved_key = synonym_map.get(chunk_property.lower())
                if not resolved_key:
                    valid_keys = list(synonym_map.keys())
                    raise ValueError(f"Invalid chunk_property: '{chunk_property}'. Valid keys/synonyms are {valid_keys}.")
                return chunk_data[resolved_key]
            return chunk_data

        def parse_size(self, size_str):
            """Parse a size string into bytes, treating an omitted unit as MB."""
            size_str = size_str.strip()
            match = re.match(r'(\d+(?:\.\d+)?)\s*([KMGT]?B)?', size_str, re.IGNORECASE)
            if not match:
                raise ValueError("Invalid size format. Please use formats like '4MB', '4.0MB', etc.")
            size, unit = match.groups()
            size = float(size)
            unit = unit if unit else 'MB'  # Default to MB if no unit is provided
            unit_map = {'B': 0, 'KB': 1, 'MB': 2, 'GB': 3, 'TB': 4}
            bytes_size = size * (1024 ** unit_map[unit.upper()])
            return int(bytes_size)

        def convert_size(self, size):
            """Convert a byte count to the selected unit or a formatted size."""
            if isinstance(size, str):
                return self.parse_size(size)
            if size == 0:
                return "0B"
            size_names = ("B", "KB", "MB", "GB", "TB")
            i = int(math.floor(math.log(size, 1024)))
            p = math.pow(1024, i)
            s = round(size / p, 2)
            return f"{s} {size_names[i]}"

        def size(self, source, units=False):
            """
            Retrieves the size of a data source, such as a file, directory, or in-memory buffer.

            Parameters:
                source (str or io.BytesIO): The data source, which can be:
                    - A path to a file or directory.
                    - An in-memory buffer (e.g., io.BytesIO).
                units (bool, optional): If True, returns the size in a human-readable format (e.g., "1.5 GB").
                                        If False, returns the size in bytes. Default is False.

            Returns:
                int or str: The size of the data source in bytes or a human-readable format.

            Raises:
                ValueError: If the source type is unsupported.
            """
            if isinstance(source, str):
                if os.path.isfile(source):
                    total_size = os.path.getsize(source)
                elif os.path.isdir(source):
                    total_size = 0
                    for dirpath, dirnames, filenames in os.walk(source):
                        for f in filenames:
                            fp = os.path.join(dirpath, f)
                            if not os.path.islink(fp):
                                total_size += os.path.getsize(fp)
                else:
                    raise ValueError(f"Invalid path: '{source}' does not refer to a valid file or directory.")
            elif isinstance(source, io.BytesIO):
                total_size = source.getbuffer().nbytes
            elif isinstance(source, bytes):
                total_size = len(source)        
            else:
                raise ValueError("Unsupported source type. Must be a path or io.BytesIO buffer.")
            return self.convert_size(total_size) if units else total_size

        def readlocal(self, file_source):
            """Read a local file into memory."""
            if not self._ismemory(file_source):
                chunk_size = self.chunk_calculation(file_source, chunk_property="preferred")
                with open(file_source, 'rb') as file:
                    file_contents = b''
                    while True:
                        chunk = file.read(chunk_size)
                        if not chunk:
                            break
                        file_contents += chunk
                return file_contents
            else:
                raise ValueError("The provided source is an in-memory type. Expected types should not be bytes or io.BytesIO.")

        def read_in_chunks(self, source, chunk_size):
            """Yield chunks of the requested size from a file."""
            if isinstance(source, (io.BytesIO, bytes)):
                source.seek(0)
                while True:
                    chunk = source.read(chunk_size)
                    if not chunk:
                        break
                    yield chunk
            else:
                with open(source, 'rb') as file:
                    while True:
                        chunk = file.read(chunk_size)
                        if not chunk:
                            break
                        yield chunk

        def __dir__(self):
            return ['parse_size', 'convert_size', 'size', 'chunk_calculation', 'readlocal', 'read_in_chunks']  


    class StreamClass:
        """Read, write, and edit local text files."""
        def __init__(self, encoding: str = 'utf-8'):
            """
            Initializes the StreamClass with a specified encoding.

            Args:
                encoding (str): The encoding to use for file operations. Default is 'utf-8'.
            """
            self.encoding = encoding

        def _to_path(self, file: Union[str, PathlibPath]) -> PathlibPath:
            """
            Converts a file input to a pathlib.PathlibPath object if it isn't one already.
            """
            if not isinstance(file, PathlibPath):
                file = PathlibPath(file)
            return file

        def save(self, file: Union[str, PathlibPath], s: str, overwrite: bool = True) -> None:
            """
            Writes a string to a file. Can overwrite or append based on the mode.

            Args:
                file (Union[str, PathlibPath]): PathlibPath to the file.
                s (str): The string content to write to the file.
                overwrite (bool): Whether to overwrite the file. Default is True.
            """
            file = self._to_path(file)
            mode = 'w' if overwrite else 'a'
            try:
                with file.open(mode, encoding=self.encoding) as writer:
                    writer.write(s)
            except OSError as e:
                raise IOError(f"Error writing to file {file}: {e}") from e

        def add(self, file: Union[str, PathlibPath], s: str) -> None:
            """
            Appends a string to the end of a file. Creates the file if it doesn't exist.

            Args:
                file (Union[str, PathlibPath]): PathlibPath to the file.
                s (str): The string content to append to the file.
            """
            file = self._to_path(file)
            try:
                if not file.exists():
                    self.save(file, s, overwrite=True)
                else:
                    with file.open('a', encoding=self.encoding) as appender:
                        appender.write(s)
            except OSError as e:
                raise IOError(f"Error appending to file {file}: {e}") from e

        def insert(self, file: Union[str, PathlibPath], s: str, line: int) -> None:
            """
            Inserts a string at a specific line in the file. Adds to the end if the line number is invalid.

            Args:
                file (Union[str, PathlibPath]): PathlibPath to the file.
                s (str): The string to insert.
                line (int): The line number to insert the string at. Use -1 for appending.
            """
            file = self._to_path(file)
            try:
                with file.open('r', encoding=self.encoding) as reader:
                    lines = reader.readlines()
            except FileNotFoundError:
                self.save(file, s + '\n', overwrite=True)
                return
            except OSError as e:
                raise IOError(f"Error reading file {file}: {e}") from e

            if line == len(lines) or line == -1:
                lines.append(s + '\n')
            else:
                # Negative indices count from the end.
                if line < 0:
                    line = max(len(lines) + line, 0)
                lines.insert(line, s + '\n')

            try:
                with file.open('w', encoding=self.encoding) as writer:
                    writer.writelines(lines)
            except OSError as e:
                raise IOError(f"Error writing to file {file}: {e}") from e

        def fetch(self, file: Union[str, PathlibPath], silent: bool = False) -> str:
            """
            Reads the content of a file.

            Args:
                file (Union[str, PathlibPath]): PathlibPath to the file.
                silent (bool): If True, returns an empty string if the file doesn't exist. Default is False.

            Returns:
                str: The content of the file.
            """
            file = self._to_path(file)
            if not file.exists():
                if silent:
                    return ''
                else:
                    raise FileNotFoundError(f"File not found: {file}")
            try:
                with file.open('r', encoding=self.encoding) as reader:
                    return reader.read()
            except OSError as e:
                raise IOError(f"Error reading file {file}: {e}") from e

        def modify(self, file: Union[str, PathlibPath], new: str, old: str = None, pattern: str = None) -> None:
            """
            Modifies the content of a file by replacing a specific string or matching a pattern.

            Args:
                file (Union[str, PathlibPath]): PathlibPath to the file.
                new (str): The new string to replace with.
                old (str, optional): The string to replace. Either 'old' or 'pattern' must be provided.
                pattern (str, optional): A regex pattern to match for replacement.
            """
            if old is None and pattern is None:
                raise ValueError("Either 'old' or 'pattern' must be provided for replacement.")
            content = self.fetch(file)
            if old is not None:
                content = content.replace(old, new)
            if pattern is not None:
                content = re.sub(pattern, new, content)
            self.save(file, content, overwrite=True)

        def __dir__(self):
            return ['add', 'fetch', 'insert', 'modify', 'save']

    def __dir__(self):
        __additional_methods__ = []
        __methods__ = ['base_path', 'chdir', 'resetdir', 'cleanup_cache', 'create_data_directory', 'files', 'full_path', 'invalidate_cache', 'search', 'directories', 'path_type', 'is_path_struct', 'join', 'kind', 'get_module_dir']
        __inner_classes__ = ['Text', 'File', 'Dirs', 'Save', 'Buffer', 'Stream', 'Load']
        __properties__ = ['delete']        
        __base_methods__ = __methods__ + __inner_classes__ + __properties__ 
        if sys.platform.startswith("win"):
            __additional_methods__.append('roaming_path')
            __additional_methods__.append('create_shortcut')            
        return list(set(__base_methods__ + __additional_methods__))


host = pathdir()

__all__ = ['host']
def __dir__(): return __all__
