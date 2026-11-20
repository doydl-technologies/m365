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
"""Windows directory lookup and shortcut creation helpers.

Directory locators use Windows shell and system APIs. Qlnk constructs shortcut
properties through PowerShell and validates configured executable and icon paths.
These helpers do not change system policies.
"""
import os
import sys
import subprocess
from pathlib import Path as PathlibPath

import ctypes
from ctypes import wintypes, c_int, windll, POINTER


class _windows_dir:
    def __init__(self):    
        self._download = self._get_download()
        self._system32 = self._get_system32()
        self._programfiles = self._get_programfiles()
        self._programfiles86 = self._get_programfiles86()     
        self._cdrive = self._get_base_drive_root()        

    def _get_download(self):
        from ctypes import wintypes, POINTER, byref, windll, create_unicode_buffer
        from ctypes.wintypes import DWORD, WORD, BYTE
        HRESULT = ctypes.c_long
        WCHAR = ctypes.c_wchar
        class GUID(ctypes.Structure):
            _fields_ = [("Data1", DWORD), ("Data2", WORD), ("Data3", WORD), ("Data4", BYTE * 8)]
        SHGetKnownFolderPath = windll.shell32.SHGetKnownFolderPath
        SHGetKnownFolderPath.argtypes = [POINTER(GUID), DWORD, wintypes.HANDLE, POINTER(POINTER(WCHAR))]
        SHGetKnownFolderPath.restype = HRESULT
        FOLDERID_Downloads = GUID(0x374DE290, 0x123F, 0x4565, (0x91, 0x64, 0x39, 0xC4, 0x92, 0x5E, 0x46, 0x7B))
        path_ptr = POINTER(WCHAR)()
        result = SHGetKnownFolderPath(byref(FOLDERID_Downloads), 0, 0, byref(path_ptr))
        try:
            if result == 0:
                return ctypes.wstring_at(path_ptr)
            else:
                raise OSError(f"SHGetKnownFolderPath failed with error code: {result}")
        finally:
            if path_ptr:
                windll.ole32.CoTaskMemFree(path_ptr)

    def _get_system32(self):
        CSIDL_SYSTEM = 37
        shfolder = ctypes.windll.shell32.SHGetFolderPathW
        shfolder.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR]
        path_buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        result = shfolder(0, CSIDL_SYSTEM, 0, 0, path_buf)
        if result == 0:
            return path_buf.value
        else:
            return None
        
    # C:\Program Files         
    def _get_programfiles(self):  	
        CSIDL_PROGRAM_FILES = 38
        shfolder = ctypes.windll.shell32.SHGetFolderPathW
        shfolder.argtypes = [wintypes.HWND, c_int, wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR]
        path_buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        result = shfolder(0, CSIDL_PROGRAM_FILES, 0, 0, path_buf)
        if result == 0:
            return path_buf.value  
                  
    # C:\Program Files (x86)                  
    def _get_programfiles86(self):  	
        CSIDL_PROGRAM_FILESX86 = 42 
        shfolder = ctypes.windll.shell32.SHGetFolderPathW
        shfolder.argtypes = [wintypes.HWND, c_int, wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR]
        path_buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        result = shfolder(0, CSIDL_PROGRAM_FILESX86, 0, 0, path_buf)
        if result == 0:
            return path_buf.value

    def _get_base_drive_root(self):
        CSIDL_PROGRAM_FILES = 38
        SHGetFolderPath = ctypes.windll.shell32.SHGetFolderPathW
        SHGetFolderPath.argtypes = [wintypes.HWND, c_int, wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR]
        path_buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        result = SHGetFolderPath(0, CSIDL_PROGRAM_FILES, 0, 0, path_buf)
        if result == 0:
            # Fall back to the conventional C: drive when system lookup fails.
            return str(PathlibPath(path_buf.value).joinpath('C:\\'))

    @property
    def Downloads(self):
        return self._download
    @property
    def System32(self):
        return self._system32
    @property
    def ProgramFiles(self):
        return self._programfiles
    @property
    def ProgramFiles86(self):
        return self._programfiles86
       
    def CDrive(self, directory=None):
        """Return the path to the C: drive or a specific directory under it."""
        if directory:
            windowsDir = self._cdrive
            directory = PathlibPath(directory)
            fpath = PathlibPath(windowsDir).parent
            dirpath = fpath / directory
            if dirpath.exists() and dirpath.is_dir():
                return str(dirpath)
            else:
                raise FileNotFoundError(f"The specified directory '{directory}' does not exist or is not a directory.")
        else:
            return self._cdrive    

    def __dir__(self):
        if sys.platform.startswith("win"):    
            return ['Downloads', 'System32', 'ProgramFiles', 'ProgramFiles86', 'CDrive']     





class _app_data_path:
    def __init__(self):
        self._appdata = self._get_app_data_path()
        self._local = self._get_local_path()
        self._locallow = self._get_local_low_path()
        self._roaming = self._get_roaming_path()        

    def _get_app_data_path(self):
        CSIDL_APPDATA = 0x001a
        SHGFP_TYPE_CURRENT = 0
        path_buffer = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        sh_get_folder_path = windll.shell32.SHGetFolderPathW
        sh_get_folder_path.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.HANDLE, ctypes.c_uint, wintypes.LPWSTR]
        result = sh_get_folder_path(0, CSIDL_APPDATA, 0, SHGFP_TYPE_CURRENT, path_buffer)
        if result == 0:
            app_data = path_buffer.value
            base_app_data = PathlibPath(app_data).parent
            if base_app_data.exists() and base_app_data.is_dir():
                return str(base_app_data)
            else:
                raise OSError("The specified directory does not exist or is not a directory.")           
        else:
            raise OSError(f"Failed with error code: {result}")

    def _get_local_path(self):
        CSIDL_LOCAL_APPDATA = 0x001c
        SHGFP_TYPE_CURRENT = 0
        sh_get_folder_path = ctypes.windll.shell32.SHGetFolderPathW
        sh_get_folder_path.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.HANDLE, ctypes.c_uint, wintypes.LPCWSTR]
        path_buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        result = sh_get_folder_path(0, CSIDL_LOCAL_APPDATA, 0, SHGFP_TYPE_CURRENT, path_buf)
        if result == 0:
            return path_buf.value

    def _get_local_low_path(self):
        from ctypes.wintypes import HANDLE, LPWSTR, DWORD    	
        class GUID(ctypes.Structure):
            _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD), ("Data3", wintypes.WORD), ("Data4", ctypes.c_ubyte * 8)]
        FOLDERID_LocalAppDataLow = GUID(0xA520A1A4, 0x1780, 0x4FF6, (0xBD, 0x18, 0x16, 0x73, 0x43, 0xC5, 0xAF, 0x16))
        SHGetKnownFolderPath = windll.shell32.SHGetKnownFolderPath
        SHGetKnownFolderPath.argtypes = [ctypes.POINTER(GUID), wintypes.DWORD, wintypes.HANDLE, ctypes.POINTER(wintypes.LPWSTR)]
        SHGetKnownFolderPath.restype = ctypes.HRESULT
        path_ptr = wintypes.LPWSTR()
        result = SHGetKnownFolderPath(ctypes.byref(FOLDERID_LocalAppDataLow), 0, 0, ctypes.byref(path_ptr))
        try:
            if result == 0:
                return path_ptr.value
            else:
                raise OSError(f"SHGetKnownFolderPath failed with error code: {result}")
        finally:
            if path_ptr:
                windll.ole32.CoTaskMemFree(path_ptr)

    def _get_roaming_path(self):
        CSIDL_APPDATA = 0x001a
        SHGFP_TYPE_CURRENT = 0
        sh_get_folder_path = ctypes.windll.shell32.SHGetFolderPathW
        sh_get_folder_path.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.HANDLE, ctypes.c_uint, wintypes.LPCWSTR]
        path_buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        result = sh_get_folder_path(0, CSIDL_APPDATA, 0, SHGFP_TYPE_CURRENT, path_buf)
        if result == 0:
            return path_buf.value
                  
    @property
    def AppData(self):
        return self._appdata
    @property
    def Local(self):
        return self._local
    @property
    def LocalLow(self):
        return self._locallow
    @property
    def Roaming(self):
        return self._roaming

    def __dir__(self):
        """ Lists the public methods of the class. """
        return ['AppData', 'Local', 'LocalLow', 'Roaming']


class _app_data_wrapper:
    def __init__(self, paths: _app_data_path):
        self._paths = paths

    def __str__(self):
        return self._paths.AppData

    def __repr__(self):
        return self._paths.AppData

    def __getattr__(self, item):
        if item in ["Local", "LocalLow", "Roaming"]:
            return getattr(self._paths, item)
        return None         



class Qlnk:
    """Create Windows shortcuts through a selected PowerShell executable."""

    def __init__(self,
                 target_file=None,
                 shortcut_path=None,
                 shortcut_description=None,
                 shortcut_iconLocation=None,
                 ps_version=None):
        """
        Initializes the Qlnk object with optional arguments. Arguments can be
        set or updated later using the `create` method or the `_set_arguments`
        method.

        Args:
            target_file (str, optional): The absolute path of the target file
                for the shortcut.
            shortcut_path (str, optional): The absolute path where the shortcut
                file will be created.
            shortcut_description (str, optional): A descriptive string for
                the shortcut. Defaults to "My Shortcut".
            shortcut_iconLocation (str, optional): The path to the icon file.
                Defaults to the target file's icon.
            ps_version (str, optional): Specifies the PowerShell version to
                use for executing the command. Defaults to 'powershell.exe'.
                Supported values include:
                - 'powershell.exe' (Windows PowerShell)
                - 'powershell' (alias for Windows PowerShell)
                - 'pwsh' (PowerShell Core, cross-platform).
        """
        if not sys.platform.startswith('win'):
            raise NotImplementedError(
                "Qlnk is only supported on Windows operating systems."
            )

        self.ps_version = self._validate_ps_version(ps_version or 'powershell.exe')
        self.target_path = None
        self.shortcut_path = None
        self.shortcut_description = shortcut_description or "My Shortcut"
        self.icon_location = None
        self.script = None
        self.command = None

        if target_file and shortcut_path:
            self._set_arguments(
                target_file=target_file,
                shortcut_path=shortcut_path,
                shortcut_description=shortcut_description,
                shortcut_iconLocation=shortcut_iconLocation,
                ps_version=self.ps_version
            )

    def _validate_ps_version(self, ps_version):
        """
        Validate the requested PowerShell version. If invalid, raise ValueError.

        Args:
            ps_version (str): The requested PowerShell executable name.

        Returns:
            str: The validated (or default) PowerShell version.
        """
        valid_versions = ['powershell.exe', 'powershell', 'pwsh']
        if ps_version.lower() not in valid_versions:
            raise ValueError(
                f"Invalid PowerShell version '{ps_version}'. "
                f"Supported values: {', '.join(valid_versions)}"
            )
        return ps_version

    def _set_arguments(self,
                       target_file,
                       shortcut_path,
                       shortcut_description=None,
                       shortcut_iconLocation=None,
                       ps_version=None):
        """
        Sets the arguments for the Qlnk instance. This allows initializing the
        class without parameters and setting them later.

        Args:
            target_file (str): The absolute (or relative) path of the target file for the shortcut.
            shortcut_path (str): The absolute (or relative) path where the
                shortcut file will be created.
            shortcut_description (str, optional): A descriptive string for the
                shortcut. Defaults to "My Shortcut".
            shortcut_iconLocation (str, optional): The path to the icon file.
                Defaults to the target file's icon.
            ps_version (str, optional): Specifies the PowerShell version to use
                for executing the command. Supported values include:
                - 'powershell.exe' (Windows PowerShell)
                - 'powershell' (alias for Windows PowerShell)
                - 'pwsh' (PowerShell Core, cross-platform).

        Raises:
            FileNotFoundError: If the target file does not exist.
            ValueError: If an invalid icon file type is provided or if the
                PowerShell version is not recognized.
        """
        if ps_version:
            self.ps_version = self._validate_ps_version(ps_version)

        self.target_path = PathlibPath(target_file).expanduser().resolve()
        if not self.target_path.exists():
            raise FileNotFoundError(f"The target file does not exist: {self.target_path}")

        self.shortcut_path = PathlibPath(shortcut_path).expanduser().resolve()
        if self.shortcut_path.suffix.lower() != '.lnk':
            self.shortcut_path = self.shortcut_path.with_suffix('.lnk')

        if not self.shortcut_path.parent.exists():
            self.shortcut_path.parent.mkdir(parents=True, exist_ok=True)

        self.shortcut_description = shortcut_description or "My Shortcut"
        if shortcut_iconLocation:
            icon_path = PathlibPath(shortcut_iconLocation).expanduser().resolve()
            if not self._is_icon(icon_path):
                raise ValueError(f"Invalid icon file type: {icon_path}")
            self.icon_location = str(icon_path)
        else:
            # Default icon to the target's own icon
            self.icon_location = f"{self.target_path},0"

        self._prepare_script()

    def _prepare_script(self):
        """Prepare the PowerShell script for the current shortcut settings."""
        if not self.target_path or not self.shortcut_path:
            raise ValueError("Both target_path and shortcut_path must be set "
                             "before creating the shortcut.")

        self.script = f'''
        # Define the path of the original file and the location of the shortcut
        $TargetFile = "{str(self.target_path)}"
        $ShortcutPath = "{str(self.shortcut_path)}"

        # Create a Shortcut object
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)

        # Configure the shortcut
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.WorkingDirectory = [System.IO.Path]::GetDirectoryName($TargetFile)
        $Shortcut.WindowStyle = 1  # 1 = Normal window, 7 = Minimized, 3 = Maximized
        $Shortcut.Description = "{self.shortcut_description}"
        $Shortcut.IconLocation = "{self.icon_location}"

        # Save the shortcut
        $Shortcut.Save()
        '''.strip()

        self.command = [self.ps_version, '-NoProfile', '-NonInteractive', '-Command', self.script]

    def _is_icon(self, icon_path):
        """
        Validates whether the provided icon file path has an acceptable file extension.

        Args:
            icon_path (Path): The path to the icon file.

        Returns:
            bool: True if the file extension is valid, False otherwise.
        """
        valid_extensions = {'.ico', '.exe', '.dll', '.cpl'}
        return icon_path.suffix.lower() in valid_extensions

    def _is_lnk(self):
        """
        Verifies the existence of the shortcut file at the specified path.

        Returns:
            bool: True if the shortcut file exists, False otherwise.
        """
        return self.shortcut_path and self.shortcut_path.exists()

    def create(self,
               target_file=None,
               shortcut_path=None,
               shortcut_description=None,
               shortcut_iconLocation=None,
               ps_version=None):
        """
        Executes the PowerShell command to create the shortcut and verifies its
        successful creation.

        Args:
            target_file (str, optional): The absolute (or relative) path of the
                target file for the shortcut.
            shortcut_path (str, optional): The absolute (or relative) path
                where the shortcut file will be created.
            shortcut_description (str, optional): A descriptive string for the
                shortcut.
            shortcut_iconLocation (str, optional): The path to the icon file.
            ps_version (str, optional): Specifies the PowerShell version to use
                for executing the command. Supported values include:
                - 'powershell.exe' (Windows PowerShell)
                - 'powershell' (alias for Windows PowerShell)
                - 'pwsh' (PowerShell Core, cross-platform).

        Returns:
            tuple: (success: bool, message: str)
                success: True if the shortcut was created successfully,
                         False otherwise.
                message: String describing success or failure reason.
        """
        if target_file and shortcut_path:
            self._set_arguments(
                target_file=target_file,
                shortcut_path=shortcut_path,
                shortcut_description=shortcut_description,
                shortcut_iconLocation=shortcut_iconLocation,
                ps_version=ps_version
            )

        if not self.script:
            raise ValueError("Script is not prepared. Ensure arguments are "
                             "set before calling 'create'.")

        process = subprocess.run(
            self.command,
            text=True,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if process.returncode == 0 and self._is_lnk():
            return True, f"Shortcut created successfully at: {self.shortcut_path}"
        else:
            error_message = (f"Failed to create shortcut. Return code: {process.returncode}. "
                             f"Stderr: {process.stderr.strip()}")
            if not self._is_lnk():
                error_message += " The shortcut file does not exist."
            return False, error_message

    def __dir__(self):
        """Lists the public methods and properties of the class."""
        return [
            'create',
            'target_path',
            'shortcut_path',
            'shortcut_description',
            'icon_location',
            '_set_arguments',
        ]


__all__ = ['_app_data_path', '_windows_dir', '_app_data_wrapper', 'Qlnk']
def __dir__(): return __all__
