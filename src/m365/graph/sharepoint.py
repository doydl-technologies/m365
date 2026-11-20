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
"""SharePoint site and document-library operations through Microsoft Graph.

Site binds a caller-supplied token provider to one selected library. Search,
upload, download and move helpers share bounded, redacted HTTP transport.
Site and drive identifiers are opaque Graph values. Paths are relative to the
selected library; name searches enumerate within configured budgets unless
indexed search is explicitly selected.

Token reads and metadata rebinding use separate locks. These are not guarantees
of arbitrary concurrent or cross-process use. Cache options remain accepted
for compatibility; the shared transport does not cache HTTP responses.
"""
import sys
from itertools import islice
import re
import time
from copy import deepcopy
from urllib.parse import urlparse, parse_qs, urlunparse, quote, unquote
import os
from pathlib import Path
import io
import logging
from typing import Optional, Union, Sequence, Literal, Dict, Any, TYPE_CHECKING

# Import AuthProvider only for static type checking to avoid a runtime cycle.
if TYPE_CHECKING:
    try:
        from .auth_types import AuthProvider
    except ImportError:    
        from .auth_types import AuthProvider
    
import requests
from azauthlib import Authentication as _AzAuth
from requests.adapters import HTTPAdapter
from requests.sessions import Session

from . import _syslib as syslib
from ._graph import GraphAPI
from ._site_resolution import resolve_site, rebind, library_relative
from ._legacy_paths import strip_file_extension, drop_library_prefix, relative_path as legacy_relative_path

from . import _misc
from ._preprocessing import filters
from .operations import doc_management
from .operations import doc_transfer
from .operations import doc_sync
from .operations import doc_dispatch


logger = logging.getLogger(__name__)


class SiteMeta(type):
    def __dir__(cls):
        """List class attributes without the internal helper namespaces.

        This affects discovery only; direct attribute access is unchanged.
        """
        default_names = super().__dir__()
        
        hidden_names = {
            'API', 'Operations', 'PATH', 'SearchBy', 'Tools', 'URL',
            '_check_id_exist', '_get_site_info', '_linked', '_validate_site_info',
            'update_drive_name', 'update_site_name'
        }
        return [name for name in default_names if name not in hidden_names]


class Site(metaclass=SiteMeta):
    """Bind authentication, resource resolution and document operations to one library.

    The api, url, path, tools, search_by and operations namespaces share the selected
    site and drive metadata. Token refresh belongs to the supplied provider, and
    resource grants remain subject to the tenant's policies.
    """
    def __init__(
        self,
        token_engine: Optional['AuthProvider'] = None,
        site_name: Optional[str] = None,
        drive_name: Optional[str] = None,
        use_cache: bool = False,
        cache_lifetime: int = 0,
        *,
        scopes: Optional[Union[str, Sequence[str]]] = None,
        token_path: Optional[str] = None,
        authority: Optional[str] = None,
        build: Optional[Literal["Default", "WithEnvFile", "WithOSEnv", "WithEntry"]] = None,
        build_kwargs: Optional[Dict[str, Any]] = None,
        enable_concurrency: bool = True,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
        hostname: Optional[str] = None,
        graph_url: str = "https://graph.microsoft.com/v1.0",
        request_timeout=(10, 120),
        max_retries: int = 3,
        search_mode: str = "enumerate",
        search_root: Optional[str] = None,
        search_max_items: int = 5000,
        search_max_requests: int = 200,
        search_max_depth: int = 30,
    ):
        """Initialize authentication, helper namespaces and selected-library metadata.

        A supplied token_engine provides access_token. Otherwise build and build_kwargs
        select an azauthlib builder. site_name accepts a site URL, hostname/path
        or short name; site_id accepts the complete opaque Graph identifier. drive_id
        or drive_name selects a library within that site; omission selects its default.

        request_timeout and max_retries bound transport work. search_mode selects
        bounded enumeration or indexed search; search_root and search_max_* constrain
        name traversal. Paths remain relative to the selected library.

        use_cache, cache_lifetime and enable_concurrency remain configuration parameters.
        The shared transport does not cache responses. Token access and metadata
        rebinding have separate locks, not a general cross-process guarantee.
        Repeated initialization of an initialized instance is skipped.
        """
        if not hasattr(self, "_initialized") or not self._initialized:

            import threading
            self._token_lock = threading.Lock()

            self._metadata_lock = threading.RLock()
            
             
            self.enable_concurrency = enable_concurrency

            # Concurrency disables the legacy GET cache settings.
            if enable_concurrency:
                self.use_cache = False
                self.cache_lifetime = 0
            else:
                self.use_cache = use_cache
                self.cache_lifetime = cache_lifetime

            if token_engine is None:
                self.token_engine = self._build_token_engine(
                    scopes=scopes,
                    token_path=token_path,
                    authority=authority,
                    build=build,
                    build_kwargs=build_kwargs or {},
                )
            else:
                self.token_engine = token_engine

            self._site_name = site_name
            self._drive_name = drive_name
            from urllib.parse import urlsplit
            endpoint = urlsplit(graph_url)
            if endpoint.scheme != "https" or not endpoint.hostname or endpoint.username or endpoint.password or endpoint.query or endpoint.fragment:
                raise ValueError("graph_url must be a clean HTTPS Graph endpoint")
            if not isinstance(max_retries, int) or not 0 <= max_retries <= 10:
                raise ValueError("max_retries must be between 0 and 10")
            timeouts = request_timeout if isinstance(request_timeout, (list, tuple)) else (request_timeout,)
            import math
            if not timeouts or len(timeouts) not in (1, 2) or any(not isinstance(t, (int, float)) or not math.isfinite(t) or t <= 0 for t in timeouts):
                raise ValueError("request_timeout must contain positive finite seconds")
            self.hostname = hostname
            self.request_timeout = tuple(timeouts) if len(timeouts) == 2 else timeouts[0]
            self.max_retries = max_retries
            if search_mode not in ("enumerate", "indexed"):
                raise ValueError("search_mode must be enumerate or indexed")
            if any(not isinstance(n, int) or isinstance(n, bool) or n < 1 for n in (search_max_items, search_max_requests, search_max_depth)):
                raise ValueError("Search budgets must be positive integers")
            if search_root is not None and (not isinstance(search_root, str) or not search_root.strip()):
                raise ValueError("search_root must be an item ID or None")
            self.search_mode, self.search_root = search_mode, search_root
            self.search_max_items = search_max_items
            self.search_max_requests = search_max_requests
            self.search_max_depth = search_max_depth
            self.graph_url = graph_url.rstrip("/")
            self.batch_url = f"{self.graph_url}/$batch"

            self.cache = {}

            self.api = self.API(self)
            self.url = self.URL(self)
            self.path = self.PATH(self)
            self.tools = self.Tools(self)
            self.search_by = self.SearchBy(self)
            self.operations = self.Operations(self)

            self.web_url = None
            self.drive_web_url = None
            self.site_id = site_id
            self.drive_id = drive_id

            try:
                self._check_id_exist()
            except Exception:
                self.api.close()
                raise

            self._initialized = True
        
    def get_token(self):
        """Read the provider token while holding the token-access lock."""
        with self._token_lock:
            return self.token_engine.access_token
           
    def _build_token_engine(
        self,
        *,
        scopes: Optional[Union[str, Sequence[str]]],
        token_path: Optional[str],
        authority: Optional[str],
        build: Optional[str],
        build_kwargs: Dict[str, Any],
    ):
        """Construct an azauthlib provider and request authentication.

        build selects Default, WithEnvFile, WithOSEnv or WithEntry. build_kwargs supplies
        the matching configuration; auth_flow selects device-code authentication for
        recognized device aliases, otherwise provider-managed Silent authentication.
        A nonempty access_token is required before returning. Token validity and
        permission checks remain the identity service's responsibility.
        """
        from ._storage_auth import _legacy_provider
        return _legacy_provider(_AzAuth, scopes=scopes, token_path=token_path,
                                authority=authority, build=build, build_kwargs=build_kwargs)
       
    @classmethod
    def from_default_env(cls, site_name, drive_name=None, **kwargs):
        """Bind a site using azauthlib's default credential discovery.

        Pass a site selector and, optionally, a library name. Extra keyword
        arguments such as scopes, token_path, and authority are forwarded to Site.
        The default sign-in is silent and requires a usable cached sign-in.
        """
        return cls(
            token_engine=None,
            site_name=site_name,
            drive_name=drive_name,
            build="Default",
            **kwargs,
        )

    @classmethod
    def from_env_file(cls, site_name, drive_name=None, env_path=None, key_map=None, **kwargs):
        """Bind a site using credentials loaded by azauthlib from an environment file.

        Use env_path to select the file and key_map to remap its variable names.
        Omit drive_name to select the site's default library. Additional keyword
        arguments are forwarded to Site.
        """
        bkwargs = {"env_path": env_path, "key_map": key_map} if env_path or key_map else {}

        return cls(
            token_engine=None,
            site_name=site_name,
            drive_name=drive_name,
            build="WithEnvFile",
            build_kwargs=bkwargs,
            **kwargs,
        )

    @classmethod
    def from_os_env(cls, site_name, drive_name=None, key_map=None, **kwargs):
        """Bind a site using credentials from the process environment.

        Use key_map when variable names differ from azauthlib's defaults.
        Omit drive_name to select the site's default library. Additional keyword
        arguments are forwarded to Site.
        """
        bkwargs = {"key_map": key_map} if key_map else {}

        return cls(
            token_engine=None,
            site_name=site_name,
            drive_name=drive_name,
            build="WithOSEnv",
            build_kwargs=bkwargs,
            **kwargs,
        )

    @classmethod
    def from_entry(cls, site_name, drive_name=None, client_id=None, tenant_id=None, client_secret=None, **kwargs):
        """Bind a site using application credentials supplied by the caller.

        Load client_id, tenant_id, and any client_secret from an approved source.
        These values configure the provider; they do not select an app-only flow.
        The default is silent delegated authentication. For an existing app-only
        profile, use from_profile.
        """
        bkwargs = {"client_id": client_id, "tenant_id": tenant_id, "client_secret": client_secret}

        return cls(
            token_engine=None,
            site_name=site_name,
            drive_name=drive_name,
            build="WithEntry",
            build_kwargs=bkwargs,
            **kwargs,
        )

    @property
    def site_name(self):
        """Return the configured SharePoint site selector.

        Assigning a different value resolves the new site and library against Graph.
        """
        return self._site_name

    @site_name.setter
    def site_name(self, value):
        if value != self._site_name:
            rebind(self, value, site=True)

    @property
    def drive_name(self):
        """Return the selected SharePoint document library name, if available.

        Assigning a different name resolves the library against Graph.
        """
        return self._drive_name

    @drive_name.setter
    def drive_name(self, value):
        if value != self._drive_name:
            rebind(self, value, site=False)

    def update_site_name(self, new_site_name):
        """Rebind to a different SharePoint site, or log that no change is needed."""
        if new_site_name != self._site_name:
            logger.info(f"Updating site name from {self._site_name} to {new_site_name}.")
            self.site_name = new_site_name  
        else:
            logger.info("No change in site name; update skipped.")

    def update_drive_name(self, new_drive_name):
        """Rebind to a different document library, or log that no change is needed."""
        if new_drive_name != self._drive_name:
            logger.info(f"Updating drive name from {self._drive_name} to {new_drive_name}.")
            self.drive_name = new_drive_name  
        else:
            logger.info("No change in drive name; update skipped.")

    def _check_id_exist(self):
        with self._metadata_lock:
            if not (self.site_id and self.drive_id and self.web_url):
                resolve_site(self)
          
    def _validate_site_info(self):
        if not self.site_id or not self.drive_id:
            raise ValueError("Site and drive IDs must be resolved")

    def _get_site_info(self):
        resolve_site(self)

    class AuthenticationFailed(RuntimeError):
        """Raised when authentication could not be completed successfully."""
        pass

    API = GraphAPI

    class Tools:
        def __init__(self, parent):
            self.parent = parent 

        def download_property_pattern(self, compiled=False):
            """Match the selected site's temporary download.aspx URL format.

            Download capability URLs are short-lived and must not be treated as durable
            file identifiers. Removing user permissions may not immediately invalidate
            an already-issued URL.
            """
            webUrl = self.parent.web_url
            escaped_webUrl = re.escape(webUrl) # Treat the site URL as a literal in the pattern.
            pattern = rf"^{escaped_webUrl}/_layouts/15/download\.aspx\?UniqueId=([a-f0-9\-]+)&Translate=(true|false)&tempauth=.+&ApiVersion=([0-9]+(?:\.[0-9]+)?)$"
            return re.compile(pattern).pattern if not compiled else re.compile(pattern)
           
        def __dir__(self):
            return ['download_property_pattern']

    class URL:
        def __init__(self, parent):
            self.parent = parent 
    	
        def is_url(self, url):
            pattern = r'https?://(?:[-\w.]|(?:%[\da-fA-F]{2}))+[^\s,.;!?]*'
            return bool(re.search(pattern, url))

        def is_absolute_url(self, url):
            return bool(urlparse(url).netloc)

        def encoded(self, url, decode=False):
            """Encode or decode a URL based on the decode flag."""
            return quote(url) if not decode else unquote(url)
           
        def ensure_trailing_slash(self, url):
            """Append a slash to a URL's path; leave non-URLs unchanged."""
            return urlunparse(urlparse(url)._replace(path=urlparse(url).path + ('/' if not urlparse(url).path.endswith('/') else ''))) if self.is_url(url) else url

        def normalize_sharepoint_path(self, full_url, remove_leading_slash=True):
            return library_relative(self.parent, full_url, leading_slash=not remove_leading_slash)
     
        def __dir__(self):
            return ['is_url', 'is_absolute_url', 'encoded', 'ensure_trailing_slash', 'normalize_sharepoint_path']
           
    class PATH:
        def __init__(self, parent):
            self.parent = parent
            self.url = parent.url

        def relative_path(self, webUrl, site_name=None, remove_leading_slash=True, remove_domain_extension=False, drop_drive=False):
            return legacy_relative_path(self.parent, webUrl, site_name, remove_leading_slash, remove_domain_extension, drop_drive)
           
        def drop_drivename(self, path):
            return drop_library_prefix(self.parent, path)
           
        def remove_leading_slash(self, path, first_only=False):
            return path[1:] if first_only and path.startswith('/') else path.lstrip('/')
               
           
        def strip_domain_extension(self, url):
            return strip_file_extension(url)

        def is_sharepoint_folderpath(self, path):
            try:
                return bool(syslib.path_type(unquote(path), posix=False, cloud=True)["folder"])
            except:
                return bool(syslib.path_type(unquote(path), posix=False, cloud=True)["dir"])

        def is_sharepoint_filepath(self, path):
            return bool(syslib.path_type(unquote(path), posix=False, cloud=True)["file"])

        def is_path_struct(self, path, posix=False, cloud=True):
            return bool(syslib.is_path_struct(unquote(path), posix=posix, cloud=cloud))

        def file_ext(self, file):
            """Return the filename extension."""
            extension = syslib.File.Extension.get(file, posix=False, cloud=True)
            return extension if syslib.path_type(file, posix=False, cloud=True)["file"] else None

        def join_file_ext(self, file, extension="xlsx"):
            """Join a filename with its extension."""
            if not syslib.File.Extension.exists(file, posix=False, cloud=True):
                return syslib.File.Extension.join(extension, file, posix=False, cloud=True)
            return file
           
        def __dir__(self):
            return ['drop_drivename', 'strip_domain_extension', 'is_sharepoint_folderpath', 'is_sharepoint_filepath', 'relative_path', 'file_ext', 'join_file_ext', 'is_path_struct']
                   
    class SearchBy:
        def __init__(self, sharepoint_instance):
            self.sharepoint_instance = sharepoint_instance
            self.Folder = doc_management.Folder            
            self.File = doc_management.File
            self.FileFolder = doc_management.FileFolder            

        def foldername(self, search_term, include_children=False, exact_match=True):
            """Find folders in the selected library using the configured search mode."""
            return self.Folder(self.sharepoint_instance).find(search_term, include_children, exact_match)
           
        def filename(self, search_term, exact_match=True, sort_order=None, num_files=None):
            """Find files in the selected library using the configured search mode."""
            return self.File(self.sharepoint_instance).find(search_term, exact_match, sort_order, num_files)
           
        def filefoldername(self, file_search, folder_search):
            """Resolve exactly one source file and one destination folder."""
            return self.FileFolder(self.sharepoint_instance).find(file_search, folder_search)

        def __dir__(self):
            return ["filename", "foldername", "filefoldername"]

    class Operations:
        def __init__(self, sharepoint_instance):
            self.sharepoint_instance = sharepoint_instance
            self.Move = doc_dispatch.Move             
            self.Upload = doc_sync.Upload 
            self.Download = doc_transfer.Download             
            
        def Mover(self, sharepoint_filename, sharepoint_foldername, new_sharepoint_filename=None):
            return self.Move(self.sharepoint_instance).move(sharepoint_filename, sharepoint_foldername, new_sharepoint_filename)
           
        def Downloader(self, max_workers=10, preserve_dir_structure=False, compress=False):
            return self.Download(self.sharepoint_instance, max_workers, preserve_dir_structure, compress)   

        def Uploader(self, file_source, sharepoint_folder_path, new_file_name=None, chunk_verbose=True, max_file_size="4MB", *, conflict_behavior="fail"):
            return self.Upload(self.sharepoint_instance, chunk_verbose, max_file_size, conflict_behavior=conflict_behavior).upload(file_source, sharepoint_folder_path, new_file_name)
           
        def __dir__(self):
            return ["Mover", "Downloader", "Uploader"]           

    @classmethod
    def from_profile(cls, profile, site_name=None, drive_name=None, *, scopes=None,
                     allow_interactive=False, **client_options):
        """Load an existing GUI profile using the same policy as Drive.from_profile."""
        import inspect
        from ._storage import _options
        from ._storage_auth import _profile_provider
        inspect.signature(cls.__init__).bind(None, token_engine=object(), site_name=site_name,
                                            drive_name=drive_name, scopes=scopes, **client_options)
        if not site_name and not client_options.get("site_id"):
            raise ValueError("Supply site_name or site_id")
        if site_name is not None and not isinstance(site_name, str):
            raise ValueError("site_name must be a string")
        if client_options.get("search_mode", "enumerate") not in ("enumerate", "indexed"):
            raise ValueError("search_mode must be enumerate or indexed")
        depth = client_options.get("search_max_depth", 30)
        if isinstance(depth, bool) or not isinstance(depth, int) or depth < 1:
            raise ValueError("Site search_max_depth must be a positive integer")
        search_root = client_options.get("search_root")
        if search_root is not None and (not isinstance(search_root, str) or not search_root.strip()):
            raise ValueError("search_root must be an item ID or None")
        options = _options(
            client_options.get("drive_id") or "validation",
            graph_url=client_options.get("graph_url", "https://graph.microsoft.com/v1.0"),
            request_timeout=client_options.get("request_timeout", (10, 120)),
            max_retries=client_options.get("max_retries", 3),
            max_requests=client_options.get("search_max_requests", 200),
            max_items=client_options.get("search_max_items", 5000),
            max_depth=client_options.get("search_max_depth", 30))
        auth = _profile_provider(profile, scopes=scopes, allow_interactive=allow_interactive,
                                 graph_url=options["graph_url"])
        return cls(token_engine=auth, site_name=site_name, drive_name=drive_name,
                   scopes=scopes, **client_options)

    def as_drive(self, folder_id=None):
        """Create an independently owned, immutable snapshot of the selected library.

        The returned Drive owns its HTTP sessions and must be closed. It remains
        bound to the captured library if this Site is rebound or closed later.
        The default anchor respects search_root. An explicit conflicting anchor
        cannot silently widen a Site configured with search_root.
        """
        from .onedrive import Drive
        with self._metadata_lock:
            if getattr(self, "_storage_closed", False):
                raise ValueError("Site client is closed")
            if self.search_root is not None and folder_id not in (None, self.search_root):
                raise ValueError("An explicit folder cannot override the configured search_root")
            provider, drive_id = self.token_engine, self.drive_id
            options = dict(folder_id=self.search_root if folder_id is None else folder_id,
                           graph_url=self.graph_url, request_timeout=self.request_timeout,
                           max_retries=self.max_retries, max_requests=self.search_max_requests,
                           max_items=self.search_max_items, max_depth=self.search_max_depth)
        return Drive(provider, drive_id, **options)

    def _storage_call(self, method, *args, **kwargs):
        with self.as_drive() as storage:
            return getattr(storage, method)(*args, **kwargs)

    def get_item(self, path=""):
        """Read metadata using the shared selected-drive storage engine."""
        return self._storage_call("get_item", path)

    def list_children(self, folder_path=""):
        return self._storage_call("list_children", folder_path)

    def find_files(self, term, *, folder_path="", exact_match=False, recursive=True,
                   limit=None, sort_order=None):
        """Modern filename search; legacy search_by behavior remains unchanged."""
        return self._storage_call("find_files", term, folder_path=folder_path,
                                  exact_match=exact_match, recursive=recursive,
                                  limit=limit, sort_order=sort_order)

    def find_folders(self, term, *, folder_path="", exact_match=False, recursive=True,
                     limit=None, sort_order=None):
        return self._storage_call("find_folders", term, folder_path=folder_path,
                                  exact_match=exact_match, recursive=recursive,
                                  limit=limit, sort_order=sort_order)

    def find_file_folder(self, file_term, folder_term, *, folder_path="",
                         exact_match=True, recursive=True):
        return self._storage_call("find_file_folder", file_term, folder_term,
                                  folder_path=folder_path, exact_match=exact_match,
                                  recursive=recursive)

    def upload(self, source, folder_path="", new_file_name=None, *, conflict_behavior="fail",
               chunk_size=10 * 327680):
        return self._storage_call("upload", source, folder_path, new_file_name,
                                  conflict_behavior=conflict_behavior, chunk_size=chunk_size)

    def download_file(self, path, directory):
        return self._storage_call("download_file", path, directory)

    def download_folder(self, folder_path, directory, *, compress=False, max_workers=4):
        return self._storage_call("download_folder", folder_path, directory,
                                  compress=compress, max_workers=max_workers)

    def create_folder(self, name, folder_path=""):
        return self._storage_call("create_folder", name, folder_path)

    def move(self, path, destination_folder="", new_name=None, *, if_match=None):
        return self._storage_call("move", path, destination_folder, new_name, if_match=if_match)

    def rename(self, path, new_name, *, if_match=None):
        return self._storage_call("rename", path, new_name, if_match=if_match)

    def clear_cache(self):
        logger.info("Clearing the cache")
        self.cache = {}

    def close(self):
        """Release HTTP connections; does not revoke or alter tenant permissions."""
        self._storage_closed = True
        self.api.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    def __dir__(self):        
        _instance_vars = ['drive_id', 'site_id', 'token_engine']
        _properties = ['drive_name', 'site_name']
        _methods = ['clear_cache', 'close', 'as_drive', 'get_item', 'list_children',
                    'find_files', 'find_folders', 'find_file_folder', 'upload',
                    'download_file', 'download_folder', 'create_folder', 'move', 'rename']
        _inner_classes = ['search_by', 'url', 'path', 'operations', 'tools']
        _public_attrs = _instance_vars + _methods + _inner_classes + _properties
        return sorted(_public_attrs)

def __dir__():
    """List module metadata and the public Site class."""
    mod = sys.modules[__name__]

    builtins = [name for name in mod.__dict__ if name.startswith('__') and name.endswith('__')]

    if 'Site' not in builtins:
        builtins.append('Site')

    return sorted(builtins)
