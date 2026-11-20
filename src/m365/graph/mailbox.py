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
"""Microsoft Graph mailbox operations with pluggable authentication.

Mail exposes api, tools, search_by and operations namespaces for folders,
messages, drafts, attachments and delta queries. The shared transport preserves
Outlook preferences and separates bearer-authenticated requests from transfer
capability URLs. Ordinary raw JSON calls retain the error-dictionary contract;
high-level reads and streaming operations raise on failure.

Cache configuration remains accepted for compatibility; the shared transport
does not cache HTTP responses. Token-access locks do not guarantee safety for
all concurrent operations or cross-process client sharing.
"""
import io
import json
import logging
from ._mail_graph import MailboxAPI
import os
import re
import shutil
import sys
import tempfile
import time
import unicodedata
from itertools import islice
from pathlib import Path
from typing import TYPE_CHECKING, Any, Dict, Literal, Optional, Sequence, Union
from urllib.parse import quote

# Import AuthProvider only for static type checking to avoid a runtime cycle.
if TYPE_CHECKING:
    try:
        from .auth_types import AuthProvider
    except ImportError:
        from .auth_types import AuthProvider

import requests
from azauthlib import Authentication as _AzAuth
from requests.adapters import HTTPAdapter

try:
    from .operations import mail_compose
    from .operations import mail_dispatch
    from .operations import mail_management
    from .operations import mail_transfer
except ImportError:
    from .operations import mail_compose
    from .operations import mail_dispatch
    from .operations import mail_management
    from .operations import mail_transfer

logger = logging.getLogger(__name__)


DEFAULT_SCOPES = ["Mail.ReadWrite", "Mail.Send", "User.Read"]
# Filter class discovery without restricting direct attribute access.
class MailboxMeta(type):
    def __dir__(cls):
        default_names = super().__dir__()
        hidden_names = {
            'API', 'SearchBy', 'Operations', 'Tools', '_check_mailbox_access'
        }
        return [name for name in default_names if name not in hidden_names]
# Bind mailbox operations to one authentication provider and resource.
# Token reads use a lock; cache options do not enable caching in the transport.
# Helper namespaces share the client state and Outlook ID preferences.
class Mail(metaclass=MailboxMeta):
    """
    Microsoft Graph-backed Outlook mailbox client.

    Parameters
    ----------
    token_engine:
        Authentication engine implementing the shared ``AuthProvider`` protocol.
        If omitted, the Mail will build one through ``azauthlib``.
    mailbox:
        Either ``"me"`` or a user identifier recognized by Microsoft Graph,
        usually a UPN or SMTP address.
    use_immutable_ids:
        If true, requests Outlook immutable IDs where applicable.
    """
    def __init__(
        self,
        token_engine: Optional['AuthProvider'],
        mailbox: str = "me",
        use_cache: bool = False,
        cache_lifetime: int = 0,
        *,
        scopes: Optional[Union[str, Sequence[str]]] = None,
        token_path: Optional[str] = None,
        authority: Optional[str] = None,
        build: Optional[Literal["Default", "WithEnvFile", "WithOSEnv", "WithEntry"]] = None,
        build_kwargs: Optional[Dict[str, Any]] = None,
        enable_concurrency: bool = True,
        use_immutable_ids: bool = True,
    ):
        if not hasattr(self, "_initialized") or not self._initialized:
            import threading; self._token_lock = threading.Lock(); self._metadata_lock = threading.Lock()
            self.enable_concurrency = enable_concurrency
            if enable_concurrency:
                self.use_cache = False; self.cache_lifetime = 0
            else:
                self.use_cache = use_cache; self.cache_lifetime = cache_lifetime
            if token_engine is None:
                self.token_engine = self._build_token_engine(
                    scopes=scopes,
                    token_path=token_path,
                    authority=authority,
                    build=build,
                    build_kwargs=build_kwargs or {},
                )
            else: self.token_engine = token_engine
            self.graph_url = "https://graph.microsoft.com/v1.0"; self.batch_url = f"{self.graph_url}/$batch"; self._mailbox = mailbox or "me"; self.use_immutable_ids = bool(use_immutable_ids); self.cache: Dict[str, Dict[str, Any]] = {}; self.mailbox_profile: Dict[str, Any] = {}; self.request_timeout = (10, 60); self.batch_request_timeout = (10, 120)
            self.api = self.API(self); self.tools = self.Tools(self); self.search_by = self.SearchBy(self); self.operations = self.Operations(self)
            self._check_mailbox_access(); self._initialized = True
    @property
    def mailbox(self) -> str: return self._mailbox
    @mailbox.setter
    def mailbox(self, value: str):
        value = value or "me"
        if value != self._mailbox:
            original = self._mailbox; original_profile = dict(self.mailbox_profile)
            try:
                self._mailbox = value; self.mailbox_profile = {}; self._check_mailbox_access(force=True)
            except Exception:
                self._mailbox = original; self.mailbox_profile = original_profile; raise
    @property
    def base_resource(self) -> str: return "me" if self.mailbox.lower() == "me" else f"users/{self.mailbox}"
    def get_token(self):
        with self._token_lock: return self.token_engine.access_token
    def initiate_device_flow(
        self,
        scopes: Optional[Union[str, Sequence[str]]] = None,
        webbrowser_enabled: bool = False,
    ):
        """Explicitly run Azure device code flow against the current token engine."""
        if not hasattr(self.token_engine, "DeviceCodeFlow"):
            raise AttributeError(
                "The configured token engine does not expose DeviceCodeFlow(). "
                "Use azauthlib.Authentication or provide a compatible auth wrapper."
            )
        with self._token_lock:
            try:
                self.token_engine.DeviceCodeFlow(
                    scopes=scopes,
                    webbrowser_enabled=webbrowser_enabled,
                )
            except TypeError:
                self.token_engine.DeviceCodeFlow(scopes=scopes)
            if not self.token_engine.access_token:
                raise RuntimeError(
                    "Authentication failed: no access token acquired via device code flow."
                )
        self.clear_cache(); return self.token_engine.access_token
    def authenticate_device_flow(
        self,
        scopes: Optional[Union[str, Sequence[str]]] = None,
        webbrowser_enabled: bool = False,
    ):
        """Alias for :meth:`initiate_device_flow`."""
        return self.initiate_device_flow(
            scopes=scopes,
            webbrowser_enabled=webbrowser_enabled,
        )
    def _build_token_engine(
        self,
        *,
        scopes: Optional[Union[str, Sequence[str]]],
        token_path: Optional[str],
        authority: Optional[str],
        build: Optional[str],
        build_kwargs: Dict[str, Any],
    ):
        build = build or "Default"; auth = _AzAuth(scopes=scopes, token_path=token_path)
        bkwargs = dict(build_kwargs or {}); auth_flow = str(bkwargs.pop("auth_flow", "interactive") or "interactive").lower(); webbrowser_enabled = bool(bkwargs.pop("webbrowser_enabled", False))
        builder = getattr(auth.Build, build, None)
        if not callable(builder):
            raise ValueError(
                f"Unknown build mode '{build}'. "
                "Use one of: Default, WithEnvFile, WithOSEnv, WithEntry."
            )
        if build == "Default": builder(authority=authority)
        else: builder(authority=authority, **bkwargs)
        if auth_flow in {"device", "device_code", "devicecode", "device-flow"}:
            try:
                auth.DeviceCodeFlow(
                    scopes=scopes,
                    webbrowser_enabled=webbrowser_enabled,
                )
            except TypeError:
                auth.DeviceCodeFlow(scopes=scopes)
        else: auth.Silent(scopes=scopes)
        if not auth.access_token: raise RuntimeError("Authentication failed: no access token acquired.")
        return auth
    @classmethod
    def from_default_env(cls, mailbox: str = "me", **kwargs): return cls(token_engine=None, mailbox=mailbox, build="Default", **kwargs)
    @classmethod
    def from_env_file(cls, mailbox: str = "me", env_path=None, key_map=None, **kwargs):
        bkwargs = {"env_path": env_path, "key_map": key_map} if env_path or key_map else {}
        return cls(
            token_engine=None,
            mailbox=mailbox,
            build="WithEnvFile",
            build_kwargs=bkwargs,
            **kwargs,
        )
    @classmethod
    def from_os_env(cls, mailbox: str = "me", key_map=None, **kwargs):
        bkwargs = {"key_map": key_map} if key_map else {}
        return cls(
            token_engine=None,
            mailbox=mailbox,
            build="WithOSEnv",
            build_kwargs=bkwargs,
            **kwargs,
        )
    @classmethod
    def from_entry(
        cls,
        mailbox: str = "me",
        client_id=None,
        tenant_id=None,
        client_secret=None,
        **kwargs,
    ):
        bkwargs = {
            "client_id": client_id,
            "tenant_id": tenant_id,
            "client_secret": client_secret,
        }
        return cls(
            token_engine=None,
            mailbox=mailbox,
            build="WithEntry",
            build_kwargs=bkwargs,
            **kwargs,
        )
    def _check_mailbox_access(self, force: bool = False):
        if self.enable_concurrency:
            with self._metadata_lock: self._probe_mailbox(force=force)
        else: self._probe_mailbox(force=force)
    def _probe_mailbox(self, force: bool = False):
        if self.mailbox_profile and not force: return
        response = self.api.graph_call(resource_url=f"{self.base_resource}/mailFolders?$top=1")
        if isinstance(response, dict) and 'error' in response: raise ValueError(f"Unable to access mailbox '{self.mailbox}': {response['error']}")
        user_response = self.api.graph_call(resource_url=f"{self.base_resource}?$select=id,displayName,mail,userPrincipalName")
        if isinstance(user_response, dict) and 'error' not in user_response: self.mailbox_profile = user_response
        else: self.mailbox_profile = {"id": None, "mailbox": self.mailbox}
    class AuthenticationFailed(RuntimeError):
        """Raised when authentication could not be completed successfully."""; pass
    API = MailboxAPI
    class Tools:
        def __init__(self, parent): self.parent = parent
        def recipient(self, value: str) -> Dict[str, Dict[str, str]]: return {"emailAddress": {"address": value}}
        def recipients(self, values: Optional[Union[str, Sequence[str]]]):
            if values is None: return []
            if isinstance(values, str): values = [values]
            return [self.recipient(v) for v in values if v]
        def well_known_folders(self):
            return [
                "inbox", "drafts", "sentitems", "deleteditems", "archive",
                "junkemail", "outbox", "clutter", "conversationhistory",
                "msgfolderroot", "recoverableitemsdeletions", "searchfolders",
            ]
        def __dir__(self): return ['recipient', 'recipients', 'well_known_folders']

    class SearchBy:
        def __init__(self, mailbox_instance):
            self.mailbox_instance = mailbox_instance; self.Folder = mail_management.Folder; self.Message = mail_management.Message; self.Attachment = mail_management.Attachment; self.Delta = mail_management.Delta
        def folders(self, search_term=None, include_hidden=False, include_children=False, exact_match=True, recursive=True, top=None):
            return self.Folder(self.mailbox_instance).find(
                search_term=search_term,
                include_hidden=include_hidden,
                include_children=include_children,
                exact_match=exact_match,
                recursive=recursive,
                top=top,
            )
        def messages(self, query=None, folder=None, top=25, order_by="receivedDateTime desc", select=None, unread=None, body_type="html", auto_paginate=True, max_items=None, page_limit=None):
            return self.Message(self.mailbox_instance).find(
                query=query,
                folder=folder,
                top=top,
                order_by=order_by,
                select=select,
                unread=unread,
                body_type=body_type,
                auto_paginate=auto_paginate,
                max_items=max_items,
                page_limit=page_limit,
            )
        def message(self, message_id, select=None, body_type="html", mime=False):
            return self.Message(self.mailbox_instance).get(
                message_id=message_id,
                select=select,
                body_type=body_type,
                mime=mime,
            )
        def attachments(self, message_id, auto_paginate=True, max_items=None, page_limit=None):
            return self.Attachment(self.mailbox_instance).list(
                message_id,
                auto_paginate=auto_paginate,
                max_items=max_items,
                page_limit=page_limit,
            )
        def message_delta(self, folder="inbox", delta_link=None, select=None, top=None, order_by=None, change_type=None):
            return self.Delta(self.mailbox_instance).messages(
                folder=folder,
                delta_link=delta_link,
                select=select,
                top=top,
                order_by=order_by,
                change_type=change_type,
            )
        def folder_delta(self, delta_link=None, top=None, include_hidden=False):
            return self.Delta(self.mailbox_instance).folders(
                delta_link=delta_link,
                top=top,
                include_hidden=include_hidden,
            )
        def __dir__(self): return ["folders", "messages", "message", "attachments", "message_delta", "folder_delta"]

    class Operations:
        def __init__(self, mailbox_instance):
            self.mailbox_instance = mailbox_instance; self.Move = mail_dispatch.Move; self.MessageState = mail_dispatch.MessageState; self.Delete = mail_dispatch.Delete; self.Draft = mail_compose.Draft; self.Send = mail_compose.Send; self.AttachmentDownload = mail_transfer.AttachmentDownload; self.AttachmentUpload = mail_transfer.AttachmentUpload
        def Mover(self): return self.Move(self.mailbox_instance)
        def MessageUpdater(self): return self.MessageState(self.mailbox_instance)
        def Deleter(self): return self.Delete(self.mailbox_instance)
        def Drafter(self): return self.Draft(self.mailbox_instance)
        def Sender(self): return self.Send(self.mailbox_instance)
        def AttachmentDownloader(self): return self.AttachmentDownload(self.mailbox_instance)
        def AttachmentUploader(self): return self.AttachmentUpload(self.mailbox_instance)
        def __dir__(self):
            return [
                "Mover", "MessageUpdater", "Deleter", "Drafter", "Sender",
                "AttachmentDownloader", "AttachmentUploader"
            ]
    def clear_cache(self):
        logger.info("Clearing mailbox cache"); self.cache = {}
    def __dir__(self):
        _instance_vars = ['mailbox', 'mailbox_profile', 'token_engine']; _properties = []; _methods = ['clear_cache', 'initiate_device_flow', 'authenticate_device_flow']; _inner_classes = ['search_by', 'operations', 'tools']; return sorted(_instance_vars + _properties + _methods + _inner_classes)
# Wrap zero, one or several discovered directory paths.
# The file accessor lists immediate child files only; invalid directories are
# skipped rather than recursively scanned.
class _dir_result:
    def __init__(self, value): self._value = value
    def _to_list(self):
        if self._value is None: return []
        if isinstance(self._value, list): return list(self._value)
        return [self._value]
    def _collapse(self, items):
        if not items: return None
        if len(items) == 1: return items[0]
        return items
    @property
    def value(self): return self._value
    @property
    def file(self):
        files = []
        for item in self._to_list():
            try:
                p = Path(item)
                if p.is_dir(): files.extend(str(child) for child in p.iterdir() if child.is_file())
            except (OSError, ValueError, TypeError):
                continue
        return self._collapse(sorted(files))
    def __contains__(self, item):
        if self._value is None: return False
        if isinstance(self._value, str): return item in self._value
        try: return item in self._value
        except TypeError: return item == self._value
    def __iter__(self):
        if self._value is None: return iter(())
        if isinstance(self._value, str): return iter((self._value,))
        try: return iter(self._value)
        except TypeError: return iter((self._value,))
    def __repr__(self): return repr(self._value)
    def __str__(self): return str(self._value)
    def __dir__(self): return ["value", "file"]    
# Locate canonical and suffixed authentication directories.
# Directory-name matching and username checks inside token.json are separate.
# Cached paths are pruned when removed; results collapse to None, one path or
# a list. Username matching requires a readable cache file.
class _authdir:
    TOKEN_CORE = (
        r'[a-z]{8}_[a-z]{7}_[a-z]{7}_'                            # word_word_word_ -> cerulean_grafiek_aansoek_
        r'(?:[a-z]{5})_'                                          # 5-letter segment -> teken_
        r'[A-Z0-9]{4}_'                                           # 4-char uppercase/alnum code -> JEQX_
        r'[0-9]{13}_'                                             # 13-digit timestamp
        r'[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}'   # UUID
    )

    ENV_CORE = (
        r'[a-z]{8}_[a-z]{7}_[a-z]{7}_'                            # word_word_word_ -> cerulean_grafiek_aansoek_
        r'(?:[a-z]{13})_'                                         # 13-letter segment -> geloofsbriewe_
        r'[A-Z0-9]{4}_'                                           # 4-char uppercase/alnum code -> JEQX_
        r'[0-9]{13}_'                                             # 13-digit timestamp
        r'[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}'   # UUID
    )
    EMAIL_MARKER = "__EMAIL__"
    TOKEN_PATTERN = re.compile(r'^' + TOKEN_CORE + r'$', re.ASCII)
    ENV_PATTERN = re.compile(r'^' + ENV_CORE + r'$', re.ASCII)
    TOKEN_PREFIX_PATTERN = re.compile(r'^(?P<canonical>' + TOKEN_CORE + r')', re.ASCII)
    ENV_PREFIX_PATTERN = re.compile(r'^(?P<canonical>' + ENV_CORE + r')', re.ASCII)
    _CACHE_UNSET = object()
    def __init__(self, base=None, run_on_init=True, scanned_token_username=None, scanned_env_suffix=None, use_cache=False):
        self.base = Path(base) if base else Path.home() / "AppData" / "Local"
        self._token_dir = None; self._env_dir = None
        self.scanned_token_dir = None; self.scanned_env_dir = None
        self._scanned_token = None; self._scanned_env = None
        self._scanned_token_ending = self._CACHE_UNSET; self._scanned_env_ending = self._CACHE_UNSET
        if run_on_init:
            self.find(token_username=scanned_token_username, use_cache=use_cache)
            self.find_with_remaining(token_username=scanned_token_username, env_suffix=scanned_env_suffix, use_cache=use_cache)
    @property
    def token_dir(self): return _dir_result(self._token_dir)
    @token_dir.setter
    def token_dir(self, value): self._token_dir = value
    @property
    def env_dir(self): return _dir_result(self._env_dir)
    @env_dir.setter
    def env_dir(self, value): self._env_dir = value
    @property
    def scanned_token(self):
        self._scanned_token = self._prune_missing_dirs(self._scanned_token)
        return _dir_result(self._scanned_token)
    @scanned_token.setter
    def scanned_token(self, value):
        if isinstance(value, list): items = value
        elif value is None: items = []
        else: items = [value]
        self._scanned_token = self._collapse(items)
    @property
    def scanned_env(self):
        self._scanned_env = self._prune_missing_dirs(self._scanned_env)
        return _dir_result(self._scanned_env)
    @scanned_env.setter
    def scanned_env(self, value):
        if isinstance(value, list): items = value
        elif value is None: items = []
        else: items = [value]
        self._scanned_env = self._collapse(items)
    def _collapse(self, value):
        def _is_empty_item(item):
            if item is None: return True
            if isinstance(item, str): return item.strip() == ""
            return False
        if not isinstance(value, list): return value
        if len(value) == 0: return None
        if all(_is_empty_item(item) for item in value): return None
        if len(value) == 1:
            item = value[0]
            return None if _is_empty_item(item) else item
        return value
    def _to_list(self, value):
        if value is None: return []
        if isinstance(value, list): return list(value)
        return [value]
    def _prune_missing_dirs(self, value):
        kept = []
        for path in self._to_list(value):
            try:
                if isinstance(path, str) and path.strip() and Path(path).is_dir(): kept.append(path)
            except (OSError, ValueError, TypeError):
                continue
        return self._collapse(kept)
    def _scan(self, pattern):
        hits = []
        if not self.base.exists() or not self.base.is_dir(): return hits
        for entry in self.base.iterdir():
            if entry.is_dir() and pattern.match(entry.name): hits.append(str(entry))
        return sorted(hits)
    def _scan_with_remaining(self, pattern, ending=None):
        hits = []
        if not self.base.exists() or not self.base.is_dir(): return hits
        for entry in self.base.iterdir():
            if not entry.is_dir(): continue
            match = pattern.match(entry.name)
            if not match: continue
            canonical = match.group("canonical"); remaining = entry.name[len(canonical):]
            if not remaining: continue
            if ending is not None and remaining != ending: continue
            hits.append(str(entry))
        return sorted(hits)
    def _build_email_suffix(self, email):
        if not isinstance(email, str): email = str(email)
        email = email.strip()
        if not email: raise ValueError("email cannot be empty")
        return "{0}{1}".format(self.EMAIL_MARKER, quote(email, safe=""))
    def _normalize_token_ending(self, username):
        if username is None: return None
        if not isinstance(username, str): username = str(username)
        username = username.strip()
        if not username: return None
        if username.startswith(self.EMAIL_MARKER):
            payload = username[len(self.EMAIL_MARKER):].strip()
            if not payload: return None
            return self.EMAIL_MARKER + payload
        if "@" in username: return self._build_email_suffix(username)
        return self.EMAIL_MARKER + username
    def _normalize_env_ending(self, suffix):
        if suffix is None: return None
        if not isinstance(suffix, str): suffix = str(suffix)
        suffix = suffix.strip()
        if not suffix: return None
        if not suffix.startswith("__"): suffix = "__" + suffix
        return suffix
    def _read_token_username_from_dir(self, dir_path, token_file_name="token.json"):
        token_path = Path(dir_path) / token_file_name
        if not token_path.exists() or not token_path.is_file(): return None
        try:
            with token_path.open("r", encoding="utf-8") as f: cache = json.load(f)
        except (OSError, ValueError, TypeError):
            return None
        account = next(iter(cache.get("Account", {}).values()), {})
        if not isinstance(account, dict): return None
        username = account.get("username") or account.get("preferred_username") or account.get("email")
        if not username: return None
        username = str(username).strip()
        return username or None
    def _filter_token_dirs_by_username(self, hits, username, token_file_name="token.json"):
        username = str(username or "").strip()
        if not username: return hits
        kept = []
        for item in hits:
            found = self._read_token_username_from_dir(item, token_file_name=token_file_name)
            if found and found.casefold() == username.casefold(): kept.append(item)
        return kept
    def find(self, token_username=None, use_cache=True, token_file_name="token.json"):
        if use_cache and self.scanned_token_dir is not None: raw_token_hits = self._to_list(self._prune_missing_dirs(self.scanned_token_dir))
        else: raw_token_hits = self._scan(self.TOKEN_PATTERN)
        self.scanned_token_dir = self._collapse(raw_token_hits)
        if token_username: token_hits = self._filter_token_dirs_by_username(raw_token_hits, token_username, token_file_name=token_file_name)
        else: token_hits = list(raw_token_hits)
        if use_cache and self.scanned_env_dir is not None: raw_env_hits = self._to_list(self._prune_missing_dirs(self.scanned_env_dir))
        else: raw_env_hits = self._scan(self.ENV_PATTERN)
        self.scanned_env_dir = self._collapse(raw_env_hits)
        self._token_dir = self._collapse(token_hits); self._env_dir = self._collapse(raw_env_hits)
        return {"token": self.token_dir, "env": self.env_dir}
    def find_with_remaining(self, token_username=None, env_suffix=None, use_cache=True):
        token_ending = self._normalize_token_ending(token_username); env_ending = self._normalize_env_ending(env_suffix)
        if use_cache and self._scanned_token_ending == token_ending: token_hits = self._to_list(self.scanned_token.value)
        else:
            token_hits = self._scan_with_remaining(self.TOKEN_PREFIX_PATTERN, ending=token_ending)
            self.scanned_token = token_hits; self._scanned_token_ending = token_ending
        if use_cache and self._scanned_env_ending == env_ending: env_hits = self._to_list(self.scanned_env.value)
        else:
            env_hits = self._scan_with_remaining(self.ENV_PREFIX_PATTERN, ending=env_ending)
            self.scanned_env = env_hits; self._scanned_env_ending = env_ending
        return {"token": self.scanned_token, "env": self.scanned_env}
    def refresh(self):
        self.scanned_token_dir = None; self.scanned_env_dir = None
        self._scanned_token = None; self._scanned_env = None
        self._scanned_token_ending = self._CACHE_UNSET; self._scanned_env_ending = self._CACHE_UNSET
        return self.find(use_cache=False)
    def clear_cache(self):
        self._token_dir = None; self._env_dir = None
        self.scanned_token_dir = None; self.scanned_env_dir = None
        self._scanned_token = None; self._scanned_env = None
        self._scanned_token_ending = self._CACHE_UNSET; self._scanned_env_ending = self._CACHE_UNSET
    def token_paths(self, strict=False, token_file_name="token.json", return_files=False):
        """Find token-shaped paths beneath the configured base directory."""
        hits = []
        if not self.base.exists() or not self.base.is_dir(): return _dir_result(None)
        for entry in self.base.iterdir():
            if not entry.is_dir(): continue
            match = self.TOKEN_PREFIX_PATTERN.match(entry.name)
            if not match: continue
            canonical = match.group("canonical"); remaining = entry.name[len(canonical):]
            if remaining and not remaining.startswith("__"): continue
            token_file = entry / token_file_name
            if strict and not (token_file.exists() and token_file.is_file()): continue
            hits.append(str(token_file if return_files else entry))
        return _dir_result(self._collapse(sorted(hits)))
# Validate authentication-directory copies and promotions.
# Checks include sibling relationships, cache identity and the unique current
# token/env-directory invariant. Filename and full-tree limits are platform
# specific; macOS checks account for NFD normalization. Validation does not itself
# copy, promote or lock the filesystem.
class _dirops_base(object):
    EMAIL_MARKER = "__EMAIL__"
    def __init__(self, overwrite_existing=False, windows_long_paths=False, base=None):
        self.overwrite_existing = overwrite_existing; self.windows_long_paths = bool(windows_long_paths); self.base = Path(base).expanduser().resolve() if base else (Path.home() / "AppData" / "Local").resolve()
    def _build_email_suffix(self, email):
        if not isinstance(email, str): email = str(email)
        email = email.strip()
        if not email: raise ValueError("email cannot be empty.")
        return "{0}{1}".format(self.EMAIL_MARKER, quote(email, safe=""))
    def _safe_suffix(self, value):
        value = str(value or "").strip()
        if not value: raise ValueError("suffix must be a non-empty string.")
        value = re.sub(r'[<>:"/\\|?*]+', "_", value); value = re.sub(r"\s+", "_", value); value = value.rstrip("._"); value = re.sub(r"^_+", "__", value)
        if not value.startswith("__"): value = "__" + value.lstrip("_")
        if value == "__": raise ValueError("suffix became empty after sanitization.")
        return value
    def _resolve_env_suffix(self, default_suffix=None, suffix=None):
        raw = default_suffix if suffix is None else suffix; return self._safe_suffix(raw)
    def _read_token_username(self, source_dir, token_file_name="token.json"):
        token_path = Path(source_dir).expanduser().resolve() / token_file_name
        if not token_path.exists(): raise FileNotFoundError("Token file not found: {0}".format(token_path))
        if not token_path.is_file(): raise FileNotFoundError("Token path is not a file: {0}".format(token_path))
        with token_path.open("r", encoding="utf-8") as f: cache = json.load(f)
        account = next(iter(cache.get("Account", {}).values()), {})
        if not isinstance(account, dict): account = {}
        username = (
            account.get("username")
            or account.get("preferred_username")
            or account.get("email")
        )
        if not username: raise RuntimeError("Could not determine username from token cache: {0}".format(token_path))
        username = str(username).strip()
        if not username: raise RuntimeError("Token cache username was empty: {0}".format(token_path))
        return username
    def _resolve_token_suffix(self, source_dir, username=None, token_file_name="token.json"):
        if username is None: username = self._read_token_username(source_dir, token_file_name=token_file_name)
        username = str(username or "").strip()
        if not username: raise ValueError("username must be a non-empty string.")
        return self._build_email_suffix(username)
    def _validate_source_dir(self, source_dir):
        if not source_dir.exists(): raise FileNotFoundError("Source directory not found: {0}".format(source_dir))
        if not source_dir.is_dir(): raise NotADirectoryError("Source path is not a directory: {0}".format(source_dir))
    def _validate_pair(self, previous_fork_dir, current_dir):
        if previous_fork_dir == current_dir: raise ValueError("previous_fork_dir and current_dir must be different directories.")
        if previous_fork_dir.parent != current_dir.parent:
            raise ValueError(
                "previous_fork_dir and current_dir must be sibling directories under the same parent. "
                "previous_fork_dir={0}, current_dir={1}".format(previous_fork_dir, current_dir)
            )
    def _validate_promote_pair(self, previous_fork_dir, current_dir, expected_previous_name):
        self._validate_pair(previous_fork_dir, current_dir)
        if previous_fork_dir.name != expected_previous_name:
            raise ValueError(
                "previous_fork_dir does not match the expected forked name. "
                "Expected: {0}, Got: {1}".format(expected_previous_name, previous_fork_dir.name)
            )
    def _validate_destination_tree(self, source_dir, final_backup_dir, followlinks=False):
        self._validate_name_limit(final_backup_dir.name, label="backup directory name"); self._validate_path_limit(final_backup_dir, label="backup directory path")
        rules = self._platform_rules(); source_root_len = self._measure_path(source_dir, rules); target_root_len = self._measure_path(final_backup_dir, rules); delta = target_root_len - source_root_len
        for root, dirnames, filenames in os.walk(str(source_dir), followlinks=followlinks):
            root_path = Path(root); dest_root_len = self._measure_path(root_path, rules) + delta
            if dest_root_len > rules["path_limit"]:
                raise OSError(
                    "destination directory path exceeds full path limit on {0}: {1} > {2}. Source: {3}".format(
                        rules["platform"], dest_root_len, rules["path_limit"], root_path
                    )
                )
            for name in dirnames:
                src_child = root_path / name; dest_child_len = self._measure_path(src_child, rules) + delta
                if dest_child_len > rules["path_limit"]:
                    raise OSError(
                        "destination directory path exceeds full path limit on {0}: {1} > {2}. Source: {3}".format(
                            rules["platform"], dest_child_len, rules["path_limit"], src_child
                        )
                    )
            for name in filenames:
                src_child = root_path / name; dest_child_len = self._measure_path(src_child, rules) + delta
                if dest_child_len > rules["path_limit"]:
                    raise OSError(
                        "destination file path exceeds full path limit on {0}: {1} > {2}. Source: {3}".format(
                            rules["platform"], dest_child_len, rules["path_limit"], src_child
                        )
                    )
    def _validate_target_root(self, target_dir, name_label="directory name", path_label="directory path"):
        self._validate_name_limit(target_dir.name, label=name_label); self._validate_path_limit(target_dir, label=path_label)
    def _make_temp_sibling_path(self, parent_dir, prefix):
        fd, temp_path = tempfile.mkstemp(prefix=prefix, dir=str(parent_dir)); os.close(fd); os.remove(temp_path); return Path(temp_path)
    def _platform_rules(self):
        if os.name == "nt":
            return {
                "platform": "Windows",
                "name_limit": 255,
                "path_limit": 32767 if self.windows_long_paths else 260,
                "name_measure": "chars",
                "path_measure": "chars",
                "normalize_nfd": False
            }
        if sys.platform == "darwin":
            return {
                "platform": "macOS",
                "name_limit": 255,
                "path_limit": 1024,
                "name_measure": "utf16_units",
                "path_measure": "bytes",
                "normalize_nfd": True
            }
        return {
            "platform": "Linux",
            "name_limit": 255,
            "path_limit": 4096,
            "name_measure": "bytes",
            "path_measure": "bytes",
            "normalize_nfd": False
        }
    def _normalize_for_measurement(self, value, normalize_nfd): return unicodedata.normalize("NFD", value) if normalize_nfd else value
    def _measure_name(self, name, rules):
        name = self._normalize_for_measurement(name, rules["normalize_nfd"])
        if rules["name_measure"] == "chars": return len(name)
        if rules["name_measure"] == "bytes": return len(os.fsencode(name))
        if rules["name_measure"] == "utf16_units": return len(name.encode("utf-16-le")) // 2
        raise ValueError("Unsupported name measurement mode.")
    def _measure_path(self, path_value, rules):
        path_text = self._normalize_for_measurement(str(path_value), rules["normalize_nfd"])
        if rules["path_measure"] == "chars": return len(path_text)
        if rules["path_measure"] == "bytes": return len(os.fsencode(path_text))
        raise ValueError("Unsupported path measurement mode.")
    def _validate_name_limit(self, name, label="name"):
        rules = self._platform_rules(); used = self._measure_name(name, rules); limit = rules["name_limit"]
        if used > limit:
            raise OSError(
                "{0} exceeds directory/file name limit on {1}: {2} > {3}. Value: {4!r}".format(
                    label, rules["platform"], used, limit, name
                )
            )
    def _validate_path_limit(self, path_value, label="path"):
        rules = self._platform_rules(); used = self._measure_path(path_value, rules); limit = rules["path_limit"]
        if used > limit:
            raise OSError(
                "{0} exceeds full path limit on {1}: {2} > {3}. Path: {4}".format(
                    label, rules["platform"], used, limit, path_value
                )
            )
    def _as_list(self, value):
        if value is None: return []
        if isinstance(value, list): return list(value)
        return [value]
    def _scan_current_dirs(self, base_dir=None, token_file_name="token.json"):
        root = Path(base_dir).expanduser().resolve() if base_dir else self.base; tracer = _authdir(base=root, run_on_init=False); found = tracer.find(use_cache=False, token_file_name=token_file_name)
        token_dirs = [Path(p).expanduser().resolve() for p in self._as_list(found["token"].value)]; env_dirs = [Path(p).expanduser().resolve() for p in self._as_list(found["env"].value)]
        return {"token": token_dirs, "env": env_dirs}
    def _require_global_current_invariant(self, base_dir=None, token_file_name="token.json"):
        root = Path(base_dir).expanduser().resolve() if base_dir else self.base; found = self._scan_current_dirs(base_dir=root, token_file_name=token_file_name)
        token_count = len(found["token"]); env_count = len(found["env"])
        if token_count != 1 or env_count != 1:
            raise RuntimeError(
                "Current-directory invariant violated under {0}. "
                "Expected exactly 1 current unsuffixed token dir and exactly 1 current unsuffixed env dir; "
                "found token={1}, env={2}. token_dirs={3}, env_dirs={4}".format(
                    root,
                    token_count,
                    env_count,
                    [str(p) for p in found["token"]],
                    [str(p) for p in found["env"]]
                )
            )
        return found
    def _resolve_single_current_dir(self, kind, base_dir=None, token_file_name="token.json"):
        root = Path(base_dir).expanduser().resolve() if base_dir else self.base; found = self._require_global_current_invariant(base_dir=root, token_file_name=token_file_name); return found[kind][0]
    def _validate_token_username_matches(self, dir_path, username, token_file_name="token.json"):
        if username is None: return
        actual = self._read_token_username(dir_path, token_file_name=token_file_name); expected = str(username).strip()
        if actual.casefold() != expected.casefold():
            raise ValueError(
                "Token directory username mismatch. Expected: {0}, Found: {1}, Directory: {2}".format(
                    expected, actual, dir_path
                )
            )
    def _require_current_path(self, path_value, kind, token_file_name="token.json", username=None):
        path_value = Path(path_value).expanduser().resolve(); self._validate_source_dir(path_value)
        current_dir = self._resolve_single_current_dir(
            kind=kind,
            base_dir=path_value.parent,
            token_file_name=token_file_name
        )
        if path_value != current_dir:
            raise ValueError(
                "Path is not the current unsuffixed {0} directory. Expected: {1}, Got: {2}".format(
                    kind, current_dir, path_value
                )
            )
        if kind == "token": self._validate_token_username_matches(path_value, username, token_file_name=token_file_name)
        return current_dir
    def _require_fork_path(self, previous_fork_dir, current_dir, expected_trailing):
        previous_fork_dir = Path(previous_fork_dir).expanduser().resolve(); current_dir = Path(current_dir).expanduser().resolve()
        self._validate_source_dir(previous_fork_dir); expected_name = "{0}{1}".format(current_dir.name, expected_trailing)
        self._validate_promote_pair(previous_fork_dir, current_dir, expected_name); return previous_fork_dir
    def __dir__(self): return ["token_dir","env_dir","scanned_token","scanned_env","find","find_with_remaining","token_paths","refresh","clear_cache"]
# Create a retained sibling copy while keeping the live directory in place.
# The copy is staged before final rename. Identity, naming and tree-length checks
# precede mutation; overwrite and symlink behavior follow explicit options.
class _fork(_dirops_base):
    """Copy the current env/token directory beside itself."""
    def __init__(self, suffix=None, overwrite_existing=True, preserve_symlinks=True, windows_long_paths=True, base=None):
        super().__init__(
            overwrite_existing=overwrite_existing,
            windows_long_paths=windows_long_paths,
            base=base
        )
        self.suffix = self._safe_suffix(suffix) if suffix is not None else None; self.preserve_symlinks = bool(preserve_symlinks)
    # backward-compatible env alias
    def copy(self, suffix=None, source_dir=None): return self.copy_env(suffix=suffix, source_dir=source_dir)
    # backward-compatible env alias
    def preview_target(self, suffix=None, source_dir=None): return self.preview_env_target(suffix=suffix, source_dir=source_dir)
    def _resolve_env_source(self, source_dir=None):
        if source_dir is None: return self._resolve_single_current_dir(kind="env")
        return self._require_current_path(source_dir, kind="env")
    def _resolve_token_source(self, source_dir=None, username=None, token_file_name="token.json"):
        if source_dir is None:
            source_dir = self._resolve_single_current_dir(
                kind="token",
                token_file_name=token_file_name
            )
            self._validate_token_username_matches(
                source_dir,
                username,
                token_file_name=token_file_name
            )
            return source_dir
        return self._require_current_path(
            source_dir,
            kind="token",
            token_file_name=token_file_name,
            username=username
        )
    def _preview_with_trailing(self, source_dir, trailing):
        source_dir = Path(source_dir).expanduser().resolve(); self._validate_source_dir(source_dir)
        target_dir = source_dir.parent / "{0}{1}".format(source_dir.name, trailing)
        self._validate_target_root(
            target_dir,
            name_label="backup directory name",
            path_label="backup directory path"
        )
        return str(target_dir)
    def _copy_with_trailing(self, source_dir, trailing):
        source_dir = Path(source_dir).expanduser().resolve(); self._validate_source_dir(source_dir)
        parent_dir = source_dir.parent; final_backup_dir = parent_dir / "{0}{1}".format(source_dir.name, trailing)
        self._validate_destination_tree(
            source_dir,
            final_backup_dir,
            followlinks=not self.preserve_symlinks
        )
        if final_backup_dir.exists():
            if not self.overwrite_existing: raise FileExistsError("Backup directory already exists: {0}".format(final_backup_dir))
            shutil.rmtree(str(final_backup_dir))
        temp_backup_dir = None
        try:
            temp_backup_dir = self._make_temp_sibling_path(parent_dir, prefix="._f_"); self._validate_path_limit(temp_backup_dir, label="temporary backup path")
            shutil.copytree(
                str(source_dir),
                str(temp_backup_dir),
                symlinks=self.preserve_symlinks,
                copy_function=shutil.copy2
            )
            os.replace(str(temp_backup_dir), str(final_backup_dir)); temp_backup_dir = None; return str(final_backup_dir)
        except Exception:
            if temp_backup_dir and temp_backup_dir.exists(): shutil.rmtree(str(temp_backup_dir), ignore_errors=True)
            raise
    def preview_env_target(self, suffix=None, source_dir=None):
        source_dir = self._resolve_env_source(source_dir=source_dir); trailing = self._resolve_env_suffix(default_suffix=self.suffix, suffix=suffix); return self._preview_with_trailing(source_dir, trailing)
    def preview_token_target(self, username=None, token_file_name="token.json", source_dir=None):
        source_dir = self._resolve_token_source(
            source_dir=source_dir,
            username=username,
            token_file_name=token_file_name
        )
        trailing = self._resolve_token_suffix(
            source_dir,
            username=username,
            token_file_name=token_file_name
        )
        return self._preview_with_trailing(source_dir, trailing)
    def copy_env(self, suffix=None, source_dir=None):
        source_dir = self._resolve_env_source(source_dir=source_dir)
        self._require_global_current_invariant(base_dir=source_dir.parent)
        trailing = self._resolve_env_suffix(default_suffix=self.suffix, suffix=suffix); result = self._copy_with_trailing(source_dir, trailing)
        self._require_global_current_invariant(base_dir=source_dir.parent); return result
    def copy_token(self, username=None, token_file_name="token.json", source_dir=None):
        source_dir = self._resolve_token_source(
            source_dir=source_dir,
            username=username,
            token_file_name=token_file_name
        )
        self._require_global_current_invariant(
            base_dir=source_dir.parent,
            token_file_name=token_file_name
        )
        trailing = self._resolve_token_suffix(
            source_dir,
            username=username,
            token_file_name=token_file_name
        )
        result = self._copy_with_trailing(source_dir, trailing)
        self._require_global_current_invariant(
            base_dir=source_dir.parent,
            token_file_name=token_file_name
        )
        return result
# Swap a validated retained sibling into the current directory name.
# The live directory becomes a suffixed backup. Identity/layout checks precede
# renames, and failures attempt to restore the prior layout. Preview methods
# validate paths without performing the swap; overwrite behavior is explicit.
class _promote(_dirops_base):
    """Promote a forked env/token directory back to the current unsuffixed name."""
    def __init__(self, suffix=None, overwrite_existing=False, windows_long_paths=False, base=None):
        super().__init__(
            overwrite_existing=overwrite_existing,
            windows_long_paths=windows_long_paths,
            base=base
        )
        self.suffix = self._safe_suffix(suffix) if suffix is not None else None
    # backward-compatible env aliases
    def roll_back(self, previous_fork_dir, suffix=None, current_dir=None):
        return self.roll_back_env(
            previous_fork_dir=previous_fork_dir,
            suffix=suffix,
            current_dir=current_dir
        )
    def copy(self, previous_fork_dir, suffix=None, current_dir=None):
        return self.roll_back_env(
            previous_fork_dir=previous_fork_dir,
            suffix=suffix,
            current_dir=current_dir
        )
    def preview_live_target(self, previous_fork_dir, suffix=None, current_dir=None):
        return self.preview_env_live_target(
            previous_fork_dir=previous_fork_dir,
            suffix=suffix,
            current_dir=current_dir
        )
    def preview_backup_target(self, suffix=None, current_dir=None):
        return self.preview_env_backup_target(
            suffix=suffix,
            current_dir=current_dir
        )
    def _resolve_env_current_dir(self, previous_fork_dir=None, current_dir=None):
        if current_dir is not None: return self._require_current_path(current_dir, kind="env")
        if previous_fork_dir is not None:
            previous_fork_dir = Path(previous_fork_dir).expanduser().resolve(); return self._resolve_single_current_dir(kind="env", base_dir=previous_fork_dir.parent)
        return self._resolve_single_current_dir(kind="env")
    def _resolve_token_current_dir(self, previous_fork_dir=None, current_dir=None, current_username=None, token_file_name="token.json"):
        if current_dir is not None:
            return self._require_current_path(
                current_dir,
                kind="token",
                token_file_name=token_file_name,
                username=current_username
            )
        if previous_fork_dir is not None:
            previous_fork_dir = Path(previous_fork_dir).expanduser().resolve()
            current_dir = self._resolve_single_current_dir(
                kind="token",
                base_dir=previous_fork_dir.parent,
                token_file_name=token_file_name
            )
        else:
            current_dir = self._resolve_single_current_dir(
                kind="token",
                token_file_name=token_file_name
            )
        self._validate_token_username_matches(
            current_dir,
            current_username,
            token_file_name=token_file_name
        )
        return current_dir
    def _resolve_previous_token_username(self, previous_fork_dir, previous_username=None, token_file_name="token.json"):
        if previous_username is not None:
            previous_username = str(previous_username).strip()
            if not previous_username: raise ValueError("previous_username must be a non-empty string.")
            return previous_username
        return self._read_token_username(previous_fork_dir, token_file_name=token_file_name)
    def _resolve_current_token_username(self, current_dir, current_username=None, token_file_name="token.json"):
        if current_username is not None:
            current_username = str(current_username).strip()
            if not current_username: raise ValueError("current_username must be a non-empty string.")
            self._validate_token_username_matches(
                current_dir,
                current_username,
                token_file_name=token_file_name
            )
            return current_username
        return self._read_token_username(current_dir, token_file_name=token_file_name)
    def _promote_with_backup(self, previous_fork_dir, current_dir, backup_current_dir):
        temp_promote_dir = None; promoted_staged = False; current_moved = False
        try:
            if backup_current_dir.exists() and backup_current_dir != previous_fork_dir:
                if not self.overwrite_existing: raise FileExistsError("Backup directory already exists: {0}".format(backup_current_dir))
                shutil.rmtree(str(backup_current_dir))
            temp_promote_dir = self._make_temp_sibling_path(current_dir.parent, prefix="._p_"); self._validate_path_limit(temp_promote_dir, label="temporary promoted path")
            os.replace(str(previous_fork_dir), str(temp_promote_dir)); promoted_staged = True
            os.replace(str(current_dir), str(backup_current_dir)); current_moved = True
            os.replace(str(temp_promote_dir), str(current_dir)); temp_promote_dir = None
            return str(current_dir)
        except Exception:
            if current_moved and backup_current_dir.exists() and not current_dir.exists():
                try:
                    os.replace(str(backup_current_dir), str(current_dir)); current_moved = False
                except Exception:
                    pass
            if promoted_staged and temp_promote_dir and temp_promote_dir.exists() and not previous_fork_dir.exists():
                try:
                    os.replace(str(temp_promote_dir), str(previous_fork_dir)); temp_promote_dir = None
                except Exception:
                    pass
            raise
    def preview_env_live_target(self, previous_fork_dir, suffix=None, current_dir=None):
        previous_fork_dir = Path(previous_fork_dir).expanduser().resolve()
        current_dir = self._resolve_env_current_dir(
            previous_fork_dir=previous_fork_dir,
            current_dir=current_dir
        )
        trailing = self._resolve_env_suffix(default_suffix=self.suffix, suffix=suffix); self._require_fork_path(previous_fork_dir, current_dir, trailing); self._validate_target_root(current_dir)
        return str(current_dir)
    def preview_env_backup_target(self, suffix=None, current_dir=None):
        current_dir = self._resolve_env_current_dir(current_dir=current_dir)
        trailing = self._resolve_env_suffix(default_suffix=self.suffix, suffix=suffix); backup_current_dir = current_dir.parent / "{0}{1}".format(current_dir.name, trailing)
        self._validate_target_root(
            backup_current_dir,
            name_label="backup directory name",
            path_label="backup directory path"
        )
        return str(backup_current_dir)
    def preview_token_live_target(self, previous_fork_dir, previous_username=None, token_file_name="token.json", current_dir=None, current_username=None):
        previous_fork_dir = Path(previous_fork_dir).expanduser().resolve()
        current_dir = self._resolve_token_current_dir(
            previous_fork_dir=previous_fork_dir,
            current_dir=current_dir,
            current_username=current_username,
            token_file_name=token_file_name
        )
        previous_username = self._resolve_previous_token_username(
            previous_fork_dir,
            previous_username=previous_username,
            token_file_name=token_file_name
        )
        previous_trailing = self._build_email_suffix(previous_username); self._require_fork_path(previous_fork_dir, current_dir, previous_trailing); self._validate_target_root(current_dir)
        return str(current_dir)
    def preview_token_backup_target(self, current_username=None, token_file_name="token.json", current_dir=None):
        current_dir = self._resolve_token_current_dir(
            current_dir=current_dir,
            current_username=current_username,
            token_file_name=token_file_name
        )
        current_username = self._resolve_current_token_username(
            current_dir,
            current_username=current_username,
            token_file_name=token_file_name
        )
        current_trailing = self._build_email_suffix(current_username); backup_current_dir = current_dir.parent / "{0}{1}".format(current_dir.name, current_trailing)
        self._validate_target_root(
            backup_current_dir,
            name_label="backup directory name",
            path_label="backup directory path"
        )
        return str(backup_current_dir)
    def roll_back_env(self, previous_fork_dir, suffix=None, current_dir=None):
        previous_fork_dir = Path(previous_fork_dir).expanduser().resolve()
        current_dir = self._resolve_env_current_dir(
            previous_fork_dir=previous_fork_dir,
            current_dir=current_dir
        )
        self._require_global_current_invariant(base_dir=current_dir.parent)
        trailing = self._resolve_env_suffix(default_suffix=self.suffix, suffix=suffix); self._require_fork_path(previous_fork_dir, current_dir, trailing)
        backup_current_dir = current_dir.parent / "{0}{1}".format(current_dir.name, trailing)
        self._validate_target_root(
            backup_current_dir,
            name_label="backup directory name",
            path_label="backup directory path"
        )
        self._validate_destination_tree(current_dir, backup_current_dir, followlinks=False)
        result = self._promote_with_backup(previous_fork_dir, current_dir, backup_current_dir)
        self._require_global_current_invariant(base_dir=current_dir.parent); return result
    def roll_back_token(self, previous_fork_dir, previous_username=None, current_username=None, token_file_name="token.json", current_dir=None):
        previous_fork_dir = Path(previous_fork_dir).expanduser().resolve()
        current_dir = self._resolve_token_current_dir(
            previous_fork_dir=previous_fork_dir,
            current_dir=current_dir,
            current_username=current_username,
            token_file_name=token_file_name
        )
        self._require_global_current_invariant(
            base_dir=current_dir.parent,
            token_file_name=token_file_name
        )
        previous_username = self._resolve_previous_token_username(
            previous_fork_dir,
            previous_username=previous_username,
            token_file_name=token_file_name
        )
        current_username = self._resolve_current_token_username(
            current_dir,
            current_username=current_username,
            token_file_name=token_file_name
        )
        previous_trailing = self._build_email_suffix(previous_username); current_trailing = self._build_email_suffix(current_username)
        self._require_fork_path(previous_fork_dir, current_dir, previous_trailing)
        backup_current_dir = current_dir.parent / "{0}{1}".format(current_dir.name, current_trailing)
        self._validate_target_root(
            backup_current_dir,
            name_label="backup directory name",
            path_label="backup directory path"
        )
        self._validate_destination_tree(current_dir, backup_current_dir, followlinks=False)
        result = self._promote_with_backup(previous_fork_dir, current_dir, backup_current_dir)
        self._require_global_current_invariant(
            base_dir=current_dir.parent,
            token_file_name=token_file_name
        )
        return result
# Read usernames from bounded, validated cache JSON.
# Search the Account section first, then other username fields. Invalid or
# unreadable files produce no match. Directory lookup deduplicates candidates;
# strict mode requires the expected cache filename.
class JsonUsernameExtractor:
    """
    Safely load a JSON file and extract a username using this priority:

    1. Find an "Account" section anywhere in the JSON and search it for "username".
    2. If not found, search the entire JSON for "username".
    3. If still not found, return None.
    """
    def __init__(self, file_path, *, case_sensitive=False, max_bytes=10 * 1024 * 1024, logger=None, authdir=None):
        self.file_path = Path(file_path); self.case_sensitive = case_sensitive; self.max_bytes = max_bytes; self.logger = logger or logging.getLogger(__name__); self.authdir = authdir
    def get_username(self):
        """Return username from JSON or None."""; data = self._load_json()
        if data is None: return None
        account_section = self._find_section(data, "Account")
        if account_section is not None:
            username = self._find_value(account_section, "username")
            if self._is_valid_username(username): return username
        username = self._find_value(data, "username")
        if self._is_valid_username(username): return username
        return None
    def _load_json(self):
        """Return parsed cache JSON, or None if it cannot be read or validated."""
        try:
            if not self.file_path.exists():
                self.logger.warning("File does not exist: %s", self.file_path); return None
            if not self.file_path.is_file():
                self.logger.warning("Path is not a file: %s", self.file_path); return None
            file_size = self.file_path.stat().st_size
            if file_size > self.max_bytes:
                self.logger.warning(
                    "File too large (%d bytes > %d bytes): %s",
                    file_size, self.max_bytes, self.file_path
                )
                return None
            with self.file_path.open("r", encoding="utf-8") as f: return json.load(f)
        except OSError as e:
            self.logger.warning("Could not read file %s: %s", self.file_path, e)
        except json.JSONDecodeError as e:
            self.logger.warning("Invalid JSON in %s: %s", self.file_path, e)
        except Exception as e:
            self.logger.exception("Unexpected error while loading %s: %s", self.file_path, e)
        return None
    def _normalize_key(self, key):
        if not isinstance(key, str): return key
        return key if self.case_sensitive else key.lower()
    def _keys_match(self, actual, expected):
        if not isinstance(actual, str): return False
        if self.case_sensitive: return actual == expected
        return actual.lower() == expected.lower()
    def _find_section(self, obj, section_name):
        """Recursively find first matching section."""
        try:
            if isinstance(obj, dict):
                for key, value in obj.items():
                    if self._keys_match(key, section_name): return value
                    nested = self._find_section(value, section_name)
                    if nested is not None: return nested
            elif isinstance(obj, list):
                for item in obj:
                    nested = self._find_section(item, section_name)
                    if nested is not None: return nested
        except Exception as e:
            self.logger.debug("Error while searching for section %r: %s", section_name, e)
        return None
    def _find_value(self, obj, target_key):
        """Recursively find first value for key."""
        try:
            if isinstance(obj, dict):
                for key, value in obj.items():
                    if self._keys_match(key, target_key): return value
                    nested = self._find_value(value, target_key)
                    if nested is not None: return nested
            elif isinstance(obj, list):
                for item in obj:
                    nested = self._find_value(item, target_key)
                    if nested is not None: return nested
        except Exception as e:
            self.logger.debug("Error while searching for key %r: %s", target_key, e)
        return None
    @staticmethod
    def _is_valid_username(value): return isinstance(value, str) and bool(value.strip())
    @staticmethod
    def _collapse_results(items):
        if not items: return None
        if len(items) == 1: return items[0]
        return items
    @staticmethod
    def _as_list(value):
        if value is None: return []
        if isinstance(value, list): return value
        return [value]
    @staticmethod
    def _usernames_match(found, expected, *, case_sensitive=False):
        if not isinstance(found, str) or not isinstance(expected, str): return False
        found = found.strip(); expected = expected.strip()
        if not found or not expected: return False
        if case_sensitive: return found == expected
        return found.casefold() == expected.casefold()
    @classmethod
    def find_directory_for_username(cls, username, authdir=None, *, token_username=None, case_sensitive=False, max_bytes=10 * 1024 * 1024, logger=None, use_cache=True, return_all=False, strict=True, token_file_name="token.json"):
        """
        Find token-shaped directories from _authdir and return those whose token JSON
        contains the requested username.
        """
        logger = logger or logging.getLogger(__name__)
        if not cls._is_valid_username(username): return None
        auth = authdir or _authdir(run_on_init=False); token_lookup = token_username if cls._is_valid_username(token_username) else username
        candidate_dirs = []; seen = set()
        def add_candidate_dirs(value):
            for item in cls._as_list(value):
                try:
                    p = Path(item)
                except (TypeError, ValueError):
                    continue
                key = str(p)
                if key in seen: continue
                try:
                    if p.is_dir():
                        candidate_dirs.append(p); seen.add(key)
                except (OSError, ValueError):
                    continue
        if hasattr(auth, "token_paths"):
            try:
                add_candidate_dirs(
                    auth.token_paths(
                        strict=False,
                        token_file_name=token_file_name,
                        return_files=False,
                    ).value
                )
            except Exception as e:
                logger.debug("auth.token_paths() failed: %s", e)
        try:
            auth.find(use_cache=use_cache, token_file_name=token_file_name); add_candidate_dirs(auth.token_dir.value)
        except TypeError:
            auth.find(use_cache=use_cache); add_candidate_dirs(auth.token_dir.value)
        except Exception as e:
            logger.debug("auth.find() failed: %s", e)
        try:
            auth.find_with_remaining(token_username=token_lookup, use_cache=use_cache); add_candidate_dirs(auth.scanned_token.value)
        except Exception as e:
            logger.debug("auth.find_with_remaining() failed: %s", e)
        matches = []
        for directory in candidate_dirs:
            files_to_check = []; token_path = directory / token_file_name
            try:
                if token_path.is_file(): files_to_check = [token_path]
                elif strict: continue
                else:
                    files_to_check = [
                        child for child in directory.iterdir()
                        if child.is_file() and child.suffix.lower() == ".json"
                    ]
            except (OSError, ValueError) as e:
                logger.debug("Could not inspect directory %s: %s", directory, e); continue
            for file_path in files_to_check:
                found_username = cls(
                    file_path,
                    case_sensitive=case_sensitive,
                    max_bytes=max_bytes,
                    logger=logger,
                    authdir=auth,
                ).get_username()
                if cls._usernames_match(found_username, username, case_sensitive=case_sensitive):
                    matches.append(str(directory)); break
        if not matches: return None
        matches = sorted(dict.fromkeys(matches))
        if return_all: return cls._collapse_results(matches)
        return matches[0]
# Select caches by mailbox identity, preferring the live cache or a validated
# retained fork. Switching validates the current layout and identity before
# promotion. Local-part matching is a configurable compatibility mode, not proof
# that accounts in different tenants represent the same identity.
class _authenticate_to_cache:
    """Select or acquire an authentication cache for a requested mailbox identity.

    The workflow coordinates cache discovery, identity checks, optional fork
    promotion and provider-supported authentication. Identity matching follows
    the configured identity_match_mode.
    """
    def __init__(
        self,
        *,
        username,
        mailbox="me",
        scopes=None,
        final_cache_dir=None,
        webbrowser_enabled=True,
        overwrite_name="token.json",
        save_as_name="token.json",
        identity_match_mode="strict",   # or "local_part"
        mailbox_cls=Mail,
        authdir_factory=_authdir,
        fork_factory=_fork,
        promote_factory=_promote,
    ):
        self.username = username; self.mailbox = mailbox; self.scopes = scopes or DEFAULT_SCOPES; self.final_cache_dir = final_cache_dir or authdir_factory().token_dir.value; self.webbrowser_enabled = webbrowser_enabled; self.overwrite_name = overwrite_name; self.save_as_name = save_as_name; self.identity_match_mode = identity_match_mode
        self.mailbox_cls = mailbox_cls; self.authdir_factory = authdir_factory; self.fork_factory = fork_factory; self.promote_factory = promote_factory
    @staticmethod
    def _normalize_email(value):
        if value is None: return None
        return str(value).strip().lower()
    @classmethod
    def _email_local_part(cls, value):
        value = cls._normalize_email(value)
        if not value: return None
        return value.split("@", 1)[0]
    @classmethod
    def _emails_match(cls, expected, actual, mode="strict"):
        expected_norm = cls._normalize_email(expected); actual_norm = cls._normalize_email(actual)
        if not expected_norm or not actual_norm: return False
        mode = (mode or "strict").strip().lower()
        if mode == "strict": return expected_norm == actual_norm
        if mode in {"local_part", "local", "username"}: return cls._email_local_part(expected_norm) == cls._email_local_part(actual_norm)
        raise ValueError("identity_match_mode must be one of: 'strict', 'local_part'")
    @staticmethod
    def _normalize_json_filename(name, fallback):
        name = (name or fallback).strip()
        if not name: name = fallback
        if not name.lower().endswith(".json"): name += ".json"
        return name
    @staticmethod
    def _read_cache_identity(token_path):
        token_path = Path(token_path)
        with token_path.open("r", encoding="utf-8") as f: cache = json.load(f)
        account = next(iter(cache.get("Account", {}).values()), {})
        if not account: raise RuntimeError(f"No Account record found in token cache: {token_path}")
        return {
            "username": account.get("username"),
            "home_account_id": account.get("home_account_id"),
            "realm": account.get("realm"),
            "environment": account.get("environment"),
        }
    @classmethod
    def _read_dir_identity(cls, token_dir, token_file_name):
        token_dir = Path(token_dir); token_path = token_dir / token_file_name
        if not token_path.is_file(): return None
        try:
            return cls._read_cache_identity(token_path)
        except Exception:
            return None
    @staticmethod
    def _install_cache_atomically(src, dst):
        dst = Path(dst); dst.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            mode="wb",
            dir=str(dst.parent),
            prefix=f".{dst.name}.",
            suffix=".tmp",
            delete=False,
        ) as f:
            staging_path = Path(f.name)
        try:
            shutil.copyfile(src, staging_path); os.replace(str(staging_path), str(dst))
        except Exception:
            try:
                staging_path.unlink()
            except FileNotFoundError:
                pass
            raise
    def _resolve_live_and_candidates(self, base_root, token_file_name):
        tracer = self.authdir_factory(base=base_root, run_on_init=False)
        current = tracer.find(use_cache=False, token_file_name=token_file_name); live_value = current["token"].value
        if not live_value: raise RuntimeError(f"No live unsuffixed token directory found under {base_root!s}")
        if isinstance(live_value, list):
            if len(live_value) != 1:
                raise RuntimeError(
                    f"Expected exactly one live unsuffixed token directory under {base_root!s}, got {live_value!r}"
                )
            live_dir = Path(live_value[0]).expanduser().resolve()
        else: live_dir = Path(live_value).expanduser().resolve()
        candidates_value = tracer.token_paths(strict=True, token_file_name=token_file_name).value
        if not candidates_value: candidates = []
        elif isinstance(candidates_value, list): candidates = [Path(p).expanduser().resolve() for p in candidates_value]
        else: candidates = [Path(candidates_value).expanduser().resolve()]
        candidates = sorted(candidates, key=lambda p: (p != live_dir, str(p).lower())); return live_dir, candidates
    def _find_matching_token_dir(self, username, base_root, token_file_name):
        live_dir, candidates = self._resolve_live_and_candidates(base_root, token_file_name)
        for token_dir in candidates:
            identity = self._read_dir_identity(token_dir, token_file_name)
            if not identity: continue
            actual_username = identity.get("username")
            if self._emails_match(
                username,
                actual_username,
                mode=self.identity_match_mode,
            ):
                return {
                    "matched_dir": token_dir,
                    "matched_identity": identity,
                    "live_dir": live_dir,
                }
        return {
            "matched_dir": None,
            "matched_identity": None,
            "live_dir": live_dir,
        }
    def _build_mailbox(self, *, token_path, build_kwargs=None):
        return self.mailbox_cls.from_default_env(
            mailbox=self.mailbox,
            scopes=self.scopes,
            token_path=str(token_path),
            build_kwargs=build_kwargs,
        )
    @staticmethod
    def _mailbox_username(mb, fallback=None):
        return (
            mb.mailbox_profile.get("mail")
            or mb.mailbox_profile.get("userPrincipalName")
            or fallback
        )
    def authenticate(self):
        username = self.username; username_norm = self._normalize_email(username)
        if not username_norm: raise ValueError("username is required.")
        token_file_name = self._normalize_json_filename(self.overwrite_name, "token.json")
        live_dir_hint = Path(self.final_cache_dir).expanduser().resolve(); base_root = live_dir_hint.parent
        scan = self._find_matching_token_dir(
            username=username,
            base_root=base_root,
            token_file_name=token_file_name,
        )
        live_token_dir = scan["live_dir"]; matched_dir = scan["matched_dir"]; matched_identity = scan["matched_identity"]
        # Reuse or promote a cache for the requested mailbox.
        if matched_dir is not None:
            if matched_dir != live_token_dir:
                current_live_identity = self._read_dir_identity(live_token_dir, token_file_name); current_live_username = (current_live_identity or {}).get("username")
                promoter = self.promote_factory(base=base_root, overwrite_existing=True)
                promoter.roll_back_token(
                    previous_fork_dir=str(matched_dir),
                    previous_username=matched_identity.get("username"),
                    current_username=current_live_username,
                    token_file_name=token_file_name,
                    current_dir=str(live_token_dir),
                )
            final_token_path = live_token_dir / token_file_name; mb = self._build_mailbox(token_path=final_token_path)
            actual_username = self._mailbox_username(
                mb,
                fallback=(matched_identity or {}).get("username"),
            )
            if not self._emails_match(
                username,
                actual_username,
                mode=self.identity_match_mode,
            ):
                raise RuntimeError(
                    "Found an existing token cache and loaded it, but the authenticated user did not match. "
                    f"Expected {username!r}, got {actual_username!r}. "
                    f"Match mode: {self.identity_match_mode!r}."
                )
            identity = {
                "username": actual_username,
                "mailbox_profile": mb.mailbox_profile,
                "cache_source": "existing_fork_promoted" if matched_dir != live_token_dir else "existing_live",
                "cache_dir": str(live_token_dir),
            }
            return mb, str(final_token_path), identity
        # Run device-code authentication when no matching cache exists.
        with tempfile.TemporaryDirectory(prefix="mailbox-auth-") as tmpdir:
            temp_token_path = Path(tmpdir) / token_file_name
            temp_mb = self._build_mailbox(
                token_path=temp_token_path,
                build_kwargs={
                    "auth_flow": "device_code",
                    "webbrowser_enabled": self.webbrowser_enabled,
                },
            )
            actual_username = self._mailbox_username(temp_mb)
            if not self._emails_match(
                username,
                actual_username,
                mode=self.identity_match_mode,
            ):
                raise RuntimeError(
                    "Authenticated user mismatch. "
                    f"Expected {username!r}, got {actual_username!r}. "
                    f"Match mode: {self.identity_match_mode!r}."
                )
            if not temp_token_path.is_file() or temp_token_path.stat().st_size == 0:
                raise RuntimeError(
                    "Authentication succeeded and the expected user was verified, "
                    "but azauthlib did not write a token cache file for this device-flow session."
                )
            current_live_identity = self._read_dir_identity(live_token_dir, token_file_name); current_live_username = (current_live_identity or {}).get("username")
            if current_live_username and not self._emails_match(
                current_live_username,
                actual_username,
                mode="strict",
            ):
                forker = self.fork_factory(base=base_root, overwrite_existing=True)
                forker.copy_token(
                    username=current_live_username,
                    token_file_name=token_file_name,
                    source_dir=str(live_token_dir),
                )
            final_token_path = live_token_dir / token_file_name; self._install_cache_atomically(temp_token_path, final_token_path)
        mb = self._build_mailbox(token_path=final_token_path)
        identity = {
            "username": actual_username,
            "mailbox_profile": mb.mailbox_profile,
            "cache_source": "new_device_flow",
            "cache_dir": str(live_token_dir),
        }
        return mb, str(final_token_path), identity



__all__ = [
    "_dir_result",
    "_authdir",
    "_dirops_base",
    "_fork",
    "_promote",
    "JsonUsernameExtractor",
    "_authenticate_to_cache",
]
def __dir__():
    mod = sys.modules[__name__]
    dunder_names = [
        name for name in mod.__dict__
        if name.startswith('__') and name.endswith('__')
    ]
    return sorted(set(dunder_names + ['Mail'] + __all__))
