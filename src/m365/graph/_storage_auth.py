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
"""Authentication adapters shared by storage clients; no tenant provisioning."""
import os
_DEFAULT_GRAPH = "https://graph.microsoft.com/v1.0"


def _profile_provider(profile, *, scopes=None, allow_interactive=False, graph_url=_DEFAULT_GRAPH):
    if not isinstance(allow_interactive, bool):
        raise ValueError("allow_interactive must be a boolean")
    from azauthlib import Authentication, InteractionRequiredError, Profile
    from azauthlib.credentials import load_public_configuration
    if not isinstance(profile, (str, Profile)) or isinstance(profile, str) and not profile.strip():
        raise ValueError("Supply a saved profile name, ID or Profile")
    if any(name in os.environ for name in ("CLIENT_ID", "TENANT_ID", "CLIENT_SECRET", "AZURE_AUTHORITY")):
        raise ValueError("Environment credentials would override the saved profile; use an explicit provider instead")
    if scopes is None:
        if graph_url != _DEFAULT_GRAPH:
            raise ValueError("Supply scopes matching the custom Graph endpoint")
        scopes = ["https://graph.microsoft.com/.default"]
    scope_values = [scopes] if isinstance(scopes, str) else scopes
    if (not isinstance(scope_values, (list, tuple)) or not scope_values
            or any(not isinstance(s, str) or not s.strip() or any(ord(c) < 32 for c in s) for s in scope_values)):
        raise ValueError("scopes must explicitly identify a nonempty authentication scope set")
    method = load_public_configuration(profile=profile)["auth_method"]
    if method not in ("client_certificate", "client_credentials", "interactive", "device_code"):
        raise ValueError("The saved profile uses an unsupported authentication method")
    auth = Authentication(profile=profile, scopes=scopes)
    auth.Build.Default()
    if method in ("client_certificate", "client_credentials"):
        auth.ClientCredentials()
    else:
        try:
            auth.Silent()
        except InteractionRequiredError:
            if not allow_interactive:
                raise
            if method == "device_code":
                auth.DeviceCodeFlow()
            else:
                auth.Interactive()
    return auth


def _legacy_provider(factory, *, scopes=None, token_path=None, authority=None,
                     build=None, build_kwargs=None):
    """Preserve the existing Site builder/silent-or-device-flow contract."""
    build = build or "Default"
    auth = factory(scopes=scopes, token_path=token_path)
    bkwargs = dict(build_kwargs or {})
    auth_flow = bkwargs.pop("auth_flow", "interactive")
    auth_flow = str(auth_flow).lower() if auth_flow is not None else "interactive"
    builder = getattr(auth.Build, build, None)
    if not callable(builder):
        raise ValueError(
            f"Unknown build mode '{build}'. "
            "Use one of: Default, WithEnvFile, WithOSEnv, WithEntry."
        )
    if build == "Default":
        builder(authority=authority)
    else:
        builder(authority=authority, **bkwargs)
    if auth_flow in {"device", "device_code", "devicecode", "device-flow"}:
        auth.DeviceCodeFlow(scopes=scopes)
    else:
        auth.Silent(scopes=scopes)
    if not auth.access_token:
        raise RuntimeError("Authentication failed: no access token acquired.")
    return auth
