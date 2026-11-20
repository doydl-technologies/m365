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
"""Opaque site/drive binding and library-relative path helpers."""
from urllib.parse import quote, unquote, urlsplit, parse_qsl
from uuid import UUID


def graph_id(value):
    if not isinstance(value, str) or not value.strip():
        raise ValueError("Graph IDs must be nonempty strings")
    return quote(value, safe=",!-")


def drive_path(path):
    """Encode each path component once; input is an UNENCODED library-relative path."""
    if not isinstance(path, str):
        raise TypeError("Drive paths must be strings")
    if "\\" in path or ":" in path or "\x00" in path:
        raise ValueError("Use a library-relative path with forward slashes")
    parts = path.strip("/").split("/") if path.strip("/") else []
    if any(p in (".", "..") or not p for p in parts):
        raise ValueError("Empty or traversal path components are not allowed")
    return "/".join(quote(p, safe="") for p in parts)


def resolve_site(client):
    if client.site_id:
        resource = "sites/" + graph_id(client.site_id)
    else:
        value = client.site_name
        if not isinstance(value, str) or not value.strip():
            raise ValueError("Supply a site URL, hostname/path, short name, or site_id")
        if "," in value and "/" not in value:
            resource = "sites/" + graph_id(value)
        else:
            host = client.hostname
            if "://" in value:
                p = urlsplit(value)
                if p.scheme != "https" or not p.hostname or p.username or p.password or p.query or p.fragment:
                    raise ValueError("Supply a clean HTTPS SharePoint site URL")
                if host and host.lower() != p.netloc.lower():
                    raise ValueError("Site URL and hostname disagree")
                host, path = p.netloc, unquote(p.path)
            elif "/" in value and not value.startswith("/") and "." in value.split("/", 1)[0]:
                host, path = value.split("/", 1)
                path = "/" + path
            else:
                path = value if value.startswith("/") else "/sites/" + value
            if not host:
                # Backwards-compatible short names: discover this provider's tenant root.
                root = client.api.graph_call(resource_url="sites/root")
                host = urlsplit(root["webUrl"]).netloc
            if any(c in host for c in "/:@?#\\") or not host:
                raise ValueError("Invalid SharePoint hostname")
            resource = "sites/{}:/{}".format(host, drive_path(path))
    site = client.api.graph_call(resource_url=resource)
    if not isinstance(site, dict) or not site.get("id") or not site.get("webUrl"):
        raise ValueError("Graph returned incomplete site metadata")
    site_id = site["id"]  # Keep the complete opaque Graph site ID.
    if client.drive_id or client.drive_name:
        drives = list(client.api.iter_collection("sites/{}/drives".format(graph_id(site_id))))
        matches = [d for d in drives if (d.get("id") == client.drive_id if client.drive_id else d.get("name") == client.drive_name)]
        if len(matches) != 1:
            raise ValueError("The requested document library did not uniquely resolve within this site")
        drive = matches[0]
        if client.drive_name and drive.get("name") != client.drive_name:
            raise ValueError("drive_id and drive_name identify different libraries")
    else:
        drive = client.api.graph_call(resource_url="sites/{}/drive".format(graph_id(site_id)))
    if not isinstance(drive, dict) or not drive.get("id"):
        raise ValueError("Graph returned incomplete document library metadata")
    # Commit all metadata together, after site AND selected library validation.
    client.site_id, client.web_url = site_id, site["webUrl"]
    client.drive_id, client._drive_name = drive["id"], drive.get("name")
    client.drive_web_url = drive.get("webUrl")


def rebind(client, value, *, site=False):
    with client._metadata_lock:
        before = (client._site_name, client._drive_name, client.site_id, client.drive_id, client.web_url, client.drive_web_url)
        try:
            if site:
                client._site_name, client.site_id = value, None
            else:
                client._drive_name = value
            client.drive_id = None
            client.clear_cache()
            resolve_site(client)
        except Exception:
            (client._site_name, client._drive_name, client.site_id, client.drive_id, client.web_url, client.drive_web_url) = before
            raise


def library_relative(client, value, leading_slash=False, *, _budget=None):
    """Convert an absolute library URL only when it belongs to the selected drive.

    Relative input is already library-relative: a folder named Documents is real,
    not an implicit request to strip a presumed default document library name.
    """
    if not isinstance(value, str):
        raise TypeError("SharePoint paths must be strings")
    if "://" not in value:
        drive_path(value)
        path = value.strip("/")
    else:
        p = urlsplit(value)
        if p.scheme != "https" or p.username or p.password or p.fragment:
            raise ValueError("Supply a clean HTTPS document library URL")
        if p.query:
            path = _document_url_path(client, p, _budget=_budget)
            return ("/" if leading_slash else "") + path
        if not client.drive_web_url:
            raise ValueError("Drive webUrl metadata is required to convert an absolute URL")
        base = urlsplit(client.drive_web_url)
        base_path = unquote(base.path).rstrip("/")
        source_path = unquote(p.path)
        if p.netloc.lower() != base.netloc.lower() or not (source_path == base_path or source_path.startswith(base_path + "/")):
            raise ValueError("URL is outside the selected document library")
        path = source_path[len(base_path):].strip("/")
        drive_path(path)
    return ("/" if leading_slash else "") + path


def _document_url_path(client, source, *, _budget=None):
    """Resolve Doc.aspx by authoritative document GUID, only in the bound drive."""
    site = urlsplit(getattr(client, "web_url", "") or "")
    expected = unquote(site.path).rstrip("/") + "/_layouts/15/Doc.aspx"
    if (not site.netloc or source.netloc.lower() != site.netloc.lower()
            or unquote(source.path).casefold() != expected.casefold()):
        raise ValueError("Browser document URLs must belong to the bound SharePoint site")
    pairs = parse_qsl(source.query, keep_blank_values=True)
    options = {}
    for key, value in pairs:
        key = key.casefold()
        if key not in {"sourcedoc", "file", "action", "mobileredirect", "web"} or key in options:
            raise ValueError("Unsupported or ambiguous browser document URL")
        options[key] = value
    try:
        wanted = UUID(options["sourcedoc"].strip("{}"))
    except (KeyError, ValueError, AttributeError):
        raise ValueError("Browser document URL requires a valid sourcedoc GUID") from None
    # Never infer identity from the optional file= display name or redeem a
    # sharing link using broader permissions. Traversal has the usual budgets.
    from ._drive_walk import walk_drive
    matches = []
    for item, path in walk_drive(client, document_ids=True, _budget=_budget):
        ids = item.get("sharepointIds") or {}
        if not isinstance(ids, dict):
            from ._graph import GraphError
            raise GraphError(code="invalid_document_metadata")
        observed = ids.get("listItemUniqueId")
        try:
            same = UUID(observed.strip("{}")) == wanted if isinstance(observed, str) else False
        except ValueError:
            same = False
        if same and "file" in item:
            matches.append((item, path))
    if len(matches) != 1:
        raise ValueError("Document GUID did not uniquely resolve in the selected library")
    item, path = matches[0]
    if "file" in options and options["file"].casefold() != item["name"].casefold():
        raise ValueError("Browser document display name contradicts its document GUID")
    drive_path(path)
    return path
