<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/doydl-technologies/m365/main/assets/png/m365-endorsed-lockup-horizontal-python-light-957x323-384dpi.png">
    <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/doydl-technologies/m365/main/assets/png/m365-endorsed-lockup-horizontal-python-dark-957x323-384dpi.png">
    <img src="https://raw.githubusercontent.com/doydl-technologies/m365/main/assets/png/m365-endorsed-lockup-horizontal-python-dark-957x323-384dpi.png" alt="m365 — powered by DOYDL Technologies" width="700">
  </picture>
</p>

# 🧩 m365: Microsoft Graph Integration Toolkit

**Build Python integrations for Microsoft 365 without coupling your business logic to tenant-specific authentication, URLs, or document-library layouts.**

`m365` provides reusable clients for SharePoint and OneDrive document management and Outlook mailbox workflows through Microsoft Graph. It is developed by DOYDL Technologies for applications, internal tools, and data pipelines that operate on explicitly authorized Microsoft 365 resources.

<p align="center">
  <a href="https://pypi.org/project/m365/"><img src="https://img.shields.io/pypi/v/m365" alt="PyPI version"></a>
  <a href="https://pypi.org/project/m365/"><img src="https://img.shields.io/pypi/pyversions/m365" alt="Supported Python versions"></a>
  <a href="https://github.com/doydl-technologies/m365/blob/main/LICENSE"><img src="https://img.shields.io/github/license/doydl-technologies/m365" alt="License"></a>
  <a href="https://doydl.studio"><img src="https://img.shields.io/badge/Powered%20by-DOYDL%20Technologies-blue" alt="Powered by DOYDL Technologies"></a>
</p>

`m365` requires Python 3.10 or later.

> [!IMPORTANT]
> This is an independent toolkit, not an official Microsoft product. A configured Microsoft Entra application, appropriate resource permissions, and a valid Microsoft Graph authentication provider are prerequisites. Installing the package does not grant access to Microsoft 365.

## Contents

- [🔍 What m365 does](#what-m365-does)
- [🏗️ Architecture and package layout](#architecture-and-package-layout)
- [📦 Requirements and installation](#requirements-and-installation)
- [🔐 Authentication and authorization](#authentication-and-authorization)
- [🚀 Quick start](#quick-start)
- [📁 SharePoint client reference](#sharepoint-client-reference)
  - [Site and library selection](#site-and-library-selection)
  - [Paths and searching](#paths-and-searching)
  - [Uploads](#uploads)
  - [Downloads](#downloads)
  - [Moving documents](#moving-documents)
- [📬 Outlook mailbox reference](#outlook-mailbox-reference)
  - [Reading messages and folders](#reading-messages-and-folders)
  - [Drafts and message operations](#drafts-and-message-operations)
  - [Attachments](#attachments)
  - [Incremental synchronization](#incremental-synchronization)
- [☁️ OneDrive integration](#onedrive-integration)
- [Shared storage operations](#shared-storage-api)
- [🌐 Shared Graph transport](#shared-graph-transport)
- [🛡️ Reliability, security, and operational limits](#reliability-security-and-operational-limits)
- [🩺 Errors and troubleshooting](#errors-and-troubleshooting)
- [💼 Business integration scenarios](#business-integration-scenarios)
- [🧪 Development and validation](#development-and-validation)
- [📋 Release and packaging notes](#release-and-packaging-notes)
- [📚 Project links and licensing](#project-links-and-licensing)

<a id="overview"></a>

## What m365 does

The toolkit organizes ordinary Microsoft Graph operations into resource-bound clients, so an application can concentrate on the workflow rather than repeatedly assembling request paths, parsing metadata, or coordinating file transfers.

| Capability | Entry point | Purpose |
| --- | --- | --- |
| 📁 SharePoint | `m365.graph.Site` | Select a site and document library |
| 🔎 Document discovery | `site.search_by` | Resolve file and folder metadata |
| 📤 Document uploads | `site.operations.Uploader(...)` | Transfer a file or byte stream |
| 📥 Document downloads | `site.operations.Downloader(...)` | Download selected items and folder trees |
| 🔄 Document moves | `site.operations.Mover(...)` | Move a selected document |
| 📬 Outlook | `m365.graph.Mail` | Work with an authorized mailbox |
| 📨 Mail search | `mailbox.search_by` | Read messages, folders, and attachments |
| 📝 Composition | `mailbox.operations.Drafter()` | Prepare and update drafts |
| 📎 Attachments | `mailbox.operations.AttachmentDownloader()` | Save selected attachments |
| ☁️ OneDrive | `m365.graph.Drive` | List, search, upload and download within an explicitly selected drive/folder |

The toolkit does not provision tenants, create permission grants, provide a job scheduler, or replace Microsoft Graph's own API contracts. OneDrive access requires existing resource permissions; creating a client never grants them.

<a id="architecture"></a>

## Architecture and package layout

The `m365` namespace loads `graph` lazily. The three resource clients share transport concepts but expose domain-specific operations.

| Layer | Module | Responsibility |
| --- | --- | --- |
| Application | Your Python project | Business rules, schedules, approvals, persistence |
| Identity | `azauthlib` or injected provider | Authentication and token renewal |
| Resource clients | `graph.sharepoint`, `graph.onedrive`, `graph.mailbox` | Bind a provider to an explicit resource |
| Operations | `graph.operations` | Document and mailbox operations |

The supported application entry points are `Site`, `Drive`, `Mail`, and `GraphError`, imported from `m365.graph`, together with their documented client operations. `_dateplumb`, `_preprocessing`, and `_syslib` are internal implementation helpers, not supported application APIs. Any module, function, method, or attribute whose name starts with `_` is internal and must not be used by applications. Internal helpers may change without public API compatibility guarantees.

<a id="installation"></a>

The former unprefixed helper import paths have been removed without compatibility
aliases. Private helpers remain packaged for client functionality; they are not
application APIs. Python privacy is a convention, not access control.
`from m365.graph import *` exports only `Site`, `Drive`, `Mail`, and `GraphError`.
If you used the former client names, update `OneDrive` to `Drive` and `Mailbox`
to `Mail` in imports and constructors. There are no old-name aliases.

## Requirements and installation

### 📋 Dependencies

| Dependency | Requirement | Role |
| --- | --- | --- |
| Python | `>=3.10` | Runtime |
| `azauthlib` | `>=6.0.6,<7` | Authentication factories |
| `requests` | Declared dependency | HTTP transport |
| `numpy` | Declared dependency | Supporting utilities |
| `tqdm` | Declared dependency | Progress reporting |
| `dately` | `>=4.7.8,<4.8` (optional) | Internal advanced-date backend; not a public API |

### 1️⃣ Create an isolated environment

```bash
python -m venv .venv
```

| Environment | Activate |
| --- | --- |
| Windows PowerShell | `.\.venv\Scripts\Activate.ps1` |
| Windows MSYS2 with standard Windows Python | `source .venv/Scripts/activate` |
| Linux/macOS | `source .venv/bin/activate` |

On Windows, if pip reports a path-length error while installing PySide6
(a dependency of `azauthlib`), create the environment at a shorter path and
install using that environment's Python. For example, in PowerShell:

```powershell
$shortEnv = Join-Path $env:LOCALAPPDATA 'm365-venv'
python -m venv $shortEnv
& (Join-Path $shortEnv 'Scripts\python.exe') -m pip install m365
```

### 2️⃣ Install

This guide documents `m365` `2.9.6`. To install the latest version published on
[PyPI](https://pypi.org/project/m365/):

```bash
python -m pip install m365
```

To install from a source checkout, run from the repository root:

```bash
python -m pip install .
```

Or install the wheel for this version after building it:

```bash
python -m pip install ./dist/m365-2.9.6-py3-none-any.whl
```

### 3️⃣ Verify the installed versions

```python
import m365
from m365 import graph

print(m365.__version__)  # 2.9.6
print(graph.__version__) # 2.9.6
```

The distribution and Graph module report the same version.
`azauthlib` is installed by default; the legacy `auth` extra adds no dependencies.

<a id="authentication"></a>

## Authentication and authorization

Clients accept a caller-supplied authentication provider. The structural contract in `m365.graph.auth_types.AuthProvider` defines an `access_token` property and `Silent(scopes=None)` method. Token renewal and credential handling belong to the provider, not the resource operation.

| Mode | Appropriate use |
| --- | --- |
| Delegated | An approved signed-in user's operations, constrained by consent and resource permissions |
| Application | An approved service identity, constrained by administrator-granted resource access |

Built-in factories on `Site`, `Drive`, and `Mail` include `from_default_env`, `from_env_file`, `from_os_env`, and `from_entry`. They delegate provider construction to `azauthlib`. Silent delegated authentication is the default; a missing usable sign-in raises `azauthlib.InteractionRequiredError`. Use an explicitly approved interactive flow when necessary. For unattended workloads, use a separately configured noninteractive provider.

> [!WARNING]
> Never commit secrets, tokens, signed assertions, token caches, or preauthenticated transfer URLs to source control. Do not use the toolkit to access resources outside your organization's authorization. Keep each tenant's provider configuration separate and request only the access needed for the intended task.

For provider-specific configuration consult [azauthlib](https://github.com/doydl-technologies/azauthlib) and your organization's approved identity setup. See the [Windows certificate example](https://github.com/doydl-technologies/m365/blob/main/STORAGE.md#windows-certificate-authentication) for an unattended storage job. A valid token alone does not establish permission to a site, drive, mailbox, or individual item.

## Quick start

Start with read-only metadata and use a resource you are authorized to access.
For example, if you saved a profile with `azauthlib`, use its profile name and
your actual SharePoint site and library:

```python
from m365.graph import Site

with Site.from_profile(
    "My approved Microsoft Graph profile",
    site_name="https://example.sharepoint.com/sites/Finance",
    drive_name="Documents",
) as site:
    print(site.get_item()["name"])
```

This does not create a profile, grant permission, or write a file. A delegated
profile needs an existing sign-in unless you explicitly opt into interactive
authentication. See [Storage API](https://github.com/doydl-technologies/m365/blob/main/STORAGE.md) for profile and `Drive` examples.

The following examples take an already configured provider from your application.
The resource URLs and names are illustrative.

```python
from m365.graph import Site

def find_report(provider):
    with Site(
        token_engine=provider,
        site_name="https://example.sharepoint.com/sites/Finance",
        drive_name="Documents",
    ) as site:
        return site.search_by.filename("Reports/Forecast.xlsx")
```

The context manager closes the SharePoint HTTP transport after use. Search returns metadata rather than downloading the workbook automatically.

```python
from m365.graph import Mail

def recent_inbox_subjects(provider):
    mailbox = Mail(token_engine=provider, mailbox="me")
    try:
        messages = mailbox.search_by.messages(
            folder="inbox", top=5, select=["id", "subject"]
        )
        return [message.get("subject") for message in messages]
    finally:
        mailbox.api.close()
```

The second example assumes approved delegated access. `me` is a signed-in user resource, not an application identity.

<a id="sharepoint"></a>

## SharePoint client reference

<a id="site-selection"></a>

### Site and library selection

`Site` resolves the target site and selected document library during construction. Use explicit resource IDs when your deployment already stores them.

| Parameter | Default | Meaning |
| --- | --- | --- |
| `token_engine` | `None` | Injected provider, or factory-built provider |
| `site_name` | `None` | Site URL, hostname/path, or short site name |
| `site_id` | `None` | Complete opaque Graph site identifier |
| `hostname` | `None` | Optional host when identifying a site |
| `drive_name` | `None` | Name of document library |
| `drive_id` | `None` | Opaque ID of a drive in the selected site |
| `graph_url` | `https://graph.microsoft.com/v1.0` | Configured HTTPS Graph endpoint |
| `request_timeout` | `(10, 120)` | Connect/read timeout, seconds |
| `max_retries` | `3` | Bounded retry configuration |
| `search_mode` | `enumerate` | Name lookup strategy |
| `search_root` | `None` | Optional folder item ID for name traversal |
| `search_max_items` | `5000` | Traversal item budget |
| `search_max_requests` | `200` | Logical page budget |
| `search_max_depth` | `30` | Traversal depth budget |

When no drive is specified, the site's default drive is selected. A named drive is resolved across its available pages. `site.site_id`, `site.drive_id`, `site.web_url`, and `site.drive_web_url` describe the resolved resource. `update_site_name` and `update_drive_name` validate rebinding; unsuccessful changes roll back. Do not retarget an active transfer client.

<a id="sharepoint-search"></a>

### Paths and searching

Use unencoded, library-relative paths with forward slashes, for example `Reports/2026/Forecast.xlsx`. A component named `Documents` is treated as a real subfolder, not automatically removed. Absolute document URLs are normalized only within the selected library. Parent folders must exist for uploads.

| Method | Purpose |
| --- | --- |
| `site.search_by.filename(term, exact_match=True, sort_order=None, num_files=None)` | File search |
| `site.search_by.foldername(term, include_children=False, exact_match=True)` | Folder search |
| `site.search_by.filefoldername(file_search, folder_search)` | Resolve one file and destination folder |
| `site.url.normalize_sharepoint_path(url)` | Convert a selected-library URL into a relative path |

Default name lookup uses bounded enumeration, compares names case-insensitively, and does not search file contents. `exact_match=False` permits substring matching. Duplicate names remain ambiguous; explicit relative paths are preferred. Remote shortcuts are not followed. Indexed lookup is an explicit optional mode requiring compatible authorization.

`search_root` limits legacy name traversal, not explicit paths, which remain
library-relative. For the modern anchor-relative methods and search error codes,
see [the storage reference](https://github.com/doydl-technologies/m365/blob/main/STORAGE.md#search-behavior).

```python
matches = site.search_by.filename("Reports/Forecast.xlsx")
folders = site.search_by.foldername("Reports", include_children=False)
```

Selected-site `Doc.aspx?sourcedoc=<GUID>` browser links can also be resolved when their GUID uniquely matches the selected library's item metadata. A display filename is not proof of identity. Prefer known paths or IDs.

### Uploads

`Uploader` accepts a nonempty local file or `io.BytesIO`. Uploads use sequential sessions and validate returned completion metadata. Uploading a directory or an empty file is not supported.

```python
result = site.operations.Uploader(
    "./exports/forecast.xlsx",
    sharepoint_folder_path="Reports",
    conflict_behavior="fail",
)
print(result["id"])
```

```python
from io import BytesIO

report = BytesIO(b"period,revenue\n2026-01,1000\n")
result = site.operations.Uploader(
    report,
    sharepoint_folder_path="Reports",
    new_file_name="sample.csv",
    conflict_behavior="fail",
)
```

| Option | Default | Behavior |
| --- | --- | --- |
| `new_file_name` | `None` | Optional destination name; required for `BytesIO` |
| `chunk_verbose` | `True` | Accepted for compatibility; does not control progress output |
| `max_file_size` | `"4MB"` | Accepted for compatibility; does not control chunk size |
| `conflict_behavior` | `fail` | Fail when remote filename conflicts |
| Other conflict values | `replace`, `rename` | Explicitly permit remote replacement or renaming |

To set the upload chunk size, use the shared `site.upload(source, folder_path, chunk_size=...)` method. The size is in bytes and must be a positive multiple of 320 KiB below 60 MiB.

The caller's `BytesIO` position is restored. Completion checks validate returned identity, filename, size and file metadata; they do not establish an end-to-end cryptographic content hash. When the final outcome is unknown, inspect the selected destination before deciding whether to retry.

### Downloads

Pass file or folder metadata from a successful search to a download helper. Local destinations are never overwritten.

```python
files = site.search_by.filename("Reports/Forecast.xlsx")
downloader = site.operations.Downloader(
    max_workers=2,
    compress=False,
)
# Use a private, approved destination directory.
paths = downloader.file("./downloads", files)
```

For an explicitly selected folder, `downloader.folder(save_dir, folder_metadata)` always preserves the folder tree on disk. `compress=True` also creates a ZIP without deleting the downloaded files. With `preserve_dir_structure=True`, the ZIP keeps relative paths; with `False` (the default for `Downloader`), entries are flattened. Empty folders are included, and duplicate ZIP entry names are rejected rather than overwritten.

Downloads stage data before publication, check the supplied size, and clean incomplete temporary files. A filesystem supporting hard links is required for atomic no-overwrite publication. Successful earlier files may remain if a later transfer fails; a batch is not a transaction.

<a id="moves"></a>

### Moving documents

`site.operations.Mover(sharepoint_filename, sharepoint_foldername, new_sharepoint_filename=None)` resolves a unique source and destination before moving a document. The result is validated against the expected item and parent metadata. This changes the remote document location; review the source and destination first. A response validation failure after the move does not imply rollback.

<a id="outlook"></a>

## Outlook mailbox reference

`Mail(token_engine, mailbox="me", use_immutable_ids=True)` binds a provider to a mailbox. The `mailbox` value can be `me` for delegated access or an authorized user identifier. Immutable IDs are requested by default where applicable. The client exposes `mailbox_profile`, `search_by`, `operations`, `tools`, and `api`.

<a id="mail-reads"></a>

### Reading messages and folders

| Entry point | Purpose |
| --- | --- |
| `search_by.folders(...)` | Resolve mailbox folders |
| `search_by.messages(...)` | Read selected messages |
| `search_by.message(message_id, ...)` | Read one message by opaque ID |
| `search_by.attachments(message_id, ...)` | List attachments |
| `search_by.message_delta(...)` | Read a message delta page |
| `search_by.folder_delta(...)` | Read a folder delta page |

```python
messages = mailbox.search_by.messages(
    folder="inbox",
    top=10,
    order_by="receivedDateTime desc",
    select=["id", "subject", "receivedDateTime", "isRead"],
    unread=True,
)
```

| Message option | Default | Interpretation |
| --- | --- | --- |
| `query` | `None` | Optional client-side text query |
| `folder` | `None` | Folder reference |
| `top` | `25` | Desired result limit, not an unlimited scan |
| `order_by` | `receivedDateTime desc` | Requested ordering |
| `select` | `None` | Fields to retrieve; a string denotes one field |
| `unread` | `None` | Optional read-state filter |
| `body_type` | `html` | Requested body format |
| `auto_paginate` | `True` | Follow continuation pages as necessary |
| `max_items` | `None` | Explicit maximum returned records |
| `page_limit` | `None` | Optional intentional page subset |

Use explicit message and folder IDs when names collide. Folder discovery supports `include_hidden`, `include_children`, `exact_match`, `recursive`, and `top` options. High-level reads raise `GraphError` on malformed or failed pages rather than disguising an incomplete read as success.

<a id="mail-operations"></a>

### Drafts and message operations

The composition API supports preparing a draft for human review separately from sending it. For example, prepare the body in your application and pass the approved recipient, subject and body to `mailbox.operations.Drafter().create(...)`. Text and HTML are supported; recipient arguments accept a single address or a sequence. Attachment, category, reply-to, importance and custom internet-header options are available subject to Graph's own restrictions.

| Operation factory | Methods | Effect |
| --- | --- | --- |
| `Drafter()` | `create`, `update`, `create_reply`, `create_reply_all`, `create_forward`, `send` | Draft creation, revision and dispatch |
| `Sender()` | `send` | Send a composed message |
| `MessageUpdater()` | `update`, `mark_read`, `mark_unread` | Modify a message's properties |
| `Mover()` | `move` | Move a selected message to a folder |
| `Deleter()` | `delete` | Delete a selected message |

> [!WARNING]
> Draft creation changes remote state; sending, updating, moving and deleting mail have additional consequences. Require explicit review in applications that perform these operations. Do not interpret a timeout as proof that a mutation did not occur.

### Attachments

Attachments have their own IDs, distinct from message IDs. The download helper accepts an exact `attachment_id` or a uniquely resolvable `attachment_name`. The upload helper accepts a selected message, file source, optional name and MIME content type.

| Method | Purpose |
| --- | --- |
| `search_by.attachments(message_id, auto_paginate=True, max_items=None, page_limit=None)` | Retrieve attachment metadata |
| `operations.AttachmentDownloader().download(message_id, save_dir='.', attachment_id=None, attachment_name=None)` | Save a selected attachment |
| `operations.AttachmentUploader().upload(message_id, file_source, file_name=None, content_type='application/octet-stream')` | Attach a file |

Download filenames are validated and existing files are rejected. Large attachment upload sessions validate intermediate offsets and final completion. An empty legitimate final HTTP 201 is represented with `completed=True`, `status_code=201`, and uploaded `size`; no ID is invented. Resumable Outlook upload recovery is not implemented.

<a id="delta"></a>

### Incremental synchronization

The delta helpers return a **single** page with three keys: `value`, `next_link`, and `delta_link`. The application owns the synchronization checkpoint, persistence, reconciliation and resync policy. Do not rewrite opaque continuation links, drop repeated events, or treat deletion tombstones as ordinary current records. A terminal delta link should be persisted only once the applicable pages have been processed successfully.

HTTP 410 delta failures use `GraphError(code="delta_resync_required")`. The library does not automatically delete local records or reset application state.

<a id="onedrive"></a>

## OneDrive integration

Use `Drive` when you already know the authorized drive ID and, for folder-scoped
access, the folder's item ID. It does not discover an account or call `/me/drive`
for you. Personal OneDrive uses delegated user authentication; an app-only
certificate is not a substitute. See [Drive binding](https://github.com/doydl-technologies/m365/blob/main/STORAGE.md#drive-binding).

## Shared storage API

`Site` and `Drive` share file/folder search, upload, download, folder creation,
move, and rename methods. See [Storage API](https://github.com/doydl-technologies/m365/blob/main/STORAGE.md) for signatures, examples,
scope boundaries, and transfer safety. The SharePoint-specific interfaces above
remain available.

<a id="transport"></a>

## Shared Graph transport

`site.api.graph_call(...)` is the lower-level request entry point for operations not represented by a convenience helper. `site.api.iter_collection(resource_url)` yields individual collection items across continuation pages. Prefer documented Microsoft Graph resource-relative paths rather than constructing arbitrary external URLs. The transport rejects unsupported off-endpoint requests and uses separate sessions for service-issued transfer capability URLs, without Graph bearer headers.

Important response contracts:

- SharePoint transport failures raise `GraphError`.
- Raw ordinary `Mail.api.graph_call` failures retain a redacted error dictionary for compatibility; high-level mailbox reads, streaming, capability operations and response-metadata calls raise errors.
- For a single non-streaming request, `response_metadata=True` returns `status_code` and `body`, not raw headers. It is not supported for batches.
- Batch subrequest failures are surfaced; successful HTTP envelope status alone does not prove all operations succeeded.
- Pagination preserves opaque continuation parameters. Cycles, malformed pages and inconsistent selected-drive metadata raise errors.
- GET-cache arguments are retained for compatibility but HTTP response caching is disabled.

<a id="reliability"></a>

## Reliability, security, and operational limits

| Concern | Behavior | Application responsibility |
| --- | --- | --- |
| Throttling | Bounded retry behavior | Schedule appropriately; respect service limits |
| Partial pagination | Fail visibly on invalid pages and cycles | Commit completeness-dependent results only after successful consumption |
| Search budgets | Limits on items, pages and depth | Narrow searches to known approved locations |
| File integrity | Metadata and size validation | Add independent verification if required |
| Local overwrite | No-clobber publication | Choose approved output paths and supported filesystem |
| Remote conflict | Explicit upload policy | Review before replacing existing content |
| Ambiguous mutation | No blind replay | Inspect current remote state |
| Credential handling | Provider-managed tokens | Use approved identity configuration and private storage |
| Concurrent access | Locks around token reads and metadata rebinding | Avoid retargeting during work; do not assume full thread/process safety |
| Delta state | Invalid checkpoints surfaced | Own checkpoint persistence and resync decisions |

Outlook's default pagination safety budget is 200 logical pages; folder traversal additionally limits depth to 30 and folder count to 5,000. SharePoint's enumeration budgets are configurable per client. A work item that exceeds a budget is not necessarily missing. Pagination does not provide snapshot isolation during concurrent remote changes.

<a id="troubleshooting"></a>

## Errors and troubleshooting

```python
from m365.graph import GraphError

try:
    report_metadata = site.search_by.filename("Reports/Forecast.xlsx")
except GraphError as exc:
    # Avoid logging sensitive URLs, payloads or provider credentials.
    print("Graph operation failed:", exc.status_code, exc.code)
    raise
```

| Symptom | Interpretation / action |
| --- | --- |
| `InteractionRequiredError` | Approved delegated authentication is needed; confirm your provider setup |
| HTTP 401 | Check token audience and provider refresh |
| HTTP 403 | Verify the existing resource-specific authorization with your administrator |
| HTTP 404 | Verify the exact site/library/folder/item resource and its visibility |
| HTTP 429 | Service throttling; respect backoff and review workload frequency |
| `pagination_cycle` | Invalid repetition in continuation state; do not accept partial results |
| `mailbox_page_limit` | Mail traversal safety budget reached |
| `delta_resync_required` | Checkpoint invalidated; follow an explicit resync policy |
| `upload_completion_unknown` | Upload outcome unresolved; inspect destination before retry |
| `batch_subrequest_redirect` | Unsupported redirected batch response |
| Existing file or unsupported hard links | Use a new destination on a compatible filesystem |

For malformed data or invalid arguments, `ValueError` can also be raised. A failed mutation response is not a guarantee of rollback. Inspect `status_code` and `code` when available rather than relying on fragile message-text matching.

<a id="business-scenarios"></a>

## Business integration scenarios

| Workflow | Example application design |
| --- | --- |
| 📊 FP&A reporting | Generate a forecast locally, validate it, then publish to an approved document library with fail-on-conflict policy |
| 🧾 Financial close | Retrieve a known authorized reconciliation workbook, validate data, and record its processing status |
| 📚 Document lifecycle | Resolve an approved contract path and move it only after review |
| 📨 Operational notifications | Prepare a report notification draft for human approval |
| 🗂️ Records processing | Check metadata, normalize timestamps, and maintain a durable application-side checkpoint |

Keep business rules outside the client. For example, a finance application should own period-close approval, workbook validation, reconciliation, audit records, and retention policy; the toolkit provides the Graph integration layer.

<a id="development"></a>

## Development and validation

From the source root, create an isolated Python environment and install in editable mode:

```bash
python -m pip install -e '.[dates]'
python -m pip install build
python -m build
```

The supplied offline test launcher supports `graph` and `combined` suites. Example commands, with directories selected by the developer:

```bash
python -B tools/test_suite.py graph --output ./.local/test-results/graph --dist ./dist
python -B tools/test_suite.py combined --output ./.local/test-results/combined --dist ./dist
```

Use a build directory containing only the wheel and source archive for the
version under test. The regression suites block live network access and do not
require Microsoft 365 credentials.

<a id="release"></a>

## Release and packaging notes

This README is the project guide on GitHub and the project description on PyPI.
The source distribution includes this README, the
[Storage API](https://github.com/doydl-technologies/m365/blob/main/STORAGE.md),
the [license](https://github.com/doydl-technologies/m365/blob/main/LICENSE), and
asset attribution notices.

<a id="project-links"></a>

## Project links and licensing

- 🏠 [Project repository](https://github.com/doydl-technologies/m365)
- 🐛 [Issue tracker](https://github.com/doydl-technologies/m365/issues)
- 🔐 [azauthlib](https://github.com/doydl-technologies/azauthlib)
- 📅 [dately](https://github.com/doydl-technologies/dately)
- 📖 [Microsoft Graph documentation](https://learn.microsoft.com/en-us/graph/overview)
- 🏢 [DOYDL Technologies](https://github.com/doydl-technologies)

The package declares the MIT license. Source headers copyright DOYDL Technologies and include MIT license terms. Third-party assets in `assets/third-party/` have separate notices. Microsoft, Microsoft Graph, Microsoft 365, Outlook, SharePoint, and OneDrive are trademarks of Microsoft Corporation. Their use does not imply affiliation, sponsorship or endorsement.

<p align="center"><strong>DOYDL Technologies</strong> · Reusable integration tools for Python</p>
