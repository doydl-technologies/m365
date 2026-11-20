# Shared storage API — m365 2.9.6

Public imports:

~~~python
from m365.graph import Site, Drive, Mail, GraphError
~~~

The packages `_dateplumb`, `_syslib`, and `_preprocessing` remain internal.
Names beginning with an underscore are internal and must not be used by applications.

## Resource roles

Site discovers a SharePoint site and selects a document library. Drive binds to
an explicit Graph drive ID, optionally anchored to an exact folder ID. A
SharePoint document library is also a Graph drive; Drive is not OneDrive-only.

Both facades use one private storage engine for the modern methods below.
There is no copied second implementation of search, upload, download or move.

## Drive binding

Use known Graph IDs; `Drive` does not discover an account, enumerate drives,
redeem sharing links, or resolve `/me/drive` automatically.

~~~python
with Drive(token_engine=auth,
           drive_id="APPROVED_DRIVE_ID",
           folder_id="APPROVED_FOLDER_ITEM_ID") as drive:
    print(drive.list_children())
    print(drive.find_files("quarterly", folder_path="Reports"))
~~~

`folder_id` is an actual item ID, not the alias `"root"`. Omitting it selects
drive root and requires root access. A failed folder binding never falls back
to root. Work or school drives use existing application or delegated grants;
personal OneDrive uses a delegated user token.

## The common API

| Method | Meaning |
|---|---|
| get_item(path="") | Copied file/folder metadata |
| list_children(folder_path="") | Complete immediate children or an explicit error |
| find_files(term, *, folder_path="", exact_match=False, recursive=True, limit=None, sort_order=None) | Bounded filename-only search |
| find_folders(...) | Equivalent folder-name search |
| find_file_folder(file_term, folder_term, *, folder_path="", exact_match=True, recursive=True) | Exactly one file and one folder, returned in that order |
| upload(source, folder_path="", new_file_name=None, *, conflict_behavior="fail", chunk_size=3276800) | Confirmed nonempty file upload |
| download_file(path, directory) | Atomic no-clobber local file |
| download_folder(folder_path, directory, *, compress=False, max_workers=4) | New folder or ZIP; preserves empty directories |
| create_folder(name, folder_path="") | One folder; fail on name conflict |
| move(path, destination_folder="", new_name=None, *, if_match=None) | Move a file/folder within the same bound drive/anchor |
| rename(path, new_name, *, if_match=None) | Rename in its existing parent |
| close() / context manager | Release owned HTTP sessions |

All paths are unencoded, relative to the selected anchor. URLs, external IDs,
absolute paths, traversal and shortcuts are not accepted as file paths.

## Site usage and ownership

~~~python
with Site(token_engine=auth, site_name="https://example.sharepoint.com/sites/Example") as site:
    metadata = site.get_item("Reports/quarterly.pdf")
    matches = site.find_files("quarterly", folder_path="Reports")
~~~

Each modern Site call creates and closes an independent drive view. It captures
the selected library under the site's metadata lock, so a later site/library
rebind cannot retarget an operation already bound to its snapshot.

For batches, keep one view open to reuse its HTTP sessions:

~~~python
with Site(token_engine=auth, site_name="https://example.sharepoint.com/sites/Example") as site:
    with site.as_drive() as storage:
        print(storage.list_children("Reports"))
        print(storage.find_files(".pdf", folder_path="Reports", sort_order="name"))
~~~

The caller owns and must close the returned Drive. Closing it does not close the
Site. A previously created view remains usable and pinned if Site is rebound
or closed; a closed Site cannot create another view through these modern methods.
The token provider still owns authentication/refresh.

Site.as_drive(folder_id=...) can explicitly select a folder. If Site has
search_root configured, modern methods use that exact anchor and refuse an
explicit conflicting folder. They never silently fall back to drive root.
Legacy search_mode="indexed" does not enable indexed search for the modern API.

## Search behavior

Modern searches match filenames/folder names case-insensitively, not file
contents or general metadata. They enumerate only the bound tree and enforce
request/item/depth budgets. Malformed pagination, foreign parents/drives, IDs
reusing known ancestors and contradictory metadata fail before returning success.

sort_order may be:

- None: traversal order, with an optional intentional early limit.
- name / name_desc: case-insensitive name order.
- recent / oldest: createdDateTime order, matching the legacy Site sort field.

Sorted searches scan all matches within the budgets before applying limit.
Missing/invalid createdDateTime raises GraphError when date sorting is requested,
including offsets that would put the UTC timestamp outside the supported range.
The error code is `missing_created_datetime` for absent timestamps and
`invalid_created_datetime` for malformed or out-of-range timestamps.
Sorting alone returns all matches, not an implicit one-item result.
limit=0 returns an empty result; limits reject booleans and negative values.

find_file_folder uses a single complete traversal and shared budget. Missing or
ambiguous matches raise; it never chooses an arbitrary duplicate name.

## Move and rename safety

~~~python
with Drive(token_engine=auth, drive_id="APPROVED_DRIVE_ID",
           folder_id="APPROVED_FOLDER_ID") as storage:
    moved = storage.move("Incoming/report.pdf", "Reviewed")
    renamed = storage.rename("Reviewed/report.pdf", "final-report.pdf")
~~~

Moves:

- Keep the same drive and anchor; cross-drive transfers are not implemented.
- Cannot move/rename the bound anchor.
- Reject moving a folder into itself or its descendants.
- Validate folder descendants against the global depth limit before moving.
- Require a source ETag from metadata, or explicit if_match supplied by the caller.
- Send If-Match and fail conflict behavior; never retry with replace or auto-rename.
- Check existing destination names before PATCH. Graph remains responsible for
  conflicts created concurrently after that check.
- Require HTTP 200 and matching item ID, kind, parent, drive, name and file size.
- Do not infer extensions: new_name is the entire desired name.
- Never automatically replay an unknown commit or attempt an inverse move.

A failed/contradictory final response does not prove that the server undid the
operation. Inspect remote state before retrying. Noop moves return current
metadata without PATCH; an explicitly stale ETag still fails.

ETags protect the source item, not an entire concurrently changing folder tree.
This API is not server snapshot isolation or a security sandbox for reuse of a
more broadly privileged token. Server permissions remain the actual boundary.
An unsupported Graph request fails visibly; it does not trigger a broader or
destructive fallback.

## Authentication factories

Both Site and Drive support from_profile with the same shared profile policy:

~~~python
with Site.from_profile("Saved Files Connection",
                       site_name="https://example.sharepoint.com/sites/Example") as site:
    print(site.list_children())

with Drive.from_profile("Saved Files Connection",
                        drive_id="APPROVED_DRIVE_ID",
                        folder_id="APPROVED_FOLDER_ID") as storage:
    print(storage.list_children())
~~~

No profile is created/edited and no permissions are granted by these factories.
Application-secret/certificate profiles use client credentials. Delegated user
profiles use the encrypted cache; browser/device UI requires allow_interactive=True.
Other authentication errors never cause interactive fallback. Environment
credential overrides, even empty ones, are rejected for named-profile loading.
A custom Graph endpoint requires explicit matching scopes.

Both also expose from_default_env, from_env_file, from_os_env and from_entry.
Resource arguments differ: Site selects site/library; Drive requires drive_id
and optional folder_id. These compatibility-style builders retain the existing
Site silent-delegated authentication contract. They do not infer app-only flow
from a supplied client secret. For app-only authentication, use from_profile or
an explicitly authenticated azauthlib provider.

Do not pass scope/credential changes into file-operation methods. Tenant selection,
consent, resource grants and scheduling remain application responsibilities.

### Windows certificate authentication

For unattended work against an authorized organization drive or SharePoint site,
an existing Entra application can use its registered public certificate. The
private key must be available to the process account in the Windows Personal
certificate store. With an RSA-PSS-capable CNG key, the key can be non-exportable;
no PFX export or client secret is needed. Use `store_location="LocalMachine"`
only when the certificate is stored there and the process can use its key.

~~~python
import os

from azauthlib import Authentication
from m365.graph import Site

auth = Authentication(scopes=["https://graph.microsoft.com/.default"])
auth.Build.WithWindowsCertificate(
    client_id=os.environ["GRAPH_CLIENT_ID"],
    tenant_id=os.environ["GRAPH_TENANT_ID"],
    thumbprint=os.environ["GRAPH_CERTIFICATE_THUMBPRINT"],
    store_location="CurrentUser",
)
auth.ClientCredentials()

with Site(token_engine=auth, site_id=os.environ["GRAPH_SITE_ID"]) as site:
    print(site.list_children())
~~~

The example only reads metadata. Reuse the authenticated provider with `Drive`
as shown under [Drive binding](#drive-binding). `.default` uses existing grants;
it does not create them. For certificate errors, check the store, key access,
expiry, and app registration. An HTTP 403 can still mean the site or folder
grant is missing. Do not log tokens or signed assertions.

## Existing Site interfaces remain supported

These still work without signature or return-shape changes:

- site.operations.Uploader(...)
- site.operations.Downloader(...).file(...) / .folder(...)
- site.operations.Mover(...)
- site.search_by.filename(...) / .foldername(...) / .filefoldername(...)
- Existing path/url helpers and authentication factories.

Legacy indexed search remains opt-in through Site's existing configuration.
Legacy filename sorting retains its implicit one-result behavior when sorting
without num_files. Legacy Mover retains extension inference. Modern methods
instead use the explicit rules above; old callers are not silently migrated.

Legacy indexed lookups, child expansion and combined lookup enforce the configured
search request/item limits across the whole call. Limits count empty pages and
nonmatching rows, and raise `search_request_limit` or `search_item_limit` rather
than returning a partial result. `exact_match` and `include_children` require
actual booleans; strings such as `"false"` raise `ValueError`.

No delete/copy API, new permission grants, tenant-specific assumptions, live
transfers, scheduled jobs or package publication are introduced.

## Transfer and operational limits

Uploads default to `conflict_behavior="fail"`; explicit replace/rename remain
available only for uploads. Replacement is not an ETag-based concurrency check.
Empty uploads remain unsupported. Upload directories must be archived or enumerated
by the caller. `BytesIO` needs `new_file_name` and retains its cursor position.
Completion checks validate the response metadata, not a content hash. If the
session reports `upload_completion_unknown`, inspect the destination before retrying.

Downloads stage data, reject unexpected partial responses, check available sizes,
and require hard-link support for atomic no-overwrite publication. ZIPs preserve
empty directories and keep the original downloads. Folder/ZIP operations are not
whole-tree transactions: completed local files can remain after a later failure,
while unrelated local content is preserved.

Metadata budgets are not total wall-clock or byte limits. They exclude streaming,
upload fragments and bounded transport retries. Direct Site calls additionally
validate a fresh bound view; use as_drive for bulk/session reuse.
Drive and Site default to 200 operation metadata requests, 5,000 items and depth 30.
Site's legacy constructor requires positive search depth; Drive also accepts zero
for root-only/empty operations. Depth includes folder prefix and filename.

Tests use synthetic HTTP, temporary local data and stateful services. The local
build must pass source/wheel/sdist suites. Live Graph move/rename behavior, other
operating systems and other Python runtimes are not certified by offline tests.
