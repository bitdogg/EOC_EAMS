# Denodo Extract — Developer & User Guide

This guide documents the Denodo extract script end-to-end and provides
operational and teaching detail for new and experienced users.

## High-level workflow
1. Select YAML config (CLI or GUI)
2. Load paths, URL, auth metadata
3. Retrieve password from keyring
4. Fetch Denodo HTML with retries
5. Extract embedded JSON
6. Write timestamped JSON and CSV

## YAML configuration
```yaml
DATALOC:
  ROOT_DIR: C:\path\to\root
  OUTPUT_DIR: output
DATA:
  EAMS_RESULT_DATA: eams_result_data.json
  EAMS_CSV_OUTPUT: eams_output.csv
  URL: https://denodo.host.example/report
AUTHENTICATION:
  SERVICE_NAME: denodo-prod
  USERNAME: myuser
```

## OneDrive/SharePoint storage options
The YAML can support multiple storage backends (local sync vs. SharePoint/Graph). A
simple pattern is to keep a single `STORAGE.BACKEND` switch and include the
configuration for each backend. Your sample config already follows that pattern.

### Option A: Local OneDrive sync path
**How it works**
- `STORAGE.BACKEND: local_sync`
- `STORAGE.LOCAL_SYNC.ROOT_DIR` points to a local OneDrive-synced folder.
- The script reads/writes via normal filesystem paths.

**Pros**
- Minimal code changes (works with existing local path handling).
- No OAuth/App Registration overhead.

**Cons**
- Requires OneDrive client installed and signed-in on every machine.
- Risk of sync conflicts if multiple writers touch the same files.
- Less reliable for headless servers or scheduled jobs.

### Option B: SharePoint/OneDrive via Microsoft Graph (recommended for servers)
**How it works**
- `STORAGE.BACKEND: sharepoint`
- Use Microsoft Graph API to read/write files in a SharePoint document library
  or OneDrive drive.
- Store a `SITE_URL`, `DOCUMENT_LIBRARY`, and optional `BASE_FOLDER` to resolve
  the remote paths, or store stable IDs (drive/item IDs) if you prefer.
- App-only auth can be provided via environment variables: `GRAPH_TENANT_ID`,
  `GRAPH_CLIENT_ID`, and `GRAPH_CLIENT_SECRET` (or `CLIENT_SECRET_ENV` in YAML
  to point at a secret environment variable).

**Pros**
- Ideal for scheduled jobs and servers (no sync client required).
- Consistent access across machines/users with app-only auth.
- SharePoint permissions, versioning, and auditing are preserved.

**Cons**
- Requires Azure AD app registration and token management.
- More code to handle API calls, throttling, and errors.

### Option C: WebDAV/Network drive mapping (legacy)
**How it works**
- Map SharePoint as a network drive or use a WebDAV URL.

**Pros**
- Filesystem-like access with minimal code changes.

**Cons**
- WebDAV support is inconsistent and increasingly discouraged.
- Authentication and MFA can be brittle for automation.

### Suggested YAML pattern
```yaml
STORAGE:
  BACKEND: local_sync   # or sharepoint

  LOCAL_SYNC:
    ROOT_DIR: C:\Users\me\OneDrive\EAMS

  SHAREPOINT:
    SITE_URL: https://contoso.sharepoint.com/sites/eams
    DOCUMENT_LIBRARY: Shared Documents
    BASE_FOLDER: EAMS
```

**Recommendation**
For a scheduled job or multi-user server scenario, prefer Microsoft Graph with
app-only authentication and store stable drive/item identifiers (or stable paths
under a known document library). Local sync is fine for single-machine or
developer use.

## Security notes
- No secrets in code or YAML
- Prefer verified TLS when possible
- Scope Denodo accounts to least privilege
