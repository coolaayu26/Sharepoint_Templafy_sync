# SharePoint → Templafy Nightly Sync

An Azure Function (Python, Timer Trigger) that mirrors SharePoint content into Templafy Libraries on a nightly schedule. Supports `.pptx`, `.docx`, and images, preserving the full folder hierarchy.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│  Azure Function App  (Timer: 2:00 AM UTC daily)                     │
│                                                                     │
│  SyncEngine                                                         │
│    ├── SharePointClient  ──→  Microsoft Graph API                   │
│    │     └── Recursive folder + file scan (etag-based delta)        │
│    ├── TemplafyClient    ──→  Templafy Public API v2                │
│    │     ├── Folder: GET/POST/PATCH children                        │
│    │     └── Asset:  POST (upload) / PATCH (replace) / DELETE       │
│    └── SyncStateStore   ──→  Azure Blob Storage (JSON)              │
│          └── SP item ID → Templafy asset ID + etag                  │
│                                                                     │
│  Secrets from Azure Key Vault (via managed identity)               │
└─────────────────────────────────────────────────────────────────────┘
```

### Library mapping

| File type | Templafy library |
|-----------|-----------------|
| `.pptx`   | presentations   |
| `.docx`   | documents       |
| `.png` `.jpg` `.jpeg` `.gif` `.svg` `.webp` `.tiff` `.bmp` | images |

Folder structure is mirrored exactly. A file at `Marketing/Q1/Deck.pptx` in SharePoint becomes `Marketing → Q1 → Deck.pptx` in the Templafy Presentations library.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| Azure subscription | For Function App, Storage, Key Vault |
| Azure AD App Registration | Needs `Sites.Read.All` + `Files.Read.All` MS Graph permissions (application, not delegated) |
| Templafy API key | Scope: `library.readwrite`. Create in Templafy Admin → API Keys |
| Templafy Space ID | Admin → Libraries → space dropdown → Copy Space ID |

---

## Setup

### 1. Azure AD App Registration (for SharePoint)

```bash
# Create app registration
az ad app create --display-name "SharePoint-Templafy-Sync"

# Note the appId (client ID) and tenantId from the output
# Add a client secret:
az ad app credential reset --id <appId> --years 2

# Grant Graph API permissions (application):
#   Microsoft Graph → Sites.Read.All
#   Microsoft Graph → Files.Read.All
# Then grant admin consent in the Azure Portal.
```

### 2. Find SharePoint IDs

```bash
# Site ID
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{siteName}

# Drive ID (document library)
GET https://graph.microsoft.com/v1.0/sites/{siteId}/drives
```

### 3. Deploy Infrastructure

```bash
az group create -n rg-sp-templafy-sync -l eastus

az deployment group create \
  -g rg-sp-templafy-sync \
  -f infra/main.bicep \
  -p baseName=sptfsync \
     graphTenantId="<tenant-id>" \
     graphClientId="<client-id>" \
     graphClientSecret="<client-secret>" \
     sharepointSiteId="<site-id>" \
     sharepointDriveId="<drive-id>" \
     templafyTenantId="<your-tenant>" \
     templafyApiKey="<api-key>" \
     templafySpaceId="<space-id>"
```

### 4. Deploy Function Code

```bash
# Publish to Azure
func azure functionapp publish <function-app-name> --python
```

---

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Copy and fill in settings
cp local.settings.json.example local.settings.json

# Run locally (triggers on startup in dev)
func start
```

Set `DRY_RUN=true` to test without writing to Templafy.

---

## Sync Behaviour

| Scenario | Action |
|---|---|
| File in SharePoint, not in state | **Upload** to Templafy, record in state |
| File in state, etag changed | **Replace** asset in Templafy, update etag |
| File in state, etag unchanged | **Skip** (no API calls) |
| File in state, not in SharePoint | **Delete** asset from Templafy, remove from state |

Change detection is etag-based — no file content hashing, minimal Graph API cost.

---

## Operational Notes

- **Schedule**: `0 0 2 * * *` (2:00 AM UTC). Edit `function.json` to change.
- **Timeout**: 30 minutes (`host.json`). For very large libraries, consider splitting root folders across multiple functions.
- **State file**: `sync-state/sharepoint-templafy-state.json` in your storage account. Back this up before manual Templafy changes.
- **Rate limits**: Templafy API has rate limits. The client retries on 429 with `Retry-After`. For initial sync of large libraries, set `MAX_FILE_SIZE_MB` conservatively.
- **First run**: All files will be uploaded. This may take time for large libraries. Consider running the first sync manually during off-hours.
- **Logs**: All runs log to Application Insights. Query: `traces | where message contains "Sync complete"`.

---

## Extending

**Add more file types**: Update `LIBRARY_MAP` in `templafy_client.py` and `supported_extensions` in `config.py`.

**Multiple SharePoint sites**: Deploy one Function App per site, each with its own state blob.

**Webhooks instead of polling**: Replace the timer trigger with an Event Grid trigger listening to SharePoint `driveItem` change notifications for near-real-time sync.
