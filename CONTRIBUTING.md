# Contributing

## Local Development

### Prerequisites

- Python 3.11
- [Azure Functions Core Tools v4](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local)
- An Azure subscription with the resources from the README set up

### Setup

```bash
# Clone the repo
git clone https://github.com/your-org/sharepoint-templafy-sync
cd sharepoint-templafy-sync

# Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Copy and fill in local settings
cp local.settings.json.example local.settings.json
# Edit local.settings.json with your values
```

### Running Locally

```bash
func start
```

The function will trigger immediately on startup in local mode. Set `DRY_RUN=true` in `local.settings.json` to test without writing to Templafy.

### Running Tests

```bash
pip install pytest pytest-cov
python -m pytest tests/ -v
```

With coverage:

```bash
python -m pytest tests/ -v --cov=sharepoint_sync --cov-report=term-missing
```

---

## Making Changes

### Adding a New File Type

No code change needed. Add the extension to the `LIBRARY_MAP` environment variable in Azure Portal:

```json
{".pptx": "slides", ".png": "images", ".docx": "documents"}
```

If `LIBRARY_MAP` is not set, the built-in defaults are used.

### Adding a New SharePoint Drive

Add the drive ID to the `SHAREPOINT_DRIVES` environment variable (JSON array) in Azure Portal. If the drive should be treated as primary (no folder prefix in Templafy), update `SHAREPOINT_PRIMARY_DRIVE` to match its name.

### Changing the Sync Schedule

Edit `function.json`:

```json
{
  "schedule": "0 0 2 * * *"
}
```

Format: `{second} {minute} {hour} {day} {month} {day-of-week}`. The above runs at 2:00 AM UTC daily.

---

## Deploying

```bash
cd /path/to/project/root
func azure functionapp publish <function-app-name> --python
```

Always deploy from the project root (the directory containing `requirements.txt` and `host.json`).

---

## Code Style

- Formatter: [black](https://black.readthedocs.io/) — `black sharepoint_sync/`
- Linter: [ruff](https://docs.astral.sh/ruff/) — `ruff check sharepoint_sync/`
- Type hints required on all new functions

---

## Secrets Management

- **Templafy API key** — static key stored in Key Vault. Update the `templafy-api-key` secret only if the key is revoked or you intentionally generate a replacement in Templafy Admin → API Keys. The Function App picks it up automatically via the Key Vault reference.
- **Graph API client secret** — rotate before expiry (default 2 years). Update `graph-client-secret` in Key Vault.
