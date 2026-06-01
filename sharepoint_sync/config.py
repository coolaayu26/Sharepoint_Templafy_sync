"""
Configuration loaded from environment variables.
Secrets should be stored in Azure Key Vault and surfaced
as App Settings via Key Vault references.
"""

import json
import os
from dataclasses import dataclass, field


def _parse_library_map(raw: str) -> dict:
    """
    Parse LIBRARY_MAP env var into a dict.
    Expected format: JSON object e.g. {".pptx": "slides", ".png": "images"}
    Returns empty dict if not set — templafy_client.py will use its built-in defaults.
    """
    if not raw:
        return {}
    try:
        result = json.loads(raw)
        if not isinstance(result, dict):
            raise ValueError("LIBRARY_MAP must be a JSON object")
        # Normalise keys to lowercase with leading dot
        return {
            (k if k.startswith(".") else f".{k}").lower(): v
            for k, v in result.items()
        }
    except (json.JSONDecodeError, ValueError) as e:
        raise EnvironmentError(f"LIBRARY_MAP is invalid: {e}")


@dataclass
class SyncConfig:
    # --- SharePoint / Graph API ---
    graph_tenant_id: str
    graph_client_id: str
    graph_client_secret: str
    sharepoint_site_id: str          # e.g. "contoso.sharepoint.com,<site-id>,<web-id>"
    sharepoint_drive_ids: list       # List of drive IDs to sync
    sharepoint_primary_drive: str    # Drive name treated as primary (no folder prefix)

    # --- Templafy ---
    templafy_tenant_id: str          # Your Templafy tenant subdomain
    templafy_api_key: str            # library.readwrite scoped API key
    templafy_space_id: str           # Target Space ID in Templafy

    # --- State store (Azure Blob Storage) ---
    state_storage_account_url: str   # e.g. "https://mystorageaccount.blob.core.windows.net"
    state_container_name: str = "sync-state"
    state_blob_name: str = "sharepoint-templafy-state.json"

    # --- Sync behaviour ---
    supported_extensions: tuple = field(default_factory=lambda: (
        ".pptx",
        ".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".tiff", ".bmp"
    ))
    dry_run: bool = False
    max_file_size_mb: int = 100
    # Optional override for extension → Templafy library mapping.
    # If None, templafy_client.py uses its built-in defaults.
    # Format: JSON object e.g. {".pptx": "slides", ".png": "images", ".docx": "documents"}
    library_map: dict = field(default_factory=dict)

    @classmethod
    def from_environment(cls) -> "SyncConfig":
        def require(key: str) -> str:
            val = os.environ.get(key)
            if not val:
                raise EnvironmentError(f"Required environment variable '{key}' is not set.")
            return val

        # SHAREPOINT_DRIVES: JSON array of drive ID strings
        # e.g. ["driveId1", "driveId2"]
        drives_raw = require("SHAREPOINT_DRIVES")
        try:
            drive_ids = json.loads(drives_raw)
            if not isinstance(drive_ids, list):
                raise ValueError("SHAREPOINT_DRIVES must be a JSON array")
        except (json.JSONDecodeError, ValueError) as e:
            raise EnvironmentError(f"SHAREPOINT_DRIVES is invalid: {e}")

        return cls(
            graph_tenant_id=require("GRAPH_TENANT_ID"),
            graph_client_id=require("GRAPH_CLIENT_ID"),
            graph_client_secret=require("GRAPH_CLIENT_SECRET"),
            sharepoint_site_id=require("SHAREPOINT_SITE_ID"),
            sharepoint_drive_ids=drive_ids,
            sharepoint_primary_drive=os.environ.get("SHAREPOINT_PRIMARY_DRIVE", "Documents"),
            templafy_tenant_id=require("TEMPLAFY_TENANT_ID"),
            templafy_api_key=require("TEMPLAFY_API_KEY"),
            templafy_space_id=require("TEMPLAFY_SPACE_ID"),
            state_storage_account_url=require("STATE_STORAGE_ACCOUNT_URL"),
            state_container_name=os.environ.get("STATE_CONTAINER_NAME", "sync-state"),
            state_blob_name=os.environ.get("STATE_BLOB_NAME", "sharepoint-templafy-state.json"),
            dry_run=os.environ.get("DRY_RUN", "false").lower() == "true",
            max_file_size_mb=int(os.environ.get("MAX_FILE_SIZE_MB", "100")),
            library_map=_parse_library_map(os.environ.get("LIBRARY_MAP", "")),
        )
