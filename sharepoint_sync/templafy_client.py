"""
Templafy Public API v3 client.

Base URL per the docs:
  https://{tenantId}.api.templafy.com/v3/libraries/{spaceId}/{libraryType}/...

Auth: API key passed as header  "Authorization: Bearer {key}"

Library types used in this sync:
  - slides   (.pptx) — individual slides inserted into presentations
  - images   (.png, .jpg, .gif, .svg, .webp, .tiff, .bmp)

Key endpoint patterns:
  GET    /libraries/{spaceId}/{libraryType}                            → library info incl. rootFolderId
  GET    /libraries/{spaceId}/{libraryType}/folders/{folderId}/folders → list child folders
  POST   /libraries/{spaceId}/{libraryType}/folders/{folderId}/folders → create folder
  POST   /libraries/{spaceId}/{libraryType}/folders/{folderId}/assets  → upload new asset
  PATCH  /libraries/{spaceId}/{libraryType}/assets/{assetId}          → replace/update asset
  DELETE /libraries/{spaceId}/{libraryType}/assets/{assetId}          → delete asset

IMPORTANT: spaceId and libraryType appear in EVERY route.
Each library type is its own isolated namespace.
"""

import logging
import time
from dataclasses import dataclass
from typing import Optional

import requests

logger = logging.getLogger(__name__)

# File extension → Templafy library type segment
EXTENSION_TO_LIBRARY: dict[str, str] = {
    ".pptx": "slides",
    ".png":  "images",
    ".jpg":  "images",
    ".jpeg": "images",
    ".gif":  "images",
    ".svg":  "images",
    ".webp": "images",
    ".tiff": "images",
    ".bmp":  "images",
}

MAX_RETRIES = 3
RETRY_BACKOFF = 2.0


@dataclass
class TemplafyFolder:
    id: str
    name: str
    parent_id: Optional[str]


class TemplafyClient:
    def __init__(self, tenant_id: str, api_key: str, space_id: str):
        self._space_id = space_id
        # Base up to and including spaceId — libraryType is always appended per call
        self._base = f"https://{tenant_id}.api.templafy.com/v3/libraries/{space_id}"
        self._session = requests.Session()
        self._session.headers.update({
            "Authorization": f"Bearer {api_key}",
        })
        # Cache: library_type → root folder ID (resolved once per sync run)
        self._root_folder_cache: dict[str, str] = {}
        # Cache: "library_type:path/to/folder" → templafy folder ID
        # Shared with SyncStateStore so it persists across runs

    # ------------------------------------------------------------------ #
    # Internal helpers                                                      #
    # ------------------------------------------------------------------ #

    def _lib_url(self, library_type: str, *segments: str) -> str:
        """Build a full URL: base/libraryType[/seg1/seg2/...]"""
        parts = [self._base, library_type] + [s.strip("/") for s in segments]
        return "/".join(p.strip("/") for p in parts)

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        delay = RETRY_BACKOFF
        last_exc: Optional[Exception] = None

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = self._session.request(method, url, timeout=60, **kwargs)
                if resp.status_code == 429:
                    retry_after = int(resp.headers.get("Retry-After", delay))
                    logger.warning(f"Rate limited. Waiting {retry_after}s (attempt {attempt}).")
                    time.sleep(retry_after)
                    continue
                resp.raise_for_status()
                return resp
            except requests.HTTPError as e:
                if attempt == MAX_RETRIES or (e.response is not None and e.response.status_code < 500):
                    raise
                logger.warning(f"HTTP {e.response.status_code} on attempt {attempt}. Retrying in {delay}s.")
                time.sleep(delay)
                delay *= 2
                last_exc = e
            except requests.RequestException as e:
                if attempt == MAX_RETRIES:
                    raise
                logger.warning(f"Request error on attempt {attempt}: {e}. Retrying in {delay}s.")
                time.sleep(delay)
                delay *= 2
                last_exc = e

        raise last_exc  # type: ignore

    # ------------------------------------------------------------------ #
    # Library info — get root folder ID                                     #
    # ------------------------------------------------------------------ #

    def get_root_folder_id(self, library_type: str) -> str:
        """
        GET /libraries/{spaceId}/{libraryType}
        Returns library metadata including rootFolderId.
        Cached per library type for the lifetime of this client instance.
        """
        if library_type in self._root_folder_cache:
            return self._root_folder_cache[library_type]

        url = self._lib_url(library_type)
        resp = self._request("GET", url)
        data = resp.json()

        root_folder_id = data.get("rootFolderId")
        if not root_folder_id:
            raise RuntimeError(
                f"No rootFolderId returned for library '{library_type}'. "
                f"Response: {data}"
            )

        self._root_folder_cache[library_type] = root_folder_id
        logger.info(f"Root folder for '{library_type}': {root_folder_id}")
        return root_folder_id

    # ------------------------------------------------------------------ #
    # Folders                                                               #
    # ------------------------------------------------------------------ #

    def list_child_folders(self, library_type: str, parent_folder_id: str) -> list[TemplafyFolder]:
        """
        GET /libraries/{spaceId}/{libraryType}/folders/{folderId}/folders
        Lists immediate child folders of the given folder.
        """
        url = self._lib_url(library_type, "folders", parent_folder_id, "folders")
        resp = self._request("GET", url)
        return [
            TemplafyFolder(
                id=f["id"],
                name=f["name"],
                parent_id=parent_folder_id,
            )
            for f in resp.json()
        ]

    def create_folder(self, library_type: str, parent_folder_id: str, name: str) -> TemplafyFolder:
        """
        POST /libraries/{spaceId}/{libraryType}/folders/{folderId}/folders
        Creates a child folder inside the given parent.
        On 409 Conflict (already exists), fetches the existing folder ID instead.
        """
        url = self._lib_url(library_type, "folders", parent_folder_id, "folders")
        try:
            resp = self._request("POST", url, json={"name": name})
        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code == 409:
                # Folder already exists — find it by listing children
                logger.info(f"[Templafy] Folder '{name}' already exists, resolving ID.")
                children = self.list_child_folders(library_type, parent_folder_id)
                for child in children:
                    if child.name == name:
                        return child
                raise RuntimeError(f"409 on create but folder '{name}' not found in children.")
            raise

        data = resp.json()
        # API may return a plain string ID or a JSON object
        if isinstance(data, str):
            folder_id = data
        else:
            folder_id = data["id"]
        logger.info(f"[Templafy] Created folder '{name}' in {library_type} under {parent_folder_id}")
        return TemplafyFolder(id=folder_id, name=name, parent_id=parent_folder_id)

    def ensure_folder_path(
        self,
        library_type: str,
        path_segments: list[str],
        folder_path_cache: dict[str, str],
    ) -> str:
        """
        Walk path segments from the library root, creating folders that don't exist.
        Returns the Templafy folder ID of the deepest folder.

        folder_path_cache keys: "{library_type}:{path/to/folder}"
        This cache is passed in from the state store so it persists across runs.

        NOTE: folder structure is per-library-type. "Marketing/Q1" in 'presentations'
        is a different set of folders than "Marketing/Q1" in 'documents'.
        Both must be created independently.
        """
        root_id = self.get_root_folder_id(library_type)

        if not path_segments:
            return root_id

        current_id = root_id
        current_path = ""

        for segment in path_segments:
            current_path = f"{current_path}/{segment}".lstrip("/")
            cache_key = f"{library_type}:{current_path}"

            if cache_key in folder_path_cache:
                current_id = folder_path_cache[cache_key]
                continue

            # Not cached — check if it exists in Templafy
            children = self.list_child_folders(library_type, current_id)
            existing = {f.name: f.id for f in children}

            if segment in existing:
                current_id = existing[segment]
            else:
                folder = self.create_folder(library_type, current_id, segment)
                current_id = folder.id

            folder_path_cache[cache_key] = current_id

        return current_id

    # ------------------------------------------------------------------ #
    # Assets                                                                #
    # ------------------------------------------------------------------ #

    def list_folder_assets(self, library_type: str, folder_id: str) -> list[dict]:
        """
        GET /libraries/{spaceId}/{libraryType}/folders/{folderId}/assets
        Lists assets in a folder. Returns list of {id, name, ...} dicts.
        """
        url = self._lib_url(library_type, "folders", folder_id, "assets")
        resp = self._request("GET", url)
        data = resp.json()
        if isinstance(data, list):
            return data
        return data.get("items", data.get("value", []))

    def upload_asset(
        self,
        library_type: str,
        folder_id: str,
        filename: str,
        content: bytes,
        external_id: Optional[str] = None,
    ) -> str:
        """
        POST /libraries/{spaceId}/{libraryType}/folders/{folderId}/assets
        Uploads a new asset. Returns the Templafy asset ID.
        On 409 Conflict (already exists), fetches the existing asset ID instead.
        external_id stores the SharePoint item ID for future reference.
        """
        url = self._lib_url(library_type, "folders", folder_id, "assets")
        files = {"file": (filename, content)}
        data = {}
        if external_id:
            data["externalData"] = f'{{"sharepoint_id": "{external_id}"}}'

        try:
            resp = self._request("POST", url, files=files, data=data or None)
        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code == 409:
                # Asset already exists — find its real ID
                logger.warning(f"[Templafy] Asset '{filename}' already exists, resolving existing ID.")
                assets = self.list_folder_assets(library_type, folder_id)
                # Try matching by SharePoint ID in externalData first (most reliable)
                import json as _json
                for asset in assets:
                    if not isinstance(asset, dict):
                        continue
                    ext_data = asset.get("externalData", "")
                    if ext_data and external_id:
                        try:
                            parsed = _json.loads(ext_data)
                            if parsed.get("sharepoint_id") == external_id:
                                return str(asset["id"])
                        except Exception:
                            pass
                # Fallback: match by name without extension
                name_without_ext = filename.rsplit(".", 1)[0] if "." in filename else filename
                for asset in assets:
                    if isinstance(asset, dict) and asset.get("name") == name_without_ext:
                        return str(asset["id"])
                logger.warning(f"[Templafy] Could not find existing asset '{filename}', using placeholder.")
                return f"existing-{folder_id}-{filename}"
            raise

        data_resp = resp.json()
        if isinstance(data_resp, str):
            asset_id = data_resp
        else:
            asset_id = data_resp["id"]
        logger.info(f"[Templafy] Uploaded '{filename}' → {library_type}/{folder_id} → assetId={asset_id}")
        return asset_id

    def replace_asset(
        self,
        library_type: str,
        asset_id: str,
        filename: str,
        content: bytes,
    ) -> None:
        """
        PATCH /libraries/{spaceId}/{libraryType}/assets/{assetId}
        Replaces the file content of an existing asset (multipart form).
        """
        url = self._lib_url(library_type, "assets", asset_id)
        files = {"file": (filename, content)}
        self._request("PATCH", url, files=files)
        logger.info(f"[Templafy] Replaced asset {asset_id} ('{filename}') in '{library_type}'")

    def delete_asset(self, library_type: str, asset_id: str) -> None:
        """
        DELETE /libraries/{spaceId}/{libraryType}/assets/{assetId}
        """
        url = self._lib_url(library_type, "assets", asset_id)
        self._request("DELETE", url)
        logger.info(f"[Templafy] Deleted asset {asset_id} from '{library_type}'")

    # ------------------------------------------------------------------ #
    # Static helpers                                                        #
    # ------------------------------------------------------------------ #

    @staticmethod
    def library_for_extension(ext: str, library_map: Optional[dict] = None) -> Optional[str]:
        """
        Map a file extension to the Templafy library type string.
        Uses library_map override if provided, falls back to built-in EXTENSION_TO_LIBRARY.
        """
        mapping = library_map if library_map else EXTENSION_TO_LIBRARY
        return mapping.get(ext.lower())
