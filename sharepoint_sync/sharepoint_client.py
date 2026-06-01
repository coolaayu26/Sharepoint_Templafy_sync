"""
SharePoint client using Microsoft Graph API.
Handles auth (client credentials), recursive folder traversal,
and streaming file downloads.
"""

import logging
import time
from dataclasses import dataclass
from typing import Generator, Optional
from urllib.parse import quote

import requests
from msal import ConfidentialClientApplication

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]


@dataclass
class SPFolder:
    id: str
    name: str
    path: str           # Relative path from sync root, e.g. "/Marketing/Q1"
    parent_path: str    # Parent's relative path


@dataclass
class SPFile:
    id: str
    name: str
    path: str           # Relative path from sync root, e.g. "/Marketing/Q1/deck.pptx"
    folder_path: str    # Parent folder relative path
    size_bytes: int
    last_modified: str  # ISO 8601
    etag: str           # Used for change detection
    download_url: str   # @microsoft.graph.downloadUrl (pre-auth, short-lived)
    mime_type: str
    drive_id: str = ""  # Drive ID this file belongs to (for fallback download)


class SharePointClient:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str,
                 site_id: str, drive_id: str):
        self._site_id = site_id
        self._drive_id = drive_id
        self._session = requests.Session()
        self._token_expiry = 0

        self._msal_app = ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
        )
        self._refresh_token()

    # ------------------------------------------------------------------ #
    # Auth                                                                  #
    # ------------------------------------------------------------------ #

    def _refresh_token(self) -> None:
        result = self._msal_app.acquire_token_for_client(scopes=GRAPH_SCOPES)
        if "access_token" not in result:
            raise RuntimeError(f"MSAL token acquisition failed: {result.get('error_description')}")
        self._session.headers.update({"Authorization": f"Bearer {result['access_token']}"})
        self._token_expiry = time.time() + result.get("expires_in", 3600) - 60

    def _ensure_token(self) -> None:
        if time.time() >= self._token_expiry:
            self._refresh_token()

    # ------------------------------------------------------------------ #
    # Graph requests                                                        #
    # ------------------------------------------------------------------ #

    def _get(self, url: str, **kwargs) -> dict:
        self._ensure_token()
        resp = self._session.get(url, **kwargs)
        resp.raise_for_status()
        return resp.json()

    def _get_paged(self, url: str) -> Generator[dict, None, None]:
        """Follow @odata.nextLink pagination automatically."""
        next_url: Optional[str] = url
        while next_url:
            data = self._get(next_url)
            yield from data.get("value", [])
            next_url = data.get("@odata.nextLink")

    # ------------------------------------------------------------------ #
    # Drive item helpers                                                    #
    # ------------------------------------------------------------------ #

    def _drive_url(self, path: str = "") -> str:
        base = f"{GRAPH_BASE}/sites/{self._site_id}/drives/{self._drive_id}"
        return f"{base}{path}"

    def _item_children_url(self, item_id: str) -> str:
        return self._drive_url(f"/items/{item_id}/children?$select=id,name,folder,file,size,lastModifiedDateTime,eTag,@microsoft.graph.downloadUrl,file")

    # ------------------------------------------------------------------ #
    # Public API                                                            #
    # ------------------------------------------------------------------ #

    def get_drive_name(self) -> str:
        """GET /drives/{driveId} — returns the drive's display name."""
        self._ensure_token()
        url = f"{GRAPH_BASE}/drives/{self._drive_id}"
        resp = self._session.get(url)
        resp.raise_for_status()
        return resp.json().get("name", self._drive_id)

    def get_root_item_id(self, root_path: str) -> str:
        """Resolve the root folder path to a Drive item ID."""
        if root_path in ("", "/"):
            data = self._get(self._drive_url("/root"))
        else:
            # Encode path segments but preserve slashes
            clean = root_path.strip("/")
            encoded = quote(clean, safe="/")
            data = self._get(self._drive_url(f"/root:/{encoded}"))
        return data["id"]

    def list_all_items(
        self,
        root_path: str,
        supported_extensions: tuple,
        max_file_size_mb: int,
    ) -> tuple[list[SPFolder], list[SPFile]]:
        """
        Recursively enumerate all folders and supported files under root_path.
        Returns (folders, files) with relative paths from root_path.
        """
        root_id = self.get_root_item_id(root_path)
        folders: list[SPFolder] = []
        files: list[SPFile] = []

        self._recurse(
            item_id=root_id,
            current_path="",
            parent_path="",
            folders=folders,
            files=files,
            supported_extensions=supported_extensions,
            max_file_size_bytes=max_file_size_mb * 1024 * 1024,
        )

        logger.info(f"SharePoint scan complete: {len(folders)} folders, {len(files)} files.")
        return folders, files

    def _recurse(
        self,
        item_id: str,
        current_path: str,
        parent_path: str,
        folders: list[SPFolder],
        files: list[SPFile],
        supported_extensions: tuple,
        max_file_size_bytes: int,
    ) -> None:
        url = self._item_children_url(item_id)
        for item in self._get_paged(url):
            name = item["name"]

            if "folder" in item:
                folder_path = f"{current_path}/{name}".lstrip("/")
                folders.append(SPFolder(
                    id=item["id"],
                    name=name,
                    path=folder_path,
                    parent_path=current_path.lstrip("/"),
                ))
                self._recurse(
                    item_id=item["id"],
                    current_path=f"{current_path}/{name}",
                    parent_path=current_path,
                    folders=folders,
                    files=files,
                    supported_extensions=supported_extensions,
                    max_file_size_bytes=max_file_size_bytes,
                )

            elif "file" in item:
                ext = "." + name.rsplit(".", 1)[-1].lower() if "." in name else ""
                size = item.get("size", 0)

                if ext not in supported_extensions:
                    logger.debug(f"Skipping unsupported type: {name}")
                    continue

                if size > max_file_size_bytes:
                    logger.warning(f"Skipping oversized file ({size/1024/1024:.1f} MB): {name}")
                    continue

                file_path = f"{current_path}/{name}".lstrip("/")
                files.append(SPFile(
                    id=item["id"],
                    name=name,
                    path=file_path,
                    folder_path=current_path.lstrip("/"),
                    size_bytes=size,
                    last_modified=item.get("lastModifiedDateTime", ""),
                    etag=item.get("eTag", ""),
                    download_url=item.get("@microsoft.graph.downloadUrl", ""),
                    mime_type=item.get("file", {}).get("mimeType", "application/octet-stream"),
                    drive_id=self._drive_id,
                ))

    def download_file(self, file: SPFile) -> bytes:
        """Download file content. Uses pre-auth URL directly (no token needed)."""
        if not file.download_url:
            # Fall back to Graph endpoint using the file's own drive ID
            self._ensure_token()
            drive_id = file.drive_id or self._drive_id
            url = f"{GRAPH_BASE}/drives/{drive_id}/items/{file.id}/content"
            resp = self._session.get(url, allow_redirects=True)
        else:
            resp = requests.get(file.download_url)

        resp.raise_for_status()
        return resp.content
