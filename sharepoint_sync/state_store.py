"""
Sync state store backed by Azure Blob Storage.

Tracks the mapping between SharePoint item IDs and Templafy asset/folder IDs,
plus the last-seen etag for change detection.

State blob schema (JSON):
{
  "files": {
    "<sharepoint_file_id>": {
      "templafy_asset_id": "...",
      "templafy_library": "presentations",
      "etag": "...",
      "path": "Marketing/Q1/deck.pptx"
    }
  },
  "folders": {
    "<sharepoint_folder_id>": {
      "templafy_folder_id": "...",
      "templafy_library": "presentations",   # one entry per library this folder exists in
      "path": "Marketing/Q1"
    }
  },
  "templafy_folder_path_cache": {
    "presentations:Marketing/Q1": "<templafy_folder_id>"
  }
}

Locking:
  A 60-second blob lease is acquired before load() and released after save().
  This prevents two concurrent function instances from clobbering each other's
  state writes. If the lease cannot be acquired within LEASE_ACQUIRE_TIMEOUT
  seconds, the run aborts rather than proceeding without a lock.
"""

import json
import logging
import time
from copy import deepcopy
from typing import Optional

from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobClient, BlobServiceClient, BlobLeaseClient

logger = logging.getLogger(__name__)

LEASE_DURATION_SECONDS = 60       # Max Azure allows: 60. Renewed automatically if run > 60s.
LEASE_ACQUIRE_TIMEOUT = 30        # How long to wait for a lock before aborting.
LEASE_RETRY_INTERVAL = 5          # Seconds between acquire attempts.


class SyncStateStore:
    def __init__(self, account_url: str, container_name: str, blob_name: str):
        self._blob_client: BlobClient = BlobServiceClient(
            account_url=account_url,
            credential=DefaultAzureCredential(),
        ).get_blob_client(container=container_name, blob=blob_name)
        self._state: dict = {"files": {}, "folders": {}, "templafy_folder_path_cache": {}}
        self._dirty = False
        self._lease: Optional[BlobLeaseClient] = None

    # ------------------------------------------------------------------ #
    # Lease management                                                      #
    # ------------------------------------------------------------------ #

    def _acquire_lease(self) -> None:
        """
        Acquire an exclusive 60-second lease on the state blob.
        Retries every LEASE_RETRY_INTERVAL seconds up to LEASE_ACQUIRE_TIMEOUT.
        Raises RuntimeError if the lease cannot be acquired in time.
        """
        deadline = time.monotonic() + LEASE_ACQUIRE_TIMEOUT
        attempt = 0
        while time.monotonic() < deadline:
            try:
                self._lease = self._blob_client.acquire_lease(
                    lease_duration=LEASE_DURATION_SECONDS
                )
                logger.info(f"Blob lease acquired (id={self._lease.id}).")
                return
            except Exception as e:
                attempt += 1
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    break
                logger.warning(
                    f"Could not acquire blob lease (attempt {attempt}): {e}. "
                    f"Retrying in {LEASE_RETRY_INTERVAL}s "
                    f"({remaining:.0f}s remaining)..."
                )
                time.sleep(LEASE_RETRY_INTERVAL)

        raise RuntimeError(
            f"Failed to acquire blob lease after {LEASE_ACQUIRE_TIMEOUT}s. "
            "Another function instance may be running. Aborting to prevent state corruption."
        )

    def _release_lease(self) -> None:
        """Release the blob lease if one is held."""
        if self._lease:
            try:
                self._lease.release()
                logger.info("Blob lease released.")
            except Exception as e:
                logger.warning(f"Failed to release blob lease: {e}")
            finally:
                self._lease = None

    def _renew_lease(self) -> None:
        """Renew the lease before it expires (call if processing takes > 60s)."""
        if self._lease:
            try:
                self._lease.renew()
                logger.debug("Blob lease renewed.")
            except Exception as e:
                logger.warning(f"Failed to renew blob lease: {e}")

    # ------------------------------------------------------------------ #
    # Load / save                                                           #
    # ------------------------------------------------------------------ #

    def load(self) -> None:
        """
        Acquire a blob lease, then load state from blob.
        The lease is held until save() or release_lease() is called.
        Initialises empty state if blob doesn't exist yet.
        """
        self._acquire_lease()
        try:
            data = self._blob_client.download_blob(
                lease=self._lease
            ).readall()
            self._state = json.loads(data)
            logger.info(
                f"Loaded sync state: "
                f"{len(self._state.get('files', {}))} files, "
                f"{len(self._state.get('folders', {}))} folders tracked."
            )
        except Exception as e:
            if "BlobNotFound" in str(e) or "404" in str(e):
                logger.info("No existing sync state found. Starting fresh.")
                self._state = {"files": {}, "folders": {}, "templafy_folder_path_cache": {}}
            else:
                self._release_lease()
                raise

    def save(self) -> None:
        """Persist state to blob (with active lease) and release the lease."""
        try:
            if not self._dirty:
                logger.info("State unchanged — skipping save.")
                return
            data = json.dumps(self._state, indent=2).encode("utf-8")
            self._blob_client.upload_blob(
                data,
                overwrite=True,
                lease=self._lease,
            )
            self._dirty = False
            logger.info("Sync state saved to blob.")
        finally:
            self._release_lease()

    def release_lease(self) -> None:
        """
        Explicit lease release for error paths where save() is never called.
        Safe to call even if no lease is held.
        """
        self._release_lease()

    # ------------------------------------------------------------------ #
    # File tracking                                                         #
    # ------------------------------------------------------------------ #

    def get_file(self, sp_file_id: str) -> Optional[dict]:
        return self._state["files"].get(sp_file_id)

    def set_file(self, sp_file_id: str, templafy_asset_id: str,
                 templafy_library: str, etag: str, path: str) -> None:
        self._state["files"][sp_file_id] = {
            "templafy_asset_id": templafy_asset_id,
            "templafy_library": templafy_library,
            "etag": etag,
            "path": path,
        }
        self._dirty = True

    def remove_file(self, sp_file_id: str) -> None:
        if sp_file_id in self._state["files"]:
            del self._state["files"][sp_file_id]
            self._dirty = True

    def all_tracked_file_ids(self) -> set[str]:
        return set(self._state["files"].keys())

    # ------------------------------------------------------------------ #
    # Folder tracking                                                       #
    # ------------------------------------------------------------------ #

    def get_folder(self, sp_folder_id: str) -> Optional[dict]:
        return self._state["folders"].get(sp_folder_id)

    def set_folder(self, sp_folder_id: str, templafy_folder_id: str,
                   templafy_library: str, path: str) -> None:
        entry = self._state["folders"].setdefault(sp_folder_id, {})
        entry.update({
            "templafy_folder_id": templafy_folder_id,
            "templafy_library": templafy_library,
            "path": path,
        })
        self._dirty = True

    def all_tracked_folder_ids(self) -> set[str]:
        return set(self._state["folders"].keys())

    # ------------------------------------------------------------------ #
    # Folder path cache (avoids redundant list calls to Templafy)           #
    # ------------------------------------------------------------------ #

    @property
    def folder_path_cache(self) -> dict[str, str]:
        """Mutable reference to the folder path cache. Mark dirty after writes."""
        return self._state["templafy_folder_path_cache"]

    def mark_dirty(self) -> None:
        self._dirty = True
