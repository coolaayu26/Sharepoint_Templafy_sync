"""
Sync Engine — orchestrates the full SharePoint → Templafy sync cycle.

Sync flow per run:
  1. Load previous state from blob storage
  2. Scan SharePoint for all folders + supported files (recursive)
  3. For each file:
     a. Not in state → UPLOAD (create asset)
     b. In state, etag changed → REPLACE (update asset)
     c. In state, etag unchanged → SKIP
  4. Files in state but not in SP scan → DELETE asset from Templafy
  5. Persist updated state
"""

import logging
from dataclasses import dataclass, field

from .config import SyncConfig
from .sharepoint_client import SharePointClient, SPFile
from .templafy_client import TemplafyClient
from .state_store import SyncStateStore

logger = logging.getLogger(__name__)

LIBRARY_ROOTS: dict[str, str] = {}  # populated per-sync from Templafy


@dataclass
class SyncResult:
    created: int = 0
    updated: int = 0
    deleted: int = 0
    skipped: int = 0
    errors: int = 0

    def as_dict(self) -> dict:
        return {
            "created": self.created,
            "updated": self.updated,
            "deleted": self.deleted,
            "skipped": self.skipped,
            "errors": self.errors,
        }


class SyncEngine:
    def __init__(self, config: SyncConfig):
        self._cfg = config
        # Build a SharePointClient per drive ID
        self._sp_clients: list = []  # list of SharePointClient
        self._sp_by_drive: dict = {}  # drive_id → SharePointClient
        for drive_id in config.sharepoint_drive_ids:
            client = SharePointClient(
                tenant_id=config.graph_tenant_id,
                client_id=config.graph_client_id,
                client_secret=config.graph_client_secret,
                site_id=config.sharepoint_site_id,
                drive_id=drive_id,
            )
            self._sp_clients.append(client)
            self._sp_by_drive[drive_id] = client

        self._tf = TemplafyClient(
            tenant_id=config.templafy_tenant_id,
            api_key=config.templafy_api_key,
            space_id=config.templafy_space_id,
        )
        self._state = SyncStateStore(
            account_url=config.state_storage_account_url,
            container_name=config.state_container_name,
            blob_name=config.state_blob_name,
        )

    # ------------------------------------------------------------------ #
    # Entry point                                                           #
    # ------------------------------------------------------------------ #

    def run(self) -> dict:
        result = SyncResult()

        # 1. Load state (acquires blob lease)
        self._state.load()

        try:
            return self._execute(result)
        except Exception:
            self._state.release_lease()
            raise

    def _execute(self, result: SyncResult) -> dict:
        # 2. Scan all SharePoint drives — auto-resolve drive name as folder prefix
        all_sp_files = []
        for sp_client in self._sp_clients:
            drive_name = sp_client.get_drive_name()
            is_primary = drive_name == self._cfg.sharepoint_primary_drive
            folder_prefix = None if is_primary else drive_name
            logger.info(f"Scanning drive '{drive_name}' "
                        f"(prefix: '{folder_prefix or 'none'}')...")
            _, drive_files = sp_client.list_all_items(
                root_path="/",
                supported_extensions=self._cfg.supported_extensions,
                max_file_size_mb=self._cfg.max_file_size_mb,
            )
            if folder_prefix:
                for f in drive_files:
                    f.folder_path = f"{folder_prefix}/{f.folder_path}".rstrip("/") if f.folder_path else folder_prefix
                    f.path = f"{folder_prefix}/{f.path}"
            all_sp_files.extend(drive_files)

        sp_files = all_sp_files
        logger.info(f"Total files to sync: {len(sp_files)}")

        # 3. Resolve Templafy root folder IDs per library
        library_roots = self._resolve_library_roots()

        # 4. Build folder path cache from state for this run
        folder_path_cache = self._state.folder_path_cache

        # Index current SP files by ID
        current_sp_file_ids = {f.id for f in sp_files}

        # 5. Process files
        for sp_file in sp_files:
            ext = "." + sp_file.name.rsplit(".", 1)[-1].lower() if "." in sp_file.name else ""
            library = self._tf.library_for_extension(ext, self._cfg.library_map or None)
            if not library:
                logger.debug(f"No library mapping for '{sp_file.name}', skipping.")
                result.skipped += 1
                continue

            root_folder_id = library_roots.get(library)
            if not root_folder_id:
                logger.warning(f"Could not resolve root folder for library '{library}'. Skipping.")
                result.errors += 1
                continue

            try:
                self._process_file(
                    sp_file=sp_file,
                    library=library,
                    root_folder_id=root_folder_id,
                    folder_path_cache=folder_path_cache,
                    result=result,
                )
            except Exception as e:
                logger.error(f"Error processing '{sp_file.path}': {e}", exc_info=True)
                result.errors += 1

        # 6. Delete assets removed from SharePoint
        tracked_ids = self._state.all_tracked_file_ids()
        removed_ids = tracked_ids - current_sp_file_ids

        for sp_id in removed_ids:
            entry = self._state.get_file(sp_id)
            if not entry:
                continue
            try:
                if not self._cfg.dry_run:
                    self._tf.delete_asset(entry["templafy_library"], entry["templafy_asset_id"])
                self._state.remove_file(sp_id)
                result.deleted += 1
                logger.info(f"Deleted asset for removed SP file: {entry['path']}")
            except Exception as e:
                logger.error(f"Error deleting asset for SP ID {sp_id}: {e}", exc_info=True)
                result.errors += 1

        # 7. Save state (also releases blob lease)
        if not self._cfg.dry_run:
            self._state.mark_dirty()
            self._state.save()
        else:
            self._state.release_lease()
            logger.info("[DRY RUN] State not saved.")

        return result.as_dict()

    # ------------------------------------------------------------------ #
    # Per-file processing                                                   #
    # ------------------------------------------------------------------ #

    def _process_file(
        self,
        sp_file: SPFile,
        library: str,
        root_folder_id: str,
        folder_path_cache: dict[str, str],
        result: SyncResult,
    ) -> None:
        existing = self._state.get_file(sp_file.id)

        if existing:
            # Check for changes via etag
            if existing["etag"] == sp_file.etag:
                logger.debug(f"Unchanged: {sp_file.path}")
                result.skipped += 1
                return

            # File changed — download and replace
            logger.info(f"Updating: {sp_file.path}")
            if not self._cfg.dry_run:
                sp_client = self._sp_by_drive.get(sp_file.drive_id, self._sp_clients[0])
                content = sp_client.download_file(sp_file)
                self._tf.replace_asset(
                    library_type=library,
                    asset_id=existing["templafy_asset_id"],
                    filename=sp_file.name,
                    content=content,
                )
                self._state.set_file(
                    sp_file_id=sp_file.id,
                    templafy_asset_id=existing["templafy_asset_id"],
                    templafy_library=library,
                    etag=sp_file.etag,
                    path=sp_file.path,
                )
            result.updated += 1

        else:
            # New file — ensure folder exists, then upload
            logger.info(f"Uploading new: {sp_file.path}")
            folder_id = self._ensure_templafy_folder(
                library=library,
                sp_folder_path=sp_file.folder_path,
                root_folder_id=root_folder_id,
                folder_path_cache=folder_path_cache,
            )

            if not self._cfg.dry_run:
                sp_client = self._sp_by_drive.get(sp_file.drive_id, self._sp_clients[0])
                content = sp_client.download_file(sp_file)
                asset_id = self._tf.upload_asset(
                    library_type=library,
                    folder_id=folder_id,
                    filename=sp_file.name,
                    content=content,
                    external_id=sp_file.id,
                )
                self._state.set_file(
                    sp_file_id=sp_file.id,
                    templafy_asset_id=asset_id,
                    templafy_library=library,
                    etag=sp_file.etag,
                    path=sp_file.path,
                )
            result.created += 1

    # ------------------------------------------------------------------ #
    # Folder resolution                                                     #
    # ------------------------------------------------------------------ #

    def _ensure_templafy_folder(
        self,
        library: str,
        sp_folder_path: str,
        root_folder_id: str,
        folder_path_cache: dict[str, str],
    ) -> str:
        """Return the Templafy folder ID for the given SP folder path, creating as needed."""
        if not sp_folder_path:
            return root_folder_id

        path_segments = [s for s in sp_folder_path.split("/") if s]
        folder_id = self._tf.ensure_folder_path(
            library_type=library,
            path_segments=path_segments,
            folder_path_cache=folder_path_cache,
        )
        self._state.mark_dirty()
        return folder_id

    def _resolve_library_roots(self) -> dict[str, str]:
        """Get the root folder ID for each library we'll be writing to."""
        needed_libraries = set()
        for ext in self._cfg.supported_extensions:
            lib = self._tf.library_for_extension(ext, self._cfg.library_map or None)
            if lib:
                needed_libraries.add(lib)

        roots = {}
        for lib in needed_libraries:
            try:
                roots[lib] = self._tf.get_root_folder_id(lib)
                logger.info(f"Templafy root folder for '{lib}': {roots[lib]}")
            except Exception as e:
                logger.error(f"Could not resolve root folder for library '{lib}': {e}")

        return roots
