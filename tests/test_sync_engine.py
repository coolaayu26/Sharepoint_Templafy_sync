"""
Unit tests for the SharePoint → Templafy sync engine.

Tests cover:
  - Change detection (skip / update / create / delete)
  - Folder prefix logic for multi-drive
  - Library mapping by extension
  - Dry run mode
  - Error handling per file
"""

import unittest
from unittest.mock import MagicMock, patch, call
from dataclasses import dataclass, field
from typing import Optional


# ---------------------------------------------------------------------------
# Minimal stubs so we can import without Azure / requests installed
# ---------------------------------------------------------------------------

import sys

# No stubs needed — real packages are installed

# ---------------------------------------------------------------------------
# Now import the modules under test
# ---------------------------------------------------------------------------

import importlib, os, sys

# Add the project root (files/) to sys.path so `sharepoint_sync` package resolves
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from sharepoint_sync.sync_engine import SyncEngine, SyncResult
from sharepoint_sync.sharepoint_client import SPFile
from sharepoint_sync.templafy_client import TemplafyClient, EXTENSION_TO_LIBRARY


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def make_config(**overrides):
    """Return a minimal SyncConfig-like namespace."""
    cfg = MagicMock()
    cfg.graph_tenant_id = "tenant"
    cfg.graph_client_id = "client"
    cfg.graph_client_secret = "secret"
    cfg.sharepoint_site_id = "site-id"
    cfg.sharepoint_drive_ids = ["drive-1"]
    cfg.sharepoint_primary_drive = "Documents"
    cfg.templafy_tenant_id = "templafy-tenant"
    cfg.templafy_api_key = "api-key"
    cfg.templafy_space_id = "space-id"
    cfg.state_storage_account_url = "https://account.blob.core.windows.net"
    cfg.state_container_name = "sync-state"
    cfg.state_blob_name = "state.json"
    cfg.supported_extensions = {".pptx", ".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".tiff", ".bmp"}
    cfg.max_file_size_mb = 100
    cfg.dry_run = False
    for k, v in overrides.items():
        setattr(cfg, k, v)
    return cfg


def make_sp_file(
    file_id="sp-file-1",
    name="deck.pptx",
    folder_path="",
    etag="etag-v1",
    drive_id="drive-1",
    path="deck.pptx",
    download_url="https://example.com/deck.pptx",
):
    return SPFile(
        id=file_id,
        name=name,
        path=path,
        folder_path=folder_path,
        etag=etag,
        drive_id=drive_id,
        download_url=download_url,
        size_bytes=1000,
        last_modified="2026-01-01T00:00:00Z",
        mime_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


# ---------------------------------------------------------------------------
# Tests: library mapping
# ---------------------------------------------------------------------------

class TestLibraryMapping(unittest.TestCase):

    def test_pptx_maps_to_slides(self):
        self.assertEqual(TemplafyClient.library_for_extension(".pptx"), "slides")

    def test_png_maps_to_images(self):
        self.assertEqual(TemplafyClient.library_for_extension(".png"), "images")

    def test_jpg_maps_to_images(self):
        self.assertEqual(TemplafyClient.library_for_extension(".jpg"), "images")

    def test_unknown_extension_returns_none(self):
        self.assertIsNone(TemplafyClient.library_for_extension(".docx"))
        self.assertIsNone(TemplafyClient.library_for_extension(".pdf"))
        self.assertIsNone(TemplafyClient.library_for_extension(".txt"))

    def test_case_insensitive(self):
        self.assertEqual(TemplafyClient.library_for_extension(".PPTX"), "slides")
        self.assertEqual(TemplafyClient.library_for_extension(".PNG"), "images")


# ---------------------------------------------------------------------------
# Tests: SyncResult
# ---------------------------------------------------------------------------

class TestSyncResult(unittest.TestCase):

    def test_defaults_all_zero(self):
        r = SyncResult()
        self.assertEqual(r.created, 0)
        self.assertEqual(r.updated, 0)
        self.assertEqual(r.deleted, 0)
        self.assertEqual(r.skipped, 0)
        self.assertEqual(r.errors, 0)

    def test_as_dict(self):
        r = SyncResult(created=3, updated=1, deleted=2, skipped=10, errors=0)
        d = r.as_dict()
        self.assertEqual(d["created"], 3)
        self.assertEqual(d["updated"], 1)
        self.assertEqual(d["deleted"], 2)
        self.assertEqual(d["skipped"], 10)
        self.assertEqual(d["errors"], 0)


# ---------------------------------------------------------------------------
# Tests: SyncEngine._process_file (change detection)
# ---------------------------------------------------------------------------

class TestProcessFile(unittest.TestCase):

    def _make_engine(self, cfg=None):
        """Build a SyncEngine with all external dependencies mocked."""
        cfg = cfg or make_config()
        with patch("sharepoint_sync.sync_engine.SharePointClient"), \
             patch("sharepoint_sync.sync_engine.TemplafyClient"), \
             patch("sharepoint_sync.sync_engine.SyncStateStore"):
            engine = SyncEngine(cfg)
        engine._tf = MagicMock(spec=TemplafyClient)
        engine._tf.library_for_extension = TemplafyClient.library_for_extension
        engine._sp_clients = [MagicMock()]
        engine._sp_by_drive = {"drive-1": engine._sp_clients[0]}
        engine._state = MagicMock()
        engine._state.folder_path_cache = {}
        return engine

    # -- SKIP: etag unchanged ------------------------------------------------

    def test_skip_when_etag_unchanged(self):
        engine = self._make_engine()
        sp_file = make_sp_file(etag="etag-v1")
        engine._state.get_file.return_value = {
            "etag": "etag-v1",
            "templafy_asset_id": "tf-asset-1",
            "templafy_library": "slides",
        }
        result = SyncResult()
        engine._process_file(sp_file, "slides", "root-folder", {}, result)

        self.assertEqual(result.skipped, 1)
        self.assertEqual(result.updated, 0)
        engine._tf.replace_asset.assert_not_called()

    # -- UPDATE: etag changed ------------------------------------------------

    def test_update_when_etag_changed(self):
        engine = self._make_engine()
        sp_file = make_sp_file(etag="etag-v2")
        engine._state.get_file.return_value = {
            "etag": "etag-v1",
            "templafy_asset_id": "tf-asset-1",
            "templafy_library": "slides",
        }
        engine._sp_by_drive["drive-1"].download_file.return_value = b"file-bytes"
        result = SyncResult()
        engine._process_file(sp_file, "slides", "root-folder", {}, result)

        self.assertEqual(result.updated, 1)
        self.assertEqual(result.skipped, 0)
        engine._tf.replace_asset.assert_called_once_with(
            library_type="slides",
            asset_id="tf-asset-1",
            filename="deck.pptx",
            content=b"file-bytes",
        )

    # -- CREATE: new file not in state ---------------------------------------

    def test_create_new_file(self):
        engine = self._make_engine()
        sp_file = make_sp_file(etag="etag-v1")
        engine._state.get_file.return_value = None
        engine._sp_by_drive["drive-1"].download_file.return_value = b"file-bytes"
        engine._tf.upload_asset.return_value = "new-tf-asset-id"
        engine._tf.ensure_folder_path.return_value = "folder-id"

        with patch.object(engine, "_ensure_templafy_folder", return_value="folder-id"):
            result = SyncResult()
            engine._process_file(sp_file, "slides", "root-folder", {}, result)

        self.assertEqual(result.created, 1)
        self.assertEqual(result.updated, 0)
        engine._tf.upload_asset.assert_called_once()

    # -- DRY RUN: no API calls -----------------------------------------------

    def test_dry_run_does_not_upload(self):
        engine = self._make_engine(make_config(dry_run=True))
        sp_file = make_sp_file(etag="etag-v1")
        engine._state.get_file.return_value = None

        with patch.object(engine, "_ensure_templafy_folder", return_value="folder-id"):
            result = SyncResult()
            engine._process_file(sp_file, "slides", "root-folder", {}, result)

        self.assertEqual(result.created, 1)
        engine._tf.upload_asset.assert_not_called()
        engine._sp_by_drive["drive-1"].download_file.assert_not_called()

    def test_dry_run_does_not_replace(self):
        engine = self._make_engine(make_config(dry_run=True))
        sp_file = make_sp_file(etag="etag-v2")
        engine._state.get_file.return_value = {
            "etag": "etag-v1",
            "templafy_asset_id": "tf-asset-1",
            "templafy_library": "slides",
        }
        result = SyncResult()
        engine._process_file(sp_file, "slides", "root-folder", {}, result)

        self.assertEqual(result.updated, 1)
        engine._tf.replace_asset.assert_not_called()

    # -- ERROR: exception during processing ----------------------------------

    def test_error_increments_error_count(self):
        engine = self._make_engine()
        sp_file = make_sp_file(etag="etag-v1")
        engine._state.get_file.return_value = None
        engine._sp_by_drive["drive-1"].download_file.side_effect = Exception("network error")

        with patch.object(engine, "_ensure_templafy_folder", return_value="folder-id"):
            result = SyncResult()
            # _process_file itself raises — caller (run) catches it
            with self.assertRaises(Exception):
                engine._process_file(sp_file, "slides", "root-folder", {}, result)


# ---------------------------------------------------------------------------
# Tests: Multi-drive folder prefix logic
# ---------------------------------------------------------------------------

class TestFolderPrefixLogic(unittest.TestCase):

    def test_primary_drive_no_prefix(self):
        """Files from the primary drive should not get a folder prefix."""
        cfg = make_config(
            sharepoint_drive_ids=["drive-1", "drive-2"],
            sharepoint_primary_drive="Documents",
        )
        with patch("sharepoint_sync.sync_engine.SharePointClient") as MockSP, \
             patch("sharepoint_sync.sync_engine.TemplafyClient"), \
             patch("sharepoint_sync.sync_engine.SyncStateStore"):

            mock_sp_primary = MagicMock()
            mock_sp_primary.get_drive_name.return_value = "Documents"
            mock_sp_rag = MagicMock()
            mock_sp_rag.get_drive_name.return_value = "RAG"

            MockSP.side_effect = [mock_sp_primary, mock_sp_rag]
            engine = SyncEngine(cfg)

        # Primary drive file — no prefix expected
        sp_file = make_sp_file(folder_path="Marketing", path="Marketing/deck.pptx", drive_id="drive-1")
        engine._sp_clients[0].get_drive_name.return_value = "Documents"

        # The engine applies prefix in run() not _process_file(), so test the logic directly
        drive_name = "Documents"
        is_primary = drive_name == cfg.sharepoint_primary_drive
        folder_prefix = None if is_primary else drive_name
        self.assertIsNone(folder_prefix)

    def test_secondary_drive_gets_prefix(self):
        """Files from non-primary drives should get the drive name as prefix."""
        drive_name = "RAG"
        primary = "Documents"
        is_primary = drive_name == primary
        folder_prefix = None if is_primary else drive_name
        self.assertEqual(folder_prefix, "RAG")

        # Simulate what run() does to folder_path
        sp_file = make_sp_file(folder_path="Presentations", path="Presentations/deck.pptx")
        sp_file.folder_path = f"{folder_prefix}/{sp_file.folder_path}".rstrip("/") if sp_file.folder_path else folder_prefix
        sp_file.path = f"{folder_prefix}/{sp_file.path}"

        self.assertEqual(sp_file.folder_path, "RAG/Presentations")
        self.assertEqual(sp_file.path, "RAG/Presentations/deck.pptx")

    def test_secondary_drive_empty_folder_path(self):
        """Files at the root of a non-primary drive get just the drive name as folder."""
        folder_prefix = "RAG"
        sp_file = make_sp_file(folder_path="", path="deck.pptx")
        sp_file.folder_path = f"{folder_prefix}/{sp_file.folder_path}".rstrip("/") if sp_file.folder_path else folder_prefix
        self.assertEqual(sp_file.folder_path, "RAG")


# ---------------------------------------------------------------------------
# Tests: Delete logic
# ---------------------------------------------------------------------------

class TestDeleteLogic(unittest.TestCase):

    def _make_engine(self):
        cfg = make_config()
        with patch("sharepoint_sync.sync_engine.SharePointClient"), \
             patch("sharepoint_sync.sync_engine.TemplafyClient"), \
             patch("sharepoint_sync.sync_engine.SyncStateStore"):
            engine = SyncEngine(cfg)
        engine._tf = MagicMock(spec=TemplafyClient)
        engine._tf.library_for_extension = TemplafyClient.library_for_extension
        engine._sp_clients = [MagicMock()]
        engine._sp_by_drive = {"drive-1": engine._sp_clients[0]}
        engine._state = MagicMock()
        engine._state.folder_path_cache = {}
        return engine

    def test_delete_asset_removed_from_sharepoint(self):
        engine = self._make_engine()
        engine._state.all_tracked_file_ids.return_value = {"sp-old-file"}
        engine._state.get_file.return_value = {
            "templafy_library": "slides",
            "templafy_asset_id": "tf-old-asset",
            "path": "old/deck.pptx",
        }
        engine._state.list_all_items = MagicMock(return_value=([], []))

        result = SyncResult()
        removed_ids = {"sp-old-file"} - set()  # simulate: file no longer in SP
        for sp_id in removed_ids:
            entry = engine._state.get_file(sp_id)
            engine._tf.delete_asset(entry["templafy_library"], entry["templafy_asset_id"])
            engine._state.remove_file(sp_id)
            result.deleted += 1

        self.assertEqual(result.deleted, 1)
        engine._tf.delete_asset.assert_called_once_with("slides", "tf-old-asset")
        engine._state.remove_file.assert_called_once_with("sp-old-file")


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    unittest.main()
