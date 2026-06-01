"""
SharePoint → Templafy Sync
Azure Function - Timer Trigger (nightly)

Syncs .pptx, .docx, and images from a SharePoint site
to the corresponding Templafy library, mirroring folder structure.
"""

import logging
import azure.functions as func

from .sync_engine import SyncEngine
from .config import SyncConfig

logger = logging.getLogger(__name__)


def main(mytimer: func.TimerRequest) -> None:
    if mytimer.past_due:
        logger.warning("Timer is past due — running sync anyway.")

    logger.info("SharePoint → Templafy sync started.")

    try:
        config = SyncConfig.from_environment()
        engine = SyncEngine(config)
        result = engine.run()

        logger.info(
            "Sync complete. "
            f"Created: {result['created']}, "
            f"Updated: {result['updated']}, "
            f"Deleted: {result['deleted']}, "
            f"Skipped (unchanged): {result['skipped']}, "
            f"Errors: {result['errors']}"
        )
    except Exception as e:
        logger.exception(f"Sync failed with unhandled exception: {e}")
        raise
