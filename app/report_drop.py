"""Write each sync's Excel report to a watched folder for email distribution (ADR-007).

This module is the app's **entire** contribution to getting the report into people's
inboxes: it writes a file. The folder is a OneDrive-synced directory bind-mounted into the
container; a Power Automate cloud flow owned by the operator watches it and sends the mail.
Nothing here knows about email, recipients, or credentials, and nothing here talks to the
database — the caller hands it finished bytes.

Disabled unless ``report_drop_dir`` is set, so the app runs unchanged for anyone without
OneDrive.
"""
import logging
import os
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

from app.config import settings

logger = logging.getLogger("gmc.report")

FILENAME_PREFIX = "gmc-report-"
FILENAME_SUFFIX = ".xlsx"
# Seconds, not minutes: two syncs in the same minute must not collide, because a
# collision would silently overwrite an archived run (ADR-007 "keep every drop").
TIMESTAMP_FORMAT = "%Y-%m-%d-%H%M%S"


def _stamp(synced_at: str | None) -> str:
    """Format a sync's ISO timestamp for the filename, falling back to now.

    ``last_synced_at`` is written by ``db.replace_disapprovals`` as a UTC ISO string, so it
    normally parses cleanly. An unparseable value is not worth failing a drop over — the
    file still needs a unique name, and the workbook carries its own provenance header.
    """
    try:
        dt = datetime.fromisoformat(synced_at) if synced_at else None
    except ValueError:
        logger.warning("unparseable last_synced_at %r; naming the drop from the clock", synced_at)
        dt = None
    return (dt or datetime.now(timezone.utc)).strftime(TIMESTAMP_FORMAT)


def drop_path(synced_at: str | None) -> Path:
    """Absolute path this sync's report would be written to."""
    return Path(settings.report_drop_dir) / f"{FILENAME_PREFIX}{_stamp(synced_at)}{FILENAME_SUFFIX}"


def write_report(buf: BytesIO, synced_at: str | None) -> Path | None:
    """Write the workbook to the drop folder; return the path, or ``None`` if disabled.

    Written to a ``.part`` sibling and then renamed. OneDrive uploads whatever appears in
    the folder, so a directly-written file can be picked up half-complete and emailed as a
    corrupt attachment; ``os.replace`` makes the file appear whole or not at all. The temp
    file is deliberately in the *same* directory — elsewhere and the rename becomes a
    cross-device copy, which is not atomic.

    Raises on any filesystem error. The caller is expected to catch: a drop failure must
    never fail the sync that produced it (ADR-007).
    """
    if not settings.report_drop_dir:
        return None

    path = drop_path(synced_at)
    path.parent.mkdir(parents=True, exist_ok=True)
    part = path.with_name(path.name + ".part")
    try:
        part.write_bytes(buf.getvalue())
        os.replace(part, path)
    except Exception:
        part.unlink(missing_ok=True)
        logger.exception("report drop failed: %s", path)
        raise

    logger.info("report dropped: %s (%d bytes)", path, path.stat().st_size)
    return path
