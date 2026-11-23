"""Application-wide logging setup.

Why this module exists: uvicorn configures only its own ``uvicorn.*`` loggers and
leaves the **root** logger with no handler. Records from ``gmc.*`` propagate to root,
find nothing there, and fall through to Python's ``lastResort`` handler — which emits
WARNING and above only. The practical effect before this module existed was that every
INFO line (including the sync summary) was silently discarded. Do not remove this as
"redundant"; nothing else attaches a root handler.

Output goes to stderr (visible via ``docker compose logs backend``) and to a rotating
file under ``settings.log_dir``, which lives inside the bind-mounted ``/app/data`` so it
survives container rebuilds. See docs/adr/ADR-005.md.
"""
import logging
import os
import sys
from logging.handlers import RotatingFileHandler

from app.config import settings

_FORMAT = "%(asctime)s %(levelname)-7s %(name)-14s %(message)s"
_LOG_FILENAME = "gmc.log"

# Guards against stacking duplicate handlers when called twice (tests, uvicorn --reload).
_configured = False


def setup_logging() -> None:
    """Attach stderr + rotating-file handlers to the root logger. Safe to call repeatedly."""
    global _configured
    if _configured:
        return

    formatter = logging.Formatter(_FORMAT)
    root = logging.getLogger()
    root.setLevel(settings.log_level.upper())

    stream = logging.StreamHandler(sys.stderr)
    stream.setFormatter(formatter)
    root.addHandler(stream)

    # File logging is a convenience, not a requirement — a read-only or unwritable
    # mount must not stop the app from starting.
    try:
        os.makedirs(settings.log_dir, exist_ok=True)
        file_handler = RotatingFileHandler(
            os.path.join(settings.log_dir, _LOG_FILENAME),
            maxBytes=settings.log_max_bytes,
            backupCount=settings.log_backup_count,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        root.addHandler(file_handler)
    except OSError:
        root.warning("file logging disabled: cannot write to %s", settings.log_dir, exc_info=True)

    # Each Merchant API call is logged by MerchantClient._get with the sub-account and
    # timing attached; httpx's own INFO line would duplicate it without that context.
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("httpcore").setLevel(logging.WARNING)

    _configured = True
