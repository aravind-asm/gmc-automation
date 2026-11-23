"""Automatic daily sync (ADR-012).

Once per KST day, at or after the daily sync time (``auto_sync_time_kst``, 13:00 KST = 09:30 IST),
the backend runs the same sync the Refresh button runs, unless a sync after that time has already
succeeded today. It checks every ``auto_sync_check_minutes``, so it catches up on its own after the
laptop wakes or the app restarts. A failed attempt is retried every ``auto_sync_retry_minutes``
until the KST day ends; a missed day is never carried over.

Every sync emails the team, so this is off unless the operator sets ``AUTO_SYNC_ENABLED=true``.
Retry state lives in memory: a restart forgets it and simply tries again (ADR-012 §4).
"""
import asyncio
import logging
from datetime import datetime, time, timedelta, timezone

from app import db
from app.api import insights
from app.config import settings

logger = logging.getLogger("gmc.auto_sync")

_last_attempt_at: datetime | None = None
_last_failed_at: datetime | None = None


def parse_sync_time(value: str) -> time:
    """``"HH:MM"`` → ``time``. Raises ``ValueError`` on anything else."""
    hours, sep, minutes = value.strip().partition(":")
    if not sep or not (hours.isdigit() and minutes.isdigit() and len(minutes) == 2):
        raise ValueError(f"auto_sync_time_kst must be HH:MM, got {value!r}")
    return time(int(hours), int(minutes))  # time() rejects 24:00, 13:60 …


def window_start(now: datetime, sync_time: time) -> datetime:
    """Today's daily sync time, today being the KST day ``now`` falls on."""
    return datetime.combine(now.astimezone(db.KST).date(), sync_time, tzinfo=db.KST)


def is_due(
    now: datetime,
    last_synced_at: str | None,
    last_failed_at: datetime | None,
    sync_time: time,
    retry: timedelta,
) -> bool:
    """Should the automatic sync run now? (ADR-012 §2.) Pure: no clock, no database.

    ``last_synced_at`` is the last **successful** sync (``sync_meta``, stored in UTC), whoever
    started it; ``last_failed_at`` the last failed automatic attempt.
    """
    start = window_start(now, sync_time)
    if now < start:
        return False  # before today's time; a Refresh earlier today does not count either
    if last_synced_at and datetime.fromisoformat(last_synced_at) >= start:
        return False  # today's sync is done
    if last_failed_at and last_failed_at >= start and now < last_failed_at + retry:
        return False  # failed today; wait for the retry gap
    return True


def _fmt(dt: datetime) -> str:
    return dt.astimezone(db.KST).strftime("%Y-%m-%d %H:%M KST")


async def check_once(now: datetime | None = None) -> None:
    """One check: run the automatic sync if it is due. Never raises (bar cancellation)."""
    global _last_attempt_at, _last_failed_at
    now = now or datetime.now(timezone.utc)
    if insights.sync_running():
        return
    sync_time = parse_sync_time(settings.auto_sync_time_kst)
    retry = timedelta(minutes=settings.auto_sync_retry_minutes)
    if not is_due(now, db.get_sync_meta()["last_synced_at"], _last_failed_at, sync_time, retry):
        return

    _last_attempt_at = now
    try:
        result = await insights.run_sync("auto")
    except insights.SyncAlreadyRunning:
        return  # a Refresh started first; if it succeeds, today's sync is done
    except Exception as exc:
        _last_failed_at = now
        if isinstance(exc, insights.SyncNotConfigured):
            logger.warning("auto sync not run: missing key or parent MCA id")
        elif not isinstance(exc, insights.SyncFailed):  # SyncFailed is logged by run_sync
            logger.exception("auto sync crashed")
        logger.warning("auto sync failed; next try at %s", _fmt(now + retry))
        return
    _last_failed_at = None
    logger.info(
        "auto sync done: %d rows, report %s",
        result["row_count"], "written" if result["report_written"] else "NOT written (no email)",
    )


def state(now: datetime | None = None) -> dict:
    """What the page's status line shows (ADR-012 §7). Times are UTC ISO; wording is the page's.

    ``state``: ``off`` · ``running`` (any sync, Refresh included) · ``done`` (today's sync is done)
    · ``retrying`` (an automatic attempt failed today) · ``waiting`` (before today's time, or due
    and about to start). No error text: Google's reasons stay in gmc.log.
    """
    out = {
        "enabled": settings.auto_sync_enabled,
        "time_kst": settings.auto_sync_time_kst,
        "state": "off",
        "next_try_at": None,
        "done_at": None,
        "done_by": None,
        "report_written": None,
        "last_attempt_at": _last_attempt_at.astimezone(timezone.utc).isoformat() if _last_attempt_at else None,
    }
    if not settings.auto_sync_enabled:
        return out
    try:
        sync_time = parse_sync_time(settings.auto_sync_time_kst)
    except ValueError:
        return out  # start() refused it too, so the timer is not running
    now = now or datetime.now(timezone.utc)
    start = window_start(now, sync_time)
    last = db.get_sync_meta()["last_synced_at"]

    if insights.sync_running():
        out["state"] = "running"
    elif last and datetime.fromisoformat(last) >= start:
        out["state"] = "done"
        out["done_at"] = last
        success = insights.last_success()
        if success and success["synced_at"] == last:  # unknown after a restart
            out["done_by"] = success["trigger"]
            out["report_written"] = success["report_written"]
    elif _last_failed_at and _last_failed_at >= start:
        out["state"] = "retrying"
        retry_at = _last_failed_at + timedelta(minutes=settings.auto_sync_retry_minutes)
        out["next_try_at"] = retry_at.astimezone(timezone.utc).isoformat()
    else:
        out["state"] = "waiting"
        if now < start:
            out["next_try_at"] = start.astimezone(timezone.utc).isoformat()
    return out


async def loop() -> None:
    """Check every ``auto_sync_check_minutes`` until cancelled at shutdown."""
    interval = settings.auto_sync_check_minutes * 60
    while True:
        try:
            await check_once()
        except asyncio.CancelledError:
            raise
        except Exception:  # a bug here must not end the timer for good
            logger.exception("auto sync check failed")
        await asyncio.sleep(interval)


def start() -> asyncio.Task | None:
    """Start the timer if switched on (called from the app's lifespan). Logs either way."""
    if not settings.auto_sync_enabled:
        logger.info("auto sync off")
        return None
    try:
        sync_time = parse_sync_time(settings.auto_sync_time_kst)
    except ValueError:
        logger.exception("auto sync off: bad AUTO_SYNC_TIME_KST")
        return None
    logger.info(
        "auto sync on, daily at %s KST, checking every %d min",
        sync_time.strftime("%H:%M"), settings.auto_sync_check_minutes,
    )
    return asyncio.create_task(loop(), name="auto-sync")
