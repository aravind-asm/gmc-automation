"""Disapproved-products API: setup status, on-demand sync, and cached reads."""
import logging
import math
from bisect import bisect_right
from datetime import date, datetime, timedelta, timezone
from io import BytesIO

from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import StreamingResponse

from app import auto_sync, db, report, report_drop, stores
from app.config import settings
from app.services import merchant_auth
from app.services.merchant_client import client

XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

logger = logging.getLogger("gmc.api")
router = APIRouter(prefix="/api", tags=["insights"])


@router.get("/setup/status")
async def setup_status() -> dict:
    """Report whether the service account + parent MCA are configured."""
    has_key = merchant_auth.key_present()
    has_parent = bool(settings.parent_mca_id)
    meta = db.get_sync_meta()
    return {
        "configured": has_key and has_parent,
        "key_present": has_key,
        "parent_mca_configured": has_parent,
        "service_account_email": merchant_auth.service_account_email() if has_key else None,
        "last_synced_at": meta["last_synced_at"],
        "cached_rows": meta["row_count"],
        "auto_sync": auto_sync.state(),
    }


class SyncNotConfigured(Exception):
    """The service-account key or the parent MCA id is missing."""


class SyncAlreadyRunning(Exception):
    """Another sync (a Refresh or the automatic one) is still running."""


class SyncFailed(Exception):
    """The Merchant API fetch failed; nothing was written and nothing was dropped."""


# One sync at a time (ADR-012 §6): every sync emails the team, so two at once would mean two
# emails and twice the ~156 Merchant API calls. A plain flag rather than an asyncio.Lock: the
# check and the set have no await between them, so on the one event loop they cannot interleave,
# and a flag is not tied to an event loop (the tests open a new loop per TestClient).
_sync_running = False
# The last successful sync's trigger and whether its report was written, for the page's status
# line (ADR-012 §7). In memory: after a restart the line simply cannot say whether it was emailed.
_last_success: dict | None = None


def sync_running() -> bool:
    return _sync_running


def last_success() -> dict | None:
    return _last_success


async def run_sync(trigger: str) -> dict:
    """Run one full sync and drop its report. Shared by the Refresh button and the automatic sync.

    ``trigger`` is ``"manual"`` or ``"auto"`` and only goes to the log. Raises
    ``SyncNotConfigured``, ``SyncAlreadyRunning`` or ``SyncFailed``.
    """
    global _sync_running, _last_success
    if not merchant_auth.key_present() or not settings.parent_mca_id:
        raise SyncNotConfigured()
    if _sync_running:
        logger.info("sync requested (trigger=%s) while one is running; not started", trigger)
        raise SyncAlreadyRunning()
    _sync_running = True
    try:
        result = await _sync_and_drop(trigger)
        _last_success = {"synced_at": result["last_synced_at"], "trigger": trigger,
                         "report_written": result["report_written"]}
        return result
    finally:
        _sync_running = False


@router.post("/sync")
async def sync() -> dict:
    """Fetch disapproved products across all sub-accounts and refresh the cache."""
    try:
        return await run_sync("manual")
    except SyncNotConfigured:
        raise HTTPException(status_code=409, detail="Not configured: missing key or parent MCA id.")
    except SyncAlreadyRunning:
        raise HTTPException(status_code=409, detail="A sync is already running.")
    except SyncFailed:
        raise HTTPException(status_code=502, detail="Failed to fetch data from Merchant API.")


async def _sync_and_drop(trigger: str) -> dict:
    logger.info("sync requested (trigger=%s)", trigger)
    try:
        rows, stats, checks, performance = await client.fetch_all_disapprovals()
    except Exception as exc:
        logger.exception("sync failed (trigger=%s)", trigger)
        raise SyncFailed() from exc
    db.replace_disapprovals(rows, stats, checks, performance)
    meta = db.get_sync_meta()
    logger.info("sync cached %d rows at %s", meta["row_count"], meta["last_synced_at"])

    # Drop the report for email pickup (ADR-007). Runs after the commit because it is
    # built from the cache, and is swallowed on failure because a full sync costs ~150
    # Merchant API calls — losing that over an unwritable folder would be absurd.
    report_written = False
    try:
        report_written = report_drop.write_report(_build_report(meta), meta["last_synced_at"]) is not None
    except Exception:
        logger.exception("report drop failed (the sync itself succeeded)")

    return {
        "synced": True,
        "row_count": meta["row_count"],
        "last_synced_at": meta["last_synced_at"],
        "report_written": report_written,
    }


@router.get("/disapprovals")
async def disapprovals(
    sub_account: str | None = Query(None),
    issue_code: str | None = Query(None),
    q: str | None = Query(None, description="Free-text over product title / offer id"),
    limit: int = Query(100, ge=1, le=5000),
    offset: int = Query(0, ge=0),
) -> dict:
    """One row per product; each product's issues are grouped into ``issues[]`` (ADR-004).

    ``total``/pagination are by product. ``issue_code`` selects products having that issue and
    returns all of their issues.
    """
    products, total = db.query_products(sub_account, issue_code, q, limit, offset)
    return {"total": total, "limit": limit, "offset": offset, "products": products}


def _rate(total_products: int, disapproved_products: int) -> float:
    """Disapproval fraction, guarded against a zero denominator (ADR-002)."""
    return disapproved_products / total_products if total_products > 0 else 0.0


# Rates are shown, and judged, to this many decimal places of a percent: 94.5 %, -2.3 pp
# (operator, 2026-11-16; was 2). Status, spikes and every shown change use the rounded figure, so
# what people read always adds up and always matches the colour beside it.
RATE_DP = 1


def _shown_pct(fraction: float) -> float:
    """A fraction as the percent people read, e.g. 0.05537 -> 5.5."""
    return round(fraction * 100, RATE_DP)


def _leaderboard_rows() -> list[dict]:
    """Latest-snapshot sub-accounts at or above the threshold, worst first.

    Shared by ``GET /api/leaderboard`` and the Excel export so the page and the sheet
    cannot disagree about who is on the board (ADR-006 S1).
    """
    threshold = settings.leaderboard_threshold
    result = []
    for r in db.latest_stats():
        rate = _rate(r["total_products"], r["disapproved_products"])
        if rate >= threshold:
            result.append(
                {
                    "sub_account_id": r["sub_account_id"],
                    "sub_account_name": r["sub_account_name"],
                    "total_products": r["total_products"],
                    "disapproved_products": r["disapproved_products"],
                    "rate": rate,
                    "last_synced_at": r["synced_at"],
                }
            )
    result.sort(key=lambda x: x["rate"], reverse=True)
    return result


def _history_rows() -> list[dict]:
    """Every recorded snapshot of every real store, oldest first, rate precomputed — the
    Approval Rate Over Time sheet (all stores, filterable: operator, 2026-11-11).

    ``since=None`` deliberately spans **all** history: the History page defaults to 90 days
    because it has a control to widen it, but the export is an archive with no such control
    (ADR-006 S5). Test stores are left out; region and country come from the Store Directory.
    Uses the same ``_point_rate`` helper as ``GET /api/history``.

    Stores with no products are hidden (ADR-011 §2): a 0-product snapshot has no rate (a gap in
    the line, never 0 %), and a store that never had products is left out altogether.
    """
    real = {sid: stores.lookup(sid) for sid in db.stats_subaccount_ids()}
    real = {sid: store for sid, store in real.items() if not store.is_test}
    rows = sorted(db.stats_history(list(real), since=None), key=lambda r: r["synced_at"])
    ever_had_products = {r["sub_account_id"] for r in rows if r["total_products"] > 0}
    return [
        {**r, "rate": _point_rate(r), "approval_rate": _approval_rate(_point_rate(r)),
         "region": real[r["sub_account_id"]].region, "country": real[r["sub_account_id"]].country}
        for r in rows
        if r["sub_account_id"] in ever_had_products
    ]


def _approval_rate(rate: float | None) -> float | None:
    """A store's approval rate: 100 % minus its disapproval rate **as shown** (``RATE_DP``), so the
    two figures always add up and the status beside it matches (ADR-011 amendment, 2026-11-12)."""
    return None if rate is None else _approval(_shown_pct(rate))


def _point_rate(snapshot: dict) -> float | None:
    """A snapshot's rate for a chart: ``None`` (a gap) when the store had no products then —
    0 / 0 is not a rate, and drawing it as 0 % would look like a perfect day (ADR-011 §2)."""
    if snapshot["total_products"] <= 0:
        return None
    return _rate(snapshot["total_products"], snapshot["disapproved_products"])


# Overview status values (ADR-009 §2). "not_synced" is a directory store with no row in the
# latest sync — its fetch failed (ADR-002 omits it) or it has never been fetched yet. Stores
# with no products are not shown at all, so they have no status (ADR-011 §2).
STATUS_ORDER = ["critical", "warning", "minor", "healthy", "not_synced"]


def _status(total_products: int, disapproved_products: int, critical_warning: bool = False) -> str:
    """Status from the rate **rounded to ``RATE_DP``**, so the colour always matches the figure
    shown (a raw 4.96 % displays as 5.0 % and must not read Healthy). Bottom bounds inclusive.

    A CRITICAL account warning makes a store Critical whatever its rate (ADR-009 §3). Only
    stores with products get a status: empty ones are hidden before this (ADR-011 §2).
    """
    if critical_warning:
        return "critical"
    shown = _shown_pct(_rate(total_products, disapproved_products))
    if shown >= _shown_pct(settings.status_critical_threshold):
        return "critical"
    if shown >= _shown_pct(settings.status_warning_threshold):
        return "warning"
    if shown >= _shown_pct(settings.leaderboard_threshold):
        return "minor"
    return "healthy"


def _region_sort_key(row: dict) -> tuple:
    unassigned = row["region"] == stores.UNASSIGNED
    return (unassigned, row["region"].lower(), row["country"].lower(), row["sub_account_name"].lower())


# Account-warning severities the Overview shows, in display order (ADR-009 §3). SUGGESTION
# is left out on purpose.
_WARNING_SEVERITIES = ["CRITICAL", "ERROR"]


def _account_warnings(check: dict | None) -> tuple[str, list[dict]]:
    """``(account_check, account_issues)`` for one store.

    ``account_check`` is ``ok``, ``failed`` (Google did not answer: status falls back to the
    rate) or ``not_checked`` (no check recorded — a sync from before warnings were fetched).
    ``account_issues`` holds CRITICAL then ERROR warnings, each group by title.
    """
    if check is None:
        return "not_checked", []
    if not check["ok"]:
        return "failed", []
    shown = [i for i in check["issues"] if i["severity"] in _WARNING_SEVERITIES]
    shown.sort(key=lambda i: (_WARNING_SEVERITIES.index(i["severity"]), i["title"].lower()))
    return "ok", shown


def _overview_rows() -> list[dict]:
    """One row per real store in the latest sync, labelled from the Store Directory.

    Shared by ``GET /api/overview`` and the Overview sheet (ADR-006 rule). Test stores and
    stores with no products are left out, without a note (ADR-011 §2); stores missing from the
    directory are kept as Unassigned; directory stores absent from the latest sync are kept as
    ``not_synced`` with no figures. Each row carries
    its account warnings (ADR-009 §3). Sorted by region, country, store, Unassigned last.
    """
    checks = db.latest_account_checks()
    rows = []
    seen: set[str] = set()
    for r in db.latest_stats():
        store = stores.lookup(r["sub_account_id"], r["sub_account_name"])
        seen.add(store.sub_account_id)  # before the skip: an empty store is not "not synced"
        if store.is_test or r["total_products"] <= 0:
            continue
        account_check, account_issues = _account_warnings(checks.get(r["sub_account_id"]))
        critical_warning = any(i["severity"] == "CRITICAL" for i in account_issues)
        rows.append(
            {
                "sub_account_id": r["sub_account_id"],
                "sub_account_name": r["sub_account_name"],
                "region": store.region,
                "country": store.country,
                "in_directory": store.in_directory,
                "total_products": r["total_products"],
                "disapproved_products": r["disapproved_products"],
                "rate": _rate(r["total_products"], r["disapproved_products"]),
                "approval_rate": _approval_rate(_rate(r["total_products"], r["disapproved_products"])),
                "status": _status(r["total_products"], r["disapproved_products"], critical_warning),
                # Red from the rate alone (a store can be red from both) — the page and sheet
                # say so in the warnings column, since Google may have sent no warning at all.
                "critical_by_rate": _status(r["total_products"], r["disapproved_products"]) == "critical",
                "account_check": account_check,
                "account_issues": account_issues,
            }
        )
    if seen:  # before the first sync there is nothing to say "not synced" against
        for store in stores.directory().values():
            if store.is_test or store.sub_account_id in seen:
                continue
            # Its products failed but its warning check may still have worked: list what we
            # know, but the status stays "not in last sync" — there are no figures to judge.
            account_check, account_issues = _account_warnings(checks.get(store.sub_account_id))
            rows.append(
                {
                    "sub_account_id": store.sub_account_id,
                    "sub_account_name": store.name,
                    "region": store.region,
                    "country": store.country,
                    "in_directory": True,
                    "total_products": None,
                    "disapproved_products": None,
                    "rate": None,
                    "approval_rate": None,
                    "status": "not_synced",
                    "critical_by_rate": False,
                    "account_check": account_check,
                    "account_issues": account_issues,
                }
            )
    rows.sort(key=_region_sort_key)
    return rows


def _overview_kpis(rows: list[dict]) -> dict:
    """KPI cards. Totals cover stores with figures (every shown store has products).

    "Average approval rate" is 1 − the unweighted mean of store rates (ADR-009 §4, ADR-011 §4):
    approved = not disapproved, so pending review counts as not approved. None when no store
    has figures, so the UI shows a dash rather than a misleading 100 %.
    """
    measured = [r for r in rows if r["status"] != "not_synced"]
    counts = {s: 0 for s in STATUS_ORDER}
    for r in rows:
        counts[r["status"]] += 1
    return {
        "store_count": len(rows),
        "total_products": sum(r["total_products"] for r in measured),
        "disapproved_products": sum(r["disapproved_products"] for r in measured),
        "average_approval_rate": (
            1 - sum(r["rate"] for r in measured) / len(measured) if measured else None
        ),
        "status_counts": counts,
    }


def _thresholds() -> dict:
    """Status thresholds as fractions — one place, for the page and the sheet."""
    return {"minor": settings.leaderboard_threshold, "warning": settings.status_warning_threshold,
            "critical": settings.status_critical_threshold}


@router.get("/overview")
async def overview() -> dict:
    """Store health by region and country, latest sync (ADR-009)."""
    rows = _overview_rows()
    return {
        "last_synced_at": db.get_sync_meta()["last_synced_at"],
        "thresholds": _thresholds(),
        "kpis": _overview_kpis(rows),
        "stores": rows,
    }


SPIKES_SHOWN = 10  # the alert card's length; the full table lists every store


def _day_label(day: str) -> str:
    """``2026-11-07`` -> ``Fri 7 Nov 2026`` (built by hand: ``%-d`` is not portable)."""
    d = date.fromisoformat(day)
    return f"{d:%a} {d.day} {d:%b %Y}"


def _spike_radar() -> dict:
    """Daily Spike Radar: each store's rate now (T-0) against the last sync of an earlier KST day
    (T-1) (ADR-009 §5). Shared by ``GET /api/spikes``, the Overview page and the sheet.

    ``state`` is ``no_sync``, ``one_day`` (nothing earlier to compare with) or ``ok``. Rows cover
    real stores with products at T-0: test stores and empty stores are left out. Changes are taken
    between the rates rounded to ``RATE_DP`` — the figures shown — so ``abs_diff_pp`` always equals what
    the reader can subtract. A store with no T-1 rate (new, failed then, or empty then) is listed
    with ``baseline_missing`` and is never a spike.

    Shown in approval terms (ADR-011 §5): ``t1_approval`` / ``t0_approval`` are 100 % minus the
    shown disapproval figures and ``approval_diff_pp`` = −``abs_diff_pp``, so before, now and
    change always add up. The spike rule is unchanged: at least ``spike_min_disapproved`` now and
    approval down by at least ``spike_min_diff_pp``. Sorted by approval gain, biggest first;
    no-baseline last. ``average_approval_diff_pp`` is the mean change of compared stores;
    ``biggest_drop`` is the compared store whose approval fell most (``None`` if nobody fell).
    """
    thresholds = {"min_disapproved": settings.spike_min_disapproved, "min_diff_pp": settings.spike_min_diff_pp}
    days = db.daily_snapshots()
    result = {"state": "no_sync", "t0": None, "t1": None, "thresholds": thresholds,
              "rows": [], "spikes": [], "spike_count": 0,
              "average_approval_diff_pp": None, "biggest_drop": None}
    if not days:
        return result
    t0 = days[0]
    result["t0"] = {**t0, "label": _day_label(t0["kst_day"])}
    if len(days) < 2:
        result["state"] = "one_day"
        return result
    t1 = days[1]
    days_back = (date.fromisoformat(t0["kst_day"]) - date.fromisoformat(t1["kst_day"])).days
    result["t1"] = {**t1, "label": _day_label(t1["kst_day"]), "days_back": days_back}
    result["state"] = "ok"

    before = {r["sub_account_id"]: r for r in db.stats_at(t1["synced_at"])}
    rows = []
    for now in db.stats_at(t0["synced_at"]):
        store = stores.lookup(now["sub_account_id"], now["sub_account_name"])
        if store.is_test or now["total_products"] <= 0:
            continue
        t0_rate = _rate(now["total_products"], now["disapproved_products"])
        shown_now = _shown_pct(t0_rate)
        prev = before.get(now["sub_account_id"])
        row = {
            "sub_account_id": now["sub_account_id"],
            "sub_account_name": now["sub_account_name"],
            "region": store.region,
            "country": store.country,
            "t0_total": now["total_products"],
            "t0_disapproved": now["disapproved_products"],
            "t0_rate": t0_rate,
            "t1_rate": None,
            "t0_approval": _approval(shown_now),
            "t1_approval": None,
            "abs_diff_pp": None,
            "approval_diff_pp": None,
            "baseline_missing": True,
            "spike": False,
        }
        if prev is not None and prev["total_products"] > 0:
            t1_rate = _rate(prev["total_products"], prev["disapproved_products"])
            shown_before = _shown_pct(t1_rate)
            diff = round(shown_now - shown_before, RATE_DP)
            row.update(
                t1_rate=t1_rate,
                t1_approval=_approval(shown_before),
                abs_diff_pp=diff,
                approval_diff_pp=round(-diff, RATE_DP) + 0.0,  # + 0.0 turns -0.0 into 0.0
                baseline_missing=False,
                spike=(now["disapproved_products"] >= settings.spike_min_disapproved
                       and diff >= round(settings.spike_min_diff_pp, RATE_DP)),
            )
        rows.append(row)

    rows.sort(key=lambda r: (r["baseline_missing"], -(r["approval_diff_pp"] or 0), r["sub_account_name"].lower()))
    spikes = [r for r in rows if r["spike"]]
    compared = [r for r in rows if not r["baseline_missing"]]
    falls = [r for r in compared if r["approval_diff_pp"] < 0]
    result.update(
        rows=rows,
        # Spikes are the biggest approval drops, so the alert lists the worst first.
        spikes=sorted(spikes, key=lambda r: (r["approval_diff_pp"], r["sub_account_name"].lower()))[:SPIKES_SHOWN],
        spike_count=len(spikes),
        average_approval_diff_pp=(
            round(sum(r["approval_diff_pp"] for r in compared) / len(compared), RATE_DP) + 0.0 if compared else None
        ),
        biggest_drop=min(falls, key=lambda r: (r["approval_diff_pp"], r["sub_account_name"].lower()))
        if falls else None,
    )
    return result


def _approval(shown_rate_pct: float) -> float:
    """Approval rate (fraction) from a disapproval rate as shown (percent, ``RATE_DP``): 100 % minus it,
    so the approval figure and the disapproval figure always add up to exactly 100 %."""
    return round((100 - shown_rate_pct) / 100, RATE_DP + 2)


@router.get("/spikes")
async def spikes() -> dict:
    """Daily Spike Radar for the latest sync (ADR-009 §5)."""
    return _spike_radar()


TOP_ISSUES_LIMIT = 10


def _most_common(counter: dict[str, set]) -> str:
    """The value carried by the most products; ties go to the alphabetically first."""
    return min(counter, key=lambda value: (-len(counter[value]), value))


def _top_issues() -> tuple[list[dict], dict]:
    """The ten issue codes affecting the most products in the latest sync (ADR-010 §3).

    Shared by ``GET /api/issues/top`` and the Top 10 Issues sheet (ADR-006 rule). A product is
    one ``(store, offer_id)`` and counts once per code however many signatures it has. Test
    stores are left out. Countries come from the Store Directory; a store outside it counts as a
    store but not as a country. Ranked by products, then stores, then code, so ties never
    reshuffle. Pending-review codes are ranked and counted like any other, but flagged.

    Returns ``(issues, totals)``; totals cover every code, not only the ten shown.
    ``top3_share`` (ADR-011 §7) = distinct products carrying any of the first three **fixable**
    (non-pending) codes ÷ disapproved products — a product with two of them counts once; None when
    there is no fixable code. ``top3_codes`` names those codes.
    """
    pending = set(settings.pending_issue_codes)
    by_code: dict[str, dict] = {}
    all_products: set[tuple[str, str]] = set()
    for r in db.product_issue_codes():
        store = stores.lookup(r["sub_account_id"])
        if store.is_test:
            continue
        product = (r["sub_account_id"], r["offer_id"])
        all_products.add(product)
        entry = by_code.setdefault(
            r["issue_code"],
            {"products": set(), "stores": set(), "countries": set(), "attributes": {}, "descriptions": {}},
        )
        entry["products"].add(product)
        entry["stores"].add(r["sub_account_id"])
        if store.in_directory:
            entry["countries"].add(store.country)
        entry["attributes"].setdefault(r["attribute"] or "", set()).add(product)
        entry["descriptions"].setdefault(r["description"] or "", set()).add(product)

    ranked = sorted(
        by_code.items(),
        key=lambda item: (-len(item[1]["products"]), -len(item[1]["stores"]), item[0]),
    )
    issues = [
        {
            "rank": rank,
            "issue_code": code,
            "attribute": _most_common(entry["attributes"]),
            "description": _most_common(entry["descriptions"]),
            "products": len(entry["products"]),
            "stores": len(entry["stores"]),
            "countries": len(entry["countries"]),
            "pending": code in pending,
        }
        for rank, (code, entry) in enumerate(ranked[:TOP_ISSUES_LIMIT], start=1)
    ]
    pending_products = set().union(*(e["products"] for code, e in by_code.items() if code in pending))
    top3 = [(code, entry) for code, entry in ranked if code not in pending][:3]
    top3_products = set().union(*(entry["products"] for _, entry in top3))
    totals = {
        "disapproved_products": len(all_products),
        "issue_codes": len(by_code),
        "pending_products": len(pending_products),  # distinct, so never double-counted
        "top3_share": len(top3_products) / len(all_products) if top3 else None,
        "top3_codes": [code for code, _ in top3],
    }
    return issues, totals


@router.get("/issues/top")
async def top_issues() -> dict:
    """Global Issue Top 10 for the latest sync (ADR-010 §3)."""
    issues, totals = _top_issues()
    return {
        "last_synced_at": db.get_sync_meta()["last_synced_at"],
        "pending_issue_codes": settings.pending_issue_codes,
        "totals": totals,
        "issues": issues,
    }


# Display and sort order (ADR-010 §3; four levels since ADR-011 §6).
PRIORITY_TAGS = ["critical", "high", "medium", "low", "pending"]
PERFORMANCE_METRICS = ("impressions", "ads_impressions", "free_impressions", "clicks", "ads_clicks", "free_clicks")


def _window_label(start: str, end: str) -> str:
    """``2026-10-11``, ``2026-11-09`` -> ``11 Oct – 9 Nov 2026`` (the year once when shared)."""
    a, b = date.fromisoformat(start), date.fromisoformat(end)
    first = f"{a.day} {a:%b}" if a.year == b.year else f"{a.day} {a:%b %Y}"
    return f"{first} – {b.day} {b:%b %Y}"


def _share_cut(share: float, n: int) -> int:
    """How many ranks a share of ``n`` groups covers, rounded up. Rounded to 9 places first:
    ``0.3 * 10`` is 3.0000000000000004 in floating point, which must not become 4."""
    return math.ceil(round(share * n, 9))


def _performance_summary(checks: dict[str, dict]) -> dict | None:
    """The window and how many real stores had their report read; None before any was fetched."""
    real = {sid: c for sid, c in checks.items() if not stores.lookup(sid).is_test}
    if not real:
        return None
    window = next((c for c in real.values() if c["window_start"] and c["window_end"]), None)
    return {
        "window_start": window["window_start"] if window else None,
        "window_end": window["window_end"] if window else None,
        "label": _window_label(window["window_start"], window["window_end"]) if window else None,
        "stores_with_data": sum(1 for c in real.values() if c["ok"]),
        "stores_without_data": sum(1 for c in real.values() if not c["ok"]),
    }


def _priority_list() -> dict:
    """Priority List: problems grouped by store × issue code, what to fix first (ADR-010 §1-3).

    Shared by ``GET /api/priority``, the Issues page and the Priority List sheet.

    Each group carries its products' **impressions and clicks** over the last 30 complete days,
    in total and split Ads / Free listings (``AFFILIATE`` in the total only). A product with no
    report row counts 0; a store whose report could not be read has ``has_performance`` False and
    no figures (None — never 0).

    Levels (ADR-011 §6): **pending** for review codes (nothing to fix — wins over everything);
    **critical** when the store has a CRITICAL account warning; otherwise groups that can be ranked
    (not pending, performance read) are ranked by impressions — ties share a rank — and
    **critical** is a rank within the top ``priority_critical_share`` of them, **high** within
    ``priority_high_share``, both rounded up, **medium** the rest with impressions; **low** is 0
    impressions. A store whose performance report could not be read is **medium** (not ranked: we
    cannot tell). Sorted Critical → High → Medium → Low → Pending, then impressions (no figures
    last), then affected products.

    ``open_days`` counts KST calendar days from the group's oldest first sighting to the latest
    sync. ``since_tracking_began`` marks a group first seen on the day tracking started: it may be
    older, so it reads "≥ N days". Test stores are left out.
    """
    meta = db.get_sync_meta()
    started = db.tracking_since()
    perf_checks = db.performance_checks()
    result = {
        "last_synced_at": meta["last_synced_at"],
        "tracking_since": {"first_seen_at": started, "label": _day_label(db.kst_day(started))} if started else None,
        "performance": _performance_summary(perf_checks),
        "shares": {"critical": settings.priority_critical_share, "high": settings.priority_high_share},
        "ranked_groups": 0,
        "tag_counts": {t: 0 for t in PRIORITY_TAGS},
        "groups": [],
    }
    if not meta["last_synced_at"]:
        return result
    latest_day = date.fromisoformat(db.kst_day(meta["last_synced_at"]))
    tracking_day = db.kst_day(started) if started else None
    pending = set(settings.pending_issue_codes)
    critical_stores = {
        sid for sid, check in db.latest_account_checks().items()
        if check["ok"] and any(i["severity"] == "CRITICAL" for i in check["issues"])
    }

    groups: dict[tuple[str, str], dict] = {}
    for r in db.priority_rows():
        store = stores.lookup(r["sub_account_id"], r["sub_account_name"])
        if store.is_test:
            continue
        key = (r["sub_account_id"], r["issue_code"])
        g = groups.setdefault(key, {
            "store": store, "name": r["sub_account_name"], "products": set(), "first_seen": None,
            "attributes": {}, "descriptions": {},
        })
        g["products"].add(r["offer_id"])
        g["attributes"].setdefault(r["attribute"] or "", set()).add(r["offer_id"])
        g["descriptions"].setdefault(r["description"] or "", set()).add(r["offer_id"])
        if r["first_seen_at"] and (g["first_seen"] is None or r["first_seen_at"] < g["first_seen"]):
            g["first_seen"] = r["first_seen_at"]

    perf = db.disapproved_product_performance()
    rows = []
    for (sid, code), g in groups.items():
        has_performance = perf_checks.get(sid, {}).get("ok", False)
        figures: dict = {m: None for m in PERFORMANCE_METRICS}
        if has_performance:
            figures = {m: 0 for m in PERFORMANCE_METRICS}
            for offer in g["products"]:
                for m, value in perf.get((sid, offer.lower()), {}).items():
                    figures[m] += value
        first = g["first_seen"]
        rows.append({
            "sub_account_id": sid,
            "sub_account_name": g["name"] or sid,
            "region": g["store"].region,
            "country": g["store"].country,
            "issue_code": code,
            "attribute": _most_common(g["attributes"]),
            "description": _most_common(g["descriptions"]),
            "products": len(g["products"]),
            "first_seen_at": first,
            "open_days": (latest_day - date.fromisoformat(db.kst_day(first))).days if first else None,
            "since_tracking_began": bool(first) and db.kst_day(first) == tracking_day,
            "has_performance": has_performance,
            **figures,
        })

    # Rank the groups someone could fix, by impressions; a tie shares the better rank.
    ranked = sorted(r["impressions"] for r in rows if r["has_performance"] and r["issue_code"] not in pending)
    critical_cut = _share_cut(settings.priority_critical_share, len(ranked))
    high_cut = _share_cut(settings.priority_high_share, len(ranked))
    for r in rows:
        if r["issue_code"] in pending:
            r["tag"] = "pending"
        elif r["sub_account_id"] in critical_stores:
            r["tag"] = "critical"
        elif not r["has_performance"]:
            r["tag"] = "medium"  # no figures to rank by
        elif r["impressions"] > 0:
            rank = 1 + len(ranked) - bisect_right(ranked, r["impressions"])  # 1 + groups with more
            r["tag"] = "critical" if rank <= critical_cut else "high" if rank <= high_cut else "medium"
        else:
            r["tag"] = "low"
    result["ranked_groups"] = len(ranked)

    rows.sort(key=lambda g: (PRIORITY_TAGS.index(g["tag"]),
                             -(g["impressions"] if g["impressions"] is not None else -1),
                             -g["products"], -(g["open_days"] or 0),
                             g["sub_account_name"].lower(), g["issue_code"]))
    for g in rows:
        result["tag_counts"][g["tag"]] += 1
    result["groups"] = rows
    return result


@router.get("/priority")
async def priority() -> dict:
    """Priority List for the latest sync (ADR-010)."""
    return _priority_list()


@router.get("/leaderboard")
async def leaderboard() -> dict:
    """Latest-snapshot sub-accounts at or above the disapproval-rate threshold, worst first."""
    return {"threshold": settings.leaderboard_threshold, "leaderboard": _leaderboard_rows()}


def _build_report(meta: dict) -> BytesIO:
    """Assemble the report workbook from the cached data of the last sync.

    The single place ``report.build_workbook`` is called. Both the browser download and
    the post-sync drop go through here so the emailed file and the downloaded file can
    never be assembled differently (ADR-007).
    """
    overview_rows = _overview_rows()
    top_issues, top_issue_totals = _top_issues()
    return report.build_workbook(
        overview=overview_rows,
        overview_kpis=_overview_kpis(overview_rows),
        thresholds=_thresholds(),
        spike_radar=_spike_radar(),
        top_issues=top_issues,
        top_issue_totals=top_issue_totals,
        pending_issue_codes=settings.pending_issue_codes,
        priority=_priority_list(),
        products=db.all_disapproval_rows(),
        history=_history_rows(),
        meta=meta,
    )


@router.get("/report.xlsx")
async def report_xlsx() -> StreamingResponse:
    """Download the multi-sheet Excel report (ADR-006).

    Availability follows **sync state**, not leaderboard state: an empty board is this
    app's goal state and must still produce a workbook, but with no sync at all there is
    no data of any kind to report on.
    """
    meta = db.get_sync_meta()
    if not meta["last_synced_at"]:
        raise HTTPException(status_code=409, detail="No sync yet — run a sync before exporting.")

    buf = _build_report(meta)
    logger.info("report downloaded (sync %s)", meta["last_synced_at"])
    return StreamingResponse(
        buf,
        media_type=XLSX_MEDIA_TYPE,
        headers={"Content-Disposition": 'attachment; filename="gmc-report.xlsx"'},
    )


@router.get("/history")
async def history(
    sub_account: list[str] = Query(default=[], description="Repeatable; one series per id"),
    days: int = Query(90, ge=1, le=3650, description="Look-back window in days"),
) -> dict:
    """Disapproval-rate time series for the selected sub-accounts, grouped one series per sub-account.

    A snapshot with no products has ``rate: null`` — a gap in the line, not 0 % (ADR-011 §2).
    """
    since = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()
    series: dict[str, dict] = {}
    order: list[str] = []
    for r in db.stats_history(sub_account, since):
        sid = r["sub_account_id"]
        if sid not in series:
            series[sid] = {
                "sub_account_id": sid,
                "sub_account_name": r["sub_account_name"],
                "points": [],
            }
            order.append(sid)
        series[sid]["points"].append(
            {
                "synced_at": r["synced_at"],
                "total_products": r["total_products"],
                "disapproved_products": r["disapproved_products"],
                "rate": _point_rate(r),
                "approval_rate": _approval_rate(_point_rate(r)),
            }
        )
    return {"days": days, "series": [series[s] for s in order]}


@router.get("/subaccounts")
async def subaccounts() -> dict:
    return {"subaccounts": db.distinct_subaccounts()}


@router.get("/issue-codes")
async def issue_codes() -> dict:
    return {"issue_codes": db.distinct_issue_codes()}
