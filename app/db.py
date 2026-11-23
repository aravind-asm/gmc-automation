"""SQLite cache for disapproved-product rows (cache-first per CONTEXT.md).

Single-operator, low-concurrency app: a module-level connection guarded by a lock
is sufficient. The DB stores only disapproved-product metadata — no secrets.
"""
import os
import sqlite3
import threading
from datetime import datetime, timedelta, timezone

from app.config import settings

_lock = threading.Lock()
_conn: sqlite3.Connection | None = None

_COLUMNS = [
    "sub_account_id",
    "sub_account_name",
    "offer_id",
    "product_title",
    "product_url",
    "reporting_context",
    "issue_code",
    "severity",
    "attribute",
    "description",
    "detail",
    "documentation_uri",
    "applicable_countries",
]


def _connect() -> sqlite3.Connection:
    global _conn
    if _conn is None:
        os.makedirs(os.path.dirname(settings.db_path) or ".", exist_ok=True)
        _conn = sqlite3.connect(settings.db_path, check_same_thread=False)
        _conn.row_factory = sqlite3.Row
        _init(_conn)
    return _conn


def _init(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS disapproved_products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sub_account_id TEXT NOT NULL,
            sub_account_name TEXT,
            offer_id TEXT,
            product_title TEXT,
            product_url TEXT,
            reporting_context TEXT,
            issue_code TEXT,
            severity TEXT,
            attribute TEXT,
            description TEXT,
            detail TEXT,
            documentation_uri TEXT,
            applicable_countries TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_dp_sub ON disapproved_products(sub_account_id);
        CREATE INDEX IF NOT EXISTS idx_dp_code ON disapproved_products(issue_code);
        CREATE TABLE IF NOT EXISTS sync_meta (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            last_synced_at TEXT,
            row_count INTEGER
        );
        -- Append-only per-sub-account rate snapshots (one row per sub-account per
        -- sync). Powers the Leaderboard (latest snapshot) and History (series).
        -- Rate is derived on read; never stored. See docs/adr/ADR-002.md.
        CREATE TABLE IF NOT EXISTS subaccount_stats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            synced_at TEXT NOT NULL,
            sub_account_id TEXT NOT NULL,
            sub_account_name TEXT,
            total_products INTEGER NOT NULL,
            disapproved_products INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_stats_sub_time
            ON subaccount_stats(sub_account_id, synced_at);
        -- Account warnings of the latest sync (ADR-009 §3), wiped and rewritten each sync
        -- like disapproved_products. One check row per store says whether its check worked
        -- (ok = 0 means "unknown", not "no warnings"); the issues hang off it.
        CREATE TABLE IF NOT EXISTS account_issue_checks (
            sub_account_id TEXT PRIMARY KEY,
            ok INTEGER NOT NULL CHECK (ok IN (0, 1))
        );
        CREATE TABLE IF NOT EXISTS account_issues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sub_account_id TEXT NOT NULL,
            severity TEXT,
            title TEXT,
            detail TEXT,
            documentation_uri TEXT
        );
        -- Performance data of the latest sync (ADR-010 §1), wiped and rewritten each sync like
        -- the account warnings: one check row per store says whether its report was read (ok = 0
        -- means "no performance data", not "no impressions") and over which days; the per-product
        -- figures hang off it. offer_id is stored lower-cased, as Google's report sends it.
        CREATE TABLE IF NOT EXISTS performance_checks (
            sub_account_id TEXT PRIMARY KEY,
            ok INTEGER NOT NULL CHECK (ok IN (0, 1)),
            window_start TEXT,
            window_end TEXT
        );
        CREATE TABLE IF NOT EXISTS product_performance (
            sub_account_id TEXT NOT NULL,
            offer_id TEXT NOT NULL,
            marketing_method TEXT NOT NULL,
            impressions INTEGER NOT NULL,
            clicks INTEGER NOT NULL,
            PRIMARY KEY (sub_account_id, offer_id, marketing_method)
        );
        """
    )
    # issue_first_seen (ADR-010 §2): when each problem — (store, product, issue code) — was first
    # and last seen. Upserted every sync and NEVER wiped: the one product-level table that
    # survives the ADR-002 refresh, so "open for N days" can be answered.
    has_first_seen = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'issue_first_seen'"
    ).fetchone()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS issue_first_seen (
            sub_account_id TEXT NOT NULL,
            offer_id TEXT NOT NULL,
            issue_code TEXT NOT NULL,
            first_seen_at TEXT NOT NULL,
            last_seen_at TEXT NOT NULL,
            PRIMARY KEY (sub_account_id, offer_id, issue_code)
        )
        """
    )
    if not has_first_seen:
        # Deployed onto a live cache: the problems already stored were seen at its last sync.
        # That is a real observation, not a guess — older history was cleared on purpose
        # (2026-11-02) and is not backfilled.
        conn.execute(
            "INSERT OR IGNORE INTO issue_first_seen "
            "(sub_account_id, offer_id, issue_code, first_seen_at, last_seen_at) "
            "SELECT DISTINCT d.sub_account_id, d.offer_id, d.issue_code, m.last_synced_at, m.last_synced_at "
            "FROM disapproved_products d JOIN sync_meta m ON m.id = 1 "
            "WHERE d.issue_code <> '' AND m.last_synced_at IS NOT NULL"
        )

    # Idempotent migration: older caches predate product_url. replace_disapprovals()
    # inserts every _COLUMNS field, so the column must exist or inserts fail.
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(disapproved_products)")}
    if "product_url" not in cols:
        conn.execute("ALTER TABLE disapproved_products ADD COLUMN product_url TEXT")
    conn.commit()


def init_db() -> None:
    with _lock:
        _connect()


def replace_disapprovals(
    rows: list[dict],
    stats: list[dict] | None = None,
    checks: list[dict] | None = None,
    performance: list[dict] | None = None,
    *,
    synced_at: str | None = None,
) -> None:
    """Atomically swap the detail cache and append this sync's stats snapshot.

    ``rows`` replaces the (wiped-each-sync) detail table; ``stats`` are appended to the
    time series in ``subaccount_stats``; ``checks`` (``{sub_account_id, ok, issues}`` per
    store) replace the account warnings. All of them, plus ``sync_meta``, are written in one
    transaction sharing a single ``synced_at`` so a sync is all-or-nothing. Warnings are
    wiped even when ``checks`` is empty, so they always describe the latest sync. Each problem
    in ``rows`` is upserted into ``issue_first_seen``: its first sighting is kept, its last
    sighting moves to this sync. ``performance`` (``{sub_account_id, ok, window_start, window_end,
    rows}`` per store) replaces the performance data, also wiped when empty. ``synced_at``
    defaults to now (tests pass fixed times).
    """
    stats = stats or []
    checks = checks or []
    performance = performance or []
    synced_at = synced_at or datetime.now(timezone.utc).isoformat()
    with _lock:
        conn = _connect()
        placeholders = ",".join(["?"] * len(_COLUMNS))
        col_list = ",".join(_COLUMNS)
        try:
            conn.execute("BEGIN")
            conn.execute("DELETE FROM disapproved_products")
            conn.executemany(
                f"INSERT INTO disapproved_products ({col_list}) VALUES ({placeholders})",
                [tuple(r.get(c, "") for c in _COLUMNS) for r in rows],
            )
            conn.executemany(
                "INSERT INTO subaccount_stats "
                "(synced_at, sub_account_id, sub_account_name, total_products, disapproved_products) "
                "VALUES (?, ?, ?, ?, ?)",
                [
                    (
                        synced_at,
                        s["sub_account_id"],
                        s.get("sub_account_name", ""),
                        int(s["total_products"]),
                        int(s["disapproved_products"]),
                    )
                    for s in stats
                ],
            )
            conn.execute("DELETE FROM account_issues")
            conn.execute("DELETE FROM account_issue_checks")
            conn.executemany(
                "INSERT INTO account_issue_checks (sub_account_id, ok) VALUES (?, ?)",
                [(c["sub_account_id"], 1 if c["ok"] else 0) for c in checks],
            )
            conn.executemany(
                "INSERT INTO account_issues "
                "(sub_account_id, severity, title, detail, documentation_uri) VALUES (?, ?, ?, ?, ?)",
                [
                    (
                        c["sub_account_id"],
                        i.get("severity", ""),
                        i.get("title", ""),
                        i.get("detail", ""),
                        i.get("documentation_uri", ""),
                    )
                    for c in checks
                    for i in c["issues"]
                ],
            )
            conn.execute("DELETE FROM product_performance")
            conn.execute("DELETE FROM performance_checks")
            conn.executemany(
                "INSERT INTO performance_checks (sub_account_id, ok, window_start, window_end) "
                "VALUES (?, ?, ?, ?)",
                [(p["sub_account_id"], 1 if p["ok"] else 0, p.get("window_start"), p.get("window_end"))
                 for p in performance],
            )
            conn.executemany(
                "INSERT INTO product_performance "
                "(sub_account_id, offer_id, marketing_method, impressions, clicks) VALUES (?, ?, ?, ?, ?) "
                # The client already sums per key; this only guards a repeat, never drops one.
                "ON CONFLICT(sub_account_id, offer_id, marketing_method) DO UPDATE SET "
                "impressions = impressions + excluded.impressions, clicks = clicks + excluded.clicks",
                [
                    (p["sub_account_id"], r["offer_id"].lower(), r["marketing_method"],
                     int(r["impressions"]), int(r["clicks"]))
                    for p in performance
                    if p["ok"]
                    for r in p["rows"]
                ],
            )
            conn.executemany(
                "INSERT INTO issue_first_seen "
                "(sub_account_id, offer_id, issue_code, first_seen_at, last_seen_at) VALUES (?, ?, ?, ?, ?) "
                "ON CONFLICT(sub_account_id, offer_id, issue_code) DO UPDATE SET last_seen_at = excluded.last_seen_at",
                [
                    (sid, offer, code, synced_at, synced_at)
                    for sid, offer, code in {
                        (r["sub_account_id"], r.get("offer_id") or "", r.get("issue_code") or "") for r in rows
                    }
                    if code
                ],
            )
            conn.execute(
                "INSERT INTO sync_meta (id, last_synced_at, row_count) VALUES (1, ?, ?) "
                "ON CONFLICT(id) DO UPDATE SET last_synced_at=excluded.last_synced_at, "
                "row_count=excluded.row_count",
                (synced_at, len(rows)),
            )
            conn.commit()
        except Exception:
            conn.rollback()
            raise


_ISSUE_FIELDS = [
    "id",
    "issue_code",
    "severity",
    "attribute",
    "description",
    "detail",
    "documentation_uri",
    "reporting_context",
    "applicable_countries",
]


def query_products(
    sub_account: str | None = None,
    issue_code: str | None = None,
    q: str | None = None,
    limit: int = 100,
    offset: int = 0,
) -> tuple[list[dict], int]:
    """Return (products, total_count): one product per ``(sub_account_id, offer_id)``, each
    carrying its issues stacked in ``issues[]``.

    The detail table is stored flat (one row per product + issue signature; see ADR-003/004);
    this groups it on read. Pagination and ``total`` are by **product**. The ``issue_code``
    filter selects products that have >=1 issue with that code, then shows **all** of that
    product's issues (product-centric view). Parameterized throughout.
    """
    where: list[str] = []
    args: list[object] = []
    if sub_account:
        where.append("sub_account_id = ?")
        args.append(sub_account)
    if q:
        where.append("(product_title LIKE ? OR offer_id LIKE ?)")
        like = f"%{q}%"
        args.extend([like, like])
    clause = f"WHERE {' AND '.join(where)}" if where else ""

    having = ""
    having_args: list[object] = []
    if issue_code:
        having = "HAVING SUM(CASE WHEN issue_code = ? THEN 1 ELSE 0 END) > 0"
        having_args.append(issue_code)

    with _lock:
        conn = _connect()
        total = conn.execute(
            f"SELECT COUNT(*) AS n FROM (SELECT 1 FROM disapproved_products {clause} "
            f"GROUP BY sub_account_id, offer_id {having})",
            [*args, *having_args],
        ).fetchone()["n"]
        keys = conn.execute(
            "SELECT sub_account_id, offer_id, MAX(sub_account_name) AS sub_account_name, "
            "MAX(product_title) AS product_title, MAX(product_url) AS product_url "
            f"FROM disapproved_products {clause} "
            f"GROUP BY sub_account_id, offer_id {having} "
            "ORDER BY sub_account_name, product_title LIMIT ? OFFSET ?",
            [*args, *having_args, limit, offset],
        ).fetchall()
        if not keys:
            return [], total
        pairs = [(k["sub_account_id"], k["offer_id"]) for k in keys]
        values = ",".join(["(?,?)"] * len(pairs))
        flat = [x for pair in pairs for x in pair]
        issue_rows = conn.execute(
            "SELECT * FROM disapproved_products "
            f"WHERE (sub_account_id, offer_id) IN (VALUES {values}) "
            "ORDER BY issue_code",
            flat,
        ).fetchall()

    products: dict[tuple, dict] = {}
    order: list[tuple] = []
    for k in keys:
        pk = (k["sub_account_id"], k["offer_id"])
        products[pk] = {
            "sub_account_id": k["sub_account_id"],
            "sub_account_name": k["sub_account_name"],
            "offer_id": k["offer_id"],
            "product_title": k["product_title"],
            "product_url": k["product_url"] or "",
            "issues": [],
        }
        order.append(pk)
    for r in issue_rows:
        p = products.get((r["sub_account_id"], r["offer_id"]))
        if p is not None:
            p["issues"].append({f: r[f] for f in _ISSUE_FIELDS})
    return [products[pk] for pk in order], total


def all_disapproval_rows() -> list[dict]:
    """Every stored disapproval row, flat and unpaginated — the shape the Excel export needs.

    ``query_products()`` groups by product and paginates for the UI (ADR-004); the export is
    whole-dataset and deliberately flat, one row per product + issue signature (ADR-006 S4),
    so it cannot reuse that reader. Ordered to match the Disapprovals page's reading order.
    """
    with _lock:
        conn = _connect()
        rows = conn.execute(
            f"SELECT {', '.join(_COLUMNS)} FROM disapproved_products "
            "ORDER BY sub_account_name, product_title, issue_code"
        ).fetchall()
    return [dict(r) for r in rows]


def product_issue_codes() -> list[dict]:
    """Each distinct (store, product, issue code, attribute, description) in the latest sync.

    The raw material for the Top 10 Issues view. A product listed once per surface-collapsed
    signature (ADR-003) can repeat here with a different attribute; the caller counts products,
    not rows. Small (one row per disapproved product and reason), so it is aggregated in Python.
    """
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT DISTINCT sub_account_id, offer_id, issue_code, attribute, description "
            "FROM disapproved_products WHERE issue_code <> ''"
        ).fetchall()
    return [dict(r) for r in rows]


def priority_rows() -> list[dict]:
    """Each distinct (store, product, issue, attribute, description) in the latest sync, with
    the problem's first sighting from ``issue_first_seen`` (None if never recorded)."""
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT DISTINCT d.sub_account_id, d.sub_account_name, d.offer_id, d.issue_code, d.attribute, "
            "d.description, f.first_seen_at "
            "FROM disapproved_products d LEFT JOIN issue_first_seen f "
            "ON f.sub_account_id = d.sub_account_id AND f.offer_id = d.offer_id AND f.issue_code = d.issue_code "
            "WHERE d.issue_code <> ''"
        ).fetchall()
    return [dict(r) for r in rows]


def performance_checks() -> dict[str, dict]:
    """The latest sync's performance fetches: ``{sub_account_id: {ok, window_start, window_end}}``.

    A store with no entry was not fetched (a sync from before performance data existed).
    """
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT sub_account_id, ok, window_start, window_end FROM performance_checks"
        ).fetchall()
    return {
        r["sub_account_id"]: {"ok": bool(r["ok"]), "window_start": r["window_start"], "window_end": r["window_end"]}
        for r in rows
    }


def disapproved_product_performance() -> dict[tuple[str, str], dict]:
    """Impressions and clicks of every disapproved product in the latest sync, keyed by
    ``(sub_account_id, lower(offer_id))``: totals plus the Ads (``ADS``) / Free listings
    (``ORGANIC``) split; ``AFFILIATE`` and any other method count in the totals only (ADR-010 §1).

    Google's report lower-cases offer IDs (plan 0008 S1), so products are matched on
    ``lower(offer_id)``. A product with no row is simply absent: the caller reads it as 0 when its
    store's report was read. Grouped in SQL: the table holds every product's reach, not only the
    disapproved ones.
    """
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT d.sub_account_id, d.offer, "
            "SUM(p.impressions) AS impressions, SUM(p.clicks) AS clicks, "
            "SUM(CASE WHEN p.marketing_method = 'ADS' THEN p.impressions ELSE 0 END) AS ads_impressions, "
            "SUM(CASE WHEN p.marketing_method = 'ORGANIC' THEN p.impressions ELSE 0 END) AS free_impressions, "
            "SUM(CASE WHEN p.marketing_method = 'ADS' THEN p.clicks ELSE 0 END) AS ads_clicks, "
            "SUM(CASE WHEN p.marketing_method = 'ORGANIC' THEN p.clicks ELSE 0 END) AS free_clicks "
            "FROM (SELECT DISTINCT sub_account_id, lower(offer_id) AS offer FROM disapproved_products) d "
            "JOIN product_performance p ON p.sub_account_id = d.sub_account_id AND p.offer_id = d.offer "
            "GROUP BY d.sub_account_id, d.offer"
        ).fetchall()
    return {
        (r["sub_account_id"], r["offer"]): {
            k: r[k] for k in ("impressions", "clicks", "ads_impressions", "free_impressions", "ads_clicks", "free_clicks")
        }
        for r in rows
    }


def tracking_since() -> str | None:
    """When ``issue_first_seen`` began: its earliest first sighting (None before any)."""
    with _lock:
        conn = _connect()
        row = conn.execute("SELECT MIN(first_seen_at) AS t FROM issue_first_seen").fetchone()
    return row["t"] if row else None


def distinct_subaccounts() -> list[dict]:
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT DISTINCT sub_account_id, sub_account_name FROM disapproved_products "
            "ORDER BY sub_account_name"
        ).fetchall()
    return [dict(r) for r in rows]


def distinct_issue_codes() -> list[str]:
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT DISTINCT issue_code FROM disapproved_products "
            "WHERE issue_code <> '' ORDER BY issue_code"
        ).fetchall()
    return [r["issue_code"] for r in rows]


def latest_stats() -> list[dict]:
    """Return each sub-account's most recent snapshot (the latest sync group).

    All rows of a sync share one ``synced_at``, so the latest snapshot is simply
    every row at ``MAX(synced_at)``. Rate is derived by the caller.
    """
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT synced_at, sub_account_id, sub_account_name, total_products, "
            "disapproved_products FROM subaccount_stats "
            "WHERE synced_at = (SELECT MAX(synced_at) FROM subaccount_stats)"
        ).fetchall()
    return [dict(r) for r in rows]


def latest_account_checks() -> dict[str, dict]:
    """The latest sync's account warnings: ``{sub_account_id: {"ok": bool, "issues": [...]}}``.

    A store with no entry was not checked (a sync from before warnings were fetched). Issues
    come back in insertion order; the Overview decides which severities count and how to sort.
    """
    with _lock:
        conn = _connect()
        checks = conn.execute("SELECT sub_account_id, ok FROM account_issue_checks").fetchall()
        issues = conn.execute(
            "SELECT sub_account_id, severity, title, detail, documentation_uri "
            "FROM account_issues ORDER BY id"
        ).fetchall()
    result = {r["sub_account_id"]: {"ok": bool(r["ok"]), "issues": []} for r in checks}
    for i in issues:
        entry = result.get(i["sub_account_id"])
        if entry is not None:
            entry["issues"].append(
                {
                    "severity": i["severity"],
                    "title": i["title"],
                    "detail": i["detail"],
                    "documentation_uri": i["documentation_uri"],
                }
            )
    return result


# Korea has one time zone and no daylight saving, so a fixed offset is exact (and needs no
# tz database in the container). The business runs on Korea time, so "a day" means a KST day
# (ADR-011 §1; it was India time until plan 0009).
KST = timezone(timedelta(hours=9))


def kst_day(synced_at: str) -> str:
    """The KST calendar day (``YYYY-MM-DD``) of a stored UTC timestamp."""
    return datetime.fromisoformat(synced_at).astimezone(KST).date().isoformat()


def daily_snapshots() -> list[dict]:
    """The last sync of each KST calendar day, newest day first: ``[{kst_day, synced_at}]``.

    Several syncs on one day collapse to the latest (ADR-009 §5).
    """
    with _lock:
        conn = _connect()
        times = [r["synced_at"] for r in conn.execute("SELECT DISTINCT synced_at FROM subaccount_stats")]
    last_per_day: dict[str, str] = {}
    for t in times:
        day = kst_day(t)
        if day not in last_per_day or t > last_per_day[day]:
            last_per_day[day] = t
    return [{"kst_day": d, "synced_at": last_per_day[d]} for d in sorted(last_per_day, reverse=True)]


def stats_at(synced_at: str) -> list[dict]:
    """Every store's snapshot from one sync."""
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT synced_at, sub_account_id, sub_account_name, total_products, disapproved_products "
            "FROM subaccount_stats WHERE synced_at = ?",
            (synced_at,),
        ).fetchall()
    return [dict(r) for r in rows]


def stats_subaccount_ids() -> list[str]:
    """Every sub-account with at least one recorded rate."""
    with _lock:
        conn = _connect()
        rows = conn.execute("SELECT DISTINCT sub_account_id FROM subaccount_stats ORDER BY sub_account_id").fetchall()
    return [r["sub_account_id"] for r in rows]


def stats_history(sub_account_ids: list[str], since: str | None = None) -> list[dict]:
    """Return the snapshot series for the given sub-accounts, oldest first.

    ``since`` is an ISO timestamp lower bound (inclusive). An empty id list
    returns nothing. Parameterized throughout — ids are bound, never interpolated.
    """
    if not sub_account_ids:
        return []
    placeholders = ",".join(["?"] * len(sub_account_ids))
    clause = f"sub_account_id IN ({placeholders})"
    args: list[object] = list(sub_account_ids)
    if since:
        clause += " AND synced_at >= ?"
        args.append(since)
    with _lock:
        conn = _connect()
        rows = conn.execute(
            "SELECT synced_at, sub_account_id, sub_account_name, total_products, "
            f"disapproved_products FROM subaccount_stats WHERE {clause} "
            "ORDER BY sub_account_id, synced_at",
            args,
        ).fetchall()
    return [dict(r) for r in rows]


def get_sync_meta() -> dict:
    with _lock:
        conn = _connect()
        row = conn.execute("SELECT last_synced_at, row_count FROM sync_meta WHERE id = 1").fetchone()
    return dict(row) if row else {"last_synced_at": None, "row_count": 0}
