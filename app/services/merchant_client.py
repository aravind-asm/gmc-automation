"""Async Merchant API (v1) client — sub-account discovery, disapproved products, account warnings.

Data model note (verified 2026-07-01 against the live account):
- Sub-accounts:  GET accounts/v1/accounts/{parent}:listSubaccounts
- Account warnings: GET accounts/v1/accounts/{sub}/issues → accountIssues[] with severity
  CRITICAL / ERROR / SUGGESTION (verified 2026-11-10 for all 49 stores, incl. the six outside
  the MCA; a store with none answers ``{}``).
- Performance:   POST reports/v1/accounts/{sub}/reports:search on product_performance_view →
  impressions and clicks per offer ID and marketing method (verified 2026-11-10, plan 0008 S1:
  numbers arrive as strings, offer IDs come back lower-cased, a blank offer ID can appear).
- Disapprovals:  GET products/v1/accounts/{sub}/products  →  each product's
  productStatus.itemLevelIssues[] carries the human-readable reason
  (description/detail/documentation). A product is disapproved when an issue's
  severity == "DISAPPROVED". (aggregateProductStatuses/productStatuses 404 in v1;
  reports product_view returns codes only, so it is not used here.)
"""
import asyncio
import logging
import time
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone

import httpx

from app import db, stores
from app.config import settings
from app.services import merchant_auth

logger = logging.getLogger("gmc.merchant")

_DISAPPROVED = "DISAPPROVED"

PERFORMANCE_DAYS = 30  # "lost impression volume (past 30 days)" — the spec (ADR-010 §1)


def performance_window(today: date) -> tuple[date, date]:
    """The last ``PERFORMANCE_DAYS`` complete days before ``today``, both ends inclusive.

    Today is left out: its figures are still filling in.
    """
    end = today - timedelta(days=1)
    return end - timedelta(days=PERFORMANCE_DAYS - 1), end


def current_performance_window(now: datetime | None = None) -> tuple[date, date]:
    """The window for a sync made at ``now`` (default: this moment). "Today" is the KST day
    (ADR-011 §1), so the window ends the KST day before the sync."""
    now = now or datetime.now(timezone.utc)
    return performance_window(now.astimezone(db.KST).date())


def _count(value) -> int:
    """A report number: Google sends int64 metrics as strings ("26454"), absent when zero."""
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


@dataclass
class CallCounter:
    """Per-sync tally of Merchant API requests, shared across the concurrent workers.

    Plain ``+= 1`` is safe: asyncio runs one task at a time and there is no ``await``
    between reading and writing, so no lock is needed.
    """

    calls: int = 0
    failures: int = 0


class MerchantClient:
    def __init__(self) -> None:
        self.base = settings.merchant_api_base.rstrip("/")

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {merchant_auth.get_token()}"}

    async def _get(
        self,
        client: httpx.AsyncClient,
        url: str,
        *,
        params: dict,
        label: str,
        counter: "CallCounter | None" = None,
    ):
        """One Merchant API GET — timed, counted, and logged as a single line.

        ``label`` describes the call in human terms (which sub-account, which page).
        Headers are built here and never logged: they carry the OAuth bearer token.
        """
        started = time.perf_counter()
        response = await client.get(url, params=params, headers=self._headers())
        elapsed_ms = (time.perf_counter() - started) * 1000
        if counter is not None:
            counter.calls += 1
            if response.status_code != 200:
                counter.failures += 1
        logger.info(
            "merchant GET %s -> HTTP %s in %.0f ms", label, response.status_code, elapsed_ms
        )
        return response

    async def _post(
        self,
        client: httpx.AsyncClient,
        url: str,
        *,
        json: dict,
        label: str,
        counter: "CallCounter | None" = None,
    ):
        """One Merchant API POST — timed, counted and logged exactly like ``_get`` (ADR-005).

        Neither the headers (they carry the bearer token) nor the body are logged.
        """
        started = time.perf_counter()
        response = await client.post(url, json=json, headers=self._headers())
        elapsed_ms = (time.perf_counter() - started) * 1000
        if counter is not None:
            counter.calls += 1
            if response.status_code != 200:
                counter.failures += 1
        logger.info(
            "merchant POST %s -> HTTP %s in %.0f ms", label, response.status_code, elapsed_ms
        )
        return response

    async def list_subaccounts(
        self, client: httpx.AsyncClient, counter: "CallCounter | None" = None
    ) -> list[dict]:
        """Return [{account_id, account_name}] for every sub-account of the MCA."""
        parent = settings.parent_mca_id
        url = f"{self.base}/accounts/v1/accounts/{parent}:listSubaccounts"
        accounts: list[dict] = []
        page_token: str | None = None
        page = 0
        while True:
            params = {"pageSize": 250}
            if page_token:
                params["pageToken"] = page_token
            page += 1
            r = await self._get(
                client, url, params=params, label=f"listSubaccounts page {page}", counter=counter
            )
            r.raise_for_status()
            data = r.json()
            for a in data.get("accounts", []):
                accounts.append(
                    {
                        "account_id": str(a.get("accountId", "")),
                        "account_name": a.get("accountName", ""),
                    }
                )
            page_token = data.get("nextPageToken")
            if not page_token:
                break
        return accounts

    async def list_disapproved_products(
        self,
        client: httpx.AsyncClient,
        account: dict,
        counter: "CallCounter | None" = None,
    ) -> tuple[list[dict], int, int, bool]:
        """Fetch a sub-account's products and summarize disapprovals.

        Returns ``(rows, total_products, disapproved_products, fully_fetched)``:
        - ``rows`` — one normalized row per disapproving item-level issue (per-issue grain).
        - ``total_products`` — every product returned across all pages (the rate denominator).
        - ``disapproved_products`` — distinct products with >=1 DISAPPROVED issue (the numerator).
        - ``fully_fetched`` — True only if every page was retrieved without error. A partial
          fetch yields plausible-but-wrong totals, so callers must NOT record stats when False.
          See docs/adr/ADR-002.md.
        """
        account_id = account["account_id"]
        url = f"{self.base}/products/v1/accounts/{account_id}/products"
        rows: list[dict] = []
        total_products = 0
        disapproved_products = 0
        fully_fetched = True
        page_token: str | None = None
        page = 0
        label = f"products account={account_id} ({account.get('account_name', '')})"
        while True:
            params = {"pageSize": 1000}
            if page_token:
                params["pageToken"] = page_token
            page += 1
            r = await self._get(
                client, url, params=params, label=f"{label} page {page}", counter=counter
            )
            if r.status_code != 200:
                # Google puts the actual reason (quota, permission, bad account) in the
                # body — without it a failed sync is undiagnosable.
                logger.warning(
                    "products.list failed for sub-account %s: HTTP %s %s",
                    account_id,
                    r.status_code,
                    getattr(r, "text", "")[:200],
                )
                fully_fetched = False
                break
            data = r.json()
            for product in data.get("products", []):
                total_products += 1
                product_rows = self._extract_disapprovals(account, product)
                if product_rows:
                    disapproved_products += 1
                    rows.extend(product_rows)
            page_token = data.get("nextPageToken")
            if not page_token:
                break
        return rows, total_products, disapproved_products, fully_fetched

    async def list_account_issues(
        self,
        client: httpx.AsyncClient,
        account: dict,
        counter: "CallCounter | None" = None,
    ) -> tuple[list[dict], bool]:
        """Fetch a store's account-level warnings (ADR-009 §3).

        Returns ``(issues, ok)``. Each issue is ``{severity, title, detail, documentation_uri}``;
        every severity is kept — the Overview decides which ones count. ``ok`` is False when any
        page failed: a failed check must read as "unknown", never as "no warnings", so callers
        discard the partial list. A clean store comes back as an empty body ``{}`` (seen live).
        """
        account_id = account["account_id"]
        url = f"{self.base}/accounts/v1/accounts/{account_id}/issues"
        issues: list[dict] = []
        page_token: str | None = None
        page = 0
        label = f"accountIssues account={account_id} ({account.get('account_name', '')})"
        while True:
            params = {"pageSize": 100, "languageCode": "en"}
            if page_token:
                params["pageToken"] = page_token
            page += 1
            r = await self._get(client, url, params=params, label=f"{label} page {page}", counter=counter)
            if r.status_code != 200:
                logger.warning(
                    "accounts.issues failed for sub-account %s: HTTP %s %s",
                    account_id,
                    r.status_code,
                    getattr(r, "text", "")[:200],
                )
                return [], False
            data = r.json()
            for issue in data.get("accountIssues", []):
                issues.append(
                    {
                        "severity": issue.get("severity", ""),
                        "title": issue.get("title", ""),
                        "detail": issue.get("detail", ""),
                        "documentation_uri": issue.get("documentationUri", ""),
                    }
                )
            page_token = data.get("nextPageToken")
            if not page_token:
                break
        return issues, True

    async def list_product_performance(
        self,
        client: httpx.AsyncClient,
        account: dict,
        window: tuple[date, date],
        counter: "CallCounter | None" = None,
    ) -> tuple[list[dict], bool]:
        """Fetch a store's impressions and clicks over ``window`` (ADR-010 §1).

        Returns ``(rows, ok)``; each row is ``{offer_id, marketing_method, impressions, clicks}``
        summed over the window. Offer IDs are lower-cased (Google already sends them so; the
        products side is matched on ``lower(offer_id)``) and a row with a blank offer ID is
        dropped — that reach belongs to no product. ``ok`` is False when any page failed: partial
        reach would rank a store's problems too low, so callers get no rows at all instead.
        """
        account_id = account["account_id"]
        url = f"{self.base}/reports/v1/accounts/{account_id}/reports:search"
        start, end = window
        # Both dates are ``date`` objects formatted here, never user input.
        query = (
            "SELECT offer_id, marketing_method, impressions, clicks "
            "FROM product_performance_view "
            f"WHERE date BETWEEN '{start.isoformat()}' AND '{end.isoformat()}'"
        )
        totals: dict[tuple[str, str], list[int]] = {}
        page_token: str | None = None
        page = 0
        label = f"performance account={account_id} ({account.get('account_name', '')})"
        while True:
            body: dict = {"query": query, "pageSize": 1000}
            if page_token:
                body["pageToken"] = page_token
            page += 1
            r = await self._post(client, url, json=body, label=f"{label} page {page}", counter=counter)
            if r.status_code != 200:
                logger.warning(
                    "reports.search failed for sub-account %s: HTTP %s %s",
                    account_id,
                    r.status_code,
                    getattr(r, "text", "")[:200],
                )
                return [], False
            data = r.json()
            for result in data.get("results", []):
                view = result.get("productPerformanceView") or {}
                offer_id = (view.get("offerId") or "").strip().lower()
                if not offer_id:
                    continue
                key = (offer_id, view.get("marketingMethod") or "")
                entry = totals.setdefault(key, [0, 0])
                entry[0] += _count(view.get("impressions"))
                entry[1] += _count(view.get("clicks"))
            page_token = data.get("nextPageToken")
            if not page_token:
                break
        rows = [
            {"offer_id": offer, "marketing_method": method, "impressions": imp, "clicks": clicks}
            for (offer, method), (imp, clicks) in totals.items()
        ]
        return rows, True

    @staticmethod
    def _extract_disapprovals(account: dict, product: dict) -> list[dict]:
        status = product.get("productStatus") or {}
        issues = status.get("itemLevelIssues") or []
        disapproving = [i for i in issues if i.get("severity") == _DISAPPROVED]
        if not disapproving:
            return []
        attrs = product.get("productAttributes") or {}
        title = attrs.get("title") or product.get("offerId") or ""
        offer_id = product.get("offerId") or ""
        product_url = attrs.get("link") or ""  # product landing page (ProductAttributes.link)

        # Collapse per-surface rows into one row per distinct issue signature.
        # The Merchant API repeats the same issue code once per reportingContext
        # (the Google surface the product runs on), so a single problem returns
        # several near-identical entries. Group by everything that makes an issue
        # "the same problem" except the surface, then merge the surfaces (and their
        # countries) into one row. See docs/adr/ADR-003.md.
        groups: dict[tuple, dict] = {}
        for issue in disapproving:
            code = issue.get("code", "")
            attribute = issue.get("attribute", "")
            description = issue.get("description", "")
            detail = issue.get("detail", "")
            documentation = issue.get("documentation", "")
            signature = (code, attribute, description, detail, documentation)
            group = groups.get(signature)
            if group is None:
                group = {
                    "sub_account_id": account["account_id"],
                    "sub_account_name": account["account_name"],
                    "offer_id": offer_id,
                    "product_title": title,
                    "product_url": product_url,
                    "issue_code": code,
                    "severity": _DISAPPROVED,
                    "attribute": attribute,
                    "description": description,
                    "detail": detail,
                    "documentation_uri": documentation,
                    "_surfaces": set(),
                    "_countries": set(),
                }
                groups[signature] = group
            context = issue.get("reportingContext")
            if context:
                group["_surfaces"].add(context)
            group["_countries"].update(issue.get("applicableCountries", []))

        rows = []
        for group in groups.values():  # dict preserves first-seen signature order
            surfaces = group.pop("_surfaces")
            countries = group.pop("_countries")
            group["reporting_context"] = ",".join(sorted(surfaces))
            group["applicable_countries"] = ",".join(sorted(countries))
            rows.append(group)
        return rows

    async def fetch_all_disapprovals(self) -> tuple[list[dict], list[dict], list[dict], list[dict]]:
        """Discover sub-accounts and gather disapprovals, account warnings and performance data.

        The stores fetched are the parent MCA's sub-accounts **plus** any store in the Store
        Directory that the MCA did not list — some of the operator's stores live outside it
        (ADR-009 §1 amendment). Each extra store costs at least one more call.

        Returns ``(rows, stats, checks, performance)``:
        - ``rows`` — flat per-issue disapproval rows (best-effort; partial fetches contribute
          whatever rows were retrieved, preserving the detail-table behavior).
        - ``stats`` — one aggregate ``{sub_account_id, sub_account_name, total_products,
          disapproved_products}`` per **fully-fetched** sub-account. Sub-accounts whose fetch
          errored are omitted so a rate is never computed from partial data (ADR-002).
        - ``checks`` — one ``{sub_account_id, ok, issues}`` per store: its account warnings
          (ADR-009 §3). A failed check is ``ok=False`` with no issues and never fails the sync.
        - ``performance`` — one ``{sub_account_id, ok, window_start, window_end, rows}`` per
          store: its impressions and clicks over the last 30 complete KST days (ADR-010 §1).
          A failed fetch is ``ok=False`` with no rows and never fails the sync either.
        """
        counter = CallCounter()
        issue_counter = CallCounter()  # kept apart so the summary can show what S4 adds
        perf_counter = CallCounter()   # ...and what S8 adds
        window = current_performance_window()
        started = time.perf_counter()
        limits = httpx.Limits(max_connections=settings.sync_concurrency * 2)
        async with httpx.AsyncClient(timeout=60, limits=limits) as client:
            subaccounts = await self.list_subaccounts(client, counter)
            extras = stores.extra_accounts({a["account_id"] for a in subaccounts})
            if extras:
                logger.info(
                    "store directory adds %d store(s) outside the parent account: %s",
                    len(extras),
                    ", ".join(a["account_id"] for a in extras),
                )
            subaccounts = subaccounts + extras
            sem = asyncio.Semaphore(settings.sync_concurrency)

            async def check_issues(acct: dict) -> dict:
                """The store's warning check, isolated: whatever goes wrong here costs this
                one signal, never the store's products or the sync (ADR-009 §3)."""
                try:
                    issues, ok = await self.list_account_issues(client, acct, issue_counter)
                except Exception:
                    logger.exception(
                        "error fetching account issues for sub-account %s", acct["account_id"]
                    )
                    issues, ok = [], False
                return {"sub_account_id": acct["account_id"], "ok": ok, "issues": issues}

            async def check_performance(acct: dict) -> dict:
                """The store's reach, isolated like the warning check: a failure costs this one
                store its impressions and clicks, never its products or the sync."""
                try:
                    perf_rows, ok = await self.list_product_performance(client, acct, window, perf_counter)
                except Exception:
                    logger.exception(
                        "error fetching performance data for sub-account %s", acct["account_id"]
                    )
                    perf_rows, ok = [], False
                return {
                    "sub_account_id": acct["account_id"],
                    "ok": ok,
                    "window_start": window[0].isoformat(),
                    "window_end": window[1].isoformat(),
                    "rows": perf_rows,
                }

            async def worker(acct: dict) -> tuple[list[dict], dict | None, dict, dict]:
                async with sem:
                    try:
                        rows, total, disapproved, ok = await self.list_disapproved_products(
                            client, acct, counter
                        )
                    except Exception:
                        logger.exception(
                            "error fetching disapprovals for sub-account %s",
                            acct["account_id"],
                        )
                        rows, ok = [], False
                    stat = None
                    if ok:
                        stat = {
                            "sub_account_id": acct["account_id"],
                            "sub_account_name": acct["account_name"],
                            "total_products": total,
                            "disapproved_products": disapproved,
                        }
                    return rows, stat, await check_issues(acct), await check_performance(acct)

            results = await asyncio.gather(*(worker(a) for a in subaccounts))
        rows = [row for sub_rows, _, _, _ in results for row in sub_rows]
        stats = [stat for _, stat, _, _ in results if stat is not None]
        checks = [check for _, _, check, _ in results]
        performance = [perf for _, _, _, perf in results]
        logger.info(
            "sync complete: %d Merchant API calls (%d failed) in %.1f s; "
            "%d disapproved rows; stats recorded for %d/%d sub-accounts; "
            "account warnings checked for %d/%d sub-accounts (%d calls, %d failed); "
            "performance data read for %d/%d sub-accounts (%d calls, %d failed)",
            counter.calls + issue_counter.calls + perf_counter.calls,
            counter.failures + issue_counter.failures + perf_counter.failures,
            time.perf_counter() - started,
            len(rows),
            len(stats),
            len(subaccounts),
            sum(1 for c in checks if c["ok"]),
            len(subaccounts),
            issue_counter.calls,
            issue_counter.failures,
            sum(1 for p in performance if p["ok"]),
            len(subaccounts),
            perf_counter.calls,
            perf_counter.failures,
        )
        return rows, stats, checks, performance


client = MerchantClient()
