"""Application configuration, loaded from environment / .env.

Auth model: GCP service account (Admin on the parent MCA). See docs/adr/ADR-001.md.
"""
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

    # Path (inside the container) to the mounted service-account JSON key.
    gcp_sa_key_path: str = "/app/secrets/sa-key.json"
    # OAuth scope required for the Merchant API.
    gcp_scope: str = "https://www.googleapis.com/auth/content"
    # Numeric account ID of the parent Multi-Client Account.
    parent_mca_id: str = ""

    # Merchant API base URL (v1 — v1beta was discontinued 2026-02-28).
    merchant_api_base: str = "https://merchantapi.googleapis.com"

    # Local SQLite cache path.
    db_path: str = "/app/data/gmc_cache.sqlite"

    # Directory (inside the container) where every successful sync drops its Excel
    # report. The host side is a OneDrive-synced folder watched by a Power Automate
    # flow that emails the file onward. Empty = feature off. See docs/adr/ADR-007.md.
    report_drop_dir: str = ""

    # --- Automatic daily sync (ADR-012) ---
    # Off by default: every sync emails the team, so only the operator switches this on, in .env
    # (AUTO_SYNC_ENABLED=true). When on, the backend syncs once per KST day at or after
    # auto_sync_time_kst ("HH:MM", Korea time; 13:00 KST = 09:30 IST), catching up later the same
    # KST day if the laptop was asleep or off, unless a sync after that time already succeeded.
    # A failed attempt is retried every auto_sync_retry_minutes until the KST day ends.
    auto_sync_enabled: bool = False
    auto_sync_time_kst: str = "13:00"
    auto_sync_retry_minutes: int = Field(30, ge=1)
    # How often the backend asks "is a sync due?". Also the longest delay after the laptop wakes.
    auto_sync_check_minutes: int = Field(5, ge=1)

    # Max concurrent per-sub-account fetches during a sync.
    sync_concurrency: int = 5

    # --- Logging (see app/logging_config.py) ---
    # Root log level: DEBUG, INFO, WARNING or ERROR.
    log_level: str = "INFO"
    # Directory (inside the container) for the rotating log file. Sits under the
    # bind-mounted /app/data, so logs survive container rebuilds like the SQLite cache.
    log_dir: str = "/app/data/logs"
    log_max_bytes: int = 5_000_000
    log_backup_count: int = 5

    # Disapproval rate above which a sub-account appears on the Leaderboard, and
    # which the History page pre-selects by default. Single source of truth for
    # the "5%" figure (see docs/adr/ADR-002.md).
    leaderboard_threshold: float = 0.05

    # --- Overview status (ADR-009 §2, four bands since ADR-011 §3) ---
    # Healthy under leaderboard_threshold (5 %) · Minor from 5 % · Warning from
    # status_warning_threshold (8 %) · Critical from status_critical_threshold (10 %).
    # All compared against the rate rounded to 1 decimal place, as displayed; bottoms inclusive.
    status_warning_threshold: float = 0.08
    status_critical_threshold: float = 0.10

    # --- Spike Radar (ADR-009 §5) ---
    # A store is a spike when it has at least this many disapproved products now AND its rate
    # rose by at least this many percentage points since the last sync of an earlier KST day.
    # Both inclusive; compared on rates rounded to 1 dp, as shown.
    spike_min_disapproved: int = 30
    spike_min_diff_pp: float = 5.0

    # Issue codes that mean "Google has not finished reviewing", not "something is wrong"
    # (ADR-010 §3). Still counted, so totals match Google's, but tagged "Pending review — not a
    # fault" so nobody spends time fixing them. In .env, a JSON list: PENDING_ISSUE_CODES=["a","b"].
    pending_issue_codes: list[str] = [
        "attribute_pending_review",
        "pending_initial_policy_review_shopping_ads",
        "image_link_pending_crawl",
    ]

    # --- Priority List (ADR-010 §3, four levels since ADR-011 §6) ---
    # Groups that can be ranked (not pending review, their store's performance data was read) are
    # ranked by impressions over the last 30 days. Critical = rank <= ceil(critical share × ranked
    # groups), High = rank <= ceil(high share × ranked groups), Medium = the rest with impressions.
    # Groups that tie share a rank, so a tie is never split across two levels; 0 impressions is
    # always Low; a store whose performance data could not be read is Medium (not ranked).
    priority_critical_share: float = 0.10
    priority_high_share: float = 0.30

    # Store Directory CSV (ADR-009 §1). Empty = the committed app/store_directory.csv.
    store_directory_path: str = ""


settings = Settings()
