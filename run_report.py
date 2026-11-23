"""GMC report — standalone runner (plan 0011, ADR-013).

The one file Windows Task Scheduler starts on a machine without Docker. Each run does what one
Refresh does on the laptop: fetch every store from the Merchant API, save the figures into the
history file, build the six-sheet Excel report and save it into the OneDrive folder, where a
Power Automate flow emails it. It has **no clock of its own**: every run makes one report, so
Task Scheduler decides when (and a run started by hand sends an email too).

It reuses the app's own code (``app/``, beside this file in the zip) without changing it, so the
report is built exactly as the laptop builds it. All settings live in ``settings.env`` next to
this file — see INSTRUCTIONS.md; nothing in this file needs editing.

    python run_report.py            fetch, save, report (the scheduled run)
    python run_report.py --check    check the setup; no Merchant API data calls, nothing saved
    python run_report.py --resave   rebuild the last run's report from the history file;
                                    no Google calls (use after a failed save)

Exit codes (Task Scheduler shows them as "Last Run Result"):
    0  done                                    (0x0)
    1  Google failed on every try; no report   (0x1)
    2  data saved, but the report file was not (0x2)
    3  not set up: key, account ID or folder   (0x3)
    4  another run is still going; not started (0x4)
    5  something unexpected broke; see the log (0x5)
"""
import argparse
import asyncio
import json
import logging
import os
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
SETTINGS_FILE = "settings.env"
VERSION_FILE = "VERSION.txt"
TRIGGER = "standalone"

EXIT_OK = 0
EXIT_FETCH_FAILED = 1
EXIT_NOT_SAVED = 2
EXIT_NOT_CONFIGURED = 3
EXIT_BUSY = 4
EXIT_CRASHED = 5

# Where things live when settings.env does not say otherwise, relative to this folder.
PATH_DEFAULTS = {
    "GCP_SA_KEY_PATH": "secrets/sa-key.json",
    "DB_PATH": "data/gmc_cache.sqlite",
    "LOG_DIR": "data/logs",
}
# Settings that are paths: a relative one is read as relative to this folder.
PATH_KEYS = (*PATH_DEFAULTS, "REPORT_DROP_DIR", "STORE_DIRECTORY_PATH")

# The runner's own settings (the app ignores them). Minimum value in brackets.
RUNNER_DEFAULTS = {
    "FETCH_TRIES": (3, 1),
    "FETCH_RETRY_MINUTES": (30, 0),
    "SAVE_TRIES": (3, 1),
    "SAVE_RETRY_MINUTES": (5, 0),
}

# The laptop's drop folder. Saving into a folder of this name on the same OneDrive account
# would make the laptop's flow email the team a second report (ADR-013).
LAPTOP_FOLDER_NAME = "GMC Reports"

sleep = time.sleep  # replaced in tests

logger = logging.getLogger("gmc.standalone")


# --- Settings ------------------------------------------------------------------------------------


def read_settings_file(path: Path) -> dict[str, str]:
    """``KEY=value`` lines from settings.env; blank values left out so the defaults apply."""
    if not path.is_file():
        return {}
    from dotenv import dotenv_values

    # utf-8-sig: Notepad and PowerShell may start the file with a byte-order mark.
    return {k: v.strip() for k, v in dotenv_values(path, encoding="utf-8-sig").items()
            if v is not None and v.strip()}


def prepare_environment(folder: Path, environ) -> None:
    """Put settings.env into ``environ`` (before ``app.config`` is imported, which reads it),
    then the folder-relative path defaults, then make every relative path absolute.

    A variable already set on the machine wins over settings.env, as usual for env files.
    """
    for key, value in read_settings_file(folder / SETTINGS_FILE).items():
        environ.setdefault(key, value)
    for key, value in PATH_DEFAULTS.items():
        if not environ.get(key, "").strip():
            environ[key] = value
    for key in PATH_KEYS:
        value = environ.get(key, "").strip()
        if value and not Path(value).is_absolute():
            environ[key] = str((folder / value).resolve())


def runner_setting(environ, key: str) -> int:
    default, minimum = RUNNER_DEFAULTS[key]
    raw = environ.get(key, "").strip()
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError:
        logger.warning("%s=%r is not a whole number; using %d", key, raw, default)
        return default
    if value < minimum:
        logger.warning("%s=%d is below %d; using %d", key, value, minimum, default)
        return default
    return value


def version(folder: Path) -> str:
    """The commit the zip was built from (VERSION.txt, written by build_zip.py)."""
    try:
        return (folder / VERSION_FILE).read_text(encoding="utf-8").splitlines()[0].strip()
    except (OSError, IndexError):
        return "development copy (no VERSION.txt)"


# --- One run at a time ---------------------------------------------------------------------------


class RunLock:
    """An operating-system lock on ``data/run.lock``. Windows releases it when the process ends,
    even after a crash, so a stale lock never blocks the next run."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self._file = None

    def acquire(self) -> bool:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        f = open(self.path, "a+b")
        try:
            if os.name == "nt":
                import msvcrt

                f.seek(0)
                msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            f.close()
            return False
        self._file = f
        return True

    def release(self) -> None:
        if self._file is None:
            return
        try:
            if os.name == "nt":
                import msvcrt

                self._file.seek(0)
                msvcrt.locking(self._file.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl

                fcntl.flock(self._file.fileno(), fcntl.LOCK_UN)
        finally:
            self._file.close()
            self._file = None


# --- The run -------------------------------------------------------------------------------------


def fetch_with_retries(tries: int, gap_minutes: int):
    """The Merchant API fetch, tried up to ``tries`` times. ``None`` if every try failed."""
    from app.services.merchant_client import client

    for attempt in range(1, tries + 1):
        try:
            return asyncio.run(client.fetch_all_disapprovals())
        except Exception:
            logger.exception("fetch failed (trigger=%s, try %d of %d)", TRIGGER, attempt, tries)
        if attempt < tries:
            logger.warning("next fetch try in %d min", gap_minutes)
            sleep(gap_minutes * 60)
    return None


def save_with_retries(meta: dict, tries: int, gap_minutes: int) -> Path | None:
    """Build the report from the history file and save it; tried up to ``tries`` times.

    Never calls Google: the figures are already in the history file, so a full folder or a
    OneDrive hiccup costs no Merchant API calls. ``None`` if every try failed.
    """
    from app import report_drop
    from app.api import insights

    for attempt in range(1, tries + 1):
        try:
            path = report_drop.write_report(insights._build_report(meta), meta["last_synced_at"])
            if path is not None:
                return path
            logger.error("report not saved: REPORT_DROP_DIR is empty")
            return None
        except Exception:
            logger.exception("report save failed (try %d of %d)", attempt, tries)
        if attempt < tries:
            logger.warning("next save try in %d min (no new Google calls)", gap_minutes)
            sleep(gap_minutes * 60)
    return None


def _missing_setup(need_google: bool) -> list[str]:
    from app.config import settings
    from app.services import merchant_auth

    missing = []
    if need_google and not merchant_auth.key_present():
        missing.append(f"Google key file not found: {settings.gcp_sa_key_path}")
    if need_google and not settings.parent_mca_id.strip():
        missing.append("PARENT_MCA_ID is empty in settings.env")
    if not settings.report_drop_dir.strip():
        missing.append("REPORT_DROP_DIR is empty in settings.env")
    return missing


def run(environ) -> int:
    """The scheduled run: fetch, save the figures, save the report."""
    from app import db

    missing = _missing_setup(need_google=True)
    if missing:
        for line in missing:
            logger.error("not set up: %s", line)
        return EXIT_NOT_CONFIGURED

    logger.info("sync requested (trigger=%s)", TRIGGER)
    result = fetch_with_retries(runner_setting(environ, "FETCH_TRIES"),
                                runner_setting(environ, "FETCH_RETRY_MINUTES"))
    if result is None:
        logger.error("standalone run failed: Google failed on every try; nothing saved, no report")
        return EXIT_FETCH_FAILED

    db.replace_disapprovals(*result)
    meta = db.get_sync_meta()
    logger.info("sync cached %d rows at %s", meta["row_count"], meta["last_synced_at"])

    path = save_with_retries(meta, runner_setting(environ, "SAVE_TRIES"),
                             runner_setting(environ, "SAVE_RETRY_MINUTES"))
    if path is None:
        logger.error("standalone run: figures saved, report NOT saved (no email). "
                     "Fix the folder, then run: python run_report.py --resave")
        return EXIT_NOT_SAVED
    logger.info("standalone run done: %d rows, report %s", meta["row_count"], path)
    return EXIT_OK


def resave(environ) -> int:
    """Rebuild and save the last run's report from the history file. No Google calls."""
    from app import db, report_drop

    missing = _missing_setup(need_google=False)
    if missing:
        for line in missing:
            logger.error("not set up: %s", line)
        return EXIT_NOT_CONFIGURED
    meta = db.get_sync_meta()
    if not meta["last_synced_at"]:
        logger.error("nothing to resave: the history file has no sync yet")
        return EXIT_NOT_SAVED
    existing = report_drop.drop_path(meta["last_synced_at"])
    if existing.exists():
        # Saving it again could make the flow email the same report twice.
        logger.info("already saved, not written again: %s", existing)
        return EXIT_OK
    path = save_with_retries(meta, runner_setting(environ, "SAVE_TRIES"),
                             runner_setting(environ, "SAVE_RETRY_MINUTES"))
    if path is None:
        return EXIT_NOT_SAVED
    logger.info("resave done: report %s (from the sync at %s)", path, meta["last_synced_at"])
    return EXIT_OK


# --- --check -------------------------------------------------------------------------------------

OK, WARN, PROBLEM = "OK     ", "WARNING", "PROBLEM"


def _google_sign_in() -> str:
    from app.services import merchant_auth

    merchant_auth.get_token()
    return "signed in to Google with the key"


def _reach_merchant_api() -> str:
    """Any HTTP answer from the Merchant API's address proves the network lets us through.
    Unauthenticated, and asks for no account data."""
    import httpx

    from app.config import settings

    answer = httpx.get(settings.merchant_api_base, timeout=15)
    return f"Merchant API address reachable ({settings.merchant_api_base}, HTTP {answer.status_code})"


def inside_onedrive(folder: Path, environ) -> bool:
    """Is ``folder`` inside a OneDrive folder? (Windows names them in these variables.)"""
    for key in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        root = environ.get(key, "").strip()
        if root and Path(folder).resolve().is_relative_to(Path(root).resolve()):
            return True
    return False


def checks(db_existed: bool, folder: Path, environ) -> list[tuple[str, str]]:
    """(level, plain-language line) for each part of the setup."""
    from app import db, stores
    from app.config import settings
    from app.services import merchant_auth

    out: list[tuple[str, str]] = []

    if sys.version_info[:2] == (3, 14):
        out.append((OK, f"Python {sys.version.split()[0]}"))
    else:
        out.append((WARN, f"Python {sys.version.split()[0]}: the app is tested on 3.14 only"))

    key_ok = merchant_auth.key_present()
    if key_ok:
        try:
            email = json.loads(Path(settings.gcp_sa_key_path).read_text(encoding="utf-8"))["client_email"]
            out.append((OK, f"Google key found ({email})"))
        except (OSError, ValueError, KeyError):
            key_ok = False
            out.append((PROBLEM, f"Google key is not a valid key file: {settings.gcp_sa_key_path}"))
    else:
        out.append((PROBLEM, f"Google key not found: {settings.gcp_sa_key_path}"))

    mca = settings.parent_mca_id.strip()
    if mca.isdigit():
        out.append((OK, f"PARENT_MCA_ID = {mca}"))
    else:
        out.append((PROBLEM, "PARENT_MCA_ID is empty or not a number in settings.env"))

    drop = settings.report_drop_dir.strip()
    if not drop:
        out.append((PROBLEM, "REPORT_DROP_DIR is empty in settings.env"))
    elif not Path(drop).is_dir():
        out.append((PROBLEM, f"report folder does not exist: {drop}"))
    elif not os.access(drop, os.W_OK):
        out.append((PROBLEM, f"report folder is not writable: {drop}"))
    else:
        out.append((OK, f"report folder: {drop}"))
        if Path(drop).name.strip().lower() == LAPTOP_FOLDER_NAME.lower():
            out.append((WARN, f"the folder is named '{LAPTOP_FOLDER_NAME}'. If another computer already "
                              "saves reports into this same OneDrive folder, the team gets two emails a day"))

    if inside_onedrive(folder, environ):
        out.append((WARN, f"this folder is inside OneDrive ({folder}): the Google key and the "
                          "history file (client data) get uploaded. Move it, e.g. to C:\\GMC\\"))

    if db_existed:
        last = db.get_sync_meta()["last_synced_at"]
        if last:
            out.append((OK, f"history file found: {settings.db_path} (last sync {last})"))
        else:
            out.append((PROBLEM, f"the history file at {settings.db_path} is empty (no earlier days): "
                                 "replace it with the history file you were given (INSTRUCTIONS.md step 6)"))
    else:
        out.append((PROBLEM, f"no history file at {settings.db_path}: copy the history file you "
                             "were given there as gmc_cache.sqlite (INSTRUCTIONS.md step 6)"))

    try:
        out.append((OK, f"account list: {len(stores.directory())} accounts"))
    except Exception as exc:
        out.append((PROBLEM, f"account list could not be read: {exc}"))

    if key_ok:
        for label, step in (("Google sign-in", _google_sign_in),
                            ("Reaching the Merchant API", _reach_merchant_api)):
            try:
                out.append((OK, step()))
            except Exception as exc:
                out.append((PROBLEM, f"{label} failed: {type(exc).__name__}: {exc}"))
    return out


def check(db_existed: bool, folder: Path, environ) -> int:
    results = checks(db_existed, folder, environ)
    for level, line in results:
        print(f"[{level}] {line}")
    problems = sum(1 for level, _ in results if level == PROBLEM)
    logger.info("check done: %d problem(s), %d warning(s)", problems,
                sum(1 for level, _ in results if level == WARN))
    print("\nAll set." if not problems else f"\n{problems} problem(s) to fix (see INSTRUCTIONS.md).")
    return EXIT_NOT_CONFIGURED if problems else EXIT_OK


# --- Entry point ---------------------------------------------------------------------------------


def main(argv=None, folder: Path = HERE, environ=None) -> int:
    environ = os.environ if environ is None else environ
    parser = argparse.ArgumentParser(description="GMC report — standalone runner")
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true", help="check the setup; saves and sends nothing")
    mode.add_argument("--resave", action="store_true",
                      help="rebuild the last report from the history file; no Google calls")
    args = parser.parse_args(argv)

    prepare_environment(folder, environ)
    if str(folder) not in sys.path:
        sys.path.insert(0, str(folder))  # the app/ folder sits beside this file in the zip

    from app import db
    from app.config import settings
    from app.logging_config import setup_logging

    setup_logging()
    logger.info("standalone runner %s (version %s, Python %s)",
                "check" if args.check else "resave" if args.resave else "run",
                version(folder), sys.version.split()[0])

    lock = RunLock(Path(settings.db_path).parent / "run.lock")
    if not lock.acquire():
        logger.warning("another run is still going; this one did not start")
        return EXIT_BUSY
    try:
        db_existed = Path(settings.db_path).is_file()
        if db_existed or not args.check:
            db.init_db()  # --check saves nothing: it never creates an empty history file
        if args.check:
            return check(db_existed, folder, environ)
        if args.resave:
            return resave(environ)
        return run(environ)
    except Exception:
        logger.exception("standalone run crashed")
        return EXIT_CRASHED
    finally:
        lock.release()


if __name__ == "__main__":
    os.chdir(HERE)  # Task Scheduler may start us anywhere; keep stray .env files out
    sys.exit(main())
