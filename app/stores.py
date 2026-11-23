"""Store Directory: which country and region each Merchant Center store belongs to (ADR-009 §1).

Merchant Center has no region field, so the mapping lives in a reviewed CSV in the repo
(``store_directory.csv``). Two jobs:

- **Labelling.** Every dashboard view looks a store up here. A store Google returns that is not
  in the file is labelled **Unassigned** — never dropped — so a new store is visible the day it
  appears, and the file is fixed afterwards.
- **Coverage.** Some stores the operator expects are not sub-accounts of the parent MCA, so
  ``listSubaccounts`` never returns them. The sync fetches every store listed here as well
  (``extra_accounts``), so the file is the single list of stores the operator expects.
"""
import csv
import logging
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from app.config import settings

logger = logging.getLogger("gmc.stores")

UNASSIGNED = "Unassigned"
DEFAULT_PATH = Path(__file__).with_name("store_directory.csv")
_HEADER = ["sub_account_id", "name", "country", "region", "is_test"]
_TRUE = {"1", "true", "yes"}
_FALSE = {"0", "false", "no", ""}


@dataclass(frozen=True)
class Store:
    sub_account_id: str
    name: str
    country: str
    region: str
    is_test: bool
    in_directory: bool = True


def _path() -> Path:
    return Path(settings.store_directory_path) if settings.store_directory_path else DEFAULT_PATH


def parse(lines) -> dict[str, Store]:
    """Parse directory CSV lines. Raises ``ValueError`` naming the bad line.

    Strict on purpose: a typo here silently mislabels a store on every view and in the
    emailed report, so a broken file must fail loudly (and the test suite loads the
    committed file to catch that before it ships).
    """
    rows = csv.reader(line for line in lines if line.strip() and not line.lstrip().startswith("#"))
    header = next(rows, None)
    if header is None or [h.strip() for h in header] != _HEADER:
        raise ValueError(f"store directory header must be {','.join(_HEADER)}; got {header}")
    stores: dict[str, Store] = {}
    for n, row in enumerate(rows, start=2):
        if len(row) != len(_HEADER):
            raise ValueError(f"store directory row {n}: expected {len(_HEADER)} fields, got {row}")
        sid, name, country, region, is_test = (c.strip() for c in row)
        if not sid.isdigit():
            raise ValueError(f"store directory row {n}: sub_account_id must be digits, got {sid!r}")
        if sid in stores:
            raise ValueError(f"store directory row {n}: duplicate sub_account_id {sid}")
        flag = is_test.lower()
        if flag not in _TRUE | _FALSE:
            raise ValueError(f"store directory row {n}: is_test must be 0 or 1, got {is_test!r}")
        test = flag in _TRUE
        if not test and (not country or not region):
            raise ValueError(f"store directory row {n}: a real store needs a country and a region")
        stores[sid] = Store(sid, name, country, region, test)
    return stores


@lru_cache(maxsize=1)
def directory() -> dict[str, Store]:
    """The directory, keyed by sub-account ID. Read once per process; restart to reload."""
    path = _path()
    with path.open(encoding="utf-8", newline="") as f:
        stores = parse(f)
    logger.info("store directory: %d stores from %s", len(stores), path.name)
    return stores


def lookup(sub_account_id: str, name: str = "") -> Store:
    """The directory entry for a store, or an Unassigned stand-in if it is not listed."""
    found = directory().get(str(sub_account_id))
    if found is not None:
        return found
    return Store(str(sub_account_id), name, UNASSIGNED, UNASSIGNED, False, in_directory=False)


def extra_accounts(known_ids: set[str]) -> list[dict]:
    """Directory stores the parent MCA did not list, shaped like ``list_subaccounts`` rows."""
    return [
        {"account_id": s.sub_account_id, "account_name": s.name}
        for s in directory().values()
        if s.sub_account_id not in known_ids
    ]
