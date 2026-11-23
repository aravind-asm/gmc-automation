"""Service-account authentication for the Merchant API.

Loads the mounted JSON key and yields a valid OAuth bearer token, refreshing it
when expired. The key is never logged. See docs/adr/ADR-001.md.
"""
import os
import threading

from google.auth.transport.requests import Request
from google.oauth2 import service_account

from app.config import settings

_lock = threading.Lock()
_credentials: service_account.Credentials | None = None


def key_present() -> bool:
    """True if the service-account key file exists and is non-empty."""
    path = settings.gcp_sa_key_path
    return os.path.isfile(path) and os.path.getsize(path) > 0


def _load_credentials() -> service_account.Credentials:
    global _credentials
    if _credentials is None:
        _credentials = service_account.Credentials.from_service_account_file(
            settings.gcp_sa_key_path, scopes=[settings.gcp_scope]
        )
    return _credentials


def get_token() -> str:
    """Return a valid access token, refreshing if necessary.

    Refresh is a blocking network call; callers should invoke this once before a
    batch of async requests rather than per-request.
    """
    with _lock:
        creds = _load_credentials()
        if not creds.valid:
            creds.refresh(Request())
        return creds.token


def service_account_email() -> str | None:
    """The SA email (from the key's client_email), or None if unavailable."""
    try:
        return _load_credentials().service_account_email
    except Exception:
        return None
