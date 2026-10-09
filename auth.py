"""
Authentication for the Village Portal.

Two kinds of users:

1. **Admin (bootstrap)** — a single built-in user named "admin" whose
   PIN is set in the ADMIN_PIN env var. The admin never appears in the
   Users sheet. Admin can access the /admin page: migrate villages,
   add/edit users, re-hash PINs.

2. **Regular users** — live in the Users tab of a master Google Sheet
   whose ID is MASTER_USERS_SHEET_ID. Row columns:

       username | password_hash | village | sheet_id | sync_pin_hash | role

Password hashes are bcrypt.
"""
from __future__ import annotations

import json
import os
import time
from typing import Dict, List, Optional

import bcrypt

import sheets

# -----------------------------------------------------------------------------
# Bootstrap admin
# -----------------------------------------------------------------------------
ADMIN_USERNAME = "admin"


def _admin_pin() -> str:
    return os.environ.get("ADMIN_PIN", "").strip()


# -----------------------------------------------------------------------------
# Regular user lookup
# -----------------------------------------------------------------------------
_CACHE: Dict[str, object] = {"users": None, "ts": 0.0}
_CACHE_TTL_SECONDS = 60


def _load_users() -> List[dict]:
    now = time.time()
    if _CACHE["users"] is not None and now - _CACHE["ts"] < _CACHE_TTL_SECONDS:
        return _CACHE["users"]  # type: ignore[return-value]

    master_id = os.environ.get("MASTER_USERS_SHEET_ID", "").strip()
    if master_id:
        users = sheets.read_users(master_id)
    else:
        # Local-dev fallback
        path = os.environ.get("USERS_JSON_PATH", "users.json")
        if not os.path.exists(path):
            users = []
        else:
            with open(path, "r", encoding="utf-8") as fh:
                users = json.load(fh)

    _CACHE["users"] = users
    _CACHE["ts"] = now
    return users


def invalidate_user_cache() -> None:
    """Force the next _load_users() call to re-read the sheet. Called
    after an admin writes to the Users sheet."""
    _CACHE["users"] = None
    _CACHE["ts"] = 0.0


def _find_user(username: str) -> Optional[dict]:
    username = username.strip().lower()
    for u in _load_users():
        if u.get("username", "").strip().lower() == username:
            return u
    return None


# -----------------------------------------------------------------------------
# Verification
# -----------------------------------------------------------------------------
def verify_login(username: str, password: str) -> Optional[dict]:
    """Return a user dict on success, None on failure.

    Admin is matched first (plain string compare against ADMIN_PIN).
    Then regular users, with bcrypt.
    """
    username = (username or "").strip().lower()
    password = password or ""

    # Admin bootstrap
    if username == ADMIN_USERNAME:
        pin = _admin_pin()
        if pin and password == pin:
            return {
                "username": ADMIN_USERNAME,
                "village": "(all villages)",
                "sheet_id": "",
                "role": "admin",
            }
        return None

    # Regular user
    user = _find_user(username)
    if not user:
        return None
    hashed = user.get("password_hash", "")
    if not hashed:
        return None
    try:
        ok = bcrypt.checkpw(password.encode("utf-8"), hashed.encode("utf-8"))
    except (ValueError, TypeError):
        return None
    return user if ok else None


def verify_sync_pin(username: str, pin: str) -> bool:
    """Verify the second-factor sync PIN. Used in Stage 3."""
    user = _find_user(username)
    if not user:
        return False
    hashed = user.get("sync_pin_hash", "")
    if not hashed:
        return False
    try:
        return bcrypt.checkpw(pin.encode("utf-8"), hashed.encode("utf-8"))
    except (ValueError, TypeError):
        return False


# -----------------------------------------------------------------------------
# Helpers used by the admin page
# -----------------------------------------------------------------------------
def hash_pin(plain: str) -> str:
    """Return a bcrypt hash for a plaintext PIN."""
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def list_users() -> List[dict]:
    """Return the current user list (admin UI only)."""
    return list(_load_users())


# -----------------------------------------------------------------------------
# CLI hash generator — kept for local dev
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    import getpass
    import sys

    pw = getpass.getpass("PIN to hash: ")
    if not pw:
        print("empty", file=sys.stderr)
        sys.exit(1)
    print(hash_pin(pw))
