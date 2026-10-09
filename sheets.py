"""
Google Sheets client — read + write.

- read_users(master_sheet_id)        → list of user dicts
- append_user(master_sheet_id, user) → add one row to the Users tab
                                        (creates header + tab if missing)
- read_village(sheet_id)             → dict in parseWorkbook() shape
- write_village(sheet_id, tabs)      → overwrite six tabs from a dict
                                        {name: [[header], [row], ...]}
                                        (called by migrate.py AND the
                                         admin page)
- list_village_sheet_titles(id)      → quick sanity check for a sheet

Credentials: GOOGLE_SERVICE_ACCOUNT_JSON (full JSON string) OR
             GOOGLE_APPLICATION_CREDENTIALS (file path).
"""
from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Optional

import gspread
from google.oauth2.service_account import Credentials


class SheetNotFound(Exception):
    pass


_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
_CLIENT: Optional[gspread.Client] = None


def _get_client() -> gspread.Client:
    global _CLIENT
    if _CLIENT is not None:
        return _CLIENT

    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw:
        info = json.loads(raw)
        creds = Credentials.from_service_account_info(info, scopes=_SCOPES)
    else:
        path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
        if not path or not os.path.exists(path):
            raise RuntimeError(
                "No Google service-account credentials. Set "
                "GOOGLE_SERVICE_ACCOUNT_JSON (full JSON string) or "
                "GOOGLE_APPLICATION_CREDENTIALS (file path)."
            )
        creds = Credentials.from_service_account_file(path, scopes=_SCOPES)

    _CLIENT = gspread.authorize(creds)
    return _CLIENT


def service_account_email() -> str:
    """Return the service-account email, for the admin UI."""
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw:
        try:
            return json.loads(raw).get("client_email", "(unknown)")
        except Exception:
            return "(unparseable JSON)"
    path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    if path and os.path.exists(path):
        try:
            with open(path) as f:
                return json.load(f).get("client_email", "(unknown)")
        except Exception:
            return "(unparseable file)"
    return "(no credentials configured)"


def _open(sheet_id: str) -> gspread.Spreadsheet:
    try:
        return _get_client().open_by_key(sheet_id)
    except gspread.exceptions.SpreadsheetNotFound as exc:
        raise SheetNotFound(sheet_id) from exc


# -----------------------------------------------------------------------------
# Users sheet
# -----------------------------------------------------------------------------
USERS_HEADERS = [
    "username", "password_hash", "village", "sheet_id",
    "sync_pin_hash", "role",
]


def _ensure_users_tab(ss: gspread.Spreadsheet) -> gspread.Worksheet:
    try:
        ws = ss.worksheet("Users")
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title="Users", rows=100, cols=len(USERS_HEADERS) + 2)
        ws.update(range_name="A1", values=[USERS_HEADERS], value_input_option="RAW")
        return ws

    first_row = ws.row_values(1)
    if first_row[: len(USERS_HEADERS)] != USERS_HEADERS:
        # Users sheet exists but headers are wrong / incomplete — write them
        ws.update(range_name="A1", values=[USERS_HEADERS], value_input_option="RAW")
    return ws


def read_users(master_sheet_id: str) -> List[dict]:
    ss = _open(master_sheet_id)
    ws = _ensure_users_tab(ss)
    rows = ws.get_all_records()
    return [
        {
            "username":      str(r.get("username", "")).strip(),
            "password_hash": str(r.get("password_hash", "")).strip(),
            "village":       str(r.get("village", "")).strip(),
            "sheet_id":      str(r.get("sheet_id", "")).strip(),
            "sync_pin_hash": str(r.get("sync_pin_hash", "")).strip(),
            "role":          (str(r.get("role", "")).strip() or "editor"),
        }
        for r in rows
        if r.get("username")
    ]


def append_user(master_sheet_id: str, user: dict) -> None:
    """Append one user row to the Users tab. Caller supplies already-hashed PINs."""
    ss = _open(master_sheet_id)
    ws = _ensure_users_tab(ss)
    row = [str(user.get(k, "")) for k in USERS_HEADERS]
    ws.append_row(row, value_input_option="RAW")


def update_user_field(master_sheet_id: str, username: str,
                      field: str, value: str) -> bool:
    """Update one field of one user. Returns True on success."""
    if field not in USERS_HEADERS:
        raise ValueError(f"unknown field {field!r}")
    ss = _open(master_sheet_id)
    ws = _ensure_users_tab(ss)
    rows = ws.get_all_values()
    header = rows[0] if rows else []
    try:
        col = header.index(field) + 1  # gspread is 1-indexed
        name_col = header.index("username") + 1
    except ValueError:
        return False
    for i, r in enumerate(rows[1:], start=2):  # row 2 = first data row
        if len(r) >= name_col and r[name_col - 1].strip().lower() == username.strip().lower():
            ws.update_cell(i, col, value)
            return True
    return False


# -----------------------------------------------------------------------------
# Village sheet — read
# -----------------------------------------------------------------------------
def _s(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _opt(v: Any) -> Optional[str]:
    if v is None or v == "":
        return None
    return str(v).strip() or None


def list_village_sheet_titles(sheet_id: str) -> List[str]:
    ss = _open(sheet_id)
    return [w.title for w in ss.worksheets()]


def read_village(sheet_id: str) -> Dict[str, Any]:
    ss = _open(sheet_id)
    tabs = {"Khewats": None, "Khatas": None, "Khasras": None,
            "Portions": None, "Mutations": None, "_Meta": None}
    for name in list(tabs):
        try:
            tabs[name] = ss.worksheet(name).get_all_records()
        except gspread.exceptions.WorksheetNotFound:
            if name in ("Mutations", "_Meta"):
                tabs[name] = []
            else:
                raise RuntimeError(f'Village sheet is missing tab "{name}"')

    khewats: Dict[str, Any] = {}
    for i, r in enumerate(tabs["Khewats"] or []):
        kno = _s(r.get("khewat_no"))
        if not kno:
            continue
        khewats[kno] = {
            "_no": kno,
            "owner":      _s(r.get("owner_raw")),
            "patti":      _s(r.get("patti")),
            "totalKanal": _opt(r.get("total_kanal")),
            "totalMarla": _opt(r.get("total_marla")),
            "orderIndex": i,
            "khatas":     {},
        }

    khata_order: Dict[str, int] = {}
    for r in tabs["Khatas"] or []:
        kewno = _s(r.get("khewat_no"))
        ktno  = _s(r.get("khata_no"))
        if not kewno or not ktno:
            continue
        kew = khewats.get(kewno)
        if kew is None:
            continue
        idx = khata_order.get(kewno, 0)
        khata_order[kewno] = idx + 1
        kew["khatas"][ktno] = {
            "_no": ktno,
            "cultivator": _s(r.get("cultivator_raw")),
            "totalKanal": _opt(r.get("total_kanal")),
            "totalMarla": _opt(r.get("total_marla")),
            "orderIndex": idx,
            "khasras":    {},
        }

    khasra_order: Dict[tuple, int] = {}
    for r in tabs["Khasras"] or []:
        kewno = _s(r.get("khewat_no"))
        ktno  = _s(r.get("khata_no"))
        khno  = _s(r.get("khasra_no_new"))
        if not (kewno and ktno and khno):
            continue
        kew = khewats.get(kewno)
        if not kew:
            continue
        kt = kew["khatas"].get(ktno)
        if not kt:
            continue
        key = (kewno, ktno)
        idx = khasra_order.get(key, 0)
        khasra_order[key] = idx + 1
        kt["khasras"][khno] = {
            "_no": khno,
            "portions":       [],
            "totalKanal":     _opt(r.get("total_kanal")),
            "totalMarla":     _opt(r.get("total_marla")),
            "mutationType":   _s(r.get("mutation_type")),
            "mutationNumber": _s(r.get("mutation_number")),
            "demand":         _s(r.get("demand")),
            "remarks":        _s(r.get("remarks")),
            "irrigation":     _s(r.get("irrigation")),
            "orderIndex":     idx,
        }

    portions = sorted(
        tabs["Portions"] or [],
        key=lambda r: (
            _s(r.get("khewat_no")), _s(r.get("khata_no")),
            _s(r.get("khasra_no")), int(r.get("portion_index") or 0),
        ),
    )
    for r in portions:
        kewno = _s(r.get("khewat_no"))
        ktno  = _s(r.get("khata_no"))
        khno  = _s(r.get("khasra_no"))
        kew = khewats.get(kewno)
        if not kew: continue
        kt = kew["khatas"].get(ktno)
        if not kt: continue
        kh = kt["khasras"].get(khno)
        if not kh: continue
        kh["portions"].append({
            "qism":  _s(r.get("qism")),
            "kanal": r.get("kanal") if r.get("kanal") is not None else 0,
            "marla": r.get("marla") if r.get("marla") is not None else 0,
        })

    return {"khewats": khewats}


# -----------------------------------------------------------------------------
# Village sheet — write (used by migrate.py AND the admin page)
# -----------------------------------------------------------------------------
def write_village(sheet_id: str, tabs: Dict[str, List[list]]) -> None:
    """Overwrite the six village tabs. tabs is {name: [[header], [row], ...]}."""
    ss = _open(sheet_id)
    for name, rows in tabs.items():
        try:
            ws = ss.worksheet(name)
            ws.clear()
        except gspread.exceptions.WorksheetNotFound:
            ws = ss.add_worksheet(
                title=name,
                rows=max(100, len(rows) + 10),
                cols=max(10, len(rows[0]) + 2 if rows else 10),
            )
        clean = [[("" if c is None else c) for c in r] for r in rows]
        ws.update(range_name="A1", values=clean, value_input_option="RAW")
