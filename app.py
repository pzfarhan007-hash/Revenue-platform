"""
Village Portal — Stage 1 (browser-only edition)
===============================================
Routes:
    GET  /login                   — login form
    POST /login                   — authenticate
    GET  /logout
    GET  /                        — Jamabandi builder (editor+ required)
    GET  /api/village-data        — JSON in parseWorkbook() shape
    GET  /admin                   — admin dashboard (admin only)
    POST /admin/migrate           — upload Zalpora xlsx + sheet ID → write 6 tabs
    POST /admin/users/add         — append a user row to the Users sheet
    POST /admin/users/update      — update one field of an existing user
    POST /admin/hash              — AJAX: hash a plaintext PIN
    GET  /healthz                 — Render healthcheck
"""
import os
import traceback
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, session,
    jsonify, url_for, flash,
)

import auth
import sheets
import migrate as migrate_mod


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")
app.permanent_session_lifetime = 60 * 60 * 8  # 8 hours

# 10 MB upload cap — a full village Excel is well under this
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024


# -----------------------------------------------------------------------------
# Decorators
# -----------------------------------------------------------------------------
def login_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if "username" not in session:
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapper


def admin_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if session.get("role") != "admin":
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapper


# -----------------------------------------------------------------------------
# Auth routes
# -----------------------------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = (request.form.get("username") or "").strip().lower()
        password = request.form.get("password") or ""
        user = auth.verify_login(username, password)
        if not user:
            return render_template(
                "login.html",
                error="Incorrect username or PIN. Try again.",
                username=username,
            )
        session.permanent = True
        session["username"] = username
        session["village"]  = user["village"]
        session["sheet_id"] = user["sheet_id"]
        session["role"]     = user.get("role", "editor")
        if user["role"] == "admin":
            return redirect(url_for("admin_home"))
        nxt = request.args.get("next") or url_for("jamabandi")
        return redirect(nxt)
    return render_template("login.html", error=None, username="")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# -----------------------------------------------------------------------------
# Main app
# -----------------------------------------------------------------------------
@app.route("/")
@login_required
def jamabandi():
    if session.get("role") == "admin":
        return redirect(url_for("admin_home"))
    return render_template(
        "jamabandi.html",
        username=session["username"],
        village=session["village"],
        role=session["role"],
    )


@app.route("/api/village-data")
@login_required
def api_village_data():
    sheet_id = session.get("sheet_id")
    if not sheet_id:
        return jsonify({"error": "No sheet configured for this user"}), 400
    try:
        data = sheets.read_village(sheet_id)
    except sheets.SheetNotFound:
        return jsonify({"error": "Village sheet not found or not shared with the service account"}), 404
    except Exception as exc:  # noqa: BLE001
        app.logger.exception("village-data read failed")
        return jsonify({"error": f"Failed to read village sheet: {exc}"}), 500
    return jsonify(data)


# -----------------------------------------------------------------------------
# Admin
# -----------------------------------------------------------------------------
@app.route("/admin")
@admin_required
def admin_home():
    master_id_set = bool(os.environ.get("MASTER_USERS_SHEET_ID", "").strip())
    try:
        users = auth.list_users() if master_id_set else []
    except Exception as exc:  # noqa: BLE001
        users = []
        flash(f"Could not read Users sheet: {exc}", "error")
    return render_template(
        "admin.html",
        username=session["username"],
        users=users,
        service_account_email=sheets.service_account_email(),
        master_sheet_id=os.environ.get("MASTER_USERS_SHEET_ID", ""),
    )


@app.route("/admin/migrate", methods=["POST"])
@admin_required
def admin_migrate():
    sheet_id = (request.form.get("sheet_id") or "").strip()
    file = request.files.get("xlsx")
    if not sheet_id:
        flash("Sheet ID is required.", "error")
        return redirect(url_for("admin_home"))
    if not file or not file.filename:
        flash("Choose an .xlsx file to migrate.", "error")
        return redirect(url_for("admin_home"))
    if not file.filename.lower().endswith(".xlsx"):
        flash("File must be .xlsx (not .xls). Open in Excel and Save As → .xlsx.", "error")
        return redirect(url_for("admin_home"))

    try:
        counts = migrate_mod.migrate_bytes(file.read(), sheet_id)
    except sheets.SheetNotFound:
        flash(
            f"Google Sheet not found: {sheet_id}. Check the ID and make sure the "
            f"sheet is shared with {sheets.service_account_email()} as Editor.",
            "error",
        )
        return redirect(url_for("admin_home"))
    except Exception as exc:  # noqa: BLE001
        app.logger.exception("migrate failed")
        flash(f"Migration failed: {exc}", "error")
        return redirect(url_for("admin_home"))

    flash(
        f"Migrated: {counts['khewats']} khewats, {counts['khatas']} khatas, "
        f"{counts['khasras']} khasras, {counts['portions']} portions, "
        f"{counts['mutations']} mutations. Sheet: {sheet_id}",
        "ok",
    )
    return redirect(url_for("admin_home"))


@app.route("/admin/users/add", methods=["POST"])
@admin_required
def admin_users_add():
    master = os.environ.get("MASTER_USERS_SHEET_ID", "").strip()
    if not master:
        flash("MASTER_USERS_SHEET_ID env var is not set.", "error")
        return redirect(url_for("admin_home"))

    username   = (request.form.get("username") or "").strip().lower()
    pin        = request.form.get("pin") or ""
    sync_pin   = request.form.get("sync_pin") or ""
    village    = (request.form.get("village") or "").strip()
    sheet_id   = (request.form.get("sheet_id") or "").strip()
    role       = (request.form.get("role") or "editor").strip() or "editor"

    if not username or not pin or not village or not sheet_id:
        flash("Username, PIN, village and sheet ID are all required.", "error")
        return redirect(url_for("admin_home"))
    if username == auth.ADMIN_USERNAME:
        flash('"admin" is reserved. Use another username.', "error")
        return redirect(url_for("admin_home"))

    try:
        sheets.append_user(master, {
            "username":      username,
            "password_hash": auth.hash_pin(pin),
            "village":       village,
            "sheet_id":      sheet_id,
            "sync_pin_hash": auth.hash_pin(sync_pin) if sync_pin else "",
            "role":          role,
        })
    except Exception as exc:  # noqa: BLE001
        app.logger.exception("add user failed")
        flash(f"Could not add user: {exc}", "error")
        return redirect(url_for("admin_home"))

    auth.invalidate_user_cache()
    flash(f"Added user '{username}' for village {village}.", "ok")
    return redirect(url_for("admin_home"))


@app.route("/admin/users/update", methods=["POST"])
@admin_required
def admin_users_update():
    master = os.environ.get("MASTER_USERS_SHEET_ID", "").strip()
    if not master:
        flash("MASTER_USERS_SHEET_ID env var is not set.", "error")
        return redirect(url_for("admin_home"))

    username = (request.form.get("username") or "").strip().lower()
    action   = (request.form.get("action") or "").strip()

    if not username:
        flash("Username is required.", "error")
        return redirect(url_for("admin_home"))

    try:
        if action == "rotate_pin":
            new_pin = request.form.get("new_pin") or ""
            if not new_pin:
                flash("New PIN cannot be empty.", "error")
                return redirect(url_for("admin_home"))
            ok = sheets.update_user_field(master, username, "password_hash", auth.hash_pin(new_pin))
        elif action == "rotate_sync_pin":
            new_pin = request.form.get("new_pin") or ""
            if not new_pin:
                flash("New sync PIN cannot be empty.", "error")
                return redirect(url_for("admin_home"))
            ok = sheets.update_user_field(master, username, "sync_pin_hash", auth.hash_pin(new_pin))
        elif action == "set_sheet":
            new_sheet = (request.form.get("new_sheet_id") or "").strip()
            ok = sheets.update_user_field(master, username, "sheet_id", new_sheet)
        else:
            flash(f"Unknown action: {action}", "error")
            return redirect(url_for("admin_home"))
    except Exception as exc:  # noqa: BLE001
        app.logger.exception("update user failed")
        flash(f"Could not update user: {exc}", "error")
        return redirect(url_for("admin_home"))

    if ok:
        auth.invalidate_user_cache()
        flash(f"Updated user '{username}'.", "ok")
    else:
        flash(f"User '{username}' not found in Users sheet.", "error")
    return redirect(url_for("admin_home"))


@app.route("/admin/hash", methods=["POST"])
@admin_required
def admin_hash():
    pin = request.form.get("pin") or ""
    if not pin:
        return jsonify({"error": "empty"}), 400
    return jsonify({"hash": auth.hash_pin(pin)})


# -----------------------------------------------------------------------------
# Health
# -----------------------------------------------------------------------------
@app.route("/healthz")
def healthz():
    return "ok", 200


# -----------------------------------------------------------------------------
# Entry point
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
