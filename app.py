"""Blockchain.com DeFi Wallet clone — Flask backend.

Routes:
  /                        redirect to /login if not authed
  /login                   login page (German)
  /logout                  logout
  /wallet/home             Startseite
  /wallet/assets           Vermögenswerte
  /wallet/currency/<sym>   TRON detail page (chart)
  /wallet/activity         Aktivität
  /admin/login             admin login
  /admin                   admin dashboard
  /admin/upload            POST xlsx upload
  /api/me                  current user info (JSON)
  /api/wallet              wallet data (JSON)
  /api/activity            activity feed (JSON)

Users are loaded from data/users.xlsx (sheet "users": email, password, name).
"""
import io
import os
import secrets
from datetime import datetime, date
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for, session,
    jsonify, flash, abort,
)
from openpyxl import load_workbook
from werkzeug.security import generate_password_hash, check_password_hash

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
USERS_XLSX = os.path.join(DATA_DIR, "users.xlsx")
SESSIONS_XLSX = os.path.join(DATA_DIR, "sessions.xlsx")
ACTIVITY_XLSX = os.path.join(DATA_DIR, "activity.xlsx")
WALLETS_XLSX = os.path.join(DATA_DIR, "wallets.xlsx")

ADMIN_USER = os.environ.get("ADMIN_USER", "admin")
ADMIN_PASS = os.environ.get("ADMIN_PASS", "change-me-now")

# Templates and static live under ./app/
APP_ROOT = os.path.dirname(os.path.abspath(__file__))
app = Flask(
    __name__,
    template_folder=os.path.join(APP_ROOT, "app", "templates"),
    static_folder=os.path.join(APP_ROOT, "app", "static"),
)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16MB upload cap

os.makedirs(DATA_DIR, exist_ok=True)


# ----------------------------------------------------------------------------
# Storage layer — xlsx-backed, simple. For ~hundreds of users this is plenty.
# ----------------------------------------------------------------------------
def _read_xlsx(path, sheet):
    if not os.path.exists(path):
        return []
    wb = load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h).strip() for h in rows[0]]
    out = []
    for r in rows[1:]:
        if r is None or all(c is None for c in r):
            continue
        out.append({header[i]: (r[i] if i < len(r) else None) for i in range(len(header))})
    return out


def _write_xlsx(path, sheet, rows, header):
    """Write a fresh xlsx with header + rows (list of dicts)."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(header)
    for r in rows:
        ws.append([r.get(h, "") for h in header])
    wb.save(path)


# ----------------------------------------------------------------------------
# Users
# ----------------------------------------------------------------------------
def _ensure_users_skeleton():
    """If users.xlsx doesn't exist, create it with an admin row."""
    if not os.path.exists(USERS_XLSX):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "users"
        ws.append(["email", "password_hash", "name", "created_at", "last_login"])
        # seed one row so the file is never empty
        ws.append([
            "demo@blockchain-demo.com",
            generate_password_hash("demo1234"),
            "Demo Nutzer",
            datetime.utcnow().isoformat(timespec="seconds"),
            "",
        ])
        wb.save(USERS_XLSX)


def _all_users():
    _ensure_users_skeleton()
    return _read_xlsx(USERS_XLSX, "users")


def _find_user(email):
    email = (email or "").strip().lower()
    for u in _all_users():
        if str(u.get("email", "")).strip().lower() == email:
            return u
    return None


def _update_user(email, **fields):
    users = _all_users()
    email = email.strip().lower()
    for u in users:
        if str(u.get("email", "")).strip().lower() == email:
            u.update(fields)
            _write_xlsx(
                USERS_XLSX, "users", users,
                ["email", "password_hash", "name", "created_at", "last_login"],
            )
            return
    # user not found -> append
    u = {"email": email, "created_at": datetime.utcnow().isoformat(timespec="seconds")}
    u.update(fields)
    users.append(u)
    _write_xlsx(
        USERS_XLSX, "users", users,
        ["email", "password_hash", "name", "created_at", "last_login"],
    )


# ----------------------------------------------------------------------------
# Wallets  (one balance row per user)
# ----------------------------------------------------------------------------
def _ensure_wallets_skeleton():
    if not os.path.exists(WALLETS_XLSX):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "wallets"
        ws.append(["email", "symbol", "name", "balance_eur", "balance_qty", "qty_unit"])
        # TRON row for the demo user
        ws.append(["demo@blockchain-demo.com", "TRX", "TRON", 0.28, 1.0, "TRX"])
        wb.save(WALLETS_XLSX)


def _wallets_for(email):
    _ensure_wallets_skeleton()
    email = (email or "").strip().lower()
    out = []
    for w in _read_xlsx(WALLETS_XLSX, "wallets"):
        if str(w.get("email", "")).strip().lower() == email:
            out.append(w)
    return out


def _networth(email):
    return round(sum(float(w.get("balance_eur") or 0) for w in _wallets_for(email)), 2)


# ----------------------------------------------------------------------------
# Activity feed
# ----------------------------------------------------------------------------
def _ensure_activity_skeleton():
    if not os.path.exists(ACTIVITY_XLSX):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "activity"
        ws.append(["email", "kind", "label", "amount_eur", "amount_qty", "qty_unit", "ts"])
        ws.append([
            "demo@blockchain-demo.com", "received", "Received TRX",
            0.28, 1.0, "TRX",
            datetime(2026, 6, 20, 14, 23).isoformat(timespec="seconds"),
        ])
        wb.save(ACTIVITY_XLSX)


def _activity_for(email):
    _ensure_activity_skeleton()
    email = (email or "").strip().lower()
    rows = []
    for r in _read_xlsx(ACTIVITY_XLSX, "activity"):
        if str(r.get("email", "")).strip().lower() == email:
            rows.append(r)
    rows.sort(key=lambda r: str(r.get("ts", "")), reverse=True)
    return rows


# ----------------------------------------------------------------------------
# Auth decorator
# ----------------------------------------------------------------------------
def login_required(fn):
    @wraps(fn)
    def _w(*a, **kw):
        if not session.get("user_email"):
            return redirect(url_for("login", next=request.path))
        return fn(*a, **kw)
    return _w


def admin_required(fn):
    @wraps(fn)
    def _w(*a, **kw):
        if not session.get("is_admin"):
            return redirect(url_for("admin_login", next=request.path))
        return fn(*a, **kw)
    return _w


# ----------------------------------------------------------------------------
# Public routes
# ----------------------------------------------------------------------------
@app.route("/")
def index():
    if not session.get("user_email"):
        return redirect(url_for("login"))
    return redirect(url_for("wallet_home"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = (request.form.get("email") or "").strip()
        pw = request.form.get("password") or ""

        # two-step screen accepts just email -> password screen
        if "password" not in request.form:
            user = _find_user(email)
            if not user:
                return render_template(
                    "login.html", step="email",
                    error="Ungültige E-Mail oder Wallet-ID",
                    email=email,
                )
            return render_template("login.html", step="password", email=email)

        # password step
        user = _find_user(email)
        if not user or not check_password_hash(str(user.get("password_hash") or ""), pw):
            return render_template(
                "login.html", step="password",
                error="Falsches Passwort",
                email=email,
            )

        session.clear()
        session["user_email"] = str(user["email"]).strip().lower()
        session["user_name"] = str(user.get("name") or "").strip() or session["user_email"]
        _update_user(session["user_email"], last_login=datetime.utcnow().isoformat(timespec="seconds"))

        nxt = request.args.get("next") or url_for("wallet_home")
        return redirect(nxt)

    # GET
    return render_template("login.html", step="email")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ----------------------------------------------------------------------------
# Wallet pages
# ----------------------------------------------------------------------------
def _wallet_ctx(email):
    return {
        "user_email": email,
        "user_name": session.get("user_name") or email,
        "networth": _networth(email),
        "wallets": _wallets_for(email),
    }


@app.route("/wallet/home")
@login_required
def wallet_home():
    ctx = _wallet_ctx(session["user_email"])
    ctx["activity"] = _activity_for(session["user_email"])[:5]
    return render_template("home.html", **ctx)


@app.route("/wallet/assets")
@login_required
def wallet_assets():
    return render_template("assets.html", **_wallet_ctx(session["user_email"]))


@app.route("/wallet/currency/<sym>")
@login_required
def wallet_currency(sym):
    sym = sym.upper()
    wallets = _wallets_for(session["user_email"])
    coin = next((w for w in wallets if str(w.get("symbol", "")).upper() == sym), None)
    if not coin:
        # graceful — show empty state but keep the page layout
        coin = {"symbol": sym, "name": sym, "balance_eur": 0, "balance_qty": 0, "qty_unit": sym}
    ctx = _wallet_ctx(session["user_email"])
    ctx["coin"] = coin
    return render_template("currency.html", **ctx)


@app.route("/wallet/activity")
@login_required
def wallet_activity():
    return render_template(
        "activity.html",
        **_wallet_ctx(session["user_email"]),
        activity=_activity_for(session["user_email"]),
    )


# ----------------------------------------------------------------------------
# JSON APIs (frontend can call these if you wire up live data later)
# ----------------------------------------------------------------------------
@app.route("/api/me")
@login_required
def api_me():
    return jsonify({
        "email": session["user_email"],
        "name": session.get("user_name"),
        "networth": _networth(session["user_email"]),
    })


@app.route("/api/wallet")
@login_required
def api_wallet():
    return jsonify({"networth": _networth(session["user_email"]),
                    "wallets": _wallets_for(session["user_email"])})


@app.route("/api/activity")
@login_required
def api_activity():
    return jsonify({"activity": _activity_for(session["user_email"])})


# ----------------------------------------------------------------------------
# Admin
# ----------------------------------------------------------------------------
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        u = (request.form.get("username") or "").strip()
        p = request.form.get("password") or ""
        if u == ADMIN_USER and p == ADMIN_PASS:
            session.clear()
            session["is_admin"] = True
            return redirect(url_for("admin_dashboard"))
        return render_template("admin_login.html", error="Falsche Zugangsdaten")
    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))


@app.route("/admin")
@admin_required
def admin_dashboard():
    users = _all_users()
    return render_template("admin.html", users=users, count=len(users))


@app.route("/admin/upload", methods=["POST"])
@admin_required
def admin_upload():
    """Accept xlsx with columns: email, password, name (optional)."""
    f = request.files.get("file")
    if not f or not f.filename:
        flash("Keine Datei hochgeladen.", "error")
        return redirect(url_for("admin_dashboard"))

    if not f.filename.lower().endswith(".xlsx"):
        flash("Nur .xlsx-Dateien werden akzeptiert.", "error")
        return redirect(url_for("admin_dashboard"))

    try:
        wb = load_workbook(io.BytesIO(f.read()), data_only=True)
    except Exception as e:
        flash(f"Datei konnte nicht gelesen werden: {e}", "error")
        return redirect(url_for("admin_dashboard"))

    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        flash("Leere Datei.", "error")
        return redirect(url_for("admin_dashboard"))

    header = [str(h or "").strip().lower() for h in rows[0]]
    required = {"email", "password"}
    missing = required - set(header)
    if missing:
        flash(f"Spalten fehlen: {', '.join(sorted(missing))}. Erforderlich: email, password, (name optional).", "error")
        return redirect(url_for("admin_dashboard"))

    idx = {h: i for i, h in enumerate(header)}

    added = updated = skipped = errors = 0
    error_lines = []
    for lineno, r in enumerate(rows[1:], start=2):
        if r is None or all(c is None for c in r):
            continue
        email = str(r[idx["email"]] or "").strip().lower()
        password = r[idx["password"]]
        name = str(r[idx["name"]] or "").strip() if "name" in idx else ""

        if not email or not password:
            skipped += 1
            continue

        if not email.count("@") or "." not in email.split("@")[-1]:
            errors += 1
            error_lines.append(f"Zeile {lineno}: ungültige E-Mail '{email}'")
            continue

        pw_hash = generate_password_hash(str(password))
        existing = _find_user(email)
        if existing:
            _update_user(email, password_hash=pw_hash, name=name or existing.get("name", ""))
            updated += 1
        else:
            _update_user(
                email, password_hash=pw_hash, name=name or email.split("@")[0],
                created_at=datetime.utcnow().isoformat(timespec="seconds"),
                last_login="",
            )
            added += 1

    msg = f"Import fertig: {added} hinzugefügt, {updated} aktualisiert, {skipped} übersprungen."
    if errors:
        msg += f" {errors} Fehler: " + "; ".join(error_lines[:5])
        if len(error_lines) > 5:
            msg += f" (+{len(error_lines) - 5} weitere)"
        flash(msg, "error")
    else:
        flash(msg, "success")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/sample")
@admin_required
def admin_sample():
    """Download a sample xlsx so the admin knows the format."""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "users"
    ws.append(["email", "password", "name"])
    ws.append(["max@example.com", "geheim123", "Max Mustermann"])
    ws.append(["anna@example.com", "nocheins", "Anna Schmidt"])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    from flask import send_file
    return send_file(
        buf,
        as_attachment=True,
        download_name="users_sample.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ----------------------------------------------------------------------------
# Bootstrap
# ----------------------------------------------------------------------------
_ensure_users_skeleton()
_ensure_wallets_skeleton()
_ensure_activity_skeleton()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=True)