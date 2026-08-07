"""
Web front-end for the DO lorry-assignment engine.

This is a thin, responsive web UI over the SAME assignment engine that powers
the WhatsApp bot (bot.py). It reuses the exact same handlers and rules —
nothing about the assignment logic changes; only the interface does.

Flow (mirrors the WhatsApp flow):
    1. Login (pick your name — ABI / VIVIAN / …)
    2. (ABI / VIVIAN only) upload today's master lorry file
    3. Choose Today / Tomorrow
    4. Upload the DO Excel file
    5. See the assignment result on screen and download the filled Excel

Run locally (e.g. on your office PC), then open it from any phone or PC on the
same network:

    pip install flask
    python web_app.py
    → open http://<this-pc-ip>:8000  (e.g. http://10.0.0.229:8000)

The page is responsive: it fits a desktop portal and a phone screen.
"""
from __future__ import annotations

import os
import re
import secrets
import io

import pandas as pd
from flask import (
    Flask, request, jsonify, send_file, make_response, Response, redirect,
)

import bot   # the existing engine — reused as-is

app = Flask(__name__)

# Each browser gets a random session token stored in a cookie. We map that token
# to a bot session (bot.get_session(token)) so the engine's per-user state
# machine works exactly as it does for a WhatsApp phone number.
_COOKIE = "do_sid"

# ─────────────────────────────────────────────────────────────────────────────
# Credential-based login — validate against data/credentials.xlsx
# (columns: email, password, ip, device_name). Keep this file OUT of git; it
# holds plaintext staff passwords (see .gitignore).
# ─────────────────────────────────────────────────────────────────────────────
_CRED_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "data", "credentials.xlsx")


def _load_credentials() -> dict:
    """Return {email_lower: {'password':..,'ip':..,'device_name':..}}.

    Read fresh each login so edits to the file take effect without a restart.
    """
    try:
        df = pd.read_excel(_CRED_PATH)
    except Exception:
        return {}
    df.columns = [str(c).strip().lower() for c in df.columns]

    def _cell(v):
        # Blank Excel cells read as NaN; normalise to "" instead of "nan".
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        s = str(v).strip()
        return "" if s.lower() in ("nan", "none") else s

    out = {}
    for _, r in df.iterrows():
        email = _cell(r.get("email", "")).lower()
        if not email:
            continue
        out[email] = {
            "password":    _cell(r.get("password", "")),
            "ip":          _cell(r.get("ip", "")),
            "device_name": _cell(r.get("device_name", "")),
        }
    return out


def _check_credentials(email: str, password: str, device_name: str = ""):
    """Return (ok, error_message). Password must match; device name, if the
    user typed one, must match the record too."""
    creds = _load_credentials()
    if not creds:
        return False, "No credentials configured. Add data/credentials.xlsx."
    rec = creds.get((email or "").strip().lower())
    # Reject unknown email, an account with no password set (incomplete row),
    # or a wrong password.
    if not rec or not rec["password"] or (password or "").strip() != rec["password"]:
        return False, "Invalid email or password."
    # Device name is optional. Only enforce it when BOTH the account has one and
    # the user typed one. Compare loosely: ignore case and all whitespace
    # (incl. stray spaces Excel sometimes leaves in a cell).
    def _norm(v):
        return re.sub(r"\s+", "", str(v or "")).upper()
    dev, want = _norm(device_name), _norm(rec["device_name"])
    if dev and want and dev != want:
        return False, "Device name does not match this account."
    return True, None


def _is_authed(sess) -> bool:
    return bool(sess.get("_authed"))


def _sid() -> str:
    """Return the caller's session id, creating one if absent."""
    sid = request.cookies.get(_COOKIE)
    if not sid:
        sid = "web_" + secrets.token_hex(8)
    return sid


def _with_cookie(payload, sid, status=200):
    resp = make_response(jsonify(payload), status)
    resp.set_cookie(_COOKIE, sid, samesite="Lax", max_age=60 * 60 * 8)
    return resp


def _file_bytes():
    """Read the uploaded file from the request (field name 'file')."""
    f = request.files.get("file")
    if not f:
        return None
    return f.read()


# ─────────────────────────────────────────────────────────────────────────────
# Result rendering — turn the engine's session state into structured JSON so the
# browser can draw a clean table instead of parsing chat text.
# ─────────────────────────────────────────────────────────────────────────────
_SENTINELS = {"NO_LORRY", "NO_ELIGIBLE_LORRY", "SPLIT", "SKIPPED",
              "OTHER_USER", "NOT_TODAY", "REMARKS_SKIP", "OUT_SOURCE",
              "", None}


def _result_json(sess) -> dict:
    """Group the assigned items by lorry and list what couldn't be assigned."""
    items = sess.get("items", []) or []
    engine = sess.get("engine")
    caps = {}
    if engine is not None and getattr(engine, "eligible_lorries", None) is not None:
        try:
            for _, r in engine.eligible_lorries.iterrows():
                caps[str(r["LORRY"]).strip().upper()] = float(r["TON"])
        except Exception:
            caps = {}

    lorries: dict[str, dict] = {}
    unassigned: list[dict] = []
    skipped_other = 0

    for it in items:
        lorry = it.get("LORRY")
        do = {
            "do": str(it.get("DO NUMBER", "")),
            "route": str(it.get("ROUTE", "")),
            "customer": str(it.get("CUSTOMER NAME", "")),
            "weight": round(float(it.get("WEIGHT", 0) or 0), 3),
            "state": str(it.get("STATE", "")),
        }
        if lorry in ("OTHER_USER", "NOT_TODAY", "OUT_SOURCE", "REMARKS_SKIP"):
            skipped_other += 1
            continue
        if lorry in _SENTINELS:
            _reasons = sess.get("unassigned_reasons")
            do["reason"] = _reasons.get(do["do"], "NO_LORRY") if isinstance(_reasons, dict) else "NO_LORRY"
            unassigned.append(do)
            continue
        grp = lorries.setdefault(lorry, {
            "lorry": lorry,
            "capacity": caps.get(str(lorry).strip().upper()),
            "load": 0.0,
            "dos": [],
        })
        grp["load"] = round(grp["load"] + do["weight"], 3)
        grp["dos"].append(do)

    lorry_list = sorted(lorries.values(), key=lambda g: g["lorry"])
    for g in lorry_list:
        cap = g["capacity"]
        g["util"] = round(100.0 * g["load"] / cap, 1) if cap else None

    total_dos = sum(len(g["dos"]) for g in lorry_list)
    return {
        "lorries": lorry_list,
        "unassigned": unassigned,
        "summary": {
            "lorries_used": len(lorry_list),
            "dos_assigned": total_dos,
            "dos_unassigned": len(unassigned),
            "dos_other": skipped_other,
        },
    }


# ─────────────────────────────────────────────────────────────────────────────
# API endpoints — each drives the existing engine handler for the current step.
# ─────────────────────────────────────────────────────────────────────────────
@app.before_request
def _guard():
    """All /api/* endpoints require a logged-in session."""
    if request.path.startswith("/api/"):
        sess = bot.get_session(_sid())
        if not _is_authed(sess):
            return _with_cookie({"error": "auth_required"}, _sid(), 401)


@app.route("/auth", methods=["POST"])
def auth():
    sid = _sid()
    sess = bot.get_session(sid)
    email = request.form.get("email", "")
    password = request.form.get("password", "")
    device = request.form.get("device_name", "")
    ok, err = _check_credentials(email, password, device)
    if not ok:
        return Response(_login_page(err), mimetype="text/html", status=401)
    sess["_authed"] = True
    sess["_email"] = email.strip()
    sess["_device"] = device.strip()
    resp = make_response(redirect("/"))
    resp.set_cookie(_COOKIE, sid, samesite="Lax", max_age=60 * 60 * 8)
    return resp


@app.route("/logout", methods=["POST", "GET"])
def logout():
    sid = _sid()
    bot.reset_session(sid)          # clears auth + any in-progress assignment
    resp = make_response(redirect("/"))
    resp.set_cookie(_COOKIE, sid, samesite="Lax", max_age=60 * 60 * 8)
    return resp


@app.route("/api/state")
def api_state():
    sid = _sid()
    sess = bot.get_session(sid)
    return _with_cookie({
        "state": sess.get("state", "IDLE"),
        "user": sess.get("user_id"),
        "email": sess.get("_email"),
    }, sid)


@app.route("/api/users")
def api_users():
    sid = _sid()
    try:
        users = bot._get_valid_users()
    except Exception:
        users = ["ABI", "VIVIAN"]
    # Hide non-selectable placeholder users from the web picker.
    _hidden = {"NAME", "SPARE"}
    users = [u for u in users if str(u).strip().upper() not in _hidden]
    return _with_cookie({"users": list(users)}, sid)


@app.route("/api/login", methods=["POST"])
def api_login():
    sid = _sid()
    sess = bot.get_session(sid)
    user = (request.json or {}).get("user", "")
    msgs = bot._handle_user_id(sid, sess, str(user))
    return _with_cookie({
        "messages": msgs,
        "state": sess.get("state"),
        "user": sess.get("user_id"),
        # AWAIT_MASTER_UPLOAD → needs master file; AWAIT_TRIP_DAY → skip to day
        "needs_master": sess.get("state") == "AWAIT_MASTER_UPLOAD",
    }, sid)


@app.route("/api/master", methods=["POST"])
def api_master():
    sid = _sid()
    sess = bot.get_session(sid)
    fb = _file_bytes()
    if fb is None:
        return _with_cookie({"error": "No file uploaded."}, sid, 400)
    msgs = bot._handle_master_upload(sid, sess, fb)
    return _with_cookie({
        "messages": msgs,
        "state": sess.get("state"),
        # advanced to AWAIT_TRIP_DAY on success; stayed put on validation error
        "ok": sess.get("state") == "AWAIT_TRIP_DAY",
    }, sid)


@app.route("/api/day", methods=["POST"])
def api_day():
    sid = _sid()
    sess = bot.get_session(sid)
    day = (request.json or {}).get("day", "today")
    # Keep the SCHD schedule filter ON so off-schedule DOs are surfaced and the
    # user is asked whether to assign them anyway (see /api/dos + /api/offschedule).
    sess.pop("_ignore_schedule", None)
    msgs = bot._handle_trip_day(sid, sess, str(day))
    return _with_cookie({
        "messages": msgs,
        "state": sess.get("state"),
        "ok": sess.get("state") in ("AWAIT_EXCEL", "AWAIT_TRIP_DAY"),
    }, sid)


@app.route("/api/dos", methods=["POST"])
def api_dos():
    sid = _sid()
    sess = bot.get_session(sid)
    fb = _file_bytes()
    if fb is None:
        return _with_cookie({"error": "No file uploaded."}, sid, 400)
    msgs = bot._handle_excel_upload(sid, sess, fb)
    # If some DOs aren't on the chosen day's route schedule, the engine parks in
    # AWAIT_OTHER_USER_REPLY and expects a YES/NO. Surface that as a question
    # instead of the final result.
    if sess.get("state") == "AWAIT_OTHER_USER_REPLY":
        return _with_cookie({
            "messages": msgs,
            "state": sess.get("state"),
            "offschedule": {
                "count": sess.get("not_today_pending_count", 0),
                "day": sess.get("trip_day", "today"),
            },
        }, sid)
    result = _result_json(sess) if sess.get("items") else None
    return _with_cookie({
        "messages": msgs,
        "state": sess.get("state"),
        "result": result,
    }, sid)


@app.route("/api/offschedule", methods=["POST"])
def api_offschedule():
    """User's answer to 'assign the off-schedule DOs too?' — YES re-runs the
    assignment with the schedule filter off; NO leaves them unassigned."""
    sid = _sid()
    sess = bot.get_session(sid)
    assign = bool((request.json or {}).get("assign"))
    msgs = bot._handle_other_user_reply(sid, sess, "YES" if assign else "NO")
    result = _result_json(sess) if sess.get("items") else None
    return _with_cookie({
        "messages": msgs,
        "state": sess.get("state"),
        "result": result,
    }, sid)


@app.route("/api/unassigned-lorries", methods=["POST"])
def api_unassigned_lorries():
    """Reply to 'which lorries are available for the unassigned DOs?' —
    bin-packs the still-unassigned DOs onto the named plate(s) directly
    from session state (no re-upload), same as the WhatsApp flow."""
    sid = _sid()
    sess = bot.get_session(sid)
    plates = (request.json or {}).get("plates", "")
    if not sess.get("items"):
        return _with_cookie({"error": "No active assignment in this session."}, sid, 400)
    msgs = bot._handle_unassigned_lorry_reply(sid, sess, str(plates))
    result = _result_json(sess)
    return _with_cookie({
        "messages": msgs,
        "state": sess.get("state"),
        "result": result,
    }, sid)


@app.route("/api/download")
def api_download():
    sid = _sid()
    sess = bot.get_session(sid)
    # Ensure the export bytes exist (build them if the user goes straight to
    # download after seeing the result).
    data = sess.get("export_bytes")
    if not data:
        try:
            bot._export_result(sess)
            data = sess.get("export_bytes")
        except Exception:
            data = None
    if not data:
        return _with_cookie({"error": "Nothing to download yet."}, sid, 400)
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name="DO_Assigned.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/reset", methods=["POST"])
def api_reset():
    sid = _sid()
    bot.reset_session(sid)
    return _with_cookie({"ok": True}, sid)


@app.route("/api/reset-engine", methods=["POST"])
def api_reset_engine():
    """Clear the assignment state machine (for start-over) WITHOUT logging the
    user out."""
    sid = _sid()
    sess = bot.get_session(sid)
    auth = {k: sess.get(k) for k in ("_authed", "_email", "_device")}
    bot.reset_session(sid)
    sess = bot.get_session(sid)
    sess.update(auth)
    return _with_cookie({"ok": True}, sid)


# Which engine state each wizard step corresponds to (for the Back button).
_STEP_STATE = {
    "master": "AWAIT_MASTER_UPLOAD",
    "day":    "AWAIT_TRIP_DAY",
    "dos":    "AWAIT_EXCEL",
}


@app.route("/api/back", methods=["POST"])
def api_back():
    """Step the wizard backward. Rewinding to an earlier step just moves the
    engine's state pointer back — the loaded master/engine is kept, so there's
    no slow Excel re-parse. Going all the way back to the user picker clears
    the engine (but keeps the login)."""
    sid = _sid()
    sess = bot.get_session(sid)
    target = (request.json or {}).get("target")
    if target == "login":
        auth = {k: sess.get(k) for k in ("_authed", "_email", "_device")}
        bot.reset_session(sid)
        sess = bot.get_session(sid)
        sess.update(auth)
    elif target in _STEP_STATE and sess.get("user_id"):
        sess["state"] = _STEP_STATE[target]
    return _with_cookie({"ok": True, "state": sess.get("state")}, sid)


@app.route("/")
def index():
    sess = bot.get_session(_sid())
    if not _is_authed(sess):
        return _login_response()
    return _with_cookie_html(_PAGE)


@app.route("/login")
def login_page():
    """Explicit login URL (http://<host>:8000/login). If already logged in,
    go straight to the app."""
    sess = bot.get_session(_sid())
    if _is_authed(sess):
        return make_response(redirect("/"))
    return _login_response()


def _no_cache(resp):
    # Never let the browser cache the app HTML/JS, so a git pull + restart is
    # picked up on the next load without needing a hard refresh.
    resp.headers["Cache-Control"] = "no-store, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    return resp


def _login_response():
    resp = make_response(Response(_login_page(), mimetype="text/html"))
    resp.set_cookie(_COOKIE, _sid(), samesite="Lax", max_age=60 * 60 * 8)
    return _no_cache(resp)


def _with_cookie_html(html):
    resp = make_response(Response(html, mimetype="text/html"))
    resp.set_cookie(_COOKIE, _sid(), samesite="Lax", max_age=60 * 60 * 8)
    return _no_cache(resp)


def _login_page(error: str = "") -> str:
    err_html = (f'<div class="alert alert-danger">{error}</div>'
                if error else "")
    return _LOGIN_PAGE.replace("<!--ERROR-->", err_html)


# Login page — styled to match the sample "Aging Reports Portal" login.
_LOGIN_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
    <title>Login - Lorry Assignment Portal</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f5f7fa; }
        .login-card { max-width: 400px; margin: 80px auto; border-radius: 12px; }
        .device-hint { font-size: 12px; color: #6c757d; }
        @media (max-width: 480px) { .login-card { margin: 32px auto; } }
    </style>
</head>
<body>
<div class="container">
    <div class="card shadow-sm p-4 login-card">
        <h3 class="mb-3 text-center">🚚 Lorry Assignment Portal</h3>
        <!--ERROR-->
        <form method="post" action="/auth">
            <div class="mb-3">
                <label class="form-label">Company Email</label>
                <input type="email" id="companyEmail" name="email" class="form-control" required placeholder="you@engsheng.com">
            </div>
            <div class="mb-3">
                <label class="form-label">Password</label>
                <input type="password" name="password" class="form-control" required>
            </div>
            <div class="mb-3">
                <label class="form-label">Device Name <small class="text-muted">(optional)</small></label>
                <input type="text" id="deviceName" name="device_name" class="form-control" placeholder="e.g. HQ-INV-XXXX">
                <div class="device-hint mt-1">
                    💡 If set on your account, it must match to log in.
                </div>
            </div>
            <button class="btn btn-primary w-100">Login</button>
        </form>
    </div>
</div>
<script>
    (function () {
        var EMAIL_KEY = 'esreports_email';
        var emailInput = document.getElementById('companyEmail');
        var deviceInput = document.getElementById('deviceName');
        var form = deviceInput.closest('form');
        // Remember the email per-browser, and the device name PER EMAIL, so a
        // previous user's device name never carries over and blocks a new user.
        function devKey(email) { return 'esreports_device::' + (email || '').trim().toLowerCase(); }
        function loadDevice() {
            try {
                var d = localStorage.getItem(devKey(emailInput.value));
                deviceInput.value = d || '';
            } catch (e) {}
        }
        try {
            var e = localStorage.getItem(EMAIL_KEY);
            if (e) emailInput.value = e;
        } catch (e) {}
        loadDevice();
        emailInput.addEventListener('input', loadDevice);   // switch device when email changes
        form.addEventListener('submit', function () {
            try {
                var em = emailInput.value.trim();
                if (em) localStorage.setItem(EMAIL_KEY, em);
                if (deviceInput.value.trim()) localStorage.setItem(devKey(em), deviceInput.value.trim());
                else localStorage.removeItem(devKey(em));
            } catch (e) {}
        });
    })();
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# The single-page responsive UI (inline so this stays one self-contained file).
# ─────────────────────────────────────────────────────────────────────────────
_PAGE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Lorry Assignment</title>
<style>
  :root{
    --bg:#0f172a; --card:#1e293b; --card2:#273449; --line:#334155;
    --ink:#e2e8f0; --muted:#94a3b8; --brand:#38bdf8; --brand2:#0ea5e9;
    --ok:#22c55e; --warn:#f59e0b; --bad:#ef4444; --radius:14px;
  }
  @media (prefers-color-scheme: light){
    :root{ --bg:#f1f5f9; --card:#ffffff; --card2:#f8fafc; --line:#e2e8f0;
      --ink:#0f172a; --muted:#64748b; --brand:#0284c7; --brand2:#0369a1; }
  }
  *{box-sizing:border-box}
  html,body{margin:0;padding:0}
  body{background:var(--bg);color:var(--ink);
    font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
    line-height:1.45;-webkit-text-size-adjust:100%}
  .wrap{max-width:960px;margin:0 auto;padding:16px}
  header{display:flex;align-items:center;gap:10px;padding:8px 0 16px}
  header .logo{font-size:26px}
  header h1{font-size:19px;margin:0;font-weight:700}
  header .who{margin-left:auto;font-size:13px;color:var(--muted)}
  .logout-btn{display:inline-flex;align-items:center;gap:6px;margin-left:14px;
    text-decoration:none;font-size:14px;font-weight:600;color:#fff;background:var(--bad);
    padding:8px 16px;border-radius:10px;white-space:nowrap;box-shadow:0 1px 2px rgba(0,0,0,.15)}
  .logout-btn:hover{filter:brightness(1.08)}
  .logout-btn:active{transform:translateY(1px)}
  .card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);
    padding:18px;margin-bottom:16px;box-shadow:0 1px 2px rgba(0,0,0,.08)}
  .step-title{font-size:13px;text-transform:uppercase;letter-spacing:.05em;
    color:var(--muted);margin:0 0 12px;font-weight:700}
  .hidden{display:none!important}
  .btn{appearance:none;border:0;border-radius:10px;padding:12px 16px;font-size:15px;
    font-weight:600;cursor:pointer;background:var(--brand);color:#04222f;width:100%}
  .btn:active{transform:translateY(1px)}
  .btn.secondary{background:var(--card2);color:var(--ink);border:1px solid var(--line)}
  .btn.back{background:transparent;color:var(--muted);border:1px solid var(--line);
    width:auto;margin-top:14px;padding:9px 16px;font-size:14px}
  .btn.back:hover{color:var(--ink);border-color:var(--brand)}
  .btn:disabled{opacity:.5;cursor:not-allowed}
  .row{display:flex;gap:10px;flex-wrap:wrap}
  .row>*{flex:1 1 140px}
  .grid-users{display:flex;flex-wrap:wrap;justify-content:center;gap:10px}
  .userbtn{background:var(--card2);border:1px solid var(--line);border-radius:12px;
    padding:16px 28px;min-width:120px;font-size:16px;font-weight:700;cursor:pointer;
    color:var(--ink);text-align:center;flex:0 0 auto}
  .userbtn:active{border-color:var(--brand)}
  .drop{display:flex;flex-direction:column;align-items:center;justify-content:center;
    box-sizing:border-box;width:100%;min-height:150px;gap:6px;
    border:2px dashed var(--line);border-radius:14px;padding:28px 16px;text-align:center;
    color:var(--muted);cursor:pointer;background:var(--card2);transition:border-color .15s,background .15s}
  .drop:hover{border-color:var(--brand)}
  .drop.hot{border-color:var(--brand);color:var(--ink);background:transparent}
  .drop input{display:none}
  .msg{font-size:14px;white-space:pre-wrap;background:var(--card2);border:1px solid var(--line);
    border-radius:10px;padding:12px;margin-top:12px;color:var(--muted)}
  .msg.err{color:var(--bad);border-color:var(--bad)}
  .stat{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px}
  .pill{background:var(--card2);border:1px solid var(--line);border-radius:999px;
    padding:6px 12px;font-size:13px;font-weight:600}
  .pill b{color:var(--brand)}
  .pill.bad b{color:var(--bad)}
  .lorry{border:1px solid var(--line);border-radius:12px;margin-bottom:12px;overflow:hidden}
  .lorry h3{margin:0;padding:12px 14px;background:var(--card2);font-size:15px;
    display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .lorry h3 .cap{margin-left:auto;font-size:13px;color:var(--muted);font-weight:600}
  .bar{height:6px;background:var(--line);border-radius:6px;overflow:hidden;margin:0 14px 10px}
  .bar>i{display:block;height:100%;background:var(--ok)}
  .bar>i.hi{background:var(--warn)}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{text-align:left;padding:8px 14px;border-top:1px solid var(--line);vertical-align:top}
  th{color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.04em}
  td.w{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}
  .scroll{overflow-x:auto}
  .unassigned h3{color:var(--bad)}
  .foot{color:var(--muted);font-size:12px;text-align:center;padding:8px 0 24px}
  .spin{display:inline-block;width:16px;height:16px;border:2px solid var(--line);
    border-top-color:var(--brand);border-radius:50%;animation:sp .7s linear infinite;vertical-align:-3px}
  @keyframes sp{to{transform:rotate(360deg)}}
  a.dl{display:inline-block;text-decoration:none}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <span class="logo">🚚</span>
    <h1>Lorry Assignment</h1>
    <span class="who" id="who"></span>
    <a href="/logout" class="logout-btn">⎋ Logout</a>
  </header>

  <!-- Step 1: login -->
  <div class="card" id="card-login">
    <p class="step-title">Step 1 · Who are you?</p>
    <div class="grid-users" id="users"><span class="spin"></span></div>
    <div class="msg hidden" id="login-msg"></div>
  </div>

  <!-- Step 2: master lorry file -->
  <div class="card hidden" id="card-master">
    <p class="step-title">Step 2 · Upload master lorry file</p>
    <label class="drop" id="drop-master">
      <div>📄 Tap to choose the <b>master lorry</b> .xlsx<br><small>or drag &amp; drop</small></div>
      <input type="file" id="file-master" accept=".xlsx">
    </label>
    <div class="msg hidden" id="master-msg"></div>
    <button class="btn back" data-back="login">← Back</button>
  </div>

  <!-- Step 3: day -->
  <div class="card hidden" id="card-day">
    <p class="step-title">Step 3 · Which day?</p>
    <div class="row">
      <button class="btn" data-day="today">Today</button>
      <button class="btn secondary" data-day="tomorrow">Tomorrow</button>
    </div>
    <div class="msg hidden" id="day-msg"></div>
    <button class="btn back" data-back="afterlogin">← Back</button>
  </div>

  <!-- Step 4: DO file -->
  <div class="card hidden" id="card-dos">
    <p class="step-title">Step 4 · Upload DO file</p>
    <label class="drop" id="drop-dos">
      <div>📎 Tap to choose the <b>DO</b> .xlsx<br><small>or drag &amp; drop</small></div>
      <input type="file" id="file-dos" accept=".xlsx">
    </label>
    <div class="msg hidden" id="dos-msg"></div>
    <button class="btn back" data-back="day">← Back</button>
  </div>

  <!-- Step 4b: off-schedule question -->
  <div class="card hidden" id="card-offsched">
    <p class="step-title">One moment · Off-schedule DOs</p>
    <div class="msg" id="offsched-msg"></div>
    <div class="row" style="margin-top:14px">
      <button class="btn" id="offsched-yes">Yes, assign them too</button>
      <button class="btn secondary" id="offsched-no">No, leave them blank</button>
    </div>
    <button class="btn back" data-back="dos">← Back (re-upload DO file)</button>
  </div>

  <!-- Step 5: result -->
  <div class="card hidden" id="card-result">
    <p class="step-title">Result</p>
    <div class="stat" id="result-stat"></div>
    <div style="margin-bottom:14px">
      <a class="dl" href="/api/download"><button class="btn">⬇️ Download filled Excel</button></a>
    </div>
    <div id="result-body"></div>
    <div class="card hidden" id="unassigned-actions" style="margin:0 0 8px">
      <div style="font-size:13px;color:var(--muted);margin-bottom:8px">
        🚚 Reply with available lorry plate(s) to assign the unassigned DOs now
        (e.g. <code>VJN9910 BQX9983</code>) — no need to re-upload:
      </div>
      <div class="row">
        <input id="unassigned-plates" type="text" placeholder="Plate(s), space or comma separated"
               style="flex:2 1 200px;padding:12px 14px;border-radius:10px;border:1px solid var(--line);
                      background:var(--card2);color:var(--ink);font-size:15px">
        <button class="btn" id="btn-assign-unassigned" style="flex:0 0 auto;width:auto;padding:12px 18px">Assign</button>
      </div>
      <div class="msg hidden" id="unassigned-msg"></div>
    </div>
    <button class="btn secondary" id="btn-restart" style="margin-top:8px">Start over</button>
  </div>

  <div class="foot">Same assignment engine as the WhatsApp bot · works on phone &amp; desktop</div>
</div>

<script>
const $ = s => document.querySelector(s);
const show = (id,on=true)=>{ $(id).classList.toggle('hidden',!on); };
const setMsg = (id,txt,err=false)=>{ const e=$(id); if(!txt){e.classList.add('hidden');return;}
  e.textContent=Array.isArray(txt)?txt.join('\n'):txt; e.classList.toggle('err',err); e.classList.remove('hidden'); };

// If the session was lost (server returns 401), send the user back to the
// login page to re-authenticate rather than showing a confusing error.
function _checkAuth(r){ if(r.status===401){ window.location.href='/login'; throw new Error('auth'); } return r; }
async function jpost(url,body){ const r=_checkAuth(await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body||{})})); return r.json(); }
async function fpost(url,file){ const fd=new FormData(); fd.append('file',file); const r=_checkAuth(await fetch(url,{method:'POST',body:fd})); return r.json(); }

// ---- Wizard state (so Back can rewind the engine and replay prior steps) ----
let selUser=null, masterFile=null, selDay=null, needsMaster=false;
const ALL_CARDS=['#card-login','#card-master','#card-day','#card-dos','#card-offsched','#card-result'];
function hideAll(){ ALL_CARDS.forEach(id=>show(id,false)); }
function clearMsgs(){ ['#login-msg','#master-msg','#day-msg','#dos-msg','#offsched-msg'].forEach(id=>setMsg(id,null)); }

// Rewind to an earlier step. This only moves the engine's state pointer back
// (server keeps the loaded master), so it's instant — no re-upload, no blank
// screen. We switch the visible card only after the quick request returns.
async function goTo(step){
  clearMsgs();
  await jpost('/api/back',{target:step});
  const inpM=document.getElementById('file-master'), inpD=document.getElementById('file-dos');
  if(step==='login'){ selUser=null; masterFile=null; selDay=null; if(inpM)inpM.value=''; if(inpD)inpD.value=''; }
  if(step==='master'){ masterFile=null; if(inpM)inpM.value=''; }
  if(step==='day'){ selDay=null; }
  if(step==='dos'){ if(inpD)inpD.value=''; }
  hideAll();
  show('#card-'+(step==='login'?'login':step), true);
}
function wireBackButtons(){
  document.querySelectorAll('[data-back]').forEach(b=>b.onclick=()=>{
    let t=b.dataset.back;
    if(t==='afterlogin') t = needsMaster ? 'master' : 'login';
    goTo(t);
  });
}

// ---- Step 1: users ----
async function loadUsers(){
  const d = await (await fetch('/api/users')).json();
  const box=$('#users'); box.innerHTML='';
  (d.users||[]).forEach(u=>{
    const b=document.createElement('button'); b.className='userbtn'; b.textContent=u;
    b.onclick=()=>login(u); box.appendChild(b);
  });
}
async function login(user){
  setMsg('#login-msg','Logging in… ',false);
  const d = await jpost('/api/login',{user});
  if(d.user){ selUser=d.user; needsMaster=!!d.needs_master;
    $('#who').textContent='Logged in as '+d.user; setMsg('#login-msg',null); show('#card-login',false);
    if(d.needs_master){ show('#card-master',true); } else { show('#card-day',true); } }
  else { setMsg('#login-msg',d.messages||'Login failed',true); }
}

// ---- file drop helper ----
function wireDrop(dropId,inputId,onFile){
  const drop=$(dropId), input=$(inputId);
  input.addEventListener('change',()=>{ if(input.files[0]) onFile(input.files[0]); });
  ['dragover','dragenter'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('hot');}));
  ['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('hot');}));
  drop.addEventListener('drop',e=>{ const f=e.dataTransfer.files[0]; if(f) onFile(f); });
}

// ---- Step 2: master ----
wireDrop('#drop-master','#file-master',async(f)=>{
  setMsg('#master-msg','Reading master file… ',false);
  const d=await fpost('/api/master',f);
  setMsg('#master-msg',d.messages||d.error,!d.ok);
  if(d.ok){ masterFile=f; show('#card-master',false); show('#card-day',true); }
});

// ---- Step 3: day ----
document.querySelectorAll('[data-day]').forEach(b=>b.onclick=async()=>{
  setMsg('#day-msg','Setting day… ',false);
  selDay=b.dataset.day;
  const d=await jpost('/api/day',{day:b.dataset.day});
  setMsg('#day-msg',d.messages,false);
  show('#card-day',false); show('#card-dos',true);
});

// ---- Step 4: DOs ----
wireDrop('#drop-dos','#file-dos',async(f)=>{
  setMsg('#dos-msg','Assigning lorries… this can take a moment ',false);
  const d=await fpost('/api/dos',f);
  if(d.offschedule){
    const os=d.offschedule;
    const dayTxt = os.day==='tomorrow' ? 'tomorrow' : 'today';
    setMsg('#offsched-msg',
      `⏭ ${os.count} DO(s) are NOT on ${dayTxt}'s route schedule.\n\n`+
      `Do you still want to assign lorries to them?`);
    show('#card-dos',false); show('#card-offsched',true); setMsg('#dos-msg',null);
  }
  else if(d.result){ renderResult(d.result); show('#card-dos',false); show('#card-result',true); setMsg('#dos-msg',null); }
  else { setMsg('#dos-msg',d.messages||d.error||'Could not process file',true); }
});

async function answerOffsched(assign){
  setMsg('#offsched-msg', assign?'Assigning off-schedule DOs… ':'Finalising… ');
  const d=await jpost('/api/offschedule',{assign});
  if(d.result){ renderResult(d.result); show('#card-offsched',false); show('#card-result',true); }
  else { setMsg('#offsched-msg', d.messages||d.error||'Something went wrong',true); }
}
document.getElementById('offsched-yes').onclick=()=>answerOffsched(true);
document.getElementById('offsched-no').onclick=()=>answerOffsched(false);

// ---- Step 5: render ----
function esc(s){ return String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c])); }
let lastUnassignedCount=0;
function renderResult(r){
  const s=r.summary;
  $('#result-stat').innerHTML =
    `<span class="pill">Lorries used <b>${s.lorries_used}</b></span>`+
    `<span class="pill">DOs assigned <b>${s.dos_assigned}</b></span>`+
    (s.dos_unassigned?`<span class="pill bad">Unassigned <b>${s.dos_unassigned}</b></span>`:'')+
    (s.dos_other?`<span class="pill">Other user / skipped <b>${s.dos_other}</b></span>`:'');
  let html='';
  r.lorries.forEach(g=>{
    const cap=g.capacity!=null?g.capacity.toFixed(2)+'T':'—';
    const util=g.util!=null?g.util:0;
    html+=`<div class="lorry"><h3>🚚 ${esc(g.lorry)} <span class="cap">${g.load.toFixed(2)}T / ${cap}${g.util!=null?' · '+g.util+'%':''}</span></h3>`;
    html+=`<div class="bar"><i class="${util>100?'hi':''}" style="width:${Math.min(util,100)}%"></i></div>`;
    html+=`<div class="scroll"><table><thead><tr><th>DO</th><th>Route</th><th>Customer</th><th class="w">Weight</th></tr></thead><tbody>`;
    g.dos.forEach(d=>{ html+=`<tr><td>${esc(d.do)}</td><td>${esc(d.route)}</td><td>${esc(d.customer)}</td><td class="w">${d.weight.toFixed(3)}T</td></tr>`; });
    html+=`</tbody></table></div></div>`;
  });
  if(r.unassigned.length){
    html+=`<div class="lorry unassigned"><h3>⚠️ Unassigned (${r.unassigned.length})</h3>`;
    html+=`<div class="scroll"><table><thead><tr><th>DO</th><th>Route</th><th>Customer</th><th class="w">Weight</th><th>Reason</th></tr></thead><tbody>`;
    r.unassigned.forEach(d=>{ html+=`<tr><td>${esc(d.do)}</td><td>${esc(d.route)}</td><td>${esc(d.customer)}</td><td class="w">${d.weight.toFixed(3)}T</td><td>${esc(d.reason||'')}</td></tr>`; });
    html+=`</tbody></table></div></div>`;
  }
  $('#result-body').innerHTML=html;
  lastUnassignedCount=r.unassigned.length;
  show('#unassigned-actions', r.unassigned.length>0);
}

async function submitUnassignedPlates(){
  const inp=$('#unassigned-plates');
  const val=(inp.value||'').trim();
  if(!val) return;
  const before=lastUnassignedCount;
  const btn=$('#btn-assign-unassigned');
  btn.disabled=true;
  setMsg('#unassigned-msg','Assigning… ',false);
  try{
    const d=await jpost('/api/unassigned-lorries',{plates:val});
    if(d.result){
      const after=d.result.unassigned.length;
      const placed=before-after;
      inp.value='';
      renderResult(d.result);   // this also re-shows/hides #unassigned-actions
      if(placed>0){ setMsg('#unassigned-msg', `✅ Assigned ${placed} DO(s). ${after} still unassigned.`, false); }
      else { setMsg('#unassigned-msg', `❌ None of the unassigned DOs fit on that plate(s).`, true); }
      show('#unassigned-msg', true);
      show('#unassigned-actions', true);   // keep visible to show the outcome even if now 0 unassigned
    } else {
      setMsg('#unassigned-msg', d.messages||d.error||'Something went wrong', true);
    }
  } finally { btn.disabled=false; }
}
$('#btn-assign-unassigned').onclick=submitUnassignedPlates;
$('#unassigned-plates').addEventListener('keydown', e=>{ if(e.key==='Enter'){ e.preventDefault(); submitUnassignedPlates(); } });

$('#btn-restart').onclick=()=>goTo('login');

// Show who is logged in (from the auth session).
fetch('/api/state').then(r=>r.json()).then(d=>{
  if(d && d.email){ document.getElementById('who').textContent = d.email; }
}).catch(()=>{});

wireBackButtons();
loadUsers();
</script>
</body>
</html>"""


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    host = os.environ.get("HOST", "0.0.0.0")
    print(f"\n  DO Lorry-Assignment web app running.")
    print(f"  Open on this PC:      http://127.0.0.1:{port}/login")
    print(f"  Open from a phone:    http://<this-pc-ip>:{port}/login  "
          f"(e.g. http://10.0.0.229:{port}/login)\n")
    # Prefer waitress (a production WSGI server) for stable 24/7 running as a
    # Windows service; fall back to Flask's dev server if it isn't installed.
    try:
        from waitress import serve
        print("  Serving with waitress (production).\n")
        serve(app, host=host, port=port, threads=8)
    except ImportError:
        print("  waitress not installed — using Flask dev server. For 24/7, run"
              " 'pip install waitress'.\n")
        app.run(host=host, port=port, debug=False)
