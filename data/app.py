"""
WhatsApp Webhook Server — Meta Cloud API
Run: python app.py
Requires: pip install flask requests pandas openpyxl
"""

import os
import sys
import io
import json
import requests
from flask import Flask, request, Response

# Use the root-level bot.py (which has all planning-file and engine improvements).
# data/app.py lives one level below the project root, so we add the parent to sys.path.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import bot

app = Flask(__name__)

# ── Load config from config.txt ───────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
_CONFIG_PATH = os.path.join(_HERE, "config.txt")

def _load_config():
    if not os.path.exists(_CONFIG_PATH):
        return
    with open(_CONFIG_PATH) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, val = line.partition("=")
                key, val = key.strip(), val.strip()
                if key and key not in os.environ:
                    os.environ[key] = val

_load_config()

META_ACCESS_TOKEN   = os.environ.get("META_ACCESS_TOKEN", "")
META_PHONE_NUMBER_ID = os.environ.get("META_PHONE_NUMBER_ID", "")
META_VERIFY_TOKEN   = os.environ.get("META_VERIFY_TOKEN", "eslorrybot2026")
PUBLIC_BASE_URL     = os.environ.get("PUBLIC_BASE_URL", "http://localhost:5000")

if not META_ACCESS_TOKEN or not META_PHONE_NUMBER_ID:
    print("⚠️  WARNING: META_ACCESS_TOKEN or META_PHONE_NUMBER_ID not set in config.txt")


# ── Webhook verification (GET) — Meta calls this to verify your endpoint ──────
@app.route("/webhook", methods=["GET"])
def verify_webhook():
    mode      = request.args.get("hub.mode")
    token     = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if mode == "subscribe" and token == META_VERIFY_TOKEN:
        print(f"✅ Webhook verified by Meta")
        return Response(challenge, status=200)
    else:
        print(f"❌ Webhook verification failed — token mismatch")
        return Response("Forbidden", status=403)


# ── Incoming messages (POST) ──────────────────────────────────────────────────
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json(silent=True) or {}

    try:
        entry    = data["entry"][0]
        changes  = entry["changes"][0]["value"]
        messages = changes.get("messages", [])

        if not messages:
            return Response("OK", status=200)  # status update, not a message

        msg    = messages[0]
        phone  = msg["from"]  # e.g. "601155003102"
        msg_type = msg.get("type", "text")

        text       = None
        file_bytes = None
        file_mime  = None

        if msg_type == "text":
            text = msg["text"]["body"]

        elif msg_type == "interactive":
            interactive = msg.get("interactive", {})
            itype       = interactive.get("type", "")
            if itype == "button_reply":
                # User tapped a reply button — use button ID as command
                text = interactive.get("button_reply", {}).get("id", "")
            elif itype == "list_reply":
                # User tapped a list item — use item ID as command (e.g. "select_do 1DO6030948")
                text = interactive.get("list_reply", {}).get("id", "")
            else:
                text = ""

        elif msg_type == "document":
            doc = msg["document"]
            file_mime = doc.get("mime_type", "")
            media_id  = doc["id"]
            file_bytes = _download_media(media_id)

        elif msg_type == "image":
            img = msg["image"]
            file_mime = img.get("mime_type", "image/jpeg")
            media_id  = img["id"]
            file_bytes = _download_media(media_id)

        # Process through bot
        replies = bot.handle_message(phone, text=text,
                                     file_bytes=file_bytes, file_mime=file_mime)

        for reply in replies:
            try:
                if not isinstance(reply, dict):
                    _send_text(phone, reply)
                elif reply.get("_type") == "buttons":
                    _send_buttons(phone, reply["body"], reply["buttons"])
                elif reply.get("_type") == "do_list":
                    _send_do_list(phone, reply["header"], reply["body"],
                                  reply["button"], reply["items"])
                else:
                    _send_text(phone, str(reply))
            except Exception as e:
                print(f"⚠️ Failed to send reply: {e}")

        try:
            export_bytes = bot.get_export_bytes(phone)
            if export_bytes:
                sess = bot.sessions.get(phone, {})
                fmt  = sess.get("file_format", "LEGACY")
                fname = "ZSDOROUTEWRH_Assigned.xlsx" if fmt == "ZSDO" else "DO_Assigned_NewRows.xlsx"
                _send_file(phone, export_bytes, filename=fname)
        except Exception as e:
            print(f"⚠️ Failed to send export file: {e}")

    except (KeyError, IndexError) as e:
        print(f"Webhook parse error: {e} — data: {json.dumps(data)[:300]}")

    return Response("OK", status=200)


# ── Meta API helpers ──────────────────────────────────────────────────────────

def _send_text(to: str, body: str):
    """Send a plain text WhatsApp message via Meta Cloud API."""
    url = f"https://graph.facebook.com/v19.0/{META_PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}",
        "Content-Type": "application/json",
    }
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "text",
        "text": {"body": body},
    }
    r = requests.post(url, headers=headers, json=payload, timeout=10)
    if not r.ok:
        print(f"❌ Send text failed: {r.status_code} {r.text}")


def _send_buttons(to: str, body: str, buttons: list[dict]):
    """
    Send an interactive button message via Meta Cloud API.
    buttons: list of {"id": str, "title": str} — max 3, title max 20 chars.
    Falls back to plain text if the API call fails.
    """
    url = f"https://graph.facebook.com/v19.0/{META_PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}",
        "Content-Type": "application/json",
    }
    # Meta requires title ≤ 20 chars
    btn_list = [
        {"type": "reply", "reply": {"id": b["id"], "title": b["title"][:20]}}
        for b in buttons[:3]   # max 3 buttons per message
    ]
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "interactive",
        "interactive": {
            "type": "button",
            "body": {"text": body},
            "action": {"buttons": btn_list},
        },
    }
    r = requests.post(url, headers=headers, json=payload, timeout=10)
    if not r.ok:
        print(f"❌ Send buttons failed: {r.status_code} {r.text}")
        # Fallback: send as plain text listing the options
        fallback = body + "\n" + "\n".join(f"  • {b['title']}" for b in buttons)
        _send_text(to, fallback)


def _download_media(media_id: str) -> bytes | None:
    """Download a media file from Meta servers."""
    # Step 1: get the download URL
    url = f"https://graph.facebook.com/v19.0/{media_id}"
    headers = {"Authorization": f"Bearer {META_ACCESS_TOKEN}"}
    r = requests.get(url, headers=headers, timeout=10)
    if not r.ok:
        print(f"❌ Media URL fetch failed: {r.status_code}")
        return None
    download_url = r.json().get("url")

    # Step 2: download the actual file
    r2 = requests.get(download_url, headers=headers, timeout=30)
    if r2.ok:
        return r2.content
    print(f"❌ Media download failed: {r2.status_code}")
    return None


def _send_file(to: str, file_bytes: bytes, filename: str):
    """
    Send a file back to the user.
    Meta requires uploading the file first, then sending by media ID.
    """
    # Step 1: Upload media
    upload_url = f"https://graph.facebook.com/v19.0/{META_PHONE_NUMBER_ID}/media"
    headers = {"Authorization": f"Bearer {META_ACCESS_TOKEN}"}
    files = {
        "file": (filename, file_bytes,
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "messaging_product": (None, "whatsapp"),
    }
    r = requests.post(upload_url, headers=headers, files=files, timeout=30)
    if not r.ok:
        print(f"❌ File upload failed: {r.status_code} {r.text}")
        _send_text(to, "⚠️ Could not send the Excel file. Please check the server.")
        return
    media_id = r.json().get("id")

    # Step 2: Send as document message
    msg_url = f"https://graph.facebook.com/v19.0/{META_PHONE_NUMBER_ID}/messages"
    payload = {
        "messaging_product": "whatsapp",
        "to": to,
        "type": "document",
        "document": {
            "id": media_id,
            "filename": filename,
        },
    }
    r2 = requests.post(msg_url, headers=headers, json=payload, timeout=10)
    if not r2.ok:
        print(f"❌ File send failed: {r2.status_code} {r2.text}")


def _clean(text: str, maxlen: int) -> str:
    """Strip emojis and non-BMP chars that WhatsApp list fields reject, then truncate."""
    import re
    cleaned = re.sub(r'[^\x00-\xFF]', '', str(text)).strip()
    return cleaned[:maxlen]


def _send_do_list(to: str, header: str, body: str, button: str, items: list[dict]):
    """
    Send a WhatsApp interactive list message.
    WhatsApp supports up to 10 sections × 10 rows = 100 items max.
    Items are split into real sections of 10 — no fake "More lorries" rows.
    If the list still exceeds 100 items, the excess are sent as a plain-text fallback.
    """
    url = f"https://graph.facebook.com/v19.0/{META_PHONE_NUMBER_ID}/messages"
    hdrs = {
        "Authorization": f"Bearer {META_ACCESS_TOKEN}",
        "Content-Type":  "application/json",
    }

    # WhatsApp hard limits
    MAX_SECTIONS = 10
    ROWS_PER_SECTION = 10
    MAX_ITEMS = MAX_SECTIONS * ROWS_PER_SECTION   # 100

    sendable  = items[:MAX_ITEMS]
    overflow  = items[MAX_ITEMS:]   # anything beyond 100 — sent as text fallback

    # Build sections: WhatsApp requires at least one section.
    # Each section can hold up to 10 rows.
    section_size = ROWS_PER_SECTION
    sections = []
    for i in range(0, len(sendable), section_size):
        chunk = sendable[i : i + section_size]
        sections.append({
            "title": f"{i+1}–{i+len(chunk)}",   # e.g. "1–10"
            "rows": [
                {
                    "id":          it["id"][:200],
                    "title":       _clean(it["title"], 24),
                    "description": _clean(it.get("description", ""), 72),
                }
                for it in chunk
            ],
        })

    clean_body = body.replace("*", "").replace("_", "")[:1024]

    payload = {
        "messaging_product": "whatsapp",
        "to":   to,
        "type": "interactive",
        "interactive": {
            "type":   "list",
            "header": {"type": "text", "text": _clean(header, 60)},
            "body":   {"text": clean_body},
            "action": {
                "button":   _clean(button, 20),
                "sections": sections,
            },
        },
    }
    try:
        r = requests.post(url, headers=hdrs, json=payload, timeout=20)
        if not r.ok:
            try:
                err_data = r.json().get("error", {})
                print(f"⚠️ List message failed (code {err_data.get('code')}): {err_data.get('message')}")
                print(f"   Payload preview: {json.dumps(payload)[:400]}")
            except Exception:
                print(f"⚠️ List message failed: {r.status_code} {r.text[:400]}")
            # Fallback: numbered plain text list
            lines_fb = [body, ""]
            for idx, it in enumerate(items[:30], 1):
                lines_fb.append(f"  {idx}. {it.get('title','')}  {it.get('description','')}")
            if len(items) > 30:
                lines_fb.append(f"  ... and {len(items)-30} more")
            lines_fb.append("")
            lines_fb.append("Type the plate number directly to continue.")
            _send_text(to, "\n".join(lines_fb))
            return
        # If there were items beyond 100 send them as a plain-text addendum
        if overflow:
            extra = "\n".join(f"  • {it.get('title','')}  {it.get('description','')}"
                               for it in overflow)
            _send_text(to, f"Remaining options (type the plate to select):\n{extra}")
    except Exception as e:
        print(f"⚠️ List message exception: {e}")
        _send_text(to, f"{body}\n\nType the plate number directly to continue.")


if __name__ == "__main__":
    print(f"🚀 Starting Meta WhatsApp Bot server...")
    print(f"   Webhook URL: {PUBLIC_BASE_URL}/webhook")
    print(f"   Verify token: {META_VERIFY_TOKEN}")
    # use_reloader=False is CRITICAL — Flask watchdog restarts the server
    # whenever any file in the project dir changes (including Excel writes),
    # which kills in-flight requests and loses all session data.
    app.run(debug=True, port=5000, use_reloader=False)
