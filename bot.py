"""
WhatsApp Bot State Machine (Twilio or Meta Cloud API compatible)
Handles the full conversation flow for lorry assignment.

State flow:
  IDLE -> AWAIT_USER_ID -> AWAIT_EXCEL -> CONFIRMING -> DONE
  (Auto-assigns best lorry for all DOs silently, shows summary for confirmation)
"""

import io
import json
import os
import re
import threading
from datetime import date, datetime, time as dtime
import pandas as pd
from lorry_engine import LorryEngine

_HERE = os.path.dirname(os.path.abspath(__file__))
# Keep data files in a separate subfolder so Flask's watchdog reloader
# never sees Excel writes and restarts the server mid-request.
_DATA_DIR = os.path.join(_HERE, "data")
os.makedirs(_DATA_DIR, exist_ok=True)

MASTER_PATH    = os.path.join(_HERE, "master lorry.xlsx")      # read-only, stays in root
# History paths — engine merges both files automatically for maximum frequency data
HISTORY_PATH     = os.path.join(_DATA_DIR, "ZSDOROUTEWRH.xlsx")               # primary (new format, manual assignments)
HISTORY_PATH_ALT = os.path.join(_DATA_DIR, "ZSDOROUTEWRH-bot.xlsx")          # bot-exported (new format)
HISTORY_PATH_OLD = os.path.join(_DATA_DIR, "126-A BI(ES) TRIP ROUTE CODE.xlsx")  # legacy reference

def _resolve_history_path() -> str:
    """Return the best available history file, preferring new format."""
    for p in [HISTORY_PATH, HISTORY_PATH_ALT, HISTORY_PATH_OLD]:
        if os.path.exists(p):
            return p
    return HISTORY_PATH_OLD  # fallback even if missing — engine will warn
DAILY_LOG_PATH = os.path.join(_DATA_DIR, "daily_assignments.json")

# ── Shared UI constants ───────────────────────────────────────────────────────
_HI_BTN = {"_type": "buttons", "body": "Tap below to start a new session.",
            "buttons": [{"id": "hi", "title": "👋 Hi"}]}

# ── Daily assignment log (persists across conversations) ─────────────────────

def _today() -> str:
    return date.today().isoformat()   # e.g. "2026-05-11"

def _load_daily_log() -> dict:
    """
    Returns { "date": "YYYY-MM-DD", "assigned": ["PLATE1", ...] }
    Resets automatically if the stored date is not today.
    """
    if os.path.exists(DAILY_LOG_PATH):
        try:
            with open(DAILY_LOG_PATH, "r") as f:
                data = json.load(f)
            if data.get("date") == _today():
                return data
        except Exception:
            pass
    return {"date": _today(), "assigned": []}

def _save_daily_log(log: dict):
    # Always sanitise before saving — remove empty/blank plate strings
    if "assigned" in log:
        log["assigned"] = sorted({p for p in log["assigned"] if p and p.strip()})
    if "broken" in log:
        log["broken"] = {k: v for k, v in log["broken"].items() if k and k.strip()}
    with open(DAILY_LOG_PATH, "w") as f:
        json.dump(log, f, indent=2)

def get_assigned_today() -> set:
    """Return set of ALL lorry plates assigned today (never includes empty strings)."""
    return {p for p in _load_daily_log()["assigned"] if p and p.strip()}

def record_assignments_today(plates: list[str]):
    """Add newly confirmed plates to today's log."""
    log = _load_daily_log()
    existing = set(log["assigned"])
    for p in plates:
        if p and p != "SKIPPED":
            existing.add(p)
    log["assigned"] = sorted(existing)
    _save_daily_log(log)


def release_specific_plates(plates: list[str]) -> bool:
    """Helper to remove plates from today's assignment and broken log."""
    log = _load_daily_log()
    assigned = set(log.get("assigned", []))
    broken = log.get("broken", {})
    changed = False
    for p in plates:
        p_up = p.upper()
        if p_up in assigned:
            assigned.discard(p_up)
            broken.pop(p_up, None)
            changed = True
    if changed:
        log["assigned"] = sorted(assigned)
        log["broken"] = broken
        _save_daily_log(log)
    return changed




def get_broken_lorries() -> dict:
    """Return dict of { broken_plate: replacement_plate } for today."""
    return _load_daily_log().get("broken", {})

def record_broken_lorry(broken: str, replacement: str):
    """Mark a lorry as broken and record its replacement for today."""
    log = _load_daily_log()
    if "broken" not in log:
        log["broken"] = {}
    log["broken"][broken.upper()] = replacement.upper()
    # Also block broken lorry from being assigned
    existing = set(log["assigned"])
    existing.add(broken.upper())
    log["assigned"] = sorted(existing)
    _save_daily_log(log)

def remove_broken_lorry(broken: str):
    """Mark a previously broken lorry as fixed — removes it from broken list."""
    log = _load_daily_log()
    broken_map = log.get("broken", {})
    plate = broken.upper()
    if plate in broken_map:
        del broken_map[plate]
        log["broken"] = broken_map
        # Also unblock from assigned list
        existing = set(log["assigned"])
        existing.discard(plate)
        log["assigned"] = sorted(existing)
        _save_daily_log(log)
        return True
    return False

def clear_daily_log_for_user(engine) -> list[str]:
    """
    Remove only the plates belonging to this user's engine (owner + SPARE)
    from today's log. Returns the list of plates actually removed.
    """
    user_lorries = set(engine.eligible_lorries["LORRY"].str.upper())
    log = _load_daily_log()
    all_plates  = set(log["assigned"])
    my_plates   = all_plates & user_lorries        # intersection = this user's plates
    remaining   = sorted(all_plates - my_plates)   # keep other users' plates
    log["assigned"] = remaining
    _save_daily_log(log)
    return sorted(my_plates)

def clear_daily_log():
    """Wipe entire log (legacy/midnight reset)."""
    _save_daily_log({"date": _today(), "assigned": []})


# ── Midnight auto-reset thread ────────────────────────────────────────────────

def _seconds_until_midnight() -> float:
    now = datetime.now()
    midnight = datetime.combine(now.date(), dtime(0, 0, 0))
    from datetime import timedelta
    next_midnight = midnight + timedelta(days=1)
    return (next_midnight - now).total_seconds()

def _midnight_reset_loop():
    """Background thread: waits until 00:00, clears the daily log, repeats."""
    while True:
        wait = _seconds_until_midnight()
        threading.Event().wait(wait)          # sleep until midnight
        clear_daily_log()
        # Also clear all in-memory sessions so lorries are fresh for the new day
        sessions.clear()

# Start the background thread once when bot.py is imported
_reset_thread = threading.Thread(target=_midnight_reset_loop, daemon=True)
_reset_thread.start()


# ── Conversation session store ────────────────────────────────────────────────
sessions: dict[str, dict] = {}

def get_session(phone: str) -> dict:
    if phone not in sessions:
        sessions[phone] = {
            "state": "IDLE",
            "user_id": None,
            "engine": None,
            "pending_dos": [],
            "current_do_index": 0,
            "suggestions": [],
            "unavailable": set(),   # marked unavailable this session
            "assigned": {},         # DO_NUMBER -> LORRY
        }
    return sessions[phone]

def reset_session(phone: str):
    sessions.pop(phone, None)


# ── Message handlers ──────────────────────────────────────────────────────────

def handle_message(phone: str, text: str = None,
                   file_bytes: bytes = None, file_mime: str = None) -> list[str]:
    sess = get_session(phone)
    state = sess["state"]
    text = (text or "").strip()
    cmd_lower = text.lower().strip()

    # ── Global commands ───────────────────────────────────────────────────────
    if text.lower() in ("reset", "restart", "start over"):
        reset_session(phone)
        return ["Session reset. Send *hi* to start again."]

    if text.lower() == "clear daily log":
        sess = get_session(phone)
        # Guard: if we are mid-flow waiting for user input (e.g. broken replacement),
        # ignore the button tap and remind the user what we are waiting for.
        if sess.get("state") == "AWAIT_BROKEN_REPLACEMENT":
            broken = sess.get("pending_broken_plate", "?")
            return [f"⚠️ Still waiting for a replacement lorry for *{broken}*.\n"
                    "Reply with the replacement plate or type *none* to skip."]
        engine = sess.get("engine")
        user   = sess.get("user_id")
        if not engine or not user:
            return ["❌ Please log in first (send *hi*) before clearing the log."]
        # Derive this user's plates by intersecting today's log with their eligible lorries
        user_lorries = set(engine.eligible_lorries["LORRY"].str.upper())
        my_plates    = sorted(get_assigned_today() & user_lorries)
        plate_count  = len(my_plates)
        plate_list   = ", ".join(my_plates) if my_plates else "none"
        return [
            {
                "_type": "buttons",
                "body": (
                    f"⚠️ *Confirm Clear Your Log ({user})?*\n\n"
                    f"This will release *{plate_count}* of your lorry assignment(s) for today.\n"
                    f"Plates: {plate_list}\n\n"
                    "This cannot be undone. Are you sure?"
                ),
                "buttons": [
                    {"id": "confirm clear daily log", "title": "✅ Yes, Clear"},
                    {"id": "cancel clear",            "title": "❌ Cancel"},
                ],
            }
        ]

    if text.lower() == "confirm clear daily log":
        sess = get_session(phone)
        engine = sess.get("engine")
        user   = sess.get("user_id")
        if not engine or not user:
            return ["❌ Please log in first (send *hi*) before clearing the log."]
        removed = clear_daily_log_for_user(engine)
        sess["unavailable"] = set()
        return [
            f"\U0001f5d1\ufe0f *{user}*'s log cleared.\n"
            f"\U0001f4cb Plates released: {', '.join(removed) or 'none'}\n"
            "Your lorries are now available again.",
            {
                "_type": "buttons",
                "body": "Tap below to start a new session, or type *hi* anytime.",
                "buttons": [{"id": "hi", "title": "👋 Hi"}],
            }
        ]

    if text.lower() == "cancel clear":
        return [
            "❌ Clear cancelled. Daily log unchanged.",
            {"_type": "buttons", "body": "What would you like to do next?",
             "buttons": [{"id": "hi", "title": "👋 Hi"}]},
        ]

    # ── Step 1: User tapped a DO# → show available lorry list ──────────────
    # ── Step 1: User tapped a DO# → show lorry picker (button pages) ────────
    if text.lower().startswith("select_do "):
        parts  = text.strip().split(" ")
        do_num = parts[1].strip().upper() if len(parts) > 1 else ""
        page   = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 0
        return _lorry_picker_buttons(sess, do_num, page)

    if text.lower().startswith("select_lorry "):
        parts  = text.strip().split(" ", 2)
        if len(parts) < 3:
            return ["❌ Invalid lorry selection."]
        do_num = parts[1].strip().upper()
        plate  = parts[2].strip().upper()

        target_do = next((d for d in sess.get("pending_dos", [])
                          if d["DO NUMBER"] == do_num), None)
        if not target_do:
            return [f"❌ DO# *{do_num}* not found."]

        if plate == "__AUTO__":
            # Auto-pick: exclude current lorry so user always gets something different
            engine: LorryEngine = sess.get("engine")
            current_lorry = sess["assigned"].get(do_num, "")
            other_assigned = set()
            for d in sess.get("pending_dos", []):
                if d["DO NUMBER"] == do_num:
                    continue
                for it in d.get("ITEMS", []):
                    lv = it.get("LORRY", "")
                    if lv not in ("NO_LORRY", "SPLIT", "", None):
                        other_assigned.add(lv)
            excluded = (sess.get("unavailable", set()) | get_assigned_today() | other_assigned | {current_lorry})
            excluded.discard("")
            suggestions = engine.suggest(
                route=target_do["ROUTE"],
                total_ton=target_do["TOTAL_TON"],
                unavailable=excluded,
                top_n=1,
            )
            if not suggestions:
                return [f"⚠️ No eligible lorry found for *{do_num}*. All lorries may be assigned."]
            plate = suggestions[0]["LORRY"]

        # Validate — not already used by another DO
        already_used = set()
        for d in sess.get("pending_dos", []):
            if d["DO NUMBER"] == do_num:
                continue
            for it in d.get("ITEMS", []):
                lv = it.get("LORRY", "")
                if lv not in ("NO_LORRY", "SPLIT", "", None):
                    already_used.add(lv)
                for bin_ in (it.get("SPLIT_LORRIES") or []):
                    already_used.add(bin_["lorry"])

        blocked_today = get_assigned_today()
        old_lorry = sess["assigned"].get(do_num, "")

        if plate in already_used:
            other = next((d["DO NUMBER"] for d in sess.get("pending_dos", [])
                          for it in d.get("ITEMS", []) if it.get("LORRY") == plate), "another DO")
            return [f"❌ *{plate}* is already assigned to {other}.",
                    {"_type": "buttons", "body": "Pick another option:",
                     "buttons": [{"id": f"select_do {do_num}", "title": "🔄 Pick again"}]}]

        if plate in blocked_today and plate != old_lorry:
            return [f"❌ *{plate}* is blocked/assigned today.",
                    {"_type": "buttons", "body": "Pick another option:",
                     "buttons": [{"id": f"select_do {do_num}", "title": "🔄 Pick again"}]}]

        # Release old lorry, assign new
        if old_lorry and old_lorry not in ("NO_LORRY", "SPLIT", ""):
            sess.get("unavailable", set()).discard(old_lorry)
        for item in target_do.get("ITEMS", []):
            item["LORRY"] = plate
            item.pop("SPLIT_LORRIES", None)
        target_do["SPLIT"] = False
        target_do.pop("SPLIT_LORRIES", None)
        sess["assigned"][do_num] = plate
        sess.setdefault("unavailable", set()).add(plate)

        _sr = _build_summary(sess)
        return [f"✅ *{do_num}* → *{plate}*"] + (_sr if isinstance(_sr, list) else [_sr])

    # ══════════════════════════════════════════════════════════════════════════
    # LORRY MANAGEMENT — single consolidated implementation
    # ══════════════════════════════════════════════════════════════════════════

    def _get_engine_safe():
        """Return engine from session or reload from master."""
        _e = sess.get("engine")
        if _e is None and sess.get("user_id"):
            try:
                _e = LorryEngine(MASTER_PATH, HISTORY_PATH, owner_user=sess["user_id"])
            except Exception:
                pass
        return _e

    # ── Main menu ─────────────────────────────────────────────────────────────
    if cmd_lower == "manage lorry":
        return [{
            "_type": "buttons",
            "body":  "🚛 *Lorry Management*\nChoose an action:",
            "buttons": [
                {"id": "lorry_maint", "title": "🔧 Maintenance"},
                {"id": "hi",          "title": "🏠 Main Menu"},
            ]
        }]

    # ── Maintenance sub-menu ──────────────────────────────────────────────────
    if cmd_lower == "lorry_maint":
        # Use a list message — supports 4 actions (WhatsApp buttons cap at 3)
        return [{
            "_type":  "do_list",
            "header": "Lorry Management",
            "body":   "Select the action to apply:",
            "button": "Choose Action",
            "items":  [
                {"id": "maint_block",   "title": "Block",   "description": "Mark lorry unavailable for today"},
                {"id": "maint_broken",  "title": "Broken",  "description": "Log breakdown and find replacement"},
                {"id": "maint_release", "title": "Release", "description": "Unblock lorry from today log"},
                {"id": "maint_fixed",   "title": "Fixed",   "description": "Mark broken lorry as repaired"},
            ],
        }]

    # ── Show plate picker for each action ────────────────────────────────────
    def _maint_list(action: str) -> list:
        """
        Build the lorry picker list for the given action.
        For RELEASE and FIXED: prepend a 'Done — apply X' row when batch non-empty.
        BLOCK and BROKEN: single-select, no batch needed.
        """
        engine2 = _get_engine_safe()
        if not engine2:
            return ["Please log in first. Send hi to start."]
        taken_today2 = get_assigned_today()
        broken_map2  = get_broken_lorries()
        batch        = sess.get(f"_maint_batch_{action}", [])

        all_items2   = []   # unselected lorries
        selected_rows = []  # already-batched lorries (shown at top with toggle)

        for _, r in engine2.eligible_lorries.iterrows():
            p2        = str(r["LORRY"]).upper()
            cap2      = float(r["TON"])
            is_broken2  = p2 in broken_map2
            is_blocked2 = p2 in taken_today2
            if action == "BLOCK"   and is_blocked2: continue
            if action == "RELEASE" and not is_blocked2: continue
            if action == "FIXED"   and not is_broken2:  continue
            if action == "BROKEN"  and is_broken2:  continue
            if is_broken2:
                rep2  = broken_map2[p2]
                desc2 = f"{cap2}T | Broken->{rep2}" if rep2 != "NONE" else f"{cap2}T | Broken"
            elif is_blocked2:
                desc2 = f"{cap2}T | Blocked"
            else:
                desc2 = f"{cap2}T | Available"

            if p2 in batch:
                # Already selected — show with checkmark and allow tap-to-deselect
                selected_rows.append({
                    "id":          f"maint_toggle {action} {p2}",
                    "title":       f"[X] {p2}",
                    "description": f"Tap to deselect | {desc2}"[:72],
                })
            else:
                all_items2.append({
                    "id":          f"maint_exec {action} {p2}",
                    "title":       p2,
                    "description": desc2[:72],
                })

        header_map2 = {
            "BLOCK":   "Block a Lorry",
            "BROKEN":  "Log Breakdown",
            "RELEASE": "Release Lorries",
            "FIXED":   "Mark Lorries Fixed",
        }

        if not all_items2 and not batch:
            msg_map2 = {
                "BLOCK":   "All lorries already blocked today.",
                "RELEASE": "No lorries currently blocked.",
                "FIXED":   "No lorries marked as broken.",
                "BROKEN":  "All lorries already marked as broken.",
            }
            return [
                msg_map2.get(action, "No lorries to show."),
                {"_type": "buttons", "body": "What would you like to do?",
                 "buttons": [{"id": "lorry_maint", "title": "Back"},
                             {"id": "hi",          "title": "Main Menu"}]},
            ]

        list_items2 = []

        # For RELEASE / FIXED: show Done row + selected rows at the top,
        # then unselected lorries below (paginated)
        if action in ("RELEASE", "FIXED"):
            if batch:
                selected_str = ", ".join(batch)
                list_items2.append({
                    "id":          f"maint_batch_done {action}",
                    "title":       f"Done ({len(batch)} selected)",
                    "description": f"Confirm: {selected_str}"[:72],
                })
            list_items2 += selected_rows   # checked items always visible, no pagination

        # Paginate unselected items
        # Reserve rows for: Done (1) + selected rows + Next page (1)
        reserved    = (1 if batch else 0) + len(selected_rows) + 1   # +1 for possible Next
        PER_PAGE2   = max(1, 9 - reserved)
        total2      = len(all_items2)
        page2       = int(sess.get("maint_picker_page", {}).get(action, 0))
        total_pages2 = max(1, -(-total2 // PER_PAGE2)) if total2 else 1
        page2        = max(0, min(page2, total_pages2 - 1))
        start2       = page2 * PER_PAGE2
        chunk2       = all_items2[start2:start2 + PER_PAGE2]
        list_items2 += chunk2

        if total_pages2 > 1:
            next_p = (page2 + 1) % total_pages2
            list_items2.append({
                "id":          f"maint_page {action} {next_p}",
                "title":       "Next page...",
                "description": f"Showing {start2+1}-{start2+len(chunk2)} of {total2}",
            })

        body_map2 = {
            "BLOCK":   "Select lorry to block for today:",
            "BROKEN":  "Select lorry to log as broken:",
            "RELEASE": "Tap lorries to release. [X] = selected. Tap Done when ready:",
            "FIXED":   "Tap lorries to mark as fixed. [X] = selected. Tap Done when ready:",
        }

        return [{
            "_type":  "do_list",
            "header": header_map2.get(action, action),
            "body":   body_map2.get(action, "Select lorry:"),
            "button": "Pick Lorry",
            "items":  list_items2,
        }]

    if cmd_lower in ("maint_block", "maint_broken", "maint_release", "maint_fixed"):
        action_map = {
            "maint_block":   "BLOCK",
            "maint_broken":  "BROKEN",
            "maint_release": "RELEASE",
            "maint_fixed":   "FIXED",
        }
        action = action_map[cmd_lower]
        # Clear batch when user re-enters the action from scratch
        if action in ("RELEASE", "FIXED"):
            sess.pop(f"_maint_batch_{action}", None)
        sess.setdefault("maint_picker_page", {})[action] = 0
        return _maint_list(action)

    # ── Page turn for lorry picker ───────────────────────────────────────────
    if cmd_lower.startswith("maint_page "):
        parts = text.strip().split()
        if len(parts) >= 3:
            action_p = parts[1].upper()
            page_p   = int(parts[2])
            sess.setdefault("maint_picker_page", {})[action_p] = page_p
            return _maint_list(action_p)
        return ["Invalid page selection."]

    # ── Commit the multi-select batch ────────────────────────────────────────
    if cmd_lower.startswith("maint_batch_done "):
        action    = text.strip().split(" ", 1)[1].strip().upper()
        batch_key = f"_maint_batch_{action}"
        batch     = sess.pop(batch_key, [])
        if not batch:
            return ["Nothing selected. Tap lorries first, then Done."]
        for p3 in batch:
            if action == "RELEASE":
                release_specific_plates([p3])
                sess.setdefault("unavailable", set()).discard(p3)
            elif action == "FIXED":
                remove_broken_lorry(p3)
                sess.setdefault("unavailable", set()).discard(p3)
        action_label = "released" if action == "RELEASE" else "marked as fixed"
        plates_str   = ", ".join(batch)
        in_active = sess.get("state") in ("CONFIRMING", "REVIEWING") and sess.get("pending_dos")
        follow_up = _build_summary(sess) if in_active else [{
            "_type": "buttons", "body": "Need anything else?",
            "buttons": [{"id": "lorry_maint", "title": "More Actions"},
                        {"id": "hi",          "title": "Main Menu"}]
        }]
        noun = "lorry" if len(batch) == 1 else "lorries"
        return [f"{len(batch)} {noun} {action_label}: {plates_str}"] + follow_up

    # ── Toggle (deselect) a plate that was already queued ────────────────────
    if cmd_lower.startswith("maint_toggle "):
        parts = text.strip().split()
        if len(parts) < 3:
            return ["Invalid selection."]
        action    = parts[1].upper()
        plate     = parts[2].upper()
        batch_key = f"_maint_batch_{action}"
        batch     = sess.setdefault(batch_key, [])
        if plate in batch:
            batch.remove(plate)
        return _maint_list(action)

    # ── Execute action after tapping a plate ─────────────────────────────────
    if cmd_lower.startswith("maint_exec "):
        parts = text.strip().split()
        if len(parts) < 3:
            return ["Invalid selection."]
        action = parts[1].upper()
        plate  = parts[2].upper()

        if action == "BLOCK":
            record_assignments_today([plate])
            sess.setdefault("unavailable", set()).add(plate)
            in_active = sess.get("state") in ("CONFIRMING", "REVIEWING") and sess.get("pending_dos")
            follow_up = _build_summary(sess) if in_active else [{
                "_type": "buttons", "body": "Need anything else?",
                "buttons": [{"id": "lorry_maint", "title": "More Actions"},
                            {"id": "hi",          "title": "Main Menu"}]
            }]
            return [f"{plate} blocked."] + follow_up

        elif action == "BROKEN":
            return handle_message(phone, text=f"select_broken_lorry {plate}")

        elif action in ("RELEASE", "FIXED"):
            # Add to batch and show updated list
            batch_key = f"_maint_batch_{action}"
            batch     = sess.setdefault(batch_key, [])
            if plate not in batch:
                batch.append(plate)
            return _maint_list(action)

    # ── manage [PLATE]: shortcut ──────────────────────────────────────────────
    if cmd_lower.startswith("manage ") and not cmd_lower.startswith("manage_lorry_pick"):
        shortcut_plate = text.split(" ", 1)[1].strip().upper()
        return handle_message(phone, text=f"manage_lorry_pick {shortcut_plate}")

    # ── manage_lorry_pick: show action buttons for a chosen plate ────────────
    if cmd_lower.startswith("manage_lorry_pick "):
        plate       = text.split(" ", 1)[1].strip().upper()
        taken_today = get_assigned_today()
        broken_map  = get_broken_lorries()
        is_broken   = plate in broken_map
        is_blocked  = plate in taken_today

        buttons = []
        if not is_blocked:
            buttons.append({"id": f"select_block_lorry {plate}",   "title": "🚫 Block"})
        if not is_broken:
            buttons.append({"id": f"select_broken_lorry {plate}",  "title": "🔧 Broken"})
        if is_blocked:
            buttons.append({"id": f"select_release_lorry {plate}", "title": "🔓 Release"})
        if is_broken:
            buttons.append({"id": f"select_fixed_lorry {plate}",   "title": "✅ Fixed"})
        if len(buttons) < 3:
            buttons.append({"id": "lorry_maint", "title": "↩️ Pick Another"})

        engine = _get_engine_safe()
        cap    = ""
        if engine is not None:
            row = engine.eligible_lorries[engine.eligible_lorries["LORRY"] == plate]
            if not row.empty:
                cap = f"{float(row.iloc[0]['TON'])}T"

        status_parts = []
        if is_broken:
            rep = broken_map[plate]
            status_parts.append(f"🔧 Broken — replacement: {rep}")
        if is_blocked:
            status_parts.append("⛔ Assigned/blocked today")
        if not is_broken and not is_blocked:
            status_parts.append("✅ Available")

        body = f"*{plate}*  {cap}\n" + "\n".join(status_parts) + "\n\nWhat would you like to do?"
        return [{"_type": "buttons", "body": body, "buttons": buttons[:3]}]


    # ── Tap-only lorry action handlers ──────────────────────────────────────
    # These are triggered from manage_lorry_pick buttons

    if text.lower().startswith("select_block_lorry "):
        plate  = text.split(" ", 1)[1].strip().upper()
        engine = sess.get("engine")
        all_plates = set(engine.all_lorries["LORRY"].str.upper()) if engine else set()
        if all_plates and plate not in all_plates:
            return [f"⚠️ *{plate}* not found in master list."]
        if plate in sess.get("unavailable", set()) or plate in get_assigned_today():
            return [f"⚠️ *{plate}* is already blocked today."]
        sess.setdefault("unavailable", set()).add(plate)
        record_assignments_today([plate])
        in_active = sess.get("state") in ("CONFIRMING", "REVIEWING") and sess.get("pending_dos")
        follow_up = _build_summary(sess) if in_active else [_HI_BTN]
        return [f"🚫 *{plate}* blocked for today."] + follow_up

    if text.lower().startswith("select_release_lorry "):
        plate  = text.split(" ", 1)[1].strip().upper()
        engine = sess.get("engine")
        all_plates = set(engine.all_lorries["LORRY"].str.upper()) if engine else set()
        if all_plates and plate not in all_plates:
            close = _find_close_plate(plate, all_plates)
            hint  = f"\nDid you mean *{close}*?" if close else ""
            return [f"⚠️ *{plate}* not found in master list.{hint}"]
        released = release_specific_plates([plate])
        in_active = sess.get("state") in ("CONFIRMING", "REVIEWING") and sess.get("pending_dos")
        follow_up = _build_summary(sess) if in_active else [_HI_BTN]
        if released:
            sess.setdefault("unavailable", set()).discard(plate)
            return [f"✅ *{plate}* released and available again."] + follow_up
        return [f"⚠️ *{plate}* was not in today's log (already available)."] + follow_up

    if text.lower().startswith("select_fixed_lorry "):
        plate = text.split(" ", 1)[1].strip().upper()
        in_active = sess.get("state") in ("CONFIRMING", "REVIEWING") and sess.get("pending_dos")
        follow_up = _build_summary(sess) if in_active else [_HI_BTN]
        if remove_broken_lorry(plate):
            sess.setdefault("unavailable", set()).discard(plate)
            return [f"✅ *{plate}* marked as fixed and unblocked."] + follow_up
        return [f"⚠️ *{plate}* was not in today's broken list."] + follow_up

    if text.lower().startswith("select_replacement "):
        parts         = text.strip().split(" ", 2)
        broken_plate  = parts[1].strip().upper() if len(parts) > 1 else ""
        replace_plate = parts[2].strip().upper() if len(parts) > 2 else "NONE"
        if not broken_plate:
            return ["❌ Invalid selection."]
        record_broken_lorry(broken_plate, replace_plate)
        sess.setdefault("unavailable", set()).add(broken_plate)
        sess.pop("pending_broken_plate", None)
        rep_str = f"replaced by *{replace_plate}*" if replace_plate != "NONE" else "no replacement"
        in_active = sess.get("state") in ("CONFIRMING", "REVIEWING") and sess.get("pending_dos")
        follow_up = _build_summary(sess) if in_active else [_HI_BTN]
        return [
            f"🔧 *Breakdown logged:*\n"
            f"  ❌ Broken:      *{broken_plate}*\n"
            f"  ✅ Replacement: {rep_str}\n\n"
            f"*{broken_plate}* is blocked for today.",
        ] + follow_up

    if text.lower().startswith("select_broken_lorry "):
        broken_plate = text.split(" ", 1)[1].strip().upper()
        engine = sess.get("engine")
        all_plates = set(engine.all_lorries["LORRY"].str.upper()) if engine else set()
        if all_plates and broken_plate not in all_plates:
            return [f"⚠️ *{broken_plate}* not found in master list."]
        excl = sess.get("unavailable", set()) | get_assigned_today() | {broken_plate}
        avail = []
        if engine:
            for _, r in engine.eligible_lorries.iterrows():
                if r["LORRY"] not in excl:
                    cap    = float(r["TON"])
                    status = "Blocked" if r["LORRY"] in get_assigned_today() else "Available"
                    avail.append({
                        "id":          f"select_replacement {broken_plate} {r['LORRY']}",
                        "title":       str(r["LORRY"])[:24],
                        "description": f"{cap}T | {status}",
                    })
        # WhatsApp list: max 10 rows per section, 1 section = 10 total max
        # Keep best 9 lorries + "No replacement" to always stay in one section
        avail = avail[:9]
        avail.append({
            "id":          f"select_replacement {broken_plate} NONE",
            "title":       "No replacement",
            "description": "Block only, no replacement needed",
        })
        sess["pending_broken_plate"] = broken_plate
        return [{
            "_type":  "do_list",
            "header": f"{broken_plate} broken - pick replacement",
            "body":   f"Select a replacement lorry for {broken_plate}:",
            "button": "Pick Lorry",
            "items":  avail,
        }]

    # ── Broken lorry commands ─────────────────────────────────────────────────
    # broken [PLATE]            — mark lorry as broken, bot asks for replacement
    # broken [PLATE] [REPLACE]  — mark broken + set replacement in one go
    # fixed [PLATE]             — mark lorry as repaired, remove from broken list
    # broken list               — show all broken lorries and replacements today

    if cmd_lower == "broken list":
        return _handle_broken_list(sess)

    if cmd_lower.startswith("fixed "):
        plate = text.split(" ", 1)[1].strip().upper()
        if remove_broken_lorry(plate):
            sess["unavailable"].discard(plate)
            return [
                f"✅ *{plate}* marked as *fixed* and unblocked for today.",
                {"_type": "buttons", "body": "What would you like to do next?",
                 "buttons": [{"id": "hi", "title": "👋 Hi"}]},
            ]
        return [f"⚠️ *{plate}* was not in today's broken list."]

    # ── release [PLATE1] [PLATE2...] ─────────────────────────────────────────
    # Remove specific plates from today's assigned log (lorry repaired/available again)
    # Does NOT require the plate to be in the broken list — handles any reason
    # Usage:
    #   release VJN9910            — release one plate
    #   release VJN9910 BQU3875    — release multiple plates at once
    if cmd_lower.startswith("release "):
        plates_to_release = [p.upper() for p in text.strip().split()[1:] if p]
        if not plates_to_release:
            return ["Usage: *release [PLATE1] [PLATE2...]*\ne.g. *release VJN9910* or *release VJN9910 BQU3875*"]

        engine       = sess.get("engine")
        all_plates   = set(engine.all_lorries["LORRY"].str.upper()) if engine else set()
        log          = _load_daily_log()
        assigned_set = set(log["assigned"])
        broken_map   = log.get("broken", {})

        released   = []
        not_in_log = []
        typos      = []   # (typed, suggested_correct)

        for plate in plates_to_release:
            # Step 1: check if plate exists in master list at all
            if all_plates and plate not in all_plates:
                close = _find_close_plate(plate, all_plates)
                typos.append((plate, close))
                continue
            # Step 2: check if it's in today's log
            if plate in assigned_set:
                assigned_set.discard(plate)
                broken_map.pop(plate, None)
                sess["unavailable"].discard(plate)
                released.append(plate)
            else:
                not_in_log.append(plate)

        log["assigned"] = sorted(assigned_set)
        log["broken"]   = broken_map
        _save_daily_log(log)

        lines = []
        if released:
            lines.append(f"✅ Released & available again: *{', '.join(released)}*")
        if not_in_log:
            lines.append(f"⚠️ Not in today's log (already available): *{', '.join(not_in_log)}*")
        if typos:
            for typed, close in typos:
                hint = f" → Did you mean *{close}*?" if close else ""
                lines.append(f"❌ *{typed}* not found in master list.{hint}")
        if not released and not lines:
            lines.append("No changes made.")

        # ── Re-evaluate active DOs against the newly released lorry(s) ──────────
        reassigned = []
        active_states = ("CONFIRMING", "REVIEWING")
        if released and engine and sess.get("state") in active_states and sess.get("pending_dos"):
            taken = (sess["unavailable"] | get_assigned_today()) - set(released)
            for do in sess["pending_dos"]:
                do_num = do["DO NUMBER"]
                for item in do.get("ITEMS", []):
                    current_lorry = item.get("LORRY", "")
                    if current_lorry in ("NO_LORRY", "SPLIT", "SKIPPED", ""):
                        continue
                    weight = item["WEIGHT"]
                    route  = item["ROUTE"]
                    cur_row = engine.eligible_lorries[
                        engine.eligible_lorries["LORRY"] == current_lorry
                    ]
                    if cur_row.empty:
                        continue
                    cur_surplus = float(cur_row.iloc[0]["TON"]) - weight
                    excl = (taken | sess["unavailable"]) - set(released) - {current_lorry}
                    best = engine.suggest(route=route, total_ton=weight,
                                         unavailable=excl, top_n=1)
                    if not best:
                        continue
                    best_lorry   = best[0]["LORRY"]
                    best_surplus = best[0]["SURPLUS"]
                    if best_lorry != current_lorry and best_surplus < cur_surplus - 0.001:
                        old = current_lorry
                        item["LORRY"] = best_lorry
                        sess["assigned"][do_num] = best_lorry
                        sess["unavailable"].discard(old)
                        sess["unavailable"].add(best_lorry)
                        taken.add(best_lorry)
                        taken.discard(old)
                        reassigned.append(f"  • {do_num}: {old} → *{best_lorry}* "
                                          f"({round(best_surplus,2)}T spare vs {round(cur_surplus,2)}T)")

        if reassigned:
            lines.append("\n🔄 *Better fits found after release:*\n" + "\n".join(reassigned))
        else:
            lines.append("\nThese lorries will now appear in suggestions for new DOs.")

        if sess.get("state") in active_states and sess.get("pending_dos"):
            return ["\n".join(lines), _build_summary(sess)]
        return ["\n".join(lines), {"_type": "buttons",
                                    "body": "What would you like to do next?",
                                    "buttons": [{"id": "hi", "title": "👋 Hi"}]}]

    if cmd_lower.startswith("broken "):
        parts = text.strip().split()
        broken_plate  = parts[1].upper() if len(parts) > 1 else None
        replace_plate = parts[2].upper() if len(parts) > 2 else None
        if not broken_plate:
            return ["Usage: *broken [PLATE]* or *broken [PLATE] [REPLACEMENT]*"]
        if replace_plate and broken_plate == replace_plate:
            return [f"⚠️ Replacement cannot be the same as the broken lorry (*{broken_plate}*)."]
        if replace_plate:
            record_broken_lorry(broken_plate, replace_plate)
            sess["unavailable"].add(broken_plate)
            return _broken_confirmed_reply(broken_plate, replace_plate, sess)
        else:
            sess["state_before_broken"] = sess["state"]
            sess["pending_broken_plate"] = broken_plate
            sess["state"] = "AWAIT_BROKEN_REPLACEMENT"
            engine = sess.get("engine")

            # ── Find what this broken lorry was assigned to ──────────────────
            # Look through current session items to get route + weight context
            broken_items = []
            for do in sess.get("pending_dos", []):
                for it in do.get("ITEMS", []):
                    if it.get("LORRY") == broken_plate:
                        broken_items.append((it, do))
                    elif it.get("LORRY") == "SPLIT":
                        for b in (it.get("SPLIT_LORRIES") or []):
                            if b.get("lorry") == broken_plate:
                                broken_items.append((it, do))

            # ── Auto-recommend best replacement ─────────────────────────────
            best_recs = []   # list of (plate, cap, util_pct, reason)
            excl = sess["unavailable"] | get_assigned_today() | {broken_plate}

            if engine is not None and broken_items:
                # Use first broken item's route + weight for recommendation
                ref_item, ref_do = broken_items[0]
                weight = ref_item.get("WEIGHT", 0)
                route  = ref_item.get("ROUTE", "")
                sug = engine.suggest(route=route, total_ton=weight,
                                     unavailable=excl, top_n=3)
                for s in sug:
                    cap  = s["TON_CAPACITY"]
                    util = round(weight / cap * 100, 1) if cap > 0 else 0
                    best_recs.append((s["LORRY"], cap, util, s["REASON"]))

            elif engine is not None:
                # No current assignment context — suggest by capacity only
                avail = engine.eligible_lorries[
                    ~engine.eligible_lorries["LORRY"].isin(excl)
                ].copy()
                avail = avail.sort_values("TON", ascending=False)
                for _, r in avail.head(3).iterrows():
                    best_recs.append((r["LORRY"], r["TON"], None, "Available lorry"))

            # ── Build message + buttons ──────────────────────────────────────
            lines = [f"🔧 *{broken_plate}* marked as broken."]

            if broken_items:
                items_str = ", ".join(
                    f"{do.get('CUSTOMER NAME','')[:18]} {round(it.get('WEIGHT',0),1)}T ({do.get('DO NUMBER','')})"
                    for it, do in broken_items[:3]
                )
                lines.append(f"Was assigned to: {items_str}")

            lines.append("")
            if best_recs:
                lines.append("🚛 *Recommended replacements:*")
                for plate, cap, util, reason in best_recs:
                    util_str = f" — {util}% util" if util is not None else ""
                    lines.append(f"  • *{plate}* ({cap}T){util_str}")
                lines.append("")
                lines.append("Tap a button to assign, or type any plate manually:")
            else:
                lines.append("No suitable replacement found automatically.")
                lines.append("Type a plate manually or *none* to skip.")

            # Buttons: top 2 recommendations + none option (max 3 buttons)
            btns = []
            for plate, cap, util, _ in best_recs[:2]:
                util_tag = f" {util}%" if util is not None else ""
                btns.append({"id": plate, "title": f"🚛 {plate}{util_tag}"[:20]})
            btns.append({"id": "none", "title": "⏭️ No replacement"})

            return [
                "\n".join(lines),
                {"_type": "buttons", "body": "Choose replacement:", "buttons": btns}
            ]


    if text.lower() == "show assigned today":
        today_plates = get_assigned_today()
        if not today_plates:
            return ["No lorries assigned yet today.", _HI_BTN]
        sess = get_session(phone)
        engine = sess.get("engine")
        if engine is not None:
            user_lorries = set(engine.eligible_lorries["LORRY"].str.upper())
            my_plates    = sorted(today_plates & user_lorries)
            other_plates = sorted(today_plates - user_lorries)
            lines = [f"🚛 *Lorries assigned today ({sess['user_id']}):*"]
            if my_plates:
                lines += [f"  • {p}" for p in my_plates]
            else:
                lines.append("  (none of your lorries assigned yet)")
            if other_plates:
                lines.append(f"\n_Other users: {', '.join(other_plates)}_")
            return ["\n".join(lines), _HI_BTN]
        return ["🚛 *Lorries already assigned today:*\n" +
                "\n".join(f"  • {p}" for p in sorted(today_plates)), _HI_BTN]

    if text.lower() in ("show blocked", "block list", "blocked list"):
        sess   = get_session(phone)
        engine = sess.get("engine")
        today  = get_assigned_today()
        broken_map = get_broken_lorries()   # {broken: replacement}

        # Only show this user's eligible lorries that are currently blocked
        user_lorries = set()
        if engine is not None:
            user_lorries = set(engine.eligible_lorries["LORRY"].str.upper())

        blocked_plates = sorted(today & user_lorries) if user_lorries else sorted(today)
        broken_plates  = set(broken_map.keys())

        if not blocked_plates:
            return ["✅ No blocked lorries for you today. All available.", _HI_BTN]

        lines = ["🚫 *Blocked lorries today:*\n"]
        for p in blocked_plates:
            if p in broken_plates:
                rep = broken_map[p]
                rep_str = f" → replaced by *{rep}*" if rep != "NONE" else " (no replacement)"
                lines.append(f"  🔧 *{p}* — broken{rep_str}")
            else:
                lines.append(f"  ⛔ *{p}* — assigned/blocked")

        lines.append("\nTap a plate to release it, or type *release [PLATE]*:")

        # Buttons: up to 3 plates as tappable release buttons
        # Button ID = "release PLATE" so it's handled by existing release handler
        btns = [
            {"id": f"release {p}", "title": f"🔓 {p}"[:20]}
            for p in blocked_plates[:3]
        ]

        return [
            "\n".join(lines),
            {"_type": "buttons", "body": "Tap to release:", "buttons": btns}
        ]

    if text.lower() in ("hi", "hello", "start"):
        return _start(phone, sess)

    # ── Bare plate release shortcut ───────────────────────────────────────────
    # If user types just a plate (e.g. "Wld8738") and it's in their blocked list,
    # treat it as "release [PLATE]" — works in any state
    if re.match(r'^[A-Za-z0-9]{4,10}$', text.strip()):
        candidate = text.strip().upper()
        _eng = sess.get("engine")
        if _eng is None and sess.get("user_id"):
            try:
                _hist = _resolve_history_path()
                _eng = LorryEngine(MASTER_PATH, _hist, owner_user=sess["user_id"])
            except Exception:
                pass
        _ul = set(_eng.eligible_lorries["LORRY"].str.upper()) if _eng else set()
        if candidate in (get_assigned_today() & _ul):
            return handle_message(phone, text=f"release {candidate}")

    if state == "IDLE":
        return _start(phone, sess)
    elif state == "AWAIT_USER_ID":
        return _handle_user_id(phone, sess, text)
    elif state == "AWAIT_EXCEL":
        if file_bytes:
            return _handle_excel_upload(phone, sess, file_bytes)
        return ["Please upload the DO Excel file (.xlsx) to continue."]
    elif state == "REVIEWING":
        return _handle_reviewing(phone, sess, text)
    elif state == "CONFIRMING":
        return _handle_confirming(phone, sess, text)
    elif state == "DONE":
        # After export, user may still block/change a lorry — revert to CONFIRMING,
        # process the command, then auto re-export and send the updated file
        cmd = text.lower().strip()
        if (cmd.startswith("block ") or cmd.startswith("change ") or
                cmd.startswith("release ")):
            sess["state"] = "CONFIRMING"
            msgs = _handle_confirming(phone, sess, text)
            # If the command changed assignments, re-export automatically
            if sess.get("state") == "CONFIRMING" and sess.get("pending_dos"):
                try:
                    export_msgs = _export_result(sess)
                    msgs += export_msgs
                except Exception as _e:
                    msgs.append(f"⚠️ Could not regenerate export: {_e}")
            return msgs
        return ["Your assignments have been exported. Send *hi* to start a new session."]
    elif state == "AWAIT_BROKEN_REPLACEMENT":
        return _handle_broken_replacement(phone, sess, text)

    return ["Sorry, I didn't understand that. Send *hi* to start."]


# ── State handlers ────────────────────────────────────────────────────────────

def _get_valid_users() -> list[str]:
    """Read all non-SPARE users dynamically from master lorry file."""
    try:
        df = pd.read_excel(MASTER_PATH)
        df.columns = [c.strip().upper() for c in df.columns]
        users = [u for u in df["USER"].str.strip().str.upper().unique()
                 if u != "SPARE"]
        return sorted(users)
    except Exception:
        return ["ABI", "VIVIAN", "SELAYANG", "BIG"]  # fallback


def _start(phone, sess):
    sess["state"] = "AWAIT_USER_ID"
    users = _get_valid_users()
    # SELAYANG is type-only; all others get clickable buttons (max 3)
    NO_BUTTON  = {"SELAYANG"}
    btn_users  = [u for u in users if u not in NO_BUTTON][:3]
    type_users = [u for u in users if u in NO_BUTTON or u not in btn_users]
    body = "👋 *Lorry Assignment Bot*\n\nPlease tap your name below or type it to continue."
    if type_users:
        body += f"\nOr type: {', '.join(u.title() for u in type_users)}"
    return [{
        "_type": "buttons",
        "body": body,
        "buttons": [{"id": u.lower(), "title": u.title()} for u in btn_users],
    }]


def _handle_user_id(phone, sess, text):
    valid_users = _get_valid_users()
    user = text.upper().strip()
    if user not in valid_users:
        return [f"❌ User not recognised. Please reply with one of: {', '.join(valid_users)}"]

    sess["user_id"] = user
    # Use new-format history if it exists, else fall back to old format
    _hist = _resolve_history_path()
    sess["engine"] = LorryEngine(MASTER_PATH, _hist, owner_user=user)
    sess["state"] = "AWAIT_EXCEL"

    lorries = sess["engine"].get_eligible_lorry_list()

    taken_today  = get_assigned_today()
    broken_today = get_broken_lorries()   # {plate: replacement}
    lines = []
    for _, r in lorries.iterrows():
        plate = r["LORRY"]
        ton   = r["TON"]
        user  = r["USER"]
        if plate in broken_today:
            rep = broken_today[plate]
            tag = f" 🔴 Broken→{rep}" if rep != "NONE" else " 🔴 Broken"
        elif plate in taken_today:
            tag = " ⛔ Assigned today"
        else:
            tag = " ✅ Available"
        lines.append(f"  • {plate} — {ton}T ({user}){tag}")

    lorry_text = (
        f"✅ Logged in as *{user}*\n\n"
        f"Your lorries:\n" + "\n".join(lines) + "\n\n"
        "📎 Now please upload your DO Excel file (.xlsx)."
    )
    # Quick-action buttons (max 3 for WhatsApp)
    buttons_msg = {
        "_type": "buttons",
        "body": "Quick actions:",
        "buttons": [
            {"id": "clear daily log",     "title": "🗑️ Clear Today Log"},
            {"id": "show assigned today", "title": "📋 Show Assigned"},
            {"id": "manage lorry",        "title": "🔧 Manage Lorry"},
        ],
    }

    return [lorry_text, buttons_msg]



def _handle_prefilled_excel(phone, sess, raw: "pd.DataFrame", prefilled: "pd.DataFrame") -> list[str]:
    """
    Called when the uploaded Excel already has LICENSE plates filled in.
    Reads every plate (including comma-separated split plates), registers them
    as assigned today in the daily log, and returns a clear summary message.
    """
    SENTINELS = {"SKIPPED", "NO_LORRY", "SPLIT", "", "nan", "none", "n/a", "-", None}

    # Collect all plates from LICENSE column (handles "BMN3682, WUD4927" too)
    plates_found = []
    rows_summary = []

    for _, row in prefilled.iterrows():
        lic_raw   = str(row.get("LICENSE", "")).strip()
        customer  = str(row.get("CUSTOMER NAME", "")).strip()
        do_num    = str(row.get("DO NUMBER", "")).strip()
        # Support both formats: WEIGHT(T) (old) or GROSS WEIGHT converted (new)
        weight    = row.get("WEIGHT(T)", row.get("GROSS WEIGHT", ""))
        if sess.get("is_new_format") and weight != "":
            try:
                weight = round(float(weight) / 1000, 3)
            except Exception:
                pass
        itmref    = str(row.get("ITMREF_0", "")).strip()

        # Split on comma to handle multi-lorry cells
        plates_in_cell = [p.strip().upper() for p in lic_raw.split(",")
                          if p.strip().upper() not in {s.upper() for s in SENTINELS if s}]
        if not plates_in_cell:
            continue

        plates_found.extend(plates_in_cell)
        plate_display = ", ".join(f"*{p}*" for p in plates_in_cell)
        itmref_str    = f" ({itmref})" if itmref and itmref.lower() not in ("nan","") else ""
        rows_summary.append(
            f"  🚛 {plate_display}  ←  {customer}{itmref_str}  {do_num}  {weight}T"
        )

    if not plates_found:
        # All LICENSE cells were empty despite column existing — fall through to auto-assign
        return None  # caller must handle None → proceed with normal auto-assign

    # Register in daily log
    record_assignments_today(plates_found)

    unique_plates = sorted(set(plates_found))
    lines = []
    lines.append(f"📋 *Pre-filled assignment detected!*")
    lines.append(f"Found *{len(unique_plates)} lorry plate(s)* in the uploaded file.")
    lines.append("Registered as assigned today:")
    for p in unique_plates:
        lines.append(f"  ⛔ *{p}*")
    lines.append("")
    lines.append("─────────────────────")
    lines.append("*Row details:*")
    lines.extend(rows_summary)
    lines.append("─────────────────────")
    lines.append(f"✅ These lorries are now marked *unavailable* for today's auto-assignment.")

    # Reset session — user may want to do a fresh assignment next
    reset_session(phone)

    return [
        "\n".join(lines),
        {
            "_type": "buttons",
            "body": "Start a new session or check what's assigned today.",
            "buttons": [
                {"id": "hi",                  "title": "👋 Hi"},
                {"id": "show assigned today",  "title": "📋 Show Assigned"},
            {"id": "show blocked",         "title": "🚫 Show Blocked"},
            ],
        }
    ]

def _handle_excel_upload(phone, sess, file_bytes):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        df.columns = [c.strip().upper() for c in df.columns]

        # ── Detect format: new (ZSDOROUTEWRH) vs old ────────────────────────
        # New format: GROSS WEIGHT (kg), no WEIGHT(T) or ITMREF_0
        # Old format: WEIGHT(T), ITMREF_0
        IS_NEW_FORMAT = "GROSS WEIGHT" in df.columns and "WEIGHT(T)" not in df.columns

        if IS_NEW_FORMAT:
            # New format required columns
            required = {"DO NUMBER", "ROUTE", "CUSTOMER NAME", "GROSS WEIGHT"}
            missing = required - set(df.columns)
            if missing:
                return [f"❌ Missing columns: {', '.join(missing)}\nPlease check and re-upload."]
            # Convert GROSS WEIGHT kg → tonnes, store as WEIGHT(T) internally
            df["WEIGHT(T)"] = pd.to_numeric(df["GROSS WEIGHT"], errors="coerce").fillna(0) / 1000.0
            if "CODE" not in df.columns:
                df["CODE"] = ""
            # DATE arrives as "2026-05-11 00:00:00" (string of a datetime serial).
            # Strip the time part so it exports as "2026-05-11".
            if "DATE" in df.columns:
                df["DATE"] = (
                    df["DATE"]
                    .astype(str)
                    .str.strip()
                    .str.split(" ").str[0]   # keep only "2026-05-11"
                    .replace({"nan": "", "NaT": "", "None": ""})
                )
            # Convert DATE to formatted string NOW so raw_df never stores datetime64.
            # pandas to_excel will re-serialise datetime64 as a date cell regardless
            # of any later string assignment — converting at source is the only safe fix.
            if "DATE" in df.columns:
                def _fmt_date_on_load(v):
                    if not isinstance(v, str) and pd.isna(v):
                        return ""
                    try:
                        ts = pd.to_datetime(v, errors="coerce")
                        if pd.isna(ts):
                            return str(v)
                        return ts.strftime("%-d/%-m/%y")
                    except Exception:
                        return str(v)
                df["DATE"] = df["DATE"].apply(_fmt_date_on_load).astype(str)
        else:
            # Old format required columns
            required = {"CODE", "CUSTOMER NAME", "ROUTE", "DO NUMBER", "WEIGHT(T)"}
            missing = required - set(df.columns)
            if missing:
                return [f"❌ Missing columns: {', '.join(missing)}\nPlease check and re-upload."]

        if "LICENSE" not in df.columns:
            df["LICENSE"] = ""

        raw = df.dropna(subset=["ROUTE", "DO NUMBER"]).copy()
        raw = raw.reset_index(drop=True)
        # Store format flag so export knows not to touch DATE
        sess["is_new_format"] = IS_NEW_FORMAT

        # ── Pre-filled detection ─────────────────────────────────────────────
        # If the uploaded Excel already has LICENSE plates in it, the user is
        # importing a completed assignment sheet (not asking us to auto-assign).
        # Read those plates, register them as assigned today, and show a summary.
        SENTINELS_STR = {"", "nan", "none", "n/a", "-"}
        prefilled_rows = raw[
            raw["LICENSE"].astype(str).str.strip().str.lower()
            .isin(SENTINELS_STR) == False
        ].copy()

        if not prefilled_rows.empty:
            result = _handle_prefilled_excel(phone, sess, raw, prefilled_rows)
            if result is not None:
                return result
            # result is None → all plates were sentinels, fall through to auto-assign

        # ── Build item list: one item per Excel row ─────────────────────────
        # Each row is an independent item that needs its own lorry.
        # Items with the same DO NUMBER belong to the same customer/route
        # but may end up on different lorries (e.g. 17.5T row vs 3.5T row).
        items = []
        for idx, row in raw.iterrows():
            items.append({
                "ROW_IDX":       idx,                         # original df index
                "DO NUMBER":     str(row["DO NUMBER"]).strip(),
                "CUSTOMER NAME": str(row["CUSTOMER NAME"]).strip(),
                "ROUTE":         str(row["ROUTE"]).strip(),
                "CODE":          str(row["CODE"]).strip(),
                "WEIGHT":        float(row["WEIGHT(T)"]),
                "ITMREF":        str(row.get("ITMREF_0", "")).strip(),
                # assignment fields filled below
                "LORRY":         None,   # plate string, or "SPLIT", "NO_LORRY"
                "SPLIT_LORRIES": None,   # list of bins if split
            })

        sess["items"]      = items          # row-level item list
        sess["raw_df"]     = raw
        sess["unavailable"] = set()
        sess["assigned"]   = {}             # kept for change/block compat (item ROW_IDX → lorry)
        sess["state"]      = "CONFIRMING"

        engine: LorryEngine = sess["engine"]

        # ── Auto-assign: two-pass global optimiser ─────────────────────────
        # PROBLEM with naive heaviest-first:
        #   Small lorries (e.g. VEA2818 1.07T) get consumed by slightly-heavier
        #   tiny DOs before even-lighter ones are processed, causing the lightest
        #   DO to get a much-too-large lorry.
        #
        # SOLUTION — two passes:
        #   Pass 1: loads ABOVE the smallest lorry capacity → heaviest first
        #           (these need large lorries; process early to claim them)
        #   Pass 2: loads AT OR BELOW the smallest lorry capacity → LIGHTEST first
        #           (tiny loads get the smallest available lorry — no waste)
        #
        # Within each pass, route frequency still acts as a tiebreaker.

        broken_map = get_broken_lorries()
        sess["unavailable"].update(broken_map.keys())

        eligible_caps   = sorted(engine.eligible_lorries["TON"].tolist())
        smallest_cap    = eligible_caps[0] if eligible_caps else 1.0

        # ── Rule 1: One route = one lorry (Same Route Same Day) ─────────────────
        # Each distinct route is a separate delivery run departing from and
        # returning to the same warehouse on the same day.  Merging different
        # routes onto one lorry is physically impossible without geo-distance
        # data, so we keep each route as its own group.
        # Rule 6 (max 8 stops) is enforced inside engine.suggest().
        from collections import defaultdict
        route_groups: dict[str, list] = defaultdict(list)
        for it in items:
            key = it["ROUTE"].strip().upper()
            route_groups[key].append(it)

        # Sort heaviest-route groups first so they claim the best-fit lorry
        # before lighter routes do (avoids a 14T lorry being grabbed by a
        # 0.5T route when a heavy route also needs it).
        sorted_groups = sorted(
            route_groups.values(),
            key=lambda grp: sum(it["WEIGHT"] for it in grp),
            reverse=True,
        )

        def _assign_group(group_items):
            """Assign ONE lorry (or split) to cover ALL items in the group.
            All items in the group share the same route (one route = one lorry).
            """
            total_w  = sum(it["WEIGHT"] for it in group_items)
            route    = group_items[0]["ROUTE"]
            customer = group_items[0]["CUSTOMER NAME"]

            broken_map = get_broken_lorries()
            sess["unavailable"].update(broken_map.keys())
            excluded = sess["unavailable"] | get_assigned_today()

            # Try single lorry for the whole group
            suggestions = engine.suggest(
                route=route,
                total_ton=total_w,
                unavailable=excluded,
                top_n=1,
                customer_name=customer,
                today_date_str=_today(),
            )

            if suggestions:
                single_cap  = suggestions[0]["TON_CAPACITY"]
                single_util = total_w / single_cap if single_cap > 0 else 0
                split_option = None
                if single_util < 0.60:
                    split_option = engine.suggest_split(
                        route=route,
                        total_ton=total_w,
                        unavailable=excluded,
                        single_util_threshold=0.60,
                    )
                if split_option is not None:
                    # Distribute items across split lorries
                    bins = [
                        {"lorry": s["LORRY"],
                         "rows":  [],
                         "cap":   s["TON_CAPACITY"] - s["PORTION"],
                         "remain": s["PORTION"]}
                        for s in split_option
                    ]
                    for s in split_option:
                        sess["unavailable"].add(s["LORRY"])
                    # Assign each item to a bin that still has capacity
                    for it in sorted(group_items, key=lambda x: x["WEIGHT"], reverse=True):
                        placed = False
                        for bin_ in bins:
                            if bin_["remain"] >= it["WEIGHT"] - 0.001:
                                bin_["rows"].append({"DO": it["DO NUMBER"], "W": it["WEIGHT"]})
                                bin_["remain"] -= it["WEIGHT"]
                                placed = True
                                break
                        if not placed:
                            bins[0]["rows"].append({"DO": it["DO NUMBER"], "W": it["WEIGHT"]})
                    # Write to each item
                    for it in group_items:
                        it["LORRY"]         = "SPLIT"
                        it["SPLIT_LORRIES"] = [b for b in bins if b["rows"]]
                else:
                    lorry = suggestions[0]["LORRY"]
                    sess["unavailable"].add(lorry)
                    for it in group_items:
                        it["LORRY"] = lorry
            else:
                # No single lorry — bin-pack across multiple lorries
                remain = total_w
                bins   = []
                for _ in range(10):
                    if remain <= 0:
                        break
                    excl = sess["unavailable"] | get_assigned_today()
                    sug  = engine.suggest(route=route, total_ton=min(remain, 0.01),
                                          unavailable=excl, top_n=20,
                                          customer_name=customer)
                    if not sug:
                        sug = engine.suggest(route=route, total_ton=0.001,
                                             unavailable=excl, top_n=20)
                    if not sug:
                        break
                    sug.sort(key=lambda x: x["TON_CAPACITY"], reverse=True)
                    lorry = sug[0]["LORRY"]
                    cap   = sug[0]["TON_CAPACITY"]
                    portion = min(cap, remain)
                    bins.append({"lorry": lorry, "rows": [], "cap": cap - portion})
                    sess["unavailable"].add(lorry)
                    remain = round(remain - cap, 6)

                if remain <= 0 and bins:
                    # Distribute items into bins
                    for it in sorted(group_items, key=lambda x: x["WEIGHT"], reverse=True):
                        for bin_ in bins:
                            bin_["rows"].append({"DO": it["DO NUMBER"], "W": it["WEIGHT"]})
                            break
                    for it in group_items:
                        it["LORRY"]         = "SPLIT"
                        it["SPLIT_LORRIES"] = bins
                else:
                    # Bin-pack failed — every lorry is taken or too small.
                    # Assign the largest still-available lorry as a last resort
                    # (overloaded is better than unassigned per business rule).
                    excl_final = sess["unavailable"] | get_assigned_today()
                    last_resort = engine.suggest_largest_available(
                        route, excl_final, _today())
                    if last_resort:
                        lorry = last_resort[0]["LORRY"]
                        sess["unavailable"].add(lorry)
                        for it in group_items:
                            it["LORRY"] = lorry
                    else:
                        # Truly no lorries left at all
                        for it in group_items:
                            it["LORRY"] = "NO_LORRY"

            for it in group_items:
                sess["assigned"][it["DO NUMBER"]] = it["LORRY"]

        for group in sorted_groups:
            _assign_group(group)

        for item in items:
            sess["assigned"][item["DO NUMBER"]] = item["LORRY"]

        # ── (legacy for-loop removed — replaced by _assign_one above) ────────
        # The block below was the old heaviest-first loop.  Keep a dummy
        # reference so diff is minimal.
        # ── Build display groups: group items by DO NUMBER, preserving order ─
        # pending_dos is used by _build_summary; rebuild from items
        seen_do = {}
        pending_dos = []
        for item in items:
            do_num = item["DO NUMBER"]
            if do_num not in seen_do:
                seen_do[do_num] = len(pending_dos)
                pending_dos.append({
                    "DO NUMBER":     do_num,
                    "ALL_DO_NUMBERS": [do_num],
                    "ROUTE":         item["ROUTE"],
                    "CODE":          item["CODE"],
                    "CUSTOMER NAME": item["CUSTOMER NAME"],
                    "ITEMS":         [],          # list of item dicts
                })
            pending_dos[seen_do[do_num]]["ITEMS"].append(item)

        # Compute TOTAL_TON and flatten split/single for display
        for do in pending_dos:
            do["TOTAL_TON"] = round(sum(it["WEIGHT"] for it in do["ITEMS"]), 3)

        sess["pending_dos"]   = pending_dos
        sess["change_do_page"] = 0   # reset Change DO pagination on new upload

        # ── Build and return summary ──────────────────────────────────────────
        total_items = len(items)
        header = f"✅ *{total_items} item(s) across {len(pending_dos)} DO(s) auto-assigned!*"
        _summ = _build_summary(sess)
        if isinstance(_summ, list):
            return [header + "\n\n" + _summ[0]] + _summ[1:]
        return [header + "\n\n" + _summ]

    except Exception as e:
        import traceback
        return [f"❌ Failed to read the Excel file: {e}\nPlease re-upload."]

def _suggest_current(sess) -> list[str]:
    idx = sess["current_do_index"]
    dos = sess["pending_dos"]

    if idx >= len(dos):
        return _finish_session(sess)

    do = dos[idx]
    engine: LorryEngine = sess["engine"]

    # Combine session unavailable + already assigned today
    excluded = sess["unavailable"] | get_assigned_today()

    suggestions = engine.suggest(
        route=do["ROUTE"],
        total_ton=do["TOTAL_TON"],
        unavailable=excluded,
        top_n=3,
    )
    sess["suggestions"] = suggestions
    sess["state"] = "REVIEWING"

    header = (
        f"📦 DO {idx + 1}/{len(dos)}\n"
        f"  *DO#* {do['DO NUMBER']}\n"
        f"  *Customer:* {do['CUSTOMER NAME']}\n"
        f"  *Route:* {do['ROUTE']}\n"
        f"  *Total weight:* {round(do['TOTAL_TON'], 3)} T\n"
    )

    if not suggestions:
        return [
            header +
            "\n⚠️ *No eligible lorry found* (all may be assigned today or over capacity).\n"
            "Reply *skip* to skip this DO or *custom [PLATE]* to assign manually."
        ]

    lines = [header + "\n🚛 *Suggested lorries:*"]
    for i, s in enumerate(suggestions, 1):
        lines.append(
            f"  *{i}.* {s['LORRY']} ({s['TON_CAPACITY']}T, {s['USER']})\n"
            f"     _{s['REASON']}_"
        )
    lines.append(
        "\nReply:\n"
        "  • *1 / 2 / 3* — to assign that lorry\n"
        "  • *block [PLATE]* — lorry unavailable all day\n  • *broken [PLATE] [REPLACEMENT]* — log breakdown & replacement\n"
        "  • *custom [PLATE]* — to assign any plate manually\n"
        "  • *skip* — skip this DO"
    )
    return ["\n".join(lines)]


def _broken_confirmed_reply(broken: str, replacement: str, sess: dict) -> list:
    """
    After a broken+replacement pair is confirmed:
    1. Log the breakdown message.
    2. Find every item currently assigned to the broken lorry.
    3. Re-assign each one to the replacement (or auto-pick if replacement="NONE").
    4. Return the breakdown summary + updated full assignment summary.
    """
    engine = sess.get("engine")
    cap_info = ""
    if engine is not None and replacement != "NONE":
        row = engine.eligible_lorries[engine.eligible_lorries["LORRY"] == replacement]
        if not row.empty:
            cap  = float(row.iloc[0]["TON"])
            user = str(row.iloc[0]["USER"])
            cap_info = f" ({cap}T, {user})"

    # ── Find and re-assign items that used the broken lorry ──────────────────
    reassigned = []   # list of (item, new_lorry)
    pending    = sess.get("pending_dos", [])

    for do in pending:
        for item in do.get("ITEMS", []):
            item_lorry = item.get("LORRY", "")

            # Check single-lorry assignment
            if item_lorry == broken:
                new_lorry = _pick_replacement(
                    broken, replacement, item, sess, engine
                )
                item["LORRY"]         = new_lorry
                item["SPLIT_LORRIES"] = None
                if new_lorry not in ("NO_LORRY", "SPLIT", None):
                    sess["unavailable"].add(new_lorry)
                reassigned.append((item, do, new_lorry))

            # Check split bins
            elif item_lorry == "SPLIT" and item.get("SPLIT_LORRIES"):
                for bin_ in (item.get("SPLIT_LORRIES") or []):
                    if bin_["lorry"] == broken:
                        new_lorry = _pick_replacement(
                            broken, replacement, item, sess, engine,
                            portion=sum(r["W"] for r in bin_["rows"])
                        )
                        bin_["lorry"] = new_lorry
                        if new_lorry not in ("NO_LORRY", None):
                            sess["unavailable"].add(new_lorry)
                        reassigned.append((item, do, new_lorry))

    # ── Build breakdown message ───────────────────────────────────────────────
    rep_display = f"*{replacement}*{cap_info}" if replacement != "NONE" else "none"
    lines = [
        f"🔧 *Breakdown logged:*",
        f"  ❌ Broken:      *{broken}*",
        f"  ✅ Replacement: {rep_display}",
        f"",
        f"*{broken}* is blocked for today.",
    ]

    if reassigned:
        lines.append(f"")
        lines.append(f"♻️ *{len(reassigned)} item(s) re-assigned:*")
        for item, do, new_lorry in reassigned:
            itmref  = item.get("ITMREF", "") or ""
            itmref_str = f" ({itmref})" if itmref and itmref.lower() not in ("nan","") else ""
            cust    = do.get("CUSTOMER NAME", "")[:22]
            w       = round(item.get("WEIGHT", 0), 2)
            lorry_s = f"*{new_lorry}*" if new_lorry not in ("NO_LORRY", None) else "❌ no lorry"
            lines.append(f"  {cust}{itmref_str} {w}T → {lorry_s}")
    else:
        lines.append("")
        lines.append("No items were assigned to this lorry.")

    msgs = ["\n".join(lines)]

    # ── Append updated full summary if we're in CONFIRMING state ─────────────
    if sess.get("state") in ("CONFIRMING", "REVIEWING") and pending:
        _sc = _build_summary(sess)
        if isinstance(_sc, list):
            msgs.extend(_sc)
        else:
            msgs.append(_sc)

    # ── Re-export if items were re-assigned and we already have a raw_df ──────
    # If the user had already confirmed (or we are in CONFIRMING), the previous
    # export is now stale. Re-run the export silently and store new bytes so
    # app.py will send the updated file automatically.
    if reassigned and sess.get("raw_df") is not None:
        try:
            _export_result(sess)   # updates sess["export_bytes"] in place
            msgs.append("📎 Updated assignment file is being sent.")
        except Exception as _e:
            msgs.append(f"⚠️ Could not regenerate export: {_e}")

    return msgs


def _pick_replacement(broken: str, replacement: str, item: dict,
                      sess: dict, engine, portion: float = None) -> str:
    """
    Pick the best lorry to replace broken for this item.
    - If replacement is specified and available → use it directly.
    - Otherwise auto-pick tightest fit from engine.
    - Falls back to NO_LORRY if nothing available.
    """
    weight = portion if portion is not None else item.get("WEIGHT", 0)

    # Try the nominated replacement first
    if replacement and replacement != "NONE":
        excluded = sess.get("unavailable", set()) | get_assigned_today()
        if replacement not in excluded:
            return replacement
        # replacement already taken — fall through to auto-pick

    # Auto-pick: tightest single lorry excluding already-used
    excluded = (sess.get("unavailable", set()) | get_assigned_today()) - {broken}
    if engine is not None:
        suggestions = engine.suggest(
            route=item.get("ROUTE", ""),
            total_ton=weight,
            unavailable=excluded,
            top_n=1,
        )
        if suggestions:
            return suggestions[0]["LORRY"]
        # Nothing fits by weight — use largest available as last resort
        last_resort = engine.suggest_largest_available(
            item.get("ROUTE", ""), excluded)
        if last_resort:
            return last_resort[0]["LORRY"]

    return "NO_LORRY"


def _handle_broken_list(sess: dict) -> list:
    """Show all broken lorries and their replacements for today."""
    broken_map = get_broken_lorries()
    if not broken_map:
        return [
            "✅ No broken lorries recorded today.",
            {"_type": "buttons",
             "body": "What would you like to do next?",
             "buttons": [{"id": "hi", "title": "👋 Hi"}]},
        ]
    engine = sess.get("engine")
    lines = ["🔧 *Broken lorries today:*\n"]
    for broken, replacement in sorted(broken_map.items()):
        rep_info = replacement
        if engine is not None and replacement != "NONE":
            row = engine.eligible_lorries[engine.eligible_lorries["LORRY"] == replacement]
            if not row.empty:
                cap = float(row.iloc[0]["TON"])
                rep_info = f"{replacement} ({cap}T)"
        status = f"→ replaced by *{rep_info}*" if replacement != "NONE" else "→ no replacement"
        lines.append(f"  ❌ *{broken}* {status}")
    lines.append("\nType *fixed [PLATE]* if lorry was in broken list.")
    lines.append("Type *release [PLATE1] [PLATE2...]* to unblock any lorry(s) from today's log.")
    return [
        "\n".join(lines),
        {"_type": "buttons",
         "body": "What would you like to do next?",
         "buttons": [{"id": "hi", "title": "👋 Hi"}]},
    ]


def _handle_broken_replacement(phone: str, sess: dict, text: str) -> list:
    """Handle the user's reply when we asked which lorry replaces the broken one."""
    broken = sess.get("pending_broken_plate", "")
    reply  = text.strip().upper()

    if reply in ("NONE", "NO", "-", "NIL"):
        # No replacement — just block the broken lorry
        record_broken_lorry(broken, "NONE")
        sess["unavailable"].add(broken)
        # Restore previous state
        sess["state"] = sess.pop("state_before_broken", "IDLE")
        sess.pop("pending_broken_plate", None)
        prev_state = sess["state"]
        msgs = [
            f"🔧 *{broken}* marked as broken with no replacement.\n"
            f"It is blocked for today."
        ]
        if prev_state == "REVIEWING":
            msgs.append("Continuing with your DO assignments...")
            msgs += _suggest_current(sess)
        else:
            msgs.append({"_type": "buttons",
                         "body": "What would you like to do next?",
                         "buttons": [{"id": "hi", "title": "👋 Hi"}]})
        return msgs

    replacement = reply
    if replacement == broken:
        return [f"⚠️ Replacement cannot be the same as the broken lorry (*{broken}*). Try again."]

    # Validate plate exists in master (warn but don't block)
    engine = sess.get("engine")
    plate_known = False
    if engine is not None:
        plate_known = replacement in engine.eligible_lorries["LORRY"].values or \
                      replacement in engine.all_lorries["LORRY"].values

    record_broken_lorry(broken, replacement)
    sess["unavailable"].add(broken)
    sess["state"] = sess.pop("state_before_broken", "IDLE")
    sess.pop("pending_broken_plate", None)
    prev_state = sess["state"]

    msgs = _broken_confirmed_reply(broken, replacement, sess)

    if not plate_known and engine is not None:
        msgs.insert(1, f"⚠️ Note: *{replacement}* is not in the master lorry list. "
                       "Double-check the plate number.")

    # _broken_confirmed_reply already appends the updated summary
    # when state is CONFIRMING/REVIEWING — no need to call _suggest_current

    return msgs


def _handle_reviewing(phone, sess, text):
    cmd = text.strip().lower()
    suggestions = sess["suggestions"]

    if cmd in ("1", "2", "3"):
        pick = int(cmd) - 1
        if pick < len(suggestions):
            chosen = suggestions[pick]["LORRY"]
            return _assign_and_next(sess, chosen)
        return ["Invalid selection. Reply 1, 2, or 3."]

    if cmd.startswith("block "):
        plate = text.split(" ", 1)[1].strip().upper()
        sess["unavailable"].add(plate)
        # Save to daily log so it stays blocked all day across all sessions
        record_assignments_today([plate])
        return [f"🚫 {plate} blocked for the entire day (won't appear again today)."] + _suggest_current(sess)

    if cmd.startswith("custom "):
        plate = text.split(" ", 1)[1].strip().upper()
        # Reject if this plate is already assigned to this DO (shouldn't normally happen,
        # but guard against re-submitting the same suggestion plate)
        do = sess["pending_dos"][sess["current_do_index"]]
        if plate == sess["assigned"].get(do["DO NUMBER"]):
            return [f"⚠️ *{plate}* is already assigned to this DO. Choose a different lorry."]
        return _assign_and_next(sess, plate)

    if cmd == "skip":
        idx = sess["current_do_index"]
        do = sess["pending_dos"][idx]
        sess["assigned"][do["DO NUMBER"]] = "SKIPPED"
        sess["current_do_index"] += 1
        return [f"⏭️ DO {do['DO NUMBER']} skipped."] + _suggest_current(sess)

    return ["Please reply with 1, 2, 3, *block [PLATE]*, *custom [PLATE]*, or *skip*."]


def _assign_and_next(sess, lorry_plate):
    idx = sess["current_do_index"]
    do = sess["pending_dos"][idx]
    sess["assigned"][do["DO NUMBER"]] = lorry_plate
    sess["current_do_index"] += 1
    # Block this lorry from appearing again in the same session
    sess["unavailable"].add(lorry_plate)
    return [f"✅ *{lorry_plate}* assigned to DO {do['DO NUMBER']}."] + _suggest_current(sess)



def _lorry_picker_buttons(sess: dict, do_num: str, page: int = 0) -> list:
    """
    Show up to 2 lorry options as tappable buttons + a Next/Prev navigation button.
    WhatsApp allows max 3 buttons per message, so:
      Button 1: Lorry option A
      Button 2: Lorry option B (if available)
      Button 3: "Next ▶" or "◀ Prev" (navigation)

    Each button tap sends "select_lorry [DO#] [PLATE]" back to the bot.
    Navigation sends "select_do [DO#] [page]".
    """
    engine: LorryEngine = sess.get("engine")
    if not engine:
        return ["❌ No engine loaded. Please restart with hi."]

    # Find the target DO/item
    target = None
    for it in sess.get("items", []):
        if it["DO NUMBER"] == do_num:
            target = it
            break
    if not target:
        return [f"❌ DO# {do_num} not found."]

    # Release current lorry from excluded so it appears as an option
    cur_lorry = target.get("LORRY", "")
    split_plates = set()
    if cur_lorry == "SPLIT" and target.get("SPLIT_LORRIES"):
        for b in target["SPLIT_LORRIES"]:
            split_plates.add(b["lorry"])
    excluded = (sess.get("unavailable", set()) | get_assigned_today()) - {cur_lorry} - split_plates

    weight = target["WEIGHT"]
    route  = target["ROUTE"]
    cust   = target.get("CUSTOMER NAME", "")

    # Get suggestions — fetch enough for pagination
    suggestions = engine.suggest(
        route=route, total_ton=weight,
        unavailable=excluded, top_n=20,
        customer_name=cust,
    )

    # Build option list: auto-pick first, then suggestions
    options = [{"plate": "__AUTO__", "label": "Auto-pick best", "desc": "Bot chooses optimal lorry"}]
    for s in suggestions:
        util = round((weight / s["TON_CAPACITY"]) * 100, 1) if s["TON_CAPACITY"] > 0 else 0
        freq = f"{s['FREQ']}trips" if s["FREQ"] > 0 else "new"
        options.append({
            "plate": s["LORRY"],
            "label": s["LORRY"],
            "desc":  f"{s['TON_CAPACITY']}T {util}% {freq}",
        })

    PER_PAGE = 2  # 2 lorry options + 1 nav button = 3 total
    total_pages = max(1, -(-len(options) // PER_PAGE))  # ceiling div
    page = max(0, min(page, total_pages - 1))

    slice_start = page * PER_PAGE
    slice_end   = slice_start + PER_PAGE
    page_opts   = options[slice_start:slice_end]

    # Build buttons
    buttons = []
    for opt in page_opts:
        label = opt["label"][:18] + ".." if len(opt["label"]) > 20 else opt["label"]
        buttons.append({
            "id":    f"select_lorry {do_num} {opt['plate']}",
            "title": label,
        })

    # Navigation button
    if total_pages > 1:
        if page < total_pages - 1:
            buttons.append({"id": f"select_do {do_num} {page+1}", "title": f"More ({page+1+1}/{total_pages})"})
        else:
            buttons.append({"id": f"select_do {do_num} 0",       "title": "From start"})

    # Truncate to max 3
    buttons = buttons[:3]

    # Body text: show current assignment + lorry details
    cur_label = ", ".join(split_plates) if split_plates else (cur_lorry if cur_lorry and cur_lorry not in ("NO_LORRY","SPLIT","") else "None")
    details = []
    for opt in page_opts:
        if opt["plate"] != "__AUTO__":
            details.append(f"  {opt['label']}: {opt['desc']}")
    detail_str = "\n".join(details)
    page_info  = f"Page {page+1}/{total_pages}" if total_pages > 1 else ""

    body = (
        f"DO: {do_num}\n"
        f"Weight: {round(weight,3)}T  Route: {route[:30]}\n"
        f"Current: {cur_label}\n"
        f"{page_info}\n"
        f"{detail_str}"
    ).strip()

    return [{"_type": "buttons", "body": body[:1024], "buttons": buttons}]


def _build_summary(sess) -> str:
    """Build a clean, mobile-friendly assignment summary."""
    pending  = sess["pending_dos"]   # list of DO groups, each with ITEMS list
    no_lorry = []
    lines    = []

    # Show broken lorry notice at top of summary if any are active
    broken_map = get_broken_lorries()
    if broken_map:
        broken_lines = ["🔧 *Active breakdowns today:*"]
        for bp, rp in sorted(broken_map.items()):
            rep = f"replaced by *{rp}*" if rp != "NONE" else "no replacement"
            broken_lines.append(f"  ❌ {bp} → {rep}")
        lines.append("\n".join(broken_lines))
        lines.append("─" * 20)

    taken_today   = get_assigned_today()
    broken_today  = set(get_broken_lorries().keys())   # lorries marked broken
    # Collect plates assigned in this session (already visible on item rows)
    session_plates = set(
        it.get("LORRY","") for do in sess.get("pending_dos",[]) for it in do.get("ITEMS",[])
        if it.get("LORRY") not in ("SPLIT","NO_LORRY",None,"")
    )
    # Also collect split bin lorries from this session
    for do in sess.get("pending_dos",[]):
        for it in do.get("ITEMS",[]):
            if it.get("LORRY") == "SPLIT":
                for b in (it.get("SPLIT_LORRIES") or []):
                    session_plates.add(b.get("lorry",""))
    # Blocked = plates in today's log that are NOT in this session AND NOT broken
    # (broken lorries already shown under Active breakdowns header above)
    extra_blocked = (taken_today - session_plates - broken_today) - {""}

    lines.append("📋 *ASSIGNMENT SUMMARY*")
    if extra_blocked:
        lines.append("⛔ Blocked: " + ", ".join(sorted(extra_blocked)))
    lines.append("━━━━━━━━━━━━━━━━━━━━━")
    lines.append("")

    engine = sess.get("engine")
    entry_num = 0

    for do in pending:
        do_num   = do["DO NUMBER"]
        customer = do["CUSTOMER NAME"]
        items    = do.get("ITEMS", [])

        # Shorten route: "KV04A - SUNGAI BULOH - U5 - KOTA DAMANSARA - N 4"
        # → "KV04A  SUNGAI BULOH → KOTA DAMANSARA"
        route = do["ROUTE"]
        if "-->" in route:
            route_short = route.split("-->")[-1].strip()
        else:
            parts = [p.strip() for p in route.split(" - ")]
            code  = parts[0] if parts else ""
            areas = [p for p in parts[1:]
                     if len(p) > 3 and (not p.replace(" ", "").isalnum() or len(p) > 5)]
            if len(areas) >= 2:
                route_short = f"{code}  {areas[0]} → {areas[1]}"
            elif areas:
                route_short = f"{code}  {areas[0]}"
            else:
                route_short = route[:45]

        for item in items:
            entry_num += 1
            weight = round(item["WEIGHT"], 3)
            lorry  = item["LORRY"]
            itmref = item.get("ITMREF", "")
            itmref_str = f" ({itmref})" if itmref and itmref != "nan" else ""

            # ── Entry header ──────────────────────────────────────────────
            lines.append(f"*{entry_num}. {customer}*{itmref_str}")
            lines.append(f"📍 {route_short}")
            lines.append(f"🔖 {do_num}")
            lines.append(f"⚖️  {weight}T")

            # ── Lorry assignment ──────────────────────────────────────────
            if lorry == "SPLIT" and item.get("SPLIT_LORRIES"):
                lines.append("🚛 SPLIT LOAD:")
                for bin_ in (item.get("SPLIT_LORRIES") or []):
                    bin_w = round(sum(r["W"] for r in bin_["rows"]), 3)
                    lines.append(f"    • *{bin_['lorry']}* — {bin_w}T")
                no_lorry  # keep list intact

            elif lorry == "NO_LORRY":
                lines.append("🚛 ❌ No lorry available")
                no_lorry.append(do_num)

            else:
                # Single lorry — compute utilization
                cap = None
                if engine is not None:
                    row = engine.eligible_lorries[
                        engine.eligible_lorries["LORRY"] == lorry]
                    if not row.empty:
                        cap = float(row.iloc[0]["TON"])
                if cap and cap > 0:
                    util_pct = round((weight / cap) * 100, 1)
                    if util_pct > 100:
                        util_icon = "🔴"
                        util_str  = f"{util_pct}% ⚠ OVER CAP"
                    elif util_pct < 50:
                        util_icon = "⚠️"
                        util_str  = f"{util_pct}%"
                    elif util_pct < 75:
                        util_icon = "🟡"
                        util_str  = f"{util_pct}%"
                    else:
                        util_icon = "✅"
                        util_str  = f"{util_pct}%"
                    lines.append(f"🚛 *{lorry}*  {util_icon} {util_str}")
                else:
                    lines.append(f"🚛 *{lorry}*")

            # ── Divider between entries ───────────────────────────────────
            lines.append("─────────────────────")

    # ── Footer ────────────────────────────────────────────────────────────
    all_items   = [it for do in pending for it in do.get("ITEMS", [])]
    assigned_ok = sum(1 for it in all_items if it["LORRY"] not in ("NO_LORRY", None))
    unassigned  = sum(1 for it in all_items if it["LORRY"] == "NO_LORRY")

    lines.append(f"✅ {assigned_ok} assigned   ❌ {unassigned} unassigned")
    summary_text = "\n".join(lines)

    result = [summary_text]

    # ── Yes / No confirm buttons ──────────────────────────────────────────
    result.append({
        "_type": "buttons",
        "body":  "Confirm assignments?",
        "buttons": [
            {"id": "yes", "title": "✅ Yes, Export"},
            {"id": "no",  "title": "❌ Cancel"},
        ],
    })

    # ── Change Assignment: DO picker ──────────────────────────────────────
    do_items = []
    for do in sess.get("pending_dos", []):
        do_num     = do.get("DO NUMBER", "")
        first_item = do.get("ITEMS", [{}])[0] if do.get("ITEMS") else {}
        lorry      = first_item.get("LORRY", "")
        if lorry == "SPLIT" and first_item.get("SPLIT_LORRIES"):
            lorry = "+".join(b["lorry"] for b in (first_item.get("SPLIT_LORRIES") or []))
        route = first_item.get("ROUTE", "")[:25]
        if do_num:
            status = "No lorry" if lorry in ("NO_LORRY", "", None) else lorry
            do_items.append({
                "id":          f"select_do {do_num}",
                "title":       do_num[:24],
                "description": f"{status} | {route}"[:72],
            })
    if do_items:
        # Paginate: WhatsApp only supports 9 rows reliably in a single-section list.
        # Show page from session; add "Next Page" row if more items exist.
        PAGE       = 9
        page       = sess.get("change_do_page", 0)
        start      = page * PAGE
        chunk      = do_items[start:start + PAGE]
        has_more   = (start + PAGE) < len(do_items)
        has_prev   = page > 0
        if has_more:
            chunk.append({
                "id":          f"change_do_page {page + 1}",
                "title":       "Next page...",
                "description": f"Showing {start+1}-{start+PAGE} of {len(do_items)}",
            })
        elif has_prev:
            chunk.append({
                "id":          "change_do_page 0",
                "title":       "Back to start",
                "description": f"Page {page+1} of {((len(do_items)-1)//PAGE)+1}",
            })
        result.append({
            "_type":  "do_list",
            "header": "Change Assignment",
            "body":   "Tap a DO# to pick a different lorry:",
            "button": "Change DO",
            "items":  chunk,
        })

    return result

def _handle_confirming(phone, sess, text):
    cmd = text.strip().lower()

    # Pagination for Change DO list
    if cmd.startswith("change_do_page "):
        try:
            page = int(cmd.split()[1])
            sess["change_do_page"] = page
        except (IndexError, ValueError):
            sess["change_do_page"] = 0
        return _build_summary(sess)

    # ── Propagate change to same-route DOs ───────────────────────────────────
    if cmd == "propagate yes":
        ctx = sess.pop("_propagate_ctx", None)
        if ctx:
            plate     = ctx["plate"]
            old_lorry = ctx["old_lorry"]
            dos       = ctx["dos"]
            updated   = []
            for do in sess["pending_dos"]:
                if do["DO NUMBER"] not in dos:
                    continue
                for it in do.get("ITEMS", []):
                    if it.get("LORRY") == old_lorry:
                        it["LORRY"]         = plate
                        it["SPLIT_LORRIES"] = None
                sess["assigned"][do["DO NUMBER"]] = plate
                updated.append(do["DO NUMBER"])
            msg = f"✅ Updated {len(updated)} DO(s) to *{plate}*: {', '.join(updated)}"
        else:
            msg = "✅ No propagation context found."
        return [msg] + _build_summary(sess)

    if cmd == "propagate no":
        sess.pop("_propagate_ctx", None)
        return ["✅ Change applied to selected DO only."] + _build_summary(sess)

    if cmd in ("yes", "confirm", "ok"):
        return _export_result(sess)

    if cmd in ("no", "cancel"):
        reset_session(phone)
        return ["❌ Cancelled. Send *hi* to start again."]

    # change [DO#] [PLATE1] [PLATE2] ... — reassign with 1 plate (single) or 2+ (split)
    # change [DO#]                       — auto-pick next best lorry
    if cmd.startswith("change "):
        parts = text.strip().split()
        if len(parts) < 2:
            return ["Usage: *change [DO#] [PLATE]* — single lorry\n"
                    "       *change [DO#] [PLATE1] [PLATE2]* — split across lorries\n"
                    "       *change [DO#]* — auto-pick next best"]

        do_num = parts[1].upper()

        # Find the target DO and ALL its items
        target_item = None
        target_do   = None
        for do in sess["pending_dos"]:
            if do["DO NUMBER"] == do_num:
                target_do = do
                if do.get("ITEMS"):
                    target_item = do["ITEMS"][0]   # primary ref for weight/route
                break

        if target_item is None:
            return [f"❌ DO# *{do_num}* not found. Check the number and try again."]

        engine: LorryEngine = sess["engine"]
        old_lorry = target_item["LORRY"]   # "SPLIT", "NO_LORRY", or plate string

        # ── Release old lorry(s) from unavailable pool ─────────────────────
        def _release_item_lorries(item):
            if item["LORRY"] == "SPLIT" and item.get("SPLIT_LORRIES"):
                for b in (item.get("SPLIT_LORRIES") or []):
                    sess["unavailable"].discard(b["lorry"])
            elif item["LORRY"] not in (None, "NO_LORRY", "SPLIT"):
                sess["unavailable"].discard(item["LORRY"])

        _release_item_lorries(target_item)

        plates = [p.upper() for p in parts[2:]]  # 0 = auto, 1 = single, 2+ = split

        # ── AUTO mode: no plate given ───────────────────────────────────────
        if not plates:
            # Always exclude current lorry(s) — user wants something DIFFERENT
            old_plates = set()
            if old_lorry == "SPLIT" and target_item.get("SPLIT_LORRIES"):
                old_plates = {b["lorry"] for b in (target_item.get("SPLIT_LORRIES") or [])}
            elif old_lorry not in (None, "NO_LORRY", "SPLIT"):
                old_plates = {old_lorry}

            # Merge broken lorries into unavailable — their replacements stay free
            broken_map = get_broken_lorries()
            sess["unavailable"].update(broken_map.keys())  # block broken lorries
            excluded = sess["unavailable"] | get_assigned_today() | old_plates

            # Step 1: try a different single lorry
            suggestions = engine.suggest(
                route=target_item["ROUTE"],
                total_ton=target_item["WEIGHT"],
                unavailable=excluded,
                top_n=1,
            )
            if suggestions:
                new_lorry = suggestions[0]["LORRY"]
                # Update ALL items in this DO
                for it in target_do.get("ITEMS", []):
                    it["LORRY"]         = new_lorry
                    it["SPLIT_LORRIES"] = None
                sess["assigned"][do_num]     = new_lorry
                sess["unavailable"].add(new_lorry)
                reason = suggestions[0]["REASON"]
                return [
                    f"✅ {do_num} auto-reassigned → *{new_lorry}*\n_{reason}_"
                ] + (_s2 if isinstance(_s2 := _build_summary(sess), list) else [_s2])

            # Step 2: no single lorry fits — greedy split across smaller lorries
            remain = target_item["WEIGHT"]
            bins   = []
            excl   = set(excluded)
            for _ in range(10):
                if remain <= 0:
                    break
                sug = engine.suggest(route=target_item["ROUTE"], total_ton=remain,
                                     unavailable=excl, top_n=1)
                if sug:
                    lorry_s, cap_s = sug[0]["LORRY"], sug[0]["TON_CAPACITY"]
                else:
                    all_sug = engine.suggest(route=target_item["ROUTE"], total_ton=0.01,
                                             unavailable=excl, top_n=20)
                    all_sug.sort(key=lambda x: x["TON_CAPACITY"], reverse=True)
                    if not all_sug:
                        break
                    lorry_s, cap_s = all_sug[0]["LORRY"], all_sug[0]["TON_CAPACITY"]
                portion = round(min(cap_s, remain), 6)
                bins.append({"lorry": lorry_s,
                             "rows": [{"DO": do_num, "W": portion}],
                             "cap": round(cap_s - portion, 4)})
                excl.add(lorry_s)
                remain = round(remain - cap_s, 6)

            if remain <= 0 and bins:
                for b in bins:
                    sess["unavailable"].add(b["lorry"])
                target_item["LORRY"]         = "SPLIT"
                target_item["SPLIT_LORRIES"] = bins
                sess["assigned"][do_num]     = "SPLIT"
                plate_str = " + ".join(b["lorry"] for b in bins)
                return [
                    f"✅ {do_num} → split: *{plate_str}*\n"
                    f"(no single lorry available — split across {len(bins)} lorries)"
                ] + (_s2 if isinstance(_s2 := _build_summary(sess), list) else [_s2])

            # Nothing works at all
            target_item["LORRY"]     = "NO_LORRY"
            sess["assigned"][do_num] = "NO_LORRY"
            return [f"⚠️ No alternative lorry available for DO *{do_num}*. "
                    "All eligible lorries are assigned or blocked."]

        # ── SINGLE plate ────────────────────────────────────────────────────
        if len(plates) == 1:
            plate = plates[0]
            # Check not already used elsewhere in this batch
            for do in sess["pending_dos"]:
                for it in do.get("ITEMS", []):
                    if it is target_item:
                        continue
                    if it["LORRY"] == plate:
                        return [f"❌ *{plate}* is already assigned to "
                                f"DO {do['DO NUMBER']} in this batch. Use a different lorry."]
            blocked_today = get_assigned_today()
            if plate in blocked_today:
                return [f"❌ *{plate}* is already assigned/blocked today. Use a different lorry."]

            old_lorry = target_item.get("LORRY", "")
            sess["_last_change_old_lorry"] = old_lorry

            target_item["LORRY"]         = plate
            target_item["SPLIT_LORRIES"] = None
            # Update ALL items in this DO
            for it in target_do.get("ITEMS", []):
                it["LORRY"]         = plate
                it["SPLIT_LORRIES"] = None
            sess["assigned"][do_num] = plate
            old_lorry = sess.get("_last_change_old_lorry")
            sess["unavailable"].add(plate)

            # ── Check if other DOs share the same route AND old lorry ────────
            # If so, ask user whether to propagate the change to them too
            same_route_dos = []
            changed_route  = target_do.get("ITEMS", [{}])[0].get("ROUTE", "")
            for do in sess["pending_dos"]:
                if do["DO NUMBER"] == do_num:
                    continue
                for it in do.get("ITEMS", []):
                    if (it.get("ROUTE") == changed_route and
                            it.get("LORRY") == old_lorry and
                            old_lorry not in (None, "", "NO_LORRY", "SPLIT")):
                        same_route_dos.append(do["DO NUMBER"])
                        break

            if same_route_dos:
                # Store context for propagation confirmation
                sess["_propagate_ctx"] = {
                    "plate":    plate,
                    "old_lorry": old_lorry,
                    "dos":      same_route_dos,
                }
                do_list_str = ", ".join(same_route_dos[:5])
                return [
                    f"✅ {do_num} → *{plate}*\n\n"
                    f"The following DOs share the same route with *{old_lorry}*:\n"
                    f"{do_list_str}\n\n"
                    "Apply the same change to these DOs too?",
                    {"_type": "buttons",
                     "body": "Propagate change?",
                     "buttons": [
                         {"id": "propagate yes", "title": "Yes, update all"},
                         {"id": "propagate no",  "title": "No, keep as-is"},
                     ]}
                ]

            return [
                f"✅ {do_num} → *{plate}*\n"
                "Reply *yes* to confirm or *change [DO#] [PLATE]* to adjust more."
            ] + _build_summary(sess)

        # ── SPLIT: 2 or more plates ─────────────────────────────────────────
        # Validate plates
        blocked_today = get_assigned_today()
        errors = []
        for plate in plates:
            for do in sess["pending_dos"]:
                for it in do.get("ITEMS", []):
                    if it is target_item:
                        continue
                    if it["LORRY"] == plate:
                        errors.append(f"*{plate}* already assigned to DO {do['DO NUMBER']}")
            if plate in blocked_today:
                errors.append(f"*{plate}* is blocked today")
        if errors:
            # Restore old assignment before returning error
            if old_lorry not in (None, "NO_LORRY", "SPLIT"):
                sess["unavailable"].add(old_lorry)
            elif target_item.get("SPLIT_LORRIES"):
                for b in (target_item.get("SPLIT_LORRIES") or []):
                    sess["unavailable"].add(b["lorry"])
            return ["❌ " + " | ".join(errors)]

        # Build bins: distribute weight across lorries in order given
        remain = target_item["WEIGHT"]
        bins   = []
        for plate in plates:
            # Look up this lorry's capacity from the engine
            row = engine.eligible_lorries[engine.eligible_lorries["LORRY"] == plate]
            cap = float(row.iloc[0]["TON"]) if not row.empty else remain
            portion = round(min(cap, remain), 6)
            bins.append({
                "lorry":  plate,
                "rows":   [{"DO": do_num, "W": portion}],
                "cap":    round(cap - portion, 4),
            })
            sess["unavailable"].add(plate)
            remain = round(remain - portion, 6)
            if remain <= 0:
                break

        if remain > 0:
            # Lorries given can't cover full weight
            for b in bins:
                sess["unavailable"].discard(b["lorry"])
            return [f"⚠️ The lorries given can only carry "
                    f"{round(target_item['WEIGHT'] - remain, 2)}T of "
                    f"{target_item['WEIGHT']}T. Add more lorries or choose larger ones."]

        target_item["LORRY"]         = "SPLIT"
        target_item["SPLIT_LORRIES"] = bins
        # Update ALL items in this DO
        for it in target_do.get("ITEMS", []):
            it["LORRY"]         = "SPLIT"
            it["SPLIT_LORRIES"] = bins
        sess["assigned"][do_num]     = "SPLIT"

        plate_str = " + ".join(b["lorry"] for b in bins)
        return [
            f"✅ {do_num} → *{plate_str}*\n(split load)\n"
            "Reply *yes* to confirm or *change [DO#] [PLATE]* to adjust more."
        ] + (_s2 if isinstance(_s2 := _build_summary(sess), list) else [_s2])

    # block [PLATE] — mark lorry unavailable all day and re-run auto-assign
    if cmd.startswith("block "):
        plate = text.split(" ", 1)[1].strip().upper()
        sess["unavailable"].add(plate)
        record_assignments_today([plate])
        # Re-run auto-assign for any DO currently assigned to this plate
        engine: LorryEngine = sess["engine"]
        changed = []
        for do in sess["pending_dos"]:
            do_num = do["DO NUMBER"]
            for item in do.get("ITEMS", []):
                if item["LORRY"] != plate:
                    continue
                # Merge broken lorries into unavailable — their replacements stay free
                broken_map = get_broken_lorries()
                sess["unavailable"].update(broken_map.keys())
                excluded = sess["unavailable"] | get_assigned_today()
                suggestions = engine.suggest(
                    route=item["ROUTE"],
                    total_ton=item["WEIGHT"],
                    unavailable=excluded,
                    top_n=1,
                )
                if suggestions:
                    new_lorry = suggestions[0]["LORRY"]
                    item["LORRY"] = new_lorry
                    sess["assigned"][do_num] = new_lorry
                    sess["unavailable"].add(new_lorry)
                    changed.append(f"  • {do_num} → *{new_lorry}*")
                else:
                    item["LORRY"] = "NO_LORRY"
                    sess["assigned"][do_num] = "NO_LORRY"
                    changed.append(f"  • {do_num} → *No lorry available*")
        msg = [f"🚫 *{plate}* blocked for today."]
        if changed:
            msg.append("Re-assigned affected DOs:\n" + "\n".join(changed))
        # Show full updated summary so user can review before confirming
        return ["\n".join(msg)] + (_s2 if isinstance(_s2 := _build_summary(sess), list) else [_s2])

    return ["Please reply *yes*, *change [DO#] [PLATE]*, *block [PLATE]*, or *no*."]

def _finish_session(sess) -> list[str]:
    sess["state"] = "CONFIRMING"
    assigned = sess["assigned"]
    lines = ["🎉 *All DOs reviewed!*\n\nAssignment summary:"]
    for do_num, lorry in assigned.items():
        lines.append(f"  • {do_num} → *{lorry}*")
    lines.append("\nReply *yes* to get the filled Excel file, or *no* to redo.")
    return ["\n".join(lines)]


def _export_result(sess) -> list[str]:
    try:
        return _export_result_inner(sess)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return [f"❌ Export failed: {e}\nYour assignments are saved. Try typing *yes* again or send *hi* to restart."]

def _export_result_inner(sess) -> list[str]:
    from datetime import date as _date
    today_str = _date.today().strftime("%d-%m-%Y")  # e.g. 12-05-2026
    is_new_fmt = sess.get("is_new_format", False)

    # ── Work on a copy of the raw uploaded DataFrame ─────────────────────────
    new_df: pd.DataFrame = sess["raw_df"].copy()

    # Ensure LICENSE column exists as object dtype (string-capable)
    if "LICENSE" not in new_df.columns:
        new_df["LICENSE"] = ""
    new_df["LICENSE"] = new_df["LICENSE"].astype(object)

    # Only touch DATE for old format; new format leaves DATE exactly as uploaded
    if not is_new_fmt:
        if "DATE" not in new_df.columns:
            new_df["DATE"] = ""
        new_df["DATE"] = new_df["DATE"].astype(object)

    # Blank out any stale sentinel strings in LICENSE so we only write real plates
    new_df["LICENSE"] = new_df["LICENSE"].astype(str).replace({"nan": "", "None": ""})

    SENTINELS = {"SKIPPED", "NO_LORRY", "SPLIT", "", None}
    confirmed_plates = []
    assigned_row_idxs = set()

    # ── Write LICENSE per original row index (item["ROW_IDX"]) ──────────────
    for do in sess.get("pending_dos", []):
        for item in do.get("ITEMS", []):
            lorry   = item.get("LORRY")
            row_idx = item.get("ROW_IDX")

            if lorry == "SPLIT" and item.get("SPLIT_LORRIES"):
                bins = (item.get("SPLIT_LORRIES") or [])
                if row_idx is not None and bins:
                    all_plates = ", ".join(b["lorry"] for b in bins
                                          if b["lorry"] not in SENTINELS)
                    new_df.loc[row_idx, "LICENSE"] = all_plates
                    if not is_new_fmt:
                        new_df.loc[row_idx, "DATE"] = today_str
                    assigned_row_idxs.add(row_idx)
                    for b in bins:
                        if b["lorry"] not in SENTINELS:
                            confirmed_plates.append(b["lorry"])
                continue

            if lorry in SENTINELS or row_idx is None:
                continue

            new_df.loc[row_idx, "LICENSE"] = lorry
            if not is_new_fmt:
                new_df.loc[row_idx, "DATE"] = today_str
            assigned_row_idxs.add(row_idx)
            if lorry not in SENTINELS:
                confirmed_plates.append(lorry)

    # ── For new format: enforce correct column order ──────────────────────────
    # Column N (index 13, 1-based col 14) must be LICENSE.
    # We rebuild the column order to match the required spec:
    # NO DATE DO-NUMBER CODE ROUTE CUSTOMER-NAME BRANCH GROSS-WEIGHT REMARKS
    # VALIDATED INVOICE-NO INV-DATE SITE LICENSE DRIVER LORRY-ASST1 LORRY-ASST2 DISTANCE
    NEW_FMT_COLS = [
        "NO", "DATE", "DO NUMBER", "CODE", "ROUTE", "CUSTOMER NAME",
        "BRANCH", "GROSS WEIGHT", "REMARKS", "VALIDATED", "INVOICE NO",
        "INV DATE", "SITE", "LICENSE", "DRIVER", "LORRY ASST1", "LORRY ASST2", "DISTANCE",
    ]
    # Drop internal helper columns added during processing (not in original spec)
    _INTERNAL_COLS = {"WEIGHT(T)"}

    if is_new_fmt:
        # Reorder columns: required spec first, then any extras from the upload
        ordered = [c for c in NEW_FMT_COLS if c in new_df.columns]
        extras  = [c for c in new_df.columns if c not in NEW_FMT_COLS and c not in _INTERNAL_COLS]
        new_df  = new_df[ordered + extras]

        # Rows for history append = only the newly assigned rows
        new_rows = new_df.loc[sorted(assigned_row_idxs)].copy()
    else:
        # Old format: only export assigned rows (existing behaviour)
        new_rows = new_df.loc[sorted(assigned_row_idxs)].copy()
        if "DATE" in new_rows.columns:
            new_rows["DATE"] = today_str

    # ── Append assigned rows into master history file ─────────────────────────
    _hist_path = _resolve_history_path()
    try:
        existing_df = pd.read_excel(_hist_path)
        existing_df.columns = [c.strip().upper() for c in existing_df.columns]
        # Old format history: normalise DATE strings
        if not is_new_fmt and "DATE" in existing_df.columns:
            def _fmt_date(v):
                try:
                    ts = pd.to_datetime(v, errors="coerce")
                    return str(v) if pd.isna(ts) else ts.strftime("%d-%m-%Y")
                except Exception:
                    return str(v)
            existing_df["DATE"] = existing_df["DATE"].apply(_fmt_date)
        # Align columns for concat
        hist_rows = new_rows.drop(columns=[c for c in _INTERNAL_COLS if c in new_rows.columns],
                                  errors="ignore")
        for col in hist_rows.columns:
            if col not in existing_df.columns:
                existing_df[col] = ""
        for col in existing_df.columns:
            if col not in hist_rows.columns:
                hist_rows[col] = ""
        hist_rows = hist_rows[existing_df.columns]
        combined  = pd.concat([existing_df, hist_rows], ignore_index=True)
    except Exception:
        combined = new_rows.drop(columns=[c for c in _INTERNAL_COLS if c in new_rows.columns],
                                 errors="ignore")

    if "DATE" in combined.columns:
        combined["DATE"] = combined["DATE"].astype(str)
    _tmp_path = _hist_path + "._tmp.xlsx"
    combined.to_excel(_tmp_path, index=False, engine="openpyxl")
    os.replace(_tmp_path, _hist_path)

    # ── Build export bytes ────────────────────────────────────────────────────
    # New format: send back the FULL uploaded file (all rows) with LICENSE filled
    # Old format: send only the newly assigned rows (existing behaviour)
    if is_new_fmt:
        out_df = new_df.drop(columns=[c for c in _INTERNAL_COLS if c in new_df.columns],
                             errors="ignore").copy()
        # DATE is already formatted as "d/m/yy" string at upload time (see _handle_excel_upload)
    else:
        out_df = new_rows.copy()
        if "DATE" in out_df.columns:
            out_df["DATE"] = out_df["DATE"].astype(str)

    buf = io.BytesIO()
    out_df.to_excel(buf, index=False, engine="openpyxl")

    # For new format: force DATE column cells to Text format so Excel never
    # re-interprets "11/5/26" as a date serial and shows the timestamp again.
    if is_new_fmt and "DATE" in out_df.columns:
        from openpyxl import load_workbook as _load_wb
        buf.seek(0)
        _wb = _load_wb(buf)
        _ws = _wb.active
        _date_col = out_df.columns.get_loc("DATE") + 1  # 1-based
        for _row in _ws.iter_rows(min_row=2, min_col=_date_col, max_col=_date_col):
            for _cell in _row:
                _cell.number_format = "@"  # @ = Text — prevents Excel date re-parsing
        buf = io.BytesIO()
        _wb.save(buf)

    buf.seek(0)
    sess["export_bytes"] = buf.read()

    # Persist confirmed plates to daily log
    record_assignments_today(list(set(confirmed_plates)))

    sess["state"] = "DONE"
    row_count = len(new_rows)
    total_count = len(new_df)
    if is_new_fmt:
        summary = (f"✅ *{row_count}/{total_count} rows* assigned and appended to history.\n"
                   f"📎 Sending you the complete file with LICENSE filled.")
    else:
        summary = (f"✅ *{row_count} rows* appended to the master trip route file.\n"
                   f"📎 Sending you a copy of the newly assigned rows.")
    return [
        summary,
        {
            "_type": "buttons",
            "body": "Tap below to start a new session, or type *hi* anytime.",
            "buttons": [{"id": "hi", "title": "👋 Hi"}],
        }
    ]

def get_export_bytes(phone: str) -> bytes | None:
    """Return export bytes if available (DONE or re-exported after post-yes block), then clear."""
    sess = sessions.get(phone, {})
    if sess.get("state") in ("DONE", "CONFIRMING"):
        data = sess.get("export_bytes")
        if data:
            sess["export_bytes"] = None  # clear after first retrieval
        return data
    return None