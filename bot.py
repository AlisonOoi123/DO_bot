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
# History paths — checked in priority order; .xls preferred as it contains LONGITUD GPS data
HISTORY_PATH_XLS = os.path.join(_DATA_DIR, "ZSDOROUTEWRH.xls")               # new format with LONGITUD column
HISTORY_PATH     = os.path.join(_DATA_DIR, "ZSDOROUTEWRH.xlsx")               # primary (new format, manual assignments)
HISTORY_PATH_ALT = os.path.join(_DATA_DIR, "ZSDOROUTEWRH-bot.xlsx")          # bot-exported (new format)
HISTORY_PATH_OLD = os.path.join(_DATA_DIR, "126-A BI(ES) TRIP ROUTE CODE.xlsx")  # legacy reference
ROUTE_CODES_PATH = os.path.join(_DATA_DIR, "route_codes.xlsx")                  # user→route mapping

def _load_user_route_prefixes(user: str) -> set | None:
    """Return the set of route-code prefixes (e.g. 'KV19A', 'PH09') assigned to
    *user* (case-insensitive).  Returns None if the mapping file doesn't exist,
    meaning no filtering is applied and all routes are processed.
    """
    if not os.path.exists(ROUTE_CODES_PATH):
        return None
    try:
        df = pd.read_excel(ROUTE_CODES_PATH)
        df.columns = [c.strip().upper() for c in df.columns]
        name_col  = next((c for c in df.columns if "NAME" in c), None)
        route_col = next((c for c in df.columns if "ROUTE" in c), None)
        if name_col is None or route_col is None:
            return None
        user_rows = df[df[name_col].str.strip().str.upper() == user.upper()]
        prefixes = set()
        for route in user_rows[route_col].dropna().astype(str):
            m = re.match(r'^([A-Za-z]{2,4}\d{1,2}[A-Za-z]?)', route.strip())
            if m:
                prefixes.add(m.group(1).upper())
        return prefixes if prefixes else None
    except Exception:
        return None

def _extract_route_prefix(route: str) -> str:
    """Extract the leading route code token (e.g. 'KV19A', 'PH09', 'JH09')."""
    m = re.match(r'^([A-Za-z]{2,4}\d{1,2}[A-Za-z]?)', route.strip())
    return m.group(1).upper() if m else ""

def _resolve_history_path() -> str:
    """Return the best available history file.
    Prefers the .xls version (has LONGITUD GPS column) over .xlsx fallbacks.
    """
    for p in [HISTORY_PATH_XLS, HISTORY_PATH, HISTORY_PATH_ALT, HISTORY_PATH_OLD]:
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
    elif state in ("REVIEWING", "CONFIRMING"):
        # Allow lorry-status file upload at any point during an active session
        if file_bytes:
            try:
                _df_up = pd.read_excel(io.BytesIO(file_bytes))
                _df_up.columns = [c.strip().upper() for c in _df_up.columns]
                _status_result = _handle_lorry_status_upload(phone, sess, _df_up)
                if _status_result is not None:
                    # After updating statuses, re-show summary so user sees changes
                    return _status_result + [_build_summary(sess)]
            except Exception:
                pass   # not an Excel — fall through to text handler
        if state == "REVIEWING":
            return _handle_reviewing(phone, sess, text)
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
        plate    = r["LORRY"]
        ton      = r["TON"]
        lorry_user = r["USER"]
        if plate in broken_today:
            rep = broken_today[plate]
            tag = f" 🔴 Broken→{rep}" if rep != "NONE" else " 🔴 Broken"
        elif plate in taken_today:
            tag = " ⛔ Assigned today"
        else:
            tag = " ✅ Available"
        lines.append(f"  • {plate} — {ton}T ({lorry_user}){tag}")

    lorry_text = (
        f"✅ Logged in as *{user}*\n\n"
        f"Your lorries:\n" + "\n".join(lines) + "\n\n"
        "📎 Now please upload your DO Excel file (.xlsx).\n\n"
        "_Tip: you can also upload the master lorry file (with a_ *Status* _column) to block/release lorries in bulk._"
    )
    # Quick-action buttons (max 3 for WhatsApp)
    buttons_msg = {
        "_type": "buttons",
        "body": "Quick actions:",
        "buttons": [
            {"id": "clear daily log",     "title": "🗑️ Clear Today Log"},
            {"id": "show assigned today", "title": "📋 Show Assigned"},
            {"id": "show blocked",        "title": "🚫 Show Blocked"},
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

def _handle_lorry_status_upload(phone, sess, df: "pd.DataFrame") -> list:
    """Handle a lorry-status update file.

    Accepted column names (case-insensitive):
      LORRY / PLATE / LICENSE  — plate number
      STATUS                   — "Available" or "Blocked" (or "block"/"avail")

    Reads each row and blocks or releases the lorry in today's log.
    Returns a reply message list.
    """
    # Normalise column names
    col_map = {c.upper(): c for c in df.columns}

    # Find LORRY column
    lorry_col = next((col_map[k] for k in ("LORRY", "PLATE", "LICENSE") if k in col_map), None)
    status_col = col_map.get("STATUS")

    if not lorry_col or not status_col:
        return None   # not a lorry-status file — caller should fall through

    engine = sess.get("engine")
    all_plates = set(engine.all_lorries["LORRY"].str.upper()) if engine else set()

    blocked_now  = []
    released_now = []
    unknown      = []

    for _, row in df.iterrows():
        plate  = str(row[lorry_col]).strip().upper()
        status = str(row[status_col]).strip().lower()
        if not plate or plate in ("NAN", "NONE", ""):
            continue
        if all_plates and plate not in all_plates:
            unknown.append(plate)
            continue

        if status.startswith("block"):
            record_assignments_today([plate])
            sess.setdefault("unavailable", set()).add(plate)
            blocked_now.append(plate)
        elif status.startswith("avail") or status in ("ok", "free", "available"):
            release_specific_plates([plate])
            sess.setdefault("unavailable", set()).discard(plate)
            released_now.append(plate)

    if not blocked_now and not released_now and not unknown:
        return None   # nothing actionable — fall through to DO-file handler

    lines = ["📋 *Lorry Status Updated from File*\n"]
    if blocked_now:
        lines.append(f"⛔ *Blocked ({len(blocked_now)}):* {', '.join(blocked_now)}")
    if released_now:
        lines.append(f"✅ *Released ({len(released_now)}):* {', '.join(released_now)}")
    if unknown:
        lines.append(f"⚠️ *Not found in master:* {', '.join(unknown)}")
    lines.append("\nLorry availability updated for today.")

    return [
        "\n".join(lines),
        {
            "_type": "buttons",
            "body": "What would you like to do next?",
            "buttons": [
                {"id": "hi",                  "title": "👋 Upload DOs"},
                {"id": "show assigned today",  "title": "📋 Show Blocked"},
                {"id": "manage lorry",         "title": "🔧 Manage Lorry"},
            ],
        }
    ]


def _handle_excel_upload(phone, sess, file_bytes):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        df.columns = [c.strip().upper() for c in df.columns]

        # ── Detect lorry-status file (LORRY + STATUS columns) ───────────────
        # Must be checked BEFORE DO-file detection so a master-lorry upload
        # with Status column doesn't accidentally trigger DO assignment flow.
        _lorry_status_result = _handle_lorry_status_upload(phone, sess, df)
        if _lorry_status_result is not None:
            return _lorry_status_result

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

        # Route-code filtering: only assign rows whose route prefix belongs to
        # the logged-in user.  Rows for other users are kept in items (so they
        # appear in the export) but pre-marked as OTHER_USER so they get a
        # blank LICENSE in the exported file.
        _user_prefixes = _load_user_route_prefixes(sess.get("user_id", ""))

        items = []
        _other_user_count = 0
        for idx, row in raw.iterrows():
            route_str = str(row["ROUTE"]).strip()
            _is_mine  = True
            if _user_prefixes:
                pfx = _extract_route_prefix(route_str)
                if pfx and pfx not in _user_prefixes:
                    _is_mine = False
                    _other_user_count += 1

            items.append({
                "ROW_IDX":       idx,
                "DO NUMBER":     str(row["DO NUMBER"]).strip(),
                "CUSTOMER NAME": str(row["CUSTOMER NAME"]).strip(),
                "ROUTE":         route_str,
                "CODE":          str(row["CODE"]).strip(),
                "WEIGHT":        float(row["WEIGHT(T)"]),
                "ITMREF":        str(row.get("ITMREF_0", "")).strip(),
                "DATE":          str(row.get("DATE", "")).strip(),
                "LORRY":         None if _is_mine else "OTHER_USER",
                "SPLIT_LORRIES": None,
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

        # ── Route grouping with geographic corridor merging ────────────────────
        # Step 1: one item-list per exact route code.
        # Step 2: build "corridor super-groups" — all routes that travel in the
        #   same cluster+corridor direction AND pass through a shared geographic
        #   waypoint (or have no parseable waypoints, in which case corridor
        #   match alone is sufficient).
        # Step 3: if a super-group is too heavy for the largest lorry, bin-pack
        #   it into sub-groups (heaviest routes first).
        # No "heaviest stays alone" rule — every route that goes the same way
        # should share a lorry; capacity is the only hard limit.
        from collections import defaultdict
        from lorry_engine import (
            _extract_route_intelligence,
            _routes_on_same_way,
            _route_centroid,
            _haversine_km,
            _bearing_deg,
            _bearing_diff,
            _DEPOT,
            can_share_cross_cluster,
            MAX_STOPS_PER_LORRY as _MAX_STOPS,
            MIN_UTIL_TO_ASSIGN  as _MIN_UTIL,
        )

        # Step 1 — exact-route buckets (skip other-user rows — they stay blank)
        route_buckets: dict[str, list] = defaultdict(list)
        for it in items:
            if it.get("LORRY") == "OTHER_USER":
                continue
            route_buckets[it["ROUTE"].strip().upper()].append(it)

        # Step 2 — cluster same-way buckets into corridor super-groups
        # Each super-group is a list of route-bucket lists.
        max_lorry_cap = float(engine.eligible_lorries["TON"].max()) \
            if not engine.eligible_lorries.empty else 99.0

        bucket_list = list(route_buckets.values())   # list of [item, …]
        in_group    = [False] * len(bucket_list)
        super_groups: list[list] = []                # each entry = flat item list

        for i, base_bucket in enumerate(bucket_list):
            if in_group[i]:
                continue
            base_route = base_bucket[0]["ROUTE"]
            merged_items = list(base_bucket)
            in_group[i]  = True

            for j in range(i + 1, len(bucket_list)):
                if in_group[j]:
                    continue
                cand_bucket = bucket_list[j]
                cand_route  = cand_bucket[0]["ROUTE"]

                # All routes already absorbed into merged_items
                combined_w = sum(it["WEIGHT"] for it in merged_items) + \
                             sum(it["WEIGHT"] for it in cand_bucket)
                n_distinct = len({it["ROUTE"] for it in merged_items}) + 1

                if (
                    combined_w <= max_lorry_cap
                    and n_distinct <= _MAX_STOPS
                    and _routes_on_same_way(base_route, cand_route)
                ):
                    merged_items += list(cand_bucket)
                    in_group[j]   = True

            super_groups.append(merged_items)

        # Step 3 — if a super-group is heavier than max lorry cap, bin-pack it
        # into capacity-sized sub-groups (heaviest route bucket first).
        sorted_groups: list[list] = []
        for sg in super_groups:
            total_w = sum(it["WEIGHT"] for it in sg)
            if total_w <= max_lorry_cap:
                sorted_groups.append(sg)
                continue

            # Over-capacity super-group — split into sub-groups
            sub_buckets = defaultdict(list)
            for it in sg:
                sub_buckets[it["ROUTE"]].append(it)
            sub_list = sorted(
                sub_buckets.values(),
                key=lambda b: sum(i["WEIGHT"] for i in b),
                reverse=True,
            )

            current_sub: list = []
            current_w = 0.0
            for sub_b in sub_list:
                w = sum(it["WEIGHT"] for it in sub_b)
                if current_sub and current_w + w > max_lorry_cap:
                    sorted_groups.append(current_sub)
                    current_sub = list(sub_b)
                    current_w   = w
                else:
                    current_sub += list(sub_b)
                    current_w   += w
            if current_sub:
                sorted_groups.append(current_sub)

        # ── Step 4: geographic cross-cluster merge (Nominatim/OSM, free) ─────────
        # Same-cluster corridor merging (Step 2) only joins routes within the
        # same region.  Here we try to join groups from DIFFERENT clusters when
        # Nominatim confirms their destinations are within 300 km straight-line.
        #
        # Example: PH01-03 (Pahang/Bentong, 0.275T) + TR02 (Terengganu, 6T)
        # both use the KL→East highway and their centroids are ≈160 km apart.
        # East Malaysia (Sabah/Sarawak) is ≈1 000 km from KL → always rejected.
        def _group_centroid(items):
            """Average lat/lng centroid of all routes in a group."""
            seen, lats, lons = set(), [], []
            for it in items:
                r = it["ROUTE"]
                if r in seen:
                    continue
                seen.add(r)
                c = _route_centroid(r)
                if c:
                    lats.append(c[0]); lons.append(c[1])
            if not lats:
                return None
            return (sum(lats) / len(lats), sum(lons) / len(lons))

        _cross_merged = [False] * len(sorted_groups)
        _new_groups: list[list] = []

        for i, base_sg in enumerate(sorted_groups):
            if _cross_merged[i]:
                continue
            merged      = list(base_sg)
            merged_cent = _group_centroid(merged)   # updated as we absorb groups

            for j in range(i + 1, len(sorted_groups)):
                if _cross_merged[j]:
                    continue
                cand_sg    = sorted_groups[j]
                cand_route = cand_sg[0]["ROUTE"]

                # Skip if all routes in cand share the same cluster as base
                # (same-cluster merging already handled by corridor merge)
                base_clusters = {_extract_route_intelligence(it["ROUTE"])["cluster"]
                                 for it in merged}
                cand_cluster  = _extract_route_intelligence(cand_route)["cluster"]
                if base_clusters == {cand_cluster}:
                    continue

                # KL_VALLEY / KL_CITY routes are local — never bundle with
                # outstation clusters (prevents KV03A coastal + PK Perak merges
                # that look like same direction but use different road corridors)
                _LOCAL_CLUSTERS = {"KL_VALLEY", "KL_CITY"}
                if (base_clusters & _LOCAL_CLUSTERS) or cand_cluster in _LOCAL_CLUSTERS:
                    continue

                combined_w = sum(it["WEIGHT"] for it in merged) + \
                             sum(it["WEIGHT"] for it in cand_sg)
                n_routes   = len({it["ROUTE"] for it in merged}) + \
                             len({it["ROUTE"] for it in cand_sg})

                if combined_w > max_lorry_cap:
                    continue
                if n_routes > _MAX_STOPS:
                    continue

                # Geographic check: candidate centroid vs merged group centroid
                cand_cent = _route_centroid(cand_route)
                if merged_cent is None or cand_cent is None:
                    continue
                dist_km = _haversine_km(merged_cent[0], merged_cent[1],
                                        cand_cent[0],   cand_cent[1])
                if dist_km > 180.0:
                    continue

                # Bearing check — ALL PAIRS: candidate must be directionally
                # compatible with every existing regional route in the group.
                # Using the rolling centroid alone can drift as groups merge,
                # allowing e.g. K.Selangor (304°) + Terengganu (34°) = 90°
                # to slip through after KV03A+PK02 centroid shifts to 316°.
                b_cand = _bearing_deg(_DEPOT[0], _DEPOT[1],
                                      cand_cent[0], cand_cent[1])
                depot_to_cand = _haversine_km(_DEPOT[0], _DEPOT[1],
                                              cand_cent[0], cand_cent[1])
                if depot_to_cand < 50.0:   # local route — skip bearing check
                    continue
                bearing_ok = True
                bearing_checked = False  # must have ≥1 regional route in group to merge
                for ex_route in {it["ROUTE"] for it in merged}:
                    ec = _route_centroid(ex_route)
                    if ec is None:
                        continue
                    d_ex = _haversine_km(_DEPOT[0], _DEPOT[1], ec[0], ec[1])
                    if d_ex < 50.0:        # local route — skip bearing check
                        continue
                    bearing_checked = True
                    b_ex = _bearing_deg(_DEPOT[0], _DEPOT[1], ec[0], ec[1])
                    if _bearing_diff(b_ex, b_cand) > 80.0:
                        bearing_ok = False
                        break
                # Reject if merged group has no regional routes — can't validate direction
                if not bearing_checked or not bearing_ok:
                    continue

                merged += list(cand_sg)
                _cross_merged[j] = True
                # Recompute centroid to keep it accurate for subsequent candidates
                merged_cent = _group_centroid(merged)

            _new_groups.append(merged)

        sorted_groups = _new_groups

        def _parse_date_sortkey(date_str: str) -> str:
            """Convert any date string to ISO 'YYYY-MM-DD' for correct ordering.
            Returns '9999-12-31' on failure so undated groups sort last.
            Handles formats: 'd/m/yy', 'dd/mm/yy', 'YYYY-MM-DD', Excel serials.
            """
            s = (date_str or "").strip()
            if not s or s.lower() in ("nan", "none", ""):
                return "9999-12-31"
            for fmt in ("%d/%m/%y", "%d/%m/%Y", "%Y-%m-%d", "%-d/%-m/%y"):
                try:
                    return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
                except ValueError:
                    pass
            try:
                ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
                if pd.notna(ts):
                    return ts.strftime("%Y-%m-%d")
            except Exception:
                pass
            return "9999-12-31"

        def _group_sort_key(g):
            """Primary: earliest delivery date in group (ascending).
            Secondary: total weight (descending) so heavy groups within the same
            date claim the best-fit lorries before lighter ones."""
            dates = [_parse_date_sortkey(it.get("DATE", "")) for it in g]
            earliest = min(dates) if dates else "9999-12-31"
            return (earliest, -sum(it["WEIGHT"] for it in g))

        # Date-first, then heaviest — ensures urgent DOs get lorries before later ones
        sorted_groups.sort(key=_group_sort_key)

        def _assign_group(group_items):
            """Assign ONE lorry (or split) to cover ALL items in the group.
            All items in the group share the same route (one route = one lorry).
            """
            # Pre-filter: mark items exceeding every available lorry's capacity as
            # NO_LORRY before computing total_w.  Without this a 27T item in a 35T
            # group inflates total_w so a 20T lorry ends up with only 8T of feasible
            # cargo at 44% utilisation instead of the correct 10.5T lorry at 85%.
            _max_cap = (float(engine.eligible_lorries["TON"].max())
                        if not engine.eligible_lorries.empty else 0.0)
            _all_group = list(group_items)
            if _max_cap > 0:
                for it in _all_group:
                    if it.get("LORRY") is None and it["WEIGHT"] > _max_cap:
                        it["LORRY"] = "NO_LORRY"
                group_items = [it for it in _all_group if it.get("LORRY") != "NO_LORRY"]
            else:
                _all_group = list(group_items)
            if not group_items:
                for it in _all_group:
                    sess["assigned"][it["DO NUMBER"]] = it["LORRY"]
                return

            # Rule 6b — per-lorry delivery-stop limit.
            # When a group carries too many individual DOs, split it across two
            # lorries so drivers aren't overloaded and idle ABI lorries get work.
            # MAX_STOPS_PER_LORRY (=8) was designed for route-count merging;
            # here we use a separate threshold for DO count.
            _MAX_DOS_PER_LORRY = 10
            if len(group_items) > _MAX_DOS_PER_LORRY:
                # Balance by weight: sort heaviest first, alternate between two halves
                _sorted = sorted(group_items, key=lambda x: x["WEIGHT"], reverse=True)
                half_a = _sorted[::2]   # indices 0, 2, 4, …
                half_b = _sorted[1::2]  # indices 1, 3, 5, …
                _assign_group(half_a)
                _assign_group(half_b)
                # Propagate back to _all_group items that were pre-filtered NO_LORRY
                for it in _all_group:
                    if it.get("LORRY") == "NO_LORRY":
                        sess["assigned"][it["DO NUMBER"]] = "NO_LORRY"
                return

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

                # Rule 8: tightest-fit lorry would still be <10% loaded → leave blank.
                # Don't waste a large lorry on a tiny DO; let it be manually reviewed.
                if single_util < _MIN_UTIL:
                    for it in group_items:
                        it["LORRY"] = "NO_LORRY"
                else:
                    split_option = None
                    if single_util < 0.60:
                        split_option = engine.suggest_split(
                            route=route,
                            total_ton=total_w,
                            unavailable=excluded,
                            single_util_threshold=0.60,
                        )
                    if split_option is not None:
                        # Each lorry in the split carries a portion of the items.
                        # Build bins with their allotted weight (PORTION).
                        bins = [
                            {"lorry": s["LORRY"],
                             "rows":  [],
                             "remain": s["PORTION"]}   # how much this bin can take
                            for s in split_option
                        ]
                        for s in split_option:
                            sess["unavailable"].add(s["LORRY"])
                        # Assign each item to exactly ONE bin (greedy, heaviest first)
                        item_bin: dict[str, str] = {}
                        for it in sorted(group_items, key=lambda x: x["WEIGHT"], reverse=True):
                            placed = False
                            for bin_ in bins:
                                if bin_["remain"] >= it["WEIGHT"] - 0.001:
                                    bin_["rows"].append({"DO": it["DO NUMBER"], "W": it["WEIGHT"]})
                                    bin_["remain"] -= it["WEIGHT"]
                                    item_bin[it["DO NUMBER"]] = bin_["lorry"]
                                    placed = True
                                    break
                            if not placed:
                                bins[0]["rows"].append({"DO": it["DO NUMBER"], "W": it["WEIGHT"]})
                                item_bin[it["DO NUMBER"]] = bins[0]["lorry"]
                        # Each item gets ONE lorry plate — no more "VEA2818, W3618U" for all
                        for it in group_items:
                            it["LORRY"] = item_bin.get(it["DO NUMBER"], bins[0]["lorry"])
                            it.pop("SPLIT_LORRIES", None)
                    else:
                        lorry = suggestions[0]["LORRY"]
                        sess["unavailable"].add(lorry)
                        for it in group_items:
                            it["LORRY"] = lorry
            else:
                # No single lorry fits — bin-pack across multiple lorries.
                # Build bins using tightest-fit first: ask suggest() for the
                # smallest lorry that can carry the remaining weight.  Fall
                # back to the largest available only when no single lorry
                # can carry the full remainder (partial-load pass).
                remain = total_w
                bins   = []
                for _ in range(10):
                    if remain <= 0:
                        break
                    excl = sess["unavailable"] | get_assigned_today()
                    # Tightest-fit pass: find smallest lorry that handles remain
                    sug = engine.suggest(route=route, total_ton=remain,
                                         unavailable=excl, top_n=20,
                                         customer_name=customer)
                    if not sug:
                        # No lorry can carry full remain — grab largest available
                        # for a partial load, then continue with what's left
                        sug = engine.suggest(route=route, total_ton=0.001,
                                             unavailable=excl, top_n=20)
                        if not sug:
                            break
                        sug.sort(key=lambda x: x["TON_CAPACITY"], reverse=True)
                    lorry   = sug[0]["LORRY"]
                    cap     = sug[0]["TON_CAPACITY"]
                    portion = min(cap, remain)
                    # Use full lorry capacity for the bin, not just the arithmetic
                    # portion. Items overflow BQU3875's last few scraps and must
                    # fit in the next bin — which needs its full capacity available,
                    # not just the remaining-weight arithmetic.
                    bins.append({"lorry": lorry, "rows": [], "remain": cap})
                    sess["unavailable"].add(lorry)
                    remain = round(remain - cap, 6)

                if remain <= 0 and bins:
                    # Distribute items into bins (each item → one bin, heaviest first)
                    item_bin2: dict[str, str] = {}
                    max_bin_cap = max(b["remain"] + sum(
                        x["W"] for x in b["rows"]) for b in bins) if bins else 0
                    for it in sorted(group_items, key=lambda x: x["WEIGHT"], reverse=True):
                        placed = False
                        for bin_ in bins:
                            if bin_["remain"] >= it["WEIGHT"] - 0.001:
                                bin_["rows"].append({"DO": it["DO NUMBER"], "W": it["WEIGHT"]})
                                bin_["remain"] -= it["WEIGHT"]
                                item_bin2[it["DO NUMBER"]] = bin_["lorry"]
                                placed = True
                                break
                        if not placed:
                            # Bins are full — try to grab one more lorry rather
                            # than giving up (greedy fill can leave tiny tail items
                            # stranded even when arithmetic says they should fit).
                            excl_retry = sess["unavailable"] | get_assigned_today()
                            extra_sug  = engine.suggest(
                                route=route, total_ton=it["WEIGHT"],
                                unavailable=excl_retry, top_n=1,
                                customer_name=customer,
                                today_date_str=_today(),
                            )
                            if extra_sug:
                                extra_lorry = extra_sug[0]["LORRY"]
                                extra_cap   = extra_sug[0]["TON_CAPACITY"]
                                new_bin = {"lorry": extra_lorry, "rows": [], "remain": extra_cap}
                                bins.append(new_bin)
                                sess["unavailable"].add(extra_lorry)
                                new_bin["rows"].append({"DO": it["DO NUMBER"], "W": it["WEIGHT"]})
                                new_bin["remain"] -= it["WEIGHT"]
                                item_bin2[it["DO NUMBER"]] = extra_lorry
                            else:
                                item_bin2[it["DO NUMBER"]] = "NO_LORRY"
                    for it in group_items:
                        it["LORRY"] = item_bin2.get(it["DO NUMBER"], "NO_LORRY")
                        it.pop("SPLIT_LORRIES", None)
                else:
                    # Bin-pack failed — all lorries taken or too small.
                    # Last resort: find the tightest-fitting available lorry
                    # that is NOT overloaded (combined weight ≤ capacity).
                    # If even the largest lorry can't handle the weight → NO_LORRY.
                    excl_final = sess["unavailable"] | get_assigned_today()
                    last_resort = engine.suggest_largest_available(
                        route, excl_final, _today(), total_ton=total_w)
                    if last_resort:
                        lr_cap  = last_resort[0]["TON_CAPACITY"]
                        lr_util = total_w / lr_cap if lr_cap > 0 else 0
                        if lr_util < _MIN_UTIL:
                            # Even the smallest available lorry would be <10% loaded
                            for it in group_items:
                                it["LORRY"] = "NO_LORRY"
                        else:
                            lorry = last_resort[0]["LORRY"]
                            sess["unavailable"].add(lorry)
                            for it in group_items:
                                it["LORRY"] = lorry
                    else:
                        # No lorry can carry this weight without overloading
                        for it in group_items:
                            it["LORRY"] = "NO_LORRY"

            for it in _all_group:
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
            if item.get("LORRY") == "OTHER_USER":
                continue          # keep in raw_df for export blank; hide from UI
            do_num = item["DO NUMBER"]
            if do_num not in seen_do:
                seen_do[do_num] = len(pending_dos)
                pending_dos.append({
                    "DO NUMBER":     do_num,
                    "ALL_DO_NUMBERS": [do_num],
                    "ROUTE":         item["ROUTE"],
                    "CODE":          item["CODE"],
                    "CUSTOMER NAME": item["CUSTOMER NAME"],
                    "DATE":          item.get("DATE", ""),
                    "ITEMS":         [],          # list of item dicts
                })
            pending_dos[seen_do[do_num]]["ITEMS"].append(item)

        # Compute TOTAL_TON and flatten split/single for display
        for do in pending_dos:
            do["TOTAL_TON"] = round(sum(it["WEIGHT"] for it in do["ITEMS"]), 3)

        sess["pending_dos"]   = pending_dos
        sess["change_do_page"] = 0   # reset Change DO pagination on new upload

        # ── Build and return summary ──────────────────────────────────────────
        my_items    = [it for it in items if it.get("LORRY") != "OTHER_USER"]
        total_items = len(my_items)
        header = f"✅ *{total_items} item(s) across {len(pending_dos)} DO(s) auto-assigned!*"
        if _other_user_count:
            header += f"\n📌 _{_other_user_count} row(s) from other users' routes left blank — only your route codes were assigned._"
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

    # ── Build lorry-grouped view ──────────────────────────────────────────
    # Collect all items and build per-lorry buckets (exclude other-user rows)
    all_items = [it for do in pending for it in do.get("ITEMS", [])
                 if it.get("LORRY") != "OTHER_USER"]

    # Map DO NUMBER → (customer_short, route_code, date)
    do_meta: dict[str, tuple] = {}
    for do in pending:
        dn  = do["DO NUMBER"]
        cust = do["CUSTOMER NAME"][:10]
        route_code = do["ROUTE"].split(" - ")[0].strip()[:8] if " - " in do["ROUTE"] else do["ROUTE"][:8]
        dt  = do.get("DATE", "")
        if dt and dt.lower() in ("nan", "none", ""):
            dt = ""
        do_meta[dn] = (cust, route_code, dt)

    # Group items by lorry plate
    from collections import defaultdict as _dd
    lorry_items: dict[str, list] = _dd(list)  # plate → [item, ...]
    no_lorry_items: list = []
    for it in all_items:
        if it["LORRY"] in ("NO_LORRY", None):
            no_lorry_items.append(it)
        else:
            lorry_items[it["LORRY"]].append(it)

    # Lorry capacities
    cap_map: dict[str, float] = {}
    if engine is not None:
        for _, r in engine.eligible_lorries.iterrows():
            cap_map[r["LORRY"]] = float(r["TON"])

    # Sort lorries: by earliest date among their items, then by total weight desc
    def _lorry_sort(plate):
        its = lorry_items[plate]
        dates = [_parse_date_sortkey(it.get("DATE", "")) for it in its]
        return (min(dates) if dates else "9999-12-31", -sum(i["WEIGHT"] for i in its))

    # _parse_date_sortkey may not be in scope here (defined inside _handle_excel_upload).
    # Use a local re-implementation for sorting.
    def _dsort(s):
        s = (s or "").strip()
        if not s or s.lower() in ("nan", "none", ""):
            return "9999-12-31"
        for fmt in ("%d/%m/%y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        try:
            ts = pd.to_datetime(s, dayfirst=True, errors="coerce")
            if pd.notna(ts):
                return ts.strftime("%Y-%m-%d")
        except Exception:
            pass
        return "9999-12-31"

    sorted_lorries = sorted(lorry_items.keys(),
        key=lambda p: (
            min((_dsort(it.get("DATE","")) for it in lorry_items[p]), default="9999-12-31"),
            -sum(i["WEIGHT"] for i in lorry_items[p])
        ))

    for plate in sorted_lorries:
        its   = lorry_items[plate]
        total_w = round(sum(i["WEIGHT"] for i in its), 3)
        cap     = cap_map.get(plate)

        if cap and cap > 0:
            util_pct = round(total_w / cap * 100, 1)
            if util_pct > 100:
                util_tag = f"🔴 {util_pct}% OVER"
            elif util_pct >= 75:
                util_tag = f"✅ {util_pct}%"
            elif util_pct >= 50:
                util_tag = f"🟡 {util_pct}%"
            else:
                util_tag = f"⚠️ {util_pct}%"
            cap_str = f"{cap}T"
        else:
            util_tag = ""
            cap_str  = "?"

        lines.append(f"🚛 *{plate}* ({cap_str})  {util_tag}  _{total_w}T_")

        # One line per DO under this lorry: DO# first, then route→dest, customer, weight, date
        for it in sorted(its, key=lambda x: _dsort(x.get("DATE", ""))):
            dn   = it["DO NUMBER"]
            dn_short = dn[-5:] if len(dn) >= 5 else dn
            w    = round(it["WEIGHT"], 3)
            cust, rcode, dt = do_meta.get(dn, (dn, "", ""))
            dt_tag = f" [{dt}]" if dt else ""
            lines.append(f"  {dn_short}  {rcode}  {cust}  {w}T{dt_tag}")

        lines.append("")   # blank line between lorries

    # ── No-lorry items ────────────────────────────────────────────────────
    if no_lorry_items:
        lines.append(f"❌ *NO LORRY ({len(no_lorry_items)} item(s)):*")
        for it in no_lorry_items:
            dn   = it["DO NUMBER"]
            dn_short = dn[-5:] if len(dn) >= 5 else dn
            w    = round(it["WEIGHT"], 3)
            cust, rcode, dt = do_meta.get(dn, (dn, "", ""))
            dt_tag = f" [{dt}]" if dt else ""
            lines.append(f"  {dn_short}  {rcode}  {cust}  {w}T{dt_tag}")
        lines.append("")

    # ── Footer ────────────────────────────────────────────────────────────
    assigned_ok = len(all_items) - len(no_lorry_items)
    unassigned  = len(no_lorry_items)

    lines.append(f"✅ {assigned_ok} assigned  ❌ {unassigned} unassigned  🚛 {len(sorted_lorries)} lorry(s)")
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

    SENTINELS = {"SKIPPED", "NO_LORRY", "SPLIT", "OTHER_USER", "", None}
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

    # Generate trip manifest for drivers (second file sent alongside the export)
    try:
        sess["trip_manifest_bytes"] = _generate_trip_manifest(sess)
    except Exception as _tm_err:
        print(f"⚠️ Trip manifest generation failed: {_tm_err}")
        sess["trip_manifest_bytes"] = None

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

def _generate_trip_manifest(sess) -> bytes:
    """
    Build a driver-friendly trip manifest Excel workbook.
    Stops are sorted geographically using a greedy nearest-neighbour algorithm
    starting from the depot, so the driver follows the most logical road sequence.
    Within a given date, stops are chained by proximity; dates appear in order.
    Columns: # | DATE | DO# | Customer | Route/Area | WT(T) | Dist | Remarks
    """
    import math
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
    from datetime import date as _date, datetime as _dt
    from collections import defaultdict as _dd

    wb = Workbook()
    wb.remove(wb.active)

    generated_str = _date.today().strftime("%d-%m-%Y")
    raw_df = sess.get("raw_df")

    _DEPOT = (3.0340, 101.5563)   # Eng Sheng HQ, Shah Alam

    # ── Geo helpers ───────────────────────────────────────────────────────────
    def _haversine(lat1, lon1, lat2, lon2) -> float:
        R = 6371.0
        dlat = math.radians(lat2 - lat1)
        dlon = math.radians(lon2 - lon1)
        a = (math.sin(dlat / 2) ** 2
             + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2))
             * math.sin(dlon / 2) ** 2)
        return R * 2 * math.asin(math.sqrt(min(1.0, a)))

    def _parse_latlon(row_idx) -> tuple[float, float] | None:
        """Parse 'lat lon' or 'lat, lon' from the LONGITUD column."""
        if raw_df is None or row_idx is None:
            return None
        try:
            v = str(raw_df.loc[row_idx, "LONGITUD"]).strip()
            if not v or v.lower() in ("nan", "none", ""):
                return None
            v = v.replace(",", " ")
            parts = v.split()
            if len(parts) >= 2:
                return (float(parts[0]), float(parts[1]))
        except Exception:
            pass
        return None

    def _nn_sort(pairs: list) -> list:
        """
        Greedy nearest-neighbour sort within a single date group.
        Pairs with no coordinates are appended at the end in route order.
        """
        with_coords    = [(do, it, _parse_latlon(it.get("ROW_IDX"))) for do, it in pairs]
        has_coords     = [(do, it, ll) for do, it, ll in with_coords if ll is not None]
        no_coords      = [(do, it)     for do, it, ll in with_coords if ll is None]

        result: list = []
        unvisited     = list(has_coords)
        cur           = _DEPOT

        while unvisited:
            nearest = min(unvisited, key=lambda x: _haversine(cur[0], cur[1], x[2][0], x[2][1]))
            result.append((nearest[0], nearest[1]))
            cur = nearest[2]
            unvisited.remove(nearest)

        # Append stops with no coordinates sorted by route then customer
        no_coords.sort(key=lambda x: (x[0]["ROUTE"], x[0]["CUSTOMER NAME"]))
        result.extend(no_coords)
        return result

    # ── Date helpers ──────────────────────────────────────────────────────────
    def _fmt_date(s) -> str:
        s = str(s).strip()
        if not s or s.lower() in ("nan", "none", "nat", ""):
            return ""
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"):
            try:
                return _dt.strptime(s, fmt).strftime("%d-%m-%Y")
            except ValueError:
                pass
        try:
            ts = pd.to_datetime(s, format="mixed", dayfirst=True, errors="coerce")
            if pd.notna(ts):
                return ts.strftime("%d-%m-%Y")
        except Exception:
            pass
        return s

    def _date_sortkey(s) -> str:
        """Return YYYY-MM-DD for chronological sort, '9999-12-31' on failure."""
        s = str(s or "").strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"):
            try:
                return _dt.strptime(s, fmt).strftime("%Y-%m-%d")
            except ValueError:
                pass
        try:
            ts = pd.to_datetime(s, format="mixed", dayfirst=True, errors="coerce")
            if pd.notna(ts):
                return ts.strftime("%Y-%m-%d")
        except Exception:
            pass
        return "9999-12-31"

    # ── Gather items per lorry ────────────────────────────────────────────────
    lorry_pairs: dict[str, list] = _dd(list)
    no_lorry_pairs: list         = []

    for do in sess.get("pending_dos", []):
        for it in do.get("ITEMS", []):
            lorry = it.get("LORRY") or "NO_LORRY"
            if lorry in ("NO_LORRY", None, ""):
                no_lorry_pairs.append((do, it))
            elif lorry != "SPLIT":
                lorry_pairs[lorry].append((do, it))

    engine  = sess.get("engine")
    cap_map: dict[str, float] = {}
    if engine is not None:
        for _, r in engine.eligible_lorries.iterrows():
            cap_map[r["LORRY"]] = float(r["TON"])

    sorted_lorries = sorted(
        lorry_pairs.keys(),
        key=lambda p: sum(it["WEIGHT"] for _, it in lorry_pairs[p]),
        reverse=True,
    )

    # ── Shared styles ─────────────────────────────────────────────────────────
    _thin       = Side(style="thin", color="BBBBBB")
    _brd        = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
    _TITLE_FILL = PatternFill("solid", fgColor="1F4E79")
    _TITLE_FONT = Font(color="FFFFFF", bold=True, size=11)
    _HDR_FILL   = PatternFill("solid", fgColor="2E75B6")
    _HDR_FONT   = Font(color="FFFFFF", bold=True, size=9)
    _ALT_FILL   = PatternFill("solid", fgColor="DEEBF7")
    _DATE_FILL  = PatternFill("solid", fgColor="E2EFDA")  # green tint = first row of new date
    _FOOT_FILL  = PatternFill("solid", fgColor="FCE4D6")
    _FOOT_FONT  = Font(bold=True, size=9)
    _NL_FILL    = PatternFill("solid", fgColor="C00000")

    HEADERS    = ["#", "DATE", "DO #", "CUSTOMER", "ROUTE / AREA", "WT (T)", "DIST", "REMARKS / NOTES"]
    COL_WIDTHS = [4,   11,     9,      22,          32,             8,        8,      42]

    def _apply_headers(ws, row):
        for ci, (hdr, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
            c = ws.cell(row, ci, hdr)
            c.font = _HDR_FONT; c.fill = _HDR_FILL; c.border = _brd
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[row].height = 16

    def _route_display(route_str: str) -> str:
        if " - " in route_str:
            parts = route_str.split(" - ")
            return f"{parts[0]}: {' - '.join(parts[1:3])}"[:32]
        return route_str[:32]

    def _raw_val(row_idx, col: str) -> str:
        if raw_df is None or row_idx is None:
            return ""
        try:
            v = str(raw_df.loc[row_idx, col]).strip()
            if v in ("nan", "None", "NaN", ""):
                return ""
            # Ignore GPS-coordinate strings accidentally placed in DISTANCE column
            if col == "DISTANCE" and re.match(r"^-?\d+\.\d+\s+-?\d+\.\d+", v):
                return ""
            return v
        except Exception:
            return ""

    # ── One sheet per lorry ───────────────────────────────────────────────────
    last_col = get_column_letter(len(HEADERS))
    for plate in sorted_lorries:
        pairs    = lorry_pairs[plate]
        cap      = cap_map.get(plate, 0)
        total_w  = round(sum(it["WEIGHT"] for _, it in pairs), 3)
        util_pct = round(total_w / cap * 100, 1) if cap > 0 else 0

        ws = wb.create_sheet(title=plate[:31].replace("/", "-"))
        ws.freeze_panes = "A3"

        # Title
        ws.merge_cells(f"A1:{last_col}1")
        util_icon = "✅" if util_pct >= 75 else ("🟡" if util_pct >= 50 else "⚠️")
        title_txt = (f"TRIP MANIFEST — {plate}   |   {cap}T capacity   "
                     f"|   {total_w}T loaded ({util_pct}%) {util_icon}   "
                     f"|   Generated: {generated_str}")
        t = ws.cell(1, 1, title_txt)
        t.font = _TITLE_FONT; t.fill = _TITLE_FILL
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        _apply_headers(ws, 2)

        # Sort: group by date chronologically, then nearest-neighbour within each date
        date_groups: dict[str, list] = _dd(list)
        for do, it in pairs:
            dk = _date_sortkey(do.get("DATE", ""))
            date_groups[dk].append((do, it))

        sorted_pairs: list = []
        for dk in sorted(date_groups.keys()):
            sorted_pairs.extend(_nn_sort(date_groups[dk]))

        prev_date = None
        for seq, (do, it) in enumerate(sorted_pairs, 1):
            dr      = seq + 2
            row_idx = it.get("ROW_IDX")
            dist_val    = _raw_val(row_idx, "DISTANCE")
            remarks_val = _raw_val(row_idx, "REMARKS")

            dn_short  = do["DO NUMBER"][-5:] if len(do["DO NUMBER"]) >= 5 else do["DO NUMBER"]
            date_disp = _fmt_date(do.get("DATE", ""))

            date_changed = (date_disp != prev_date)
            prev_date    = date_disp
            fill = _DATE_FILL if date_changed else (_ALT_FILL if seq % 2 == 0 else None)

            row_data = [seq, date_disp, dn_short, do["CUSTOMER NAME"][:22],
                        _route_display(do["ROUTE"]), round(it["WEIGHT"], 3),
                        dist_val, remarks_val]
            for ci, val in enumerate(row_data, 1):
                c = ws.cell(dr, ci, val)
                c.border    = _brd
                c.alignment = Alignment(vertical="top", wrap_text=(ci == len(HEADERS)))
                if fill:
                    c.fill = fill
            if remarks_val:
                ws.row_dimensions[dr].height = min(60, max(15, len(remarks_val) // 5 * 8))

        # Footer
        fr = len(sorted_pairs) + 3
        ws.merge_cells(f"A{fr}:E{fr}")
        c = ws.cell(fr, 1, f"TOTAL — {len(sorted_pairs)} stop(s)")
        c.font = _FOOT_FONT; c.fill = _FOOT_FILL; c.border = _brd
        c.alignment = Alignment(horizontal="right")

        c = ws.cell(fr, 6, total_w)
        c.font = _FOOT_FONT; c.fill = _FOOT_FILL; c.border = _brd
        c.alignment = Alignment(horizontal="center")

        ws.merge_cells(f"G{fr}:{last_col}{fr}")
        c = ws.cell(fr, 7, f"{util_icon} {util_pct}% utilisation  ({total_w}T / {cap}T)")
        c.font = _FOOT_FONT; c.fill = _FOOT_FILL; c.border = _brd
        c.alignment = Alignment(horizontal="center")

    # ── NO LORRY sheet ────────────────────────────────────────────────────────
    if no_lorry_pairs:
        ws = wb.create_sheet(title="NO LORRY")
        ws.merge_cells(f"A1:{last_col}1")
        t = ws.cell(1, 1,
            f"UNASSIGNED — {len(no_lorry_pairs)} item(s)  |  Generated: {generated_str}  — Needs manual assignment")
        t.font = Font(color="FFFFFF", bold=True, size=11)
        t.fill = _NL_FILL
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        _apply_headers(ws, 2)
        nl_sorted = _nn_sort(no_lorry_pairs)
        for seq, (do, it) in enumerate(nl_sorted, 1):
            dr = seq + 2
            row_idx     = it.get("ROW_IDX")
            remarks_val = _raw_val(row_idx, "REMARKS")
            dn_short    = do["DO NUMBER"][-5:] if len(do["DO NUMBER"]) >= 5 else do["DO NUMBER"]
            date_disp   = _fmt_date(do.get("DATE", ""))
            for ci, val in enumerate(
                [seq, date_disp, dn_short, do["CUSTOMER NAME"][:22],
                 _route_display(do["ROUTE"]), round(it["WEIGHT"], 3), "",
                 remarks_val or "⚠️ No lorry assigned"],
                1,
            ):
                c = ws.cell(dr, ci, val)
                c.border = _brd
                c.alignment = Alignment(vertical="top", wrap_text=(ci == len(HEADERS)))
        for ci, w in enumerate(COL_WIDTHS, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

    generated_str = _date.today().strftime("%d-%m-%Y")
    raw_df = sess.get("raw_df")

    # ── Normalise a date string to DD-MM-YYYY for display ────────────────────
    def _fmt_date(s) -> str:
        s = str(s).strip()
        if not s or s.lower() in ("nan", "none", "nat", ""):
            return ""
        # Pandas Timestamp str e.g. "2026-05-18 00:00:00"
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y"):
            try:
                return _dt.strptime(s, fmt).strftime("%d-%m-%Y")
            except ValueError:
                pass
        try:
            ts = pd.to_datetime(s, format="mixed", dayfirst=True, errors="coerce")
            if pd.notna(ts):
                return ts.strftime("%d-%m-%Y")
        except Exception:
            pass
        return s

    # ── Sort key: date then route ─────────────────────────────────────────────
    def _sort_key(pair):
        do, _ = pair
        s = str(do.get("DATE", "") or "").strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"):
            try:
                return (_dt.strptime(s, fmt).strftime("%Y-%m-%d"), do["ROUTE"], do["CUSTOMER NAME"])
            except ValueError:
                pass
        try:
            ts = pd.to_datetime(s, format="mixed", dayfirst=True, errors="coerce")
            if pd.notna(ts):
                return (ts.strftime("%Y-%m-%d"), do["ROUTE"], do["CUSTOMER NAME"])
        except Exception:
            pass
        return ("9999-12-31", do["ROUTE"], do["CUSTOMER NAME"])

    # ── Gather items per lorry ────────────────────────────────────────────────
    lorry_pairs: dict[str, list] = _dd(list)
    no_lorry_pairs: list         = []

    for do in sess.get("pending_dos", []):
        for it in do.get("ITEMS", []):
            lorry = it.get("LORRY") or "NO_LORRY"
            if lorry in ("NO_LORRY", None, ""):
                no_lorry_pairs.append((do, it))
            elif lorry != "SPLIT":
                lorry_pairs[lorry].append((do, it))

    engine  = sess.get("engine")
    cap_map: dict[str, float] = {}
    if engine is not None:
        for _, r in engine.eligible_lorries.iterrows():
            cap_map[r["LORRY"]] = float(r["TON"])

    sorted_lorries = sorted(
        lorry_pairs.keys(),
        key=lambda p: sum(it["WEIGHT"] for _, it in lorry_pairs[p]),
        reverse=True,
    )

    # ── Shared styles ─────────────────────────────────────────────────────────
    _thin       = Side(style="thin", color="BBBBBB")
    _brd        = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
    _TITLE_FILL = PatternFill("solid", fgColor="1F4E79")
    _TITLE_FONT = Font(color="FFFFFF", bold=True, size=11)
    _HDR_FILL   = PatternFill("solid", fgColor="2E75B6")
    _HDR_FONT   = Font(color="FFFFFF", bold=True, size=9)
    _ALT_FILL   = PatternFill("solid", fgColor="DEEBF7")
    _DATE_FILL  = PatternFill("solid", fgColor="E2EFDA")   # green tint for date change rows
    _FOOT_FILL  = PatternFill("solid", fgColor="FCE4D6")
    _FOOT_FONT  = Font(bold=True, size=9)
    _NL_FILL    = PatternFill("solid", fgColor="C00000")

    # DATE column added as column 2 (after #)
    HEADERS    = ["#", "DATE", "DO #", "CUSTOMER", "ROUTE / AREA", "WT (T)", "DIST", "REMARKS / NOTES"]
    COL_WIDTHS = [4,   11,     9,      22,          32,             8,        8,      42]

    def _apply_headers(ws, row):
        for ci, (hdr, w) in enumerate(zip(HEADERS, COL_WIDTHS), 1):
            c = ws.cell(row, ci, hdr)
            c.font = _HDR_FONT; c.fill = _HDR_FILL; c.border = _brd
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[row].height = 16

    def _route_display(route_str: str) -> str:
        if " - " in route_str:
            parts = route_str.split(" - ")
            return f"{parts[0]}: {' - '.join(parts[1:3])}"[:32]
        return route_str[:32]

    def _raw_val(row_idx, col: str) -> str:
        if raw_df is None or row_idx is None:
            return ""
        try:
            v = str(raw_df.loc[row_idx, col]).strip()
            if v in ("nan", "None", "NaN", ""):
                return ""
            # Ignore GPS-coordinate strings accidentally placed in DISTANCE column
            if col == "DISTANCE" and re.match(r"^-?\d+\.\d+\s+-?\d+\.\d+", v):
                return ""
            return v
        except Exception:
            return ""

    # ── One sheet per lorry ───────────────────────────────────────────────────
    last_col = get_column_letter(len(HEADERS))
    for plate in sorted_lorries:
        pairs    = lorry_pairs[plate]
        cap      = cap_map.get(plate, 0)
        total_w  = round(sum(it["WEIGHT"] for _, it in pairs), 3)
        util_pct = round(total_w / cap * 100, 1) if cap > 0 else 0

        ws = wb.create_sheet(title=plate[:31].replace("/", "-"))
        ws.freeze_panes = "A3"

        # Title — generated date (not DO date)
        ws.merge_cells(f"A1:{last_col}1")
        util_icon = "✅" if util_pct >= 75 else ("🟡" if util_pct >= 50 else "⚠️")
        title_txt = (f"TRIP MANIFEST — {plate}   |   {cap}T capacity   "
                     f"|   {total_w}T loaded ({util_pct}%) {util_icon}   "
                     f"|   Generated: {generated_str}")
        t = ws.cell(1, 1, title_txt)
        t.font = _TITLE_FONT; t.fill = _TITLE_FILL
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        _apply_headers(ws, 2)

        # Sort: date first so same-day deliveries are grouped, then route, then customer
        sorted_pairs = sorted(pairs, key=_sort_key)

        prev_date = None
        for seq, (do, it) in enumerate(sorted_pairs, 1):
            dr      = seq + 2
            row_idx = it.get("ROW_IDX")
            dist_val    = _raw_val(row_idx, "DISTANCE")
            remarks_val = _raw_val(row_idx, "REMARKS")

            dn_short  = do["DO NUMBER"][-5:] if len(do["DO NUMBER"]) >= 5 else do["DO NUMBER"]
            date_disp = _fmt_date(do.get("DATE", ""))

            # Shade first row of each new date group in green tint so dates are visually separated
            date_changed = (date_disp != prev_date)
            prev_date    = date_disp
            fill = _DATE_FILL if date_changed else (_ALT_FILL if seq % 2 == 0 else None)

            row_data = [
                seq,
                date_disp,
                dn_short,
                do["CUSTOMER NAME"][:22],
                _route_display(do["ROUTE"]),
                round(it["WEIGHT"], 3),
                dist_val,
                remarks_val,
            ]
            for ci, val in enumerate(row_data, 1):
                c = ws.cell(dr, ci, val)
                c.border    = _brd
                c.alignment = Alignment(vertical="top", wrap_text=(ci == len(HEADERS)))
                if fill:
                    c.fill = fill
            if remarks_val:
                ws.row_dimensions[dr].height = min(60, max(15, len(remarks_val) // 5 * 8))

        # Footer
        fr = len(sorted_pairs) + 3
        ws.merge_cells(f"A{fr}:E{fr}")
        c = ws.cell(fr, 1, f"TOTAL — {len(sorted_pairs)} stop(s)")
        c.font = _FOOT_FONT; c.fill = _FOOT_FILL; c.border = _brd
        c.alignment = Alignment(horizontal="right")

        c = ws.cell(fr, 6, total_w)
        c.font = _FOOT_FONT; c.fill = _FOOT_FILL; c.border = _brd
        c.alignment = Alignment(horizontal="center")

        ws.merge_cells(f"G{fr}:{last_col}{fr}")
        c = ws.cell(fr, 7, f"{util_icon} {util_pct}% utilisation  ({total_w}T / {cap}T)")
        c.font = _FOOT_FONT; c.fill = _FOOT_FILL; c.border = _brd
        c.alignment = Alignment(horizontal="center")

    # ── NO LORRY sheet ────────────────────────────────────────────────────────
    if no_lorry_pairs:
        ws = wb.create_sheet(title="NO LORRY")
        ws.merge_cells(f"A1:{last_col}1")
        t = ws.cell(1, 1,
            f"UNASSIGNED — {len(no_lorry_pairs)} item(s)  |  Generated: {generated_str}  — Needs manual assignment")
        t.font = Font(color="FFFFFF", bold=True, size=11)
        t.fill = _NL_FILL
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        _apply_headers(ws, 2)
        for seq, (do, it) in enumerate(sorted(no_lorry_pairs, key=_sort_key), 1):
            dr = seq + 2
            row_idx     = it.get("ROW_IDX")
            remarks_val = _raw_val(row_idx, "REMARKS")
            dn_short    = do["DO NUMBER"][-5:] if len(do["DO NUMBER"]) >= 5 else do["DO NUMBER"]
            date_disp   = _fmt_date(do.get("DATE", ""))
            for ci, val in enumerate(
                [seq, date_disp, dn_short, do["CUSTOMER NAME"][:22],
                 _route_display(do["ROUTE"]), round(it["WEIGHT"], 3), "",
                 remarks_val or "⚠️ No lorry assigned"],
                1,
            ):
                c = ws.cell(dr, ci, val)
                c.border = _brd
                c.alignment = Alignment(vertical="top", wrap_text=(ci == len(HEADERS)))
        for ci, w in enumerate(COL_WIDTHS, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def get_trip_manifest_bytes(phone: str) -> bytes | None:
    """Return trip manifest bytes if available (generated at Yes confirm), then clear."""
    sess = sessions.get(phone, {})
    if sess.get("state") in ("DONE", "CONFIRMING"):
        data = sess.get("trip_manifest_bytes")
        if data:
            sess["trip_manifest_bytes"] = None
        return data
    return None


def get_export_bytes(phone: str) -> bytes | None:
    """Return export bytes if available (DONE or re-exported after post-yes block), then clear."""
    sess = sessions.get(phone, {})
    if sess.get("state") in ("DONE", "CONFIRMING"):
        data = sess.get("export_bytes")
        if data:
            sess["export_bytes"] = None  # clear after first retrieval
        return data
    return None