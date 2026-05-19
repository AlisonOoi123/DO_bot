"""
Local CLI simulator for the WhatsApp lorry bot.
Run: python test_bot.py
No Twilio account needed — simulates a full conversation in terminal.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import bot

# The simulator phone — used for your own test messages.
# "clear daily log" wipes ALL sessions regardless of phone, so it works
# correctly even when real WhatsApp users have different numbers.
PHONE = "+60111234567"

def print_bot(msgs):
    for m in msgs:
        if isinstance(m, dict) and m.get("_type") == "buttons":
            print(f"\n🤖 Bot (buttons): {m['body']}")
            for b in m["buttons"]:
                print(f"  [{b['id']}] {b['title']}")
            print(f"  → Type the button ID to simulate a tap")
            print("─" * 50)
        else:
            print(f"\n🤖 Bot:\n{m}\n{'─'*50}")

def main():
    print("=" * 60)
    print(" WhatsApp Lorry Bot — Local Simulator")
    print(" Type your message. To upload a file, type: FILE <path>")
    print(" Type 'sessions' to see all active sessions and unavailable lorries.")
    print(" Type 'quit' to exit.")
    print("=" * 60)

    print_bot(bot.handle_message(PHONE, text="hi"))

    while True:
        user_input = input("👤 You: ").strip()
        if user_input.lower() == "quit":
            break

        if user_input.upper().startswith("FILE "):
            # Strip quotes so both:  FILE path  and  FILE "path"  both work
            path = user_input[5:].strip().strip('"').strip("'")
            if not os.path.exists(path):
                print(f"❌ File not found: {path}")
                print(f"   Check the path and try again.")
            else:
                try:
                    with open(path, "rb") as f:
                        file_bytes = f.read()
                    print(f"✅ Loaded: {os.path.basename(path)} ({len(file_bytes):,} bytes)")
                    # Detect mime type from extension
                    ext = os.path.splitext(path)[1].lower()
                    mime = (
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        if ext in (".xlsx", ".xls") else
                        "image/jpeg" if ext in (".jpg", ".jpeg") else
                        "image/png" if ext == ".png" else
                        "application/octet-stream"
                    )
                    print_bot(bot.handle_message(PHONE, file_bytes=file_bytes, file_mime=mime))
                except Exception as e:
                    print(f"❌ Error reading file: {e}")
        elif user_input.lower() == "sessions":
            # Debug: show all active in-memory sessions and their unavailable sets
            if not bot.sessions:
                print("  (no active sessions)")
            for ph, s in bot.sessions.items():
                unavail = sorted(s.get("unavailable", set()))
                state   = s.get("state", "?")
                user    = s.get("user_id", "?")
                print(f"  {ph} | user={user} | state={state} | unavailable={unavail}")
        else:
            print_bot(bot.handle_message(PHONE, text=user_input))

        export = bot.get_export_bytes(PHONE)
        if export:
            out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "DO_Assigned_NewRows.xlsx")
            with open(out, "wb") as f:
                f.write(export)
            history = os.path.join(os.path.dirname(os.path.abspath(__file__)), "126-A BI(ES) TRIP ROUTE CODE.xlsx")
            print(f"\n📎 New rows saved to: {out}")
            print(f"📂 Master history file updated: {history}")

if __name__ == "__main__":
    main()
