"""Run this from the SAME folder as bot.py on the deployed machine:
       python diag_filter.py
It tells you exactly why VIVIAN routes are/aren't being filtered."""
import os, bot
print("bot.py loaded from :", os.path.abspath(bot.__file__))
print("PLANNING_PATH      :", bot.PLANNING_PATH)
print("file exists        :", os.path.exists(bot.PLANNING_PATH))
try:
    import pandas as pd
    print("sheets in file     :", pd.ExcelFile(bot.PLANNING_PATH).sheet_names)
except Exception as e:
    print("could not open file:", e)
abi = bot._load_user_route_prefixes("ABI")
print("ABI prefixes loaded:", None if abi is None else f"{len(abi)} -> {sorted(abi)}")
print()
if not abi:
    print(">>> PROBLEM: ABI prefixes are EMPTY/None -> NO filtering -> all routes assigned.")
    print(">>> Fix the data file (needs an 'ABI ROUTE' sheet with route codes).")
else:
    print("Filter is ACTIVE. Test routes:")
    for p in ["PK02","KV18A","KV22A","PH07"]:
        print(f"   {p:7} -> {'KEEP (ABI)' if p in abi else 'BLANK (other user)'}")
