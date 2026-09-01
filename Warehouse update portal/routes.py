"""
Route code reference list for the DRN_0 column on ENGSHENG.BPDLVCUST.

DRN_0 is just a numeric code in the database - it does NOT store the
route name. The route names live only in data/route_codes.csv, which is
exported from the IT department's route master list ("Local menus" in
Sage X3). This module reads that file so the portal can show a human
route name next to each code.

The CSV is re-read on every call (it's tiny - under 100 rows) rather than
cached at import time, so IT can update data/route_codes.csv on the
server and the change shows up immediately without restarting the app.
"""
import csv
import os
from typing import Dict, List, Optional

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROUTES_CSV_PATH = os.path.join(SCRIPT_DIR, "data", "route_codes.csv")


def load_routes() -> Dict[int, str]:
    """Returns {Number: Message} straight from the CSV."""
    routes: Dict[int, str] = {}
    with open(ROUTES_CSV_PATH, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                number = int(row["Number"])
            except (KeyError, TypeError, ValueError):
                continue
            routes[number] = (row.get("Message") or "").strip()
    return routes


def get_route_list() -> List[dict]:
    """Full reference list for the route picker, ordered by Number."""
    return [
        {"number": number, "message": message}
        for number, message in sorted(load_routes().items())
    ]


def get_route_message(number) -> Optional[str]:
    """Looks up the route name for a DRN_0 value; None if not on file."""
    if number is None:
        return None
    try:
        number = int(number)
    except (TypeError, ValueError):
        return None
    return load_routes().get(number)
