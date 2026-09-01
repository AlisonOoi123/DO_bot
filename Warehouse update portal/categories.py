"""
Whitelisted columns per maintenance category.

This is the single source of truth for which DB columns each category is
allowed to touch. The API never accepts a raw column name from the client -
it only accepts a category key + a dict of {label -> checked/unchecked},
and looks up the real column name from here. This is what prevents a
tampered request from writing to an arbitrary column.

Value convention (matches the existing BPDLVCUST data):
    2 = checked / flag applies
    1 = unchecked / default
"""

CATEGORIES = {
    "day": {
        "title": "Delivery Day",
        "locked": True,
        "locked_message": (
            "Delivery Day cannot be changed here. Please contact the IT "
            "department if a customer's delivery days need to be updated."
        ),
        "columns": [
            ("UVYDAY1_0", "Monday"),
            ("UVYDAY2_0", "Tuesday"),
            ("UVYDAY3_0", "Wednesday"),
            ("UVYDAY4_0", "Thursday"),
            ("UVYDAY5_0", "Friday"),
            ("UVYDAY6_0", "Saturday"),
            ("UVYDAY7_0", "Sunday"),
        ],
    },
    "time": {
        "title": "Delivery Time",
        "columns": [
            ("ZUVYDAY15_0", "AM"),
            ("ZUVYDAY16_0", "PM"),
        ],
    },
    "lorry": {
        "title": "Lorry Size",
        "columns": [
            ("ZUVYDAY17_0", "MAX 2 TON"),
            ("ZUVYDAY18_0", "MAX 5 TON"),
            ("ZUVYDAY19_0", "MAX 11 TON"),
            ("ZUVYDAY20_0", "MAX 15 TON"),
            ("ZUVYDAY21_0", "MAX 21 TON"),
        ],
    },
    "outsource": {
        "title": "Out Source",
        "columns": [
            ("ZUVYDAY14_0", "Out Source"),
        ],
    },
    "route": {
        "title": "Route Code",
        "type": "route",
        "column": "DRN_0",
        "note": (
            "Route names are not stored in the database - they only exist in "
            "the route reference list this portal ships with. If the code you "
            "need isn't in the list, check with the IT department before "
            "adding a new one, so the route list stays consistent."
        ),
    },
}

ALL_COLUMNS = [col for cat in CATEGORIES.values() for col, _ in cat.get("columns", [])]
