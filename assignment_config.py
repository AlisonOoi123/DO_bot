"""
assignment_config.py — Single source of truth for all business-rule constants.

All lorry assignment rules are defined here.  bot.py and lorry_engine.py import
from this module instead of hard-coding values inline.  To change a rule (e.g.
add a new preferred lorry, adjust a tonnage threshold, or define a new corridor
group), edit ONLY this file.  The ASSIGNMENT_RULES.md document explains the
rationale behind each rule in plain language.
"""

from __future__ import annotations

# ── Capacity / utilisation thresholds ────────────────────────────────────────
CAPACITY_TARGET       = 0.80   # Rule 3: target ≥ 80% utilisation
MIN_UTIL_TO_ASSIGN    = 0.10   # Rule 8: don't dispatch if load < 10% capacity (outstation only)
MAX_STOPS_PER_LORRY   = 8      # Rule 6: max delivery stops per lorry per day
MERGE_DIST_THRESHOLD  = 0.25   # Rule 7: reject merge if extra distance > 25%

NAIK_FACTOR           = 1.05   # LORRY NAIK: lorries may carry up to 5% over rated tonnage
SAME_ROUTE_NAIK       = 1.10   # Same-route loads may use up to 10% overage
REBAL_THRESHOLD       = 0.50   # Partial-transfer rebalance: underloaded threshold
FILL_TARGET           = CAPACITY_TARGET   # alias used in fill-to-80% pass
MAX_DOS_PER_LORRY     = 15     # split trigger: groups with > 15 DOs are checked for weight overflow

# ── Lorry size categories (in tonnes) ────────────────────────────────────────
LORRY_LARGE_MIN_TON   = 11.0   # ≥ 11T = LARGE
LORRY_SMALL_MAX_TON   = 5.0    # ≤ 5T  = SMALL (van / small lorry)
LORRY_TINY_EXCL_TON   = 4.5    # routes with tiny items exclude lorries ≥ this tonnage

# ── Outstation minimum tonnage ────────────────────────────────────────────────
# Lorries with TON ≤ 5T are forbidden on any outstation (non-urban) route.
OUTSTATION_MIN_TON    = 5.001

# ── Geographic constants ──────────────────────────────────────────────────────
# Depot / HQ coordinates (HICOM Shah Alam)
DEPOT_LAT  = 3.0340
DEPOT_LON  = 101.5563

CROSS_BEARING_LIMIT           = 90.0   # max bearing diff (°) for same-direction merge
MAX_CITY_MERGE_KM_OUTSTATION  = 60.0   # max GPS distance (km) when merging outstation city clusters
OUTSTATION_DIST_KM            = 50.0   # routes ≥ this distance from depot are "outstation"
LOCAL_ZONE_KM                 = 8.0    # within this radius of depot → "LOCAL" octant
MAX_CROSS_CLUSTER_KM          = 180.0  # max distance between cross-cluster route centroids
MAX_CROSS_CLUSTER_BEARING     = 80.0   # max bearing diff for cross-cluster merge

# ── Tiny-item route guard ─────────────────────────────────────────────────────
# Routes whose average DO weight is at or below this threshold (tonnes) are
# "tiny-item" routes — narrow shophouse streets, cannot accept large lorries.
TINY_ITEM_AVG_WEIGHT_T = 0.15   # 150 kg

# ── Destination group minimum tonnage ────────────────────────────────────────
# Minimum lorry tonnage (TON) required per destination group.
# 0.0 = any size allowed (urban); 5.001 = outstation (small lorries excluded).
DEST_MIN_TON: dict[str, float] = {
    "LARGE_LONG":  OUTSTATION_MIN_TON,   # Pahang / TR / JH / PK / …
    "MEDIUM_LONG": OUTSTATION_MIN_TON,   # NS / Seremban / KV01A direction
    "KL":          0.0,                  # Kuala Lumpur urban — any size
    "SELANGOR":    0.0,                  # Selangor urban — any size
    "KL_SELANGOR": 0.0,                  # generic urban fallback — any size
}

# ── Destination cluster sets ──────────────────────────────────────────────────
# 2-character route-code prefixes that map to each long-haul group.
DEST_LARGE_LONG_CLUSTERS: set[str] = {
    "PH", "TR", "KB", "JH", "PK", "KD", "PN", "MC", "SB", "SR",
}
DEST_MEDIUM_LONG_CLUSTERS: set[str] = {"NS"}

# KV route codes that are classified as outstation (not urban Selangor/KL).
# KV02A (Batu Beruntung/Serendah/Rawang) stays within Selangor → urban.
# KV01A (Rawang/Tanjung Malim toward Perak) is genuinely outstation → MEDIUM_LONG.
DEST_MEDIUM_LONG_KV_CODES: set[str] = {"KV01A"}

# Urban destination groups (KL / Selangor) — accept any lorry size.
DEST_URBAN_GROUPS: set[str] = {"KL", "SELANGOR", "KL_SELANGOR"}

# Assignment priority: outstation groups must claim large lorries FIRST.
DEST_SORT_PRI: dict[str, int] = {
    "LARGE_LONG":  0,
    "MEDIUM_LONG": 1,
    "SELANGOR":    2,
    "KL":          2,
    "KL_SELANGOR": 2,
}

# ── Urban-compatible states ───────────────────────────────────────────────────
# KL and Selangor lorries serve BOTH states freely — the hard state-boundary
# rule does NOT apply within this set.
URBAN_COMPATIBLE_STATES: frozenset[str] = frozenset({
    "KUALA LUMPUR", "W.P. KUALA LUMPUR", "WILAYAH PERSEKUTUAN KUALA LUMPUR",
    "SELANGOR", "KL", "WP KL",
})

# ── State name normalisation ──────────────────────────────────────────────────
# Maps verbose or alternate state names to the canonical short form used in DO items.
STATE_NAME_NORM: dict[str, str] = {
    "KUALA LUMPUR (FEDERAL TERRITORY)": "KUALA LUMPUR",
    "WILAYAH PERSEKUTUAN KUALA LUMPUR": "KUALA LUMPUR",
    "W.P. KUALA LUMPUR":                "KUALA LUMPUR",
    "PUTRAJAYA (FEDERAL TERRITORY)":    "PUTRAJAYA",
    "WILAYAH PERSEKUTUAN PUTRAJAYA":    "PUTRAJAYA",
    "LABUAN (FEDERAL TERRITORY)":       "LABUAN",
    "WILAYAH PERSEKUTUAN LABUAN":       "LABUAN",
    "PULAU PINANG":                     "PENANG",
    "PULAU PINANG (PENANG)":            "PENANG",
    "NEGERI SEMBILAN":                  "NEGERI SEMBILAN",
}

# ── Postcode → Malaysian state lookup ────────────────────────────────────────
# Fallback when the uploaded DO file has no STATE column.
# Each tuple: (lo, hi inclusive, state_name)
POSTCODE_STATE_RANGES: list[tuple[int, int, str]] = [
    (50000, 60999, "KUALA LUMPUR"),
    (40000, 42999, "SELANGOR"),
    (43000, 43999, "SELANGOR"),
    (44000, 44999, "SELANGOR"),
    (45000, 45999, "SELANGOR"),
    (47000, 47999, "SELANGOR"),
    (48000, 48999, "SELANGOR"),
    (63000, 63999, "SELANGOR"),
    (64000, 64999, "SELANGOR"),
    (68000, 68999, "SELANGOR"),
    (70000, 73999, "NEGERI SEMBILAN"),
    (25000, 28999, "PAHANG"),
    (39000, 39999, "PAHANG"),
    (18000, 18999, "TERENGGANU"),
    (20000, 24999, "TERENGGANU"),
    (15000, 17999, "KELANTAN"),
    (80000, 83999, "JOHOR"),
    (84000, 86999, "JOHOR"),
    (30000, 34999, "PERAK"),
    (35000, 36999, "PERAK"),
    (5000,  9999,  "KEDAH"),
    (10000, 14999, "PENANG"),
    (75000, 78999, "MELAKA"),
    (88000, 91300, "SABAH"),
    (93000, 98999, "SARAWAK"),
]

# ── Schedule day-name aliases ─────────────────────────────────────────────────
# Used when parsing the SCHD sheet column headers.
SCHD_DAY_MAP: dict[str, int] = {
    "MON": 0, "MONDAY": 0,
    "TUES": 1, "TUES.": 1, "TUESDAYS": 1, "TUESDAY": 1, "TUE": 1, "TEUS": 1,
    "WED": 2, "WEDNESDAY": 2,
    "THUS": 3, "THURS": 3, "THURSDAY": 3, "THU": 3,
    "FRI": 4, "FRIDAY": 4,
    "SAT": 5, "SATURDAY": 5,
    "SUN": 6, "SUNDAY": 6,
}

# ── Remarks delivery-day parser keyword table ─────────────────────────────────
# Maps BM / English day keywords to weekday integers (Mon=0 … Sun=6).
REMARKS_KEYWORD_DAY: list[tuple[str, int]] = [
    # Malay day names
    ("ISNIN",    0), ("SENIN",   0),
    ("SELASA",   1),
    ("RABU",     2),
    ("KHAMIS",   3),
    ("JUMAAT",   4), ("JUMAT",  4),
    ("SABTU",    5),
    ("AHAD",     6), ("MINGGU", 6),
    # English day names / abbreviations
    ("MONDAY",   0), ("MON",   0),
    ("TUESDAY",  1), ("TUES",  1), ("TUE", 1),
    ("WEDNESDAY",2), ("WED",   2),
    ("THURSDAY", 3), ("THURS", 3), ("THU", 3),
    ("FRIDAY",   4), ("FRI",   4),
    ("SATURDAY", 5), ("SAT",   5),
    ("SUNDAY",   6), ("SUN",   6),
]

# ── Route corridor groups ─────────────────────────────────────────────────────
# Routes in the same group always merge before lorry lock — prevents greedy
# single-route assignment from starving adjacent routes in the same corridor.
ROUTE_CORRIDOR_GROUPS: dict[str, list[str]] = {
    "NS":       ["NS04", "NS05", "NS06", "NS07", "NS08"],
    "PH_INT":   ["PH01", "PH02", "PH03", "PH04", "PH05", "PH06", "PH07", "PH08"],
    "KV_NORTH": ["KV01A", "KV02A", "KV04A"],
    "KV_EAST":  ["KV10A", "KV11A", "KV12A"],
}

# ── Strict lorry-route reservations ──────────────────────────────────────────
# Lorries physically or contractually bound to specific route directions.
# plate → set of route-code prefixes (2–5 chars) it is ALLOWED to serve.
LORRY_STRICT_ROUTE: dict[str, set[str]] = {
    "BQU3875": {"PH"},    # 21T — Pahang routes only
    "WA6899M": {"PH"},    # 13T spare — Pahang routes only
}

# ── Preferred lorry per route ─────────────────────────────────────────────────
# Ordered list: first plate is primary, remainder are backups.
# Matched by route-code prefix (startswith).
ROUTE_PREFERRED_LORRY: dict[str, list[str]] = {
    # North KL corridor — Rawang / Tanjung Malim
    "KV01A": ["BQY7823", "BMN3682"],
    "KV02A": ["BQY7823", "BMN3682"],
    # Inner-KL tight streets (4.2T only)
    "KV06A": ["W3618U",  "W3826C"],
    "KV07A": ["W3618U",  "W3826C"],
    # Central KL
    "KV10A": ["W3826C",  "W3618U"],
    # Pudu / Ampang shophouse (van / small lorry only)
    "KV11A": ["VKN8836", "W3618U",  "W3826C"],
    # Southeast KL corridor — Cheras / Kajang / Semenyih
    "KV18A": ["BPE9788", "VJN9910", "VEA2818"],
    "KV19A": ["BPE9788", "VJN9910", "VEA2818"],
    "KV20A": ["BPE9788", "VJN9910", "VEA2818"],
    # Negeri Sembilan corridor
    "NS04":  ["BQX9983", "BMN3682"],
    "NS05":  ["BQX9983", "BMN3682"],
    "NS06":  ["BQX9983", "BMN3682"],
    # Pahang interior
    "PH01":  ["BPE9788", "WA6899M"],
    "PH02":  ["BPE9788", "WA6899M"],
    "PH03":  ["BPE9788", "WA6899M"],
    "PH04":  ["BPE9788", "WA6899M"],
    "PH05":  ["BPE9788", "WA6899M"],
    "PH06":  ["BPE9788", "WA6899M"],
    "PH07":  ["BPE9788", "WA6899M"],
    "PH08":  ["BPE9788", "WA6899M"],
    # Long-haul outstation
    "PK":    ["VJN9910", "BQY7823", "BQX9983"],
    "JH":    ["VJN9910", "BQX9983", "BQY7823"],
    "TR":    ["VER2872", "VJN9910", "BQY7823"],
    "KB":    ["VJN9910", "BQY7823", "VER2872"],
    "KD":    ["VJN9910", "BQY7823", "BQX9983"],
    "PN":    ["VJN9910", "BQY7823", "BQX9983"],
    "MC":    ["BQX9983", "VJN9910", "BQY7823"],
    "SB":    ["VJN9910", "BQX9983"],
    "SR":    ["VJN9910", "BQX9983"],
}

# ── Route intelligence maps (used by lorry_engine.py) ────────────────────────
CLUSTER_MAP: dict[str, str] = {
    "KV": "KL_VALLEY",  "KL": "KL_CITY",
    "JH": "JOHOR",      "NS": "NEGERI_SEMBILAN",
    "PH": "PAHANG",     "PK": "PERAK",
    "MC": "MELAKA",     "SB": "SABAH",
    "SR": "SARAWAK",    "KD": "KEDAH",
    "PN": "PENANG",     "TR": "TERENGGANU",
    "KB": "KELANTAN",
}

CORRIDOR_MAP: dict[str, str] = {
    "N":  "NORTH",     "S":  "SOUTH",      "E":  "EAST",       "W":  "WEST",
    "SE": "SOUTHEAST", "ES": "SOUTHEAST",
    "NE": "NORTHEAST", "EN": "NORTHEAST",
    "SW": "SOUTHWEST", "WS": "SOUTHWEST",
    "NW": "NORTHWEST", "WN": "WEST_NORTH",
    "C":  "CENTRAL",   "P":  "PORT",
}

# Rule 2: which corridors can share a lorry
ADJACENT_CORRIDORS: dict[str, set[str]] = {
    "NORTH":      {"NORTH", "WEST_NORTH", "NORTHWEST", "CENTRAL"},
    "SOUTH":      {"SOUTH", "SOUTHEAST",  "SOUTHWEST",  "CENTRAL"},
    "EAST":       {"EAST",  "NORTHEAST",  "SOUTHEAST",  "CENTRAL"},
    "WEST":       {"WEST",  "WEST_NORTH", "NORTHWEST",  "SOUTHWEST", "PORT"},
    "SOUTHEAST":  {"SOUTHEAST",  "EAST",  "SOUTH"},
    "NORTHEAST":  {"NORTHEAST",  "EAST",  "NORTH"},
    "SOUTHWEST":  {"SOUTHWEST",  "WEST",  "SOUTH"},
    "NORTHWEST":  {"NORTHWEST",  "WEST",  "NORTH", "WEST_NORTH"},
    "CENTRAL":    {"CENTRAL",    "NORTH", "SOUTH", "EAST", "WEST"},
    "WEST_NORTH": {"WEST_NORTH", "NORTH", "WEST",  "NORTHWEST"},
    "PORT":       {"PORT", "WEST"},
    "GENERAL":    {"GENERAL"},
}
