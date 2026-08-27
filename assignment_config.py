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

# TON already equals the "LORRY NAIK (5%)" column from LORRY DAILY PLANNING.xlsx
# (MUATAN sheet) — i.e. the effective MAX load per lorry already includes the 5%
# naik. So NO further overage is allowed: a lorry may never carry more than its
# LORRY NAIK (5%) value. Both factors are 1.0 (the hard ceiling is TON itself).
NAIK_FACTOR           = 1.0    # no overage beyond TON (TON IS the LORRY NAIK max
                               # from the master file — 10% naik is baked into it)
SAME_ROUTE_NAIK       = 1.0    # same-route loads also capped at the file's TON max
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

# Tiny-outstation relaxation. A VERY small outstation load on a Negeri Sembilan
# (Seremban) route may ride a small lorry instead of tying up a big one — the
# operator allows this ONLY for NS/Seremban, NEVER for far outstation (Kuantan,
# Pahang, Johor, Perak, Terengganu). If the route prefix is in the set below AND
# the group's total load is ≤ the cap, the outstation minimum tonnage is waived.
TINY_OUTSTATION_PREFIXES = {"NS"}
TINY_OUTSTATION_MAX_TON  = 2.0

# ── Urban maximum tonnage ─────────────────────────────────────────────────────
# Urban (KL / Selangor) routes visit many closely-spaced stops in one trip; a
# lorry larger than this cannot physically cover them all in one run. So any DO
# whose destination is urban is capped at this tonnage. Outstation is unaffected.
URBAN_MAX_TON         = 11.0

# ── Geographic constants ──────────────────────────────────────────────────────
# Depot / HQ coordinates (HICOM Shah Alam)
DEPOT_LAT  = 3.0340
DEPOT_LON  = 101.5563

# Geographic clustering: a lorry's stops, sorted by longitude, must form a
# chain where no gap between neighbouring stops exceeds this straight-line
# distance (in degrees: sqrt(dlat² + dlon²)).  0.29° ≈ 32 km.  Applies to ALL
# routes — items too far from their neighbours are moved to another lorry.
MAX_GEO_GAP_DEG               = 0.29

# Urban anti-chaining cap. Single-linkage under MAX_GEO_GAP_DEG lets a chain of
# close stops bridge two far urban zones (e.g. central KL → Rawang, Semenyih →
# KL). For URBAN stops of DIFFERENT route codes, the overall spread (max
# straight-line distance between any two) may not exceed this. Same route code
# stays atomic (any distance), so a legitimate multi-town route is never split.
URBAN_MERGE_SPREAD_DEG        = 0.25

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
# Far outstation (LARGE_LONG: Kuantan / Pahang / Perak / Johor / Terengganu …)
# must use a lorry OVER 11 T. Nearer outstation (MEDIUM_LONG: KV01A north /
# Rawang-Serendah direction) keeps the >5 T minimum so an 8-9 T lorry (e.g.
# BMN3682 for KV01A) can still serve it.
LARGE_LONG_MIN_TON    = 11.001
DEST_MIN_TON: dict[str, float] = {
    "LARGE_LONG":  LARGE_LONG_MIN_TON,   # Pahang / TR / JH / PK / … — must be >11T
    "MEDIUM_LONG": OUTSTATION_MIN_TON,   # NS / Seremban / KV01A direction — >5T
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

# ── REMARKS FIELD — lorry tonnage requirement (FIELD 3) ───────────────────────
# The "REMARKS FIELD" sheet in LORRY DAILY PLANNING.xlsx defines the canonical
# remark phrases planners use.  FIELD 3 specifies which lorry size a DO needs.
# A DO whose REMARKS match one of these phrases MUST be assigned to a lorry whose
# rated tonnage is at or below the cap.  When REMARKS are empty, no cap applies
# and the DO follows normal city + state + nearest-longitude grouping.
#
# phrase (UPPER, substring match) → max lorry tonnage allowed (None = any size)
REMARKS_FIELD3_TON_CAP: dict[str, float | None] = {
    "VAN":           2.0,    # van class only (≤ 2T)
    "BELOW 5 TON":   5.0,
    "BELOW 10 TON":  10.0,
    "BELOW 14 TON":  14.0,
    "BELOW 20 TON":  20.0,
    "ANY SIZE":      None,    # explicit no-cap
}

# Free-text remark phrases (mixed Malay/English) that imply a size cap but do
# not use the canonical FIELD 3 wording.  Mapped to the same tonnage caps.
REMARKS_SIZE_ALIASES: dict[str, float | None] = {
    # "small lorry" / "lorry kecil" in the remarks → small lorry class (≤ 5T).
    # The "LORRY SMALL" word-order variant is included so "only lorry small"
    # is detected; LORI/LORRY spellings are interchangeable (see _phrase_in).
    "LORRY KECIL SAHAJA":        5.0,
    "LORRY KECIL SAJA":          5.0,
    "LORRY KECIL":               5.0,
    "LORI KECIL":                5.0,
    "SMALL LORRY":               5.0,
    "LORRY SMALL":               5.0,   # "only lorry small" word order
    "LORI BESAR TIDAK BOLEH":    5.0,   # big lorry can't enter → small only
    "LORRY BESAR TIDAK BOLEH":   5.0,
    "LORI BESAR TAK BOLEH":      5.0,
    "LORRY BESAR TAK BOLEH":     5.0,
    # "X BOLEH" is texting-shorthand for "tak/tidak boleh" (X = "tak") — same
    # "big lorry can't enter" restriction, just abbreviated. Real remark seen:
    # "LORI BESAR X BOLEH MASUK". Without this, that phrase only matched the
    # bare "LORI BESAR" entry below (cap=None), so the DO got NO size cap at
    # all — even when SHIP_DETAIL explicitly gave a MAX N TON to use instead.
    "LORI BESAR X BOLEH":        5.0,
    "LORRY BESAR X BOLEH":       5.0,
    "BIG LORRY":                 None,
    "LORI BESAR":                None,
    "LORRY BESAR":               None,
}

# ── Route corridor groups ─────────────────────────────────────────────────────
# Routes in the same group always merge before lorry lock — prevents greedy
# single-route assignment from starving adjacent routes in the same corridor.
ROUTE_CORRIDOR_GROUPS: dict[str, list[str]] = {
    "NS":       ["NS04", "NS05", "NS06", "NS07", "NS08"],
    "PH_INT":   ["PH01", "PH02", "PH03", "PH04", "PH05", "PH06", "PH07", "PH08"],
    "KV_NORTH": ["KV01A", "KV02A", "KV04A"],
    "KV_EAST":  ["KV10A", "KV11A", "KV12A"],
    # South Selangor — Kajang (KV19A) and Semenyih/Bangi (KV20A) are adjacent, so
    # they may share one lorry to fill spare capacity (fully utilise the load).
    "KV_SOUTH": ["KV19A", "KV20A"],
    # Perak coastal/interior — one outstation direction. Sabak Bernam (PK01),
    # Teluk Intan (PK02), Manjung (PK06) etc. combine onto one big lorry so the
    # Perak run is fully utilised instead of split across half-empty lorries.
    "PERAK":    ["PK01", "PK02", "PK03", "PK04", "PK05", "PK06", "PK07", "PK08", "PK09"],
}

# ── Strict lorry-route reservations ──────────────────────────────────────────
# No lorry is bound to a specific route direction any more — assignment is
# purely weight/capacity based.  A large lorry serves an outstation route when
# one is needed; otherwise it is free to serve urban routes.  The ONLY hard
# rule is the outstation minimum tonnage (OUTSTATION_MIN_TON): lorries ≤5T can
# never run an outstation route, only KL/Selangor urban routes.
#
# Exception: BQU3875 is reserved for Kuantan (PH09) only, by explicit request
# — it must never serve any other route even if idle/under-utilised that day.
#
# Exception: VJN9910 is restricted to exactly the routes it's listed for in
# LORRY DAILY PLANNING.xlsx's FIT IN LORRY sheet (all outstation NS/PH codes,
# TR02, and KV01A/KV02A/KV24/KV19A/KV20A) — by explicit request, since it was
# turning up on the small-van-only urban cluster (KV04A-KV12A) where it's
# never listed, including overloaded.
LORRY_STRICT_ROUTE: dict[str, set[str]] = {
    "BQU3875": {"PH09"},
    "VJN9910": {
        "NS04", "NS05", "NS06",
        "PH01", "PH02", "PH03", "PH04", "PH05", "PH06", "PH07", "PH09", "PH10", "PH11",
        "TR02",
        "KV01A", "KV02A", "KV24", "KV19A", "KV20A",
    },
}

# ── Preferred lorry per route ─────────────────────────────────────────────────
# Ordered list: first plate is primary, remainder are backups.
# Matched by route-code prefix (startswith) — SPECIFIC route codes are listed
# FIRST so they win over the generic 2-char cluster fallbacks at the bottom.
#
# Derived from real manual-assignment history (data/ZSDOROUTEWRH.xlsx): for each
# route, the lorries the owning planner (ABI or VIVIAN) most frequently used,
# counting only owner-owned + SPARE lorries (cross-owner borrows/swaps ignored).
# When ABI is logged in only ABI+SPARE lorries are eligible, and when VIVIAN is
# logged in only VIVIAN+SPARE — so both owners' preferences can coexist here.
ROUTE_PREFERRED_LORRY: dict[str, list[str]] = {
    # All route-specific preferred lists removed — assignment is purely
    # weight/capacity based (tightest fit wins).
    # Only hard constraint: LORRY_STRICT_ROUTE above (BQU3875/WA6899M → PH only).
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
