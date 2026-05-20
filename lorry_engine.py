"""
Lorry Assignment Engine — Enhanced with AI Logistics Rules
===========================================================
Implements requirements from WhatsApp AI Logistics Suggestion System:

  Rule 1 — Same Route Same Day:       same cluster+date → same lorry
  Rule 2 — Nearby Route Merge:        adjacent corridors may share a lorry
  Rule 3 — Capacity Optimisation:     target ≥ 80% utilisation; avoid waste
  Rule 4 — Historical Assignment:     customer+route history (strongest signal)
  Rule 5 — Driver Familiarity:        prefer lorries whose driver knows the region
  Rule 6 — Multi-Drop Limit:          max 8 stops per lorry per day
  Rule 7 — Distance Efficiency:       reject merge if extra distance > 25%
  Route Intelligence:                  cluster + corridor derived from route code
"""

import re
import os
import json
import requests
import pandas as pd
from math import radians, degrees, atan2, cos, sin, asin, sqrt
from typing import Optional


# ── Route intelligence maps ───────────────────────────────────────────────────

_CLUSTER_MAP = {
    "KV": "KL_VALLEY",  "KL": "KL_CITY",
    "JH": "JOHOR",      "NS": "NEGERI_SEMBILAN",
    "PH": "PAHANG",     "PK": "PERAK",
    "MC": "MELAKA",     "SB": "SABAH",
    "SR": "SARAWAK",    "KD": "KEDAH",
    "PN": "PENANG",     "TR": "TERENGGANU",
    "KB": "KELANTAN",
}

_CORRIDOR_MAP = {
    "N": "NORTH",    "S": "SOUTH",     "E": "EAST",      "W": "WEST",
    "SE": "SOUTHEAST", "ES": "SOUTHEAST",
    "NE": "NORTHEAST", "EN": "NORTHEAST",
    "SW": "SOUTHWEST", "WS": "SOUTHWEST",
    "NW": "NORTHWEST", "WN": "WEST_NORTH",
    "C": "CENTRAL",  "P": "PORT",
}

# Rule 2: which corridors can share a lorry
_ADJACENT_CORRIDORS = {
    "NORTH":      {"NORTH", "WEST_NORTH", "NORTHWEST", "CENTRAL"},
    "SOUTH":      {"SOUTH", "SOUTHEAST", "SOUTHWEST", "CENTRAL"},
    "EAST":       {"EAST", "NORTHEAST", "SOUTHEAST", "CENTRAL"},
    "WEST":       {"WEST", "WEST_NORTH", "NORTHWEST", "SOUTHWEST", "PORT"},
    "SOUTHEAST":  {"SOUTHEAST", "EAST", "SOUTH"},
    "NORTHEAST":  {"NORTHEAST", "EAST", "NORTH"},
    "SOUTHWEST":  {"SOUTHWEST", "WEST", "SOUTH"},
    "NORTHWEST":  {"NORTHWEST", "WEST", "NORTH", "WEST_NORTH"},
    "CENTRAL":    {"CENTRAL", "NORTH", "SOUTH", "EAST", "WEST"},
    "WEST_NORTH": {"WEST_NORTH", "NORTH", "WEST", "NORTHWEST"},
    "PORT":       {"PORT", "WEST"},
    "GENERAL":    {"GENERAL"},
}

MAX_STOPS_PER_LORRY   = 8     # Rule 6
MERGE_DIST_THRESHOLD  = 0.25  # Rule 7: reject if extra dist > 25%
CAPACITY_TARGET       = 0.80  # Rule 3: target >= 80% utilisation
MIN_UTIL_TO_ASSIGN    = 0.10  # Rule 8: don't assign a lorry if load < 10% of its capacity

# ── Geographic cross-cluster merging (Nominatim/OSM + Haversine) ─────────────
# Nominatim is the geocoding service behind OpenStreetMap — completely free,
# no API key required.  Usage policy: max 1 request/second; must send a
# descriptive User-Agent.  All results are cached locally in
# data/geocode_cache.json so each unique place name is only queried once.

import time as _time

_GEOCACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "data", "geocode_cache.json")
_geocode_cache: dict = {}
_geocache_dirty = False
_last_nominatim_call = 0.0          # epoch seconds of last network request

_NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
_NOMINATIM_UA  = "DO_bot-logistics/1.0 (whatsapp-lorry-assignment)"

def _load_geocache():
    global _geocode_cache
    if os.path.exists(_GEOCACHE_PATH):
        try:
            with open(_GEOCACHE_PATH) as f:
                _geocode_cache = json.load(f)
        except Exception:
            _geocode_cache = {}

def _save_geocache():
    global _geocache_dirty
    if not _geocache_dirty:
        return
    os.makedirs(os.path.dirname(_GEOCACHE_PATH), exist_ok=True)
    with open(_GEOCACHE_PATH, "w") as f:
        json.dump(_geocode_cache, f, indent=2)
    _geocache_dirty = False

_load_geocache()

# Built-in coordinates for common Malaysian cities/towns (lat, lng).
# Covers the vast majority of route waypoints without any network call.
# Nominatim is used as a fallback only for places not listed here.
_MY_COORDS: dict[str, tuple] = {
    # Klang Valley / KL
    "KUALA LUMPUR": (3.1390, 101.6869), "KL": (3.1390, 101.6869),
    "PETALING JAYA": (3.1073, 101.6067), "PJ": (3.1073, 101.6067),
    "SHAH ALAM": (3.0733, 101.5185), "HICOM": (3.0340, 101.5563),
    "SUBANG": (3.1467, 101.5833),
    "SUBANG JAYA": (3.0596, 101.5858), "KLANG": (3.0449, 101.4455),
    "PORT KLANG": (2.9993, 101.3931), "PELABUHAN KLANG": (2.9993, 101.3931),
    "CHERAS": (3.0877, 101.7482), "AMPANG": (3.1543, 101.7571),
    "KEPONG": (3.2178, 101.6394), "SETAPAK": (3.2029, 101.7061),
    "GOMBAK": (3.2517, 101.7164), "BATU CAVES": (3.2378, 101.6836),
    "SELAYANG": (3.2521, 101.6514), "RAWANG": (3.3232, 101.5745),
    "SUNGAI BULOH": (3.2003, 101.5760), "KOTA DAMANSARA": (3.1652, 101.5870),
    "MONT KIARA": (3.1727, 101.6574), "CHOW KIT": (3.1681, 101.6980),
    "RAJA LAUT": (3.1681, 101.6920), "PUDU": (3.1319, 101.7074),
    "PANDAN": (3.1120, 101.7540), "BALAKONG": (3.0400, 101.7700),
    "SERDANG": (2.9953, 101.7151), "SUNGAI BESI": (3.0700, 101.7200),
    "BATU 9 CHERAS": (3.0033, 101.7930), "KAJANG": (2.9934, 101.7867),
    "SEMENYIH": (2.9320, 101.8426), "BANGI": (2.9592, 101.7817),
    "BERANANG": (2.8751, 101.8519), "MANTIN": (2.8037, 101.9262),
    "PUCHONG": (3.0353, 101.6175), "USJ": (3.0522, 101.5794),
    "SUNWAY": (3.0692, 101.6014), "GLENMARIE": (3.1033, 101.5697),
    "PJ OLD TOWN": (3.1075, 101.6072), "PUTRA PERDANA": (2.9811, 101.6683),
    "BANTING": (2.8128, 101.5027), "JENJAROM": (2.7978, 101.5456),
    "TELOK PANGLIMA GARANG": (2.8667, 101.4833),
    "TELUK PANGLIMA GARANG": (2.8667, 101.4833),
    "PANDAMARAN": (2.9869, 101.3975), "TELUK GONG": (2.9681, 101.3838),
    "BUKIT TINGGI": (3.3478, 101.8143), "KOTA KEMUNING": (3.0317, 101.5419),
    "SALAK TINGGI": (2.7400, 101.7217), "SEPANG": (2.7275, 101.7033),
    "SUNGAI MUDA": (3.0500, 101.4833), "KAPAR": (3.1336, 101.4586),
    "PUNCAK ALAM": (3.2628, 101.5126), "SETIA ALAM": (3.1158, 101.5044),
    "HULU LANGAT": (3.1167, 101.8500), "S.KEMBANGAN": (3.0570, 101.7282),
    "SERDANG PERDANA": (2.9855, 101.7260), "BANDAR JALIL": (3.0333, 101.7617),
    "PUTRA JAYA": (2.9264, 101.6964), "PUTRAJAYA": (2.9264, 101.6964),
    "SEKINCHAN": (3.6846, 101.0339), "TANJUNG KARANG": (3.4167, 101.0583),
    "SUNGAI BESAR": (3.6593, 100.9993), "K.SELANGOR": (3.3356, 101.2525),
    "KUALA SELANGOR": (3.3356, 101.2525),
    # Pahang
    "BENTONG": (3.5151, 101.9175), "KARAK": (3.5323, 101.9922),
    "RAUB": (3.7893, 101.8582), "BENTA": (3.9167, 101.9333),
    "JERANTUT": (3.9333, 102.3667), "LANCHANG": (3.6167, 102.0833),
    "TEMERLOH": (3.4500, 102.4167), "KUANTAN": (3.8319, 103.3322),
    "MENTAKAB": (3.5000, 102.3500), "MARAN": (3.9833, 102.7667),
    # Johor
    "JOHOR BAHRU": (1.4927, 103.7414), "JB": (1.4927, 103.7414),
    "JOHOR BHARU": (1.4927, 103.7414),
    "KULAI": (1.6600, 103.5935), "SENAI": (1.6372, 103.6697),
    "SKUDAI": (1.5264, 103.6700), "KLUANG": (2.0272, 103.3219),
    "BATU PAHAT": (1.8559, 102.9325), "MUAR": (2.0437, 102.5691),
    "PONTIAN": (1.4866, 103.3881), "MERSING": (2.4386, 103.8309),
    "SEGAMAT": (2.5128, 102.8158), "AYER HITAM": (1.9233, 103.1797),
    # Perak
    "IPOH": (4.5975, 101.0901), "TELUK INTAN": (4.0228, 101.0202),
    "LANGKAP": (4.1000, 101.0667), "SLIM RIVER": (3.8167, 101.4000),
    "TAIPING": (4.8500, 100.7333), "KAMPAR": (4.3000, 101.1500),
    "SITIAWAN": (4.2167, 100.7000), "LUMUT": (4.2278, 100.6250),
    # Terengganu
    "TERENGGANU": (5.3117, 103.1324), "KUALA TERENGGANU": (5.3117, 103.1324),
    "KEMAMAN": (4.2333, 103.4167), "DUNGUN": (4.7667, 103.4167),
    "KERTEH": (4.5167, 103.4500), "MARANG": (5.2000, 103.2167),
    # Kelantan
    "KOTA BHARU": (6.1254, 102.2380), "KB": (6.1254, 102.2380),
    "GUAL PERIOK": (6.0333, 102.2500), "PASIR MAS": (6.0450, 102.1375),
    "TANAH MERAH": (5.8076, 102.1464),
    # Negeri Sembilan
    "SEREMBAN": (2.7297, 101.9381), "PORT DICKSON": (2.5210, 101.7981),
    "NILAI": (2.8193, 101.7924), "SENAWANG": (2.7333, 101.9667),
    # Melaka
    "MELAKA": (2.1896, 102.2501), "MALACCA": (2.1896, 102.2501),
    "ALOR GAJAH": (2.3833, 102.2000), "JASIN": (2.3094, 102.4369),
    # Kedah
    "ALOR SETAR": (6.1248, 100.3673), "SUNGAI PETANI": (5.6472, 100.4888),
    "KULIM": (5.3647, 100.5619), "BALING": (5.6725, 100.9231),
    # Penang
    "GEORGE TOWN": (5.4141, 100.3288), "PENANG": (5.4141, 100.3288),
    "BUTTERWORTH": (5.3993, 100.3639), "BUKIT MERTAJAM": (5.3636, 100.4611),
    # Sabah
    "KOTA KINABALU": (5.9804, 116.0735), "KK": (5.9804, 116.0735),
    "SANDAKAN": (5.8402, 118.1179), "TAWAU": (4.2333, 117.8833),
    # Sarawak
    "KUCHING": (1.5533, 110.3592), "MIRI": (4.3995, 113.9914),
    "SIBU": (2.3010, 111.8254), "BINTULU": (3.1727, 113.0447),
    "LIMBANG": (4.7500, 115.0167),
}

def _geocode(place: str) -> tuple | None:
    """Return (lat, lng) for a Malaysian place name.
    Lookup order:
      1. Built-in _MY_COORDS dictionary (instant, no network)
      2. Local cache from previous Nominatim calls
      3. Nominatim (OpenStreetMap, free) — enforces 1 req/sec per OSM policy
    """
    global _geocache_dirty, _last_nominatim_call
    key = place.strip().upper()

    # 1. Built-in dictionary — covers the vast majority of route waypoints
    if key in _MY_COORDS:
        return _MY_COORDS[key]

    # 2. Previously cached Nominatim result
    if key in _geocode_cache:
        cached = _geocode_cache[key]
        return tuple(cached) if cached else None

    # 3. Nominatim fallback for unknown places
    elapsed = _time.time() - _last_nominatim_call
    if elapsed < 1.0:
        _time.sleep(1.0 - elapsed)

    try:
        resp = requests.get(
            _NOMINATIM_URL,
            params={
                "q":               f"{place}, Malaysia",
                "format":          "json",
                "limit":           1,
                "accept-language": "en",
                "countrycodes":    "my",
            },
            headers={"User-Agent": _NOMINATIM_UA},
            timeout=8,
        )
        _last_nominatim_call = _time.time()
        data = resp.json()
        if data:
            coords = [float(data[0]["lat"]), float(data[0]["lon"])]
            _geocode_cache[key] = coords
            _geocache_dirty = True
            _save_geocache()
            return tuple(coords)
    except Exception:
        _last_nominatim_call = _time.time()

    # Cache the miss so we don't hammer the server for unknown places
    _geocode_cache[key] = None
    _geocache_dirty = True
    _save_geocache()
    return None

# Fallback city for each cluster when no waypoints can be extracted
_CLUSTER_CITY = {
    "KL_VALLEY":       "Petaling Jaya",
    "KL_CITY":         "Kuala Lumpur",
    "JOHOR":           "Johor Bahru",
    "PAHANG":          "Kuantan",
    "PERAK":           "Ipoh",
    "MELAKA":          "Melaka",
    "SABAH":           "Kota Kinabalu",
    "SARAWAK":         "Kuching",
    "KEDAH":           "Alor Setar",
    "PENANG":          "George Town Penang",
    "TERENGGANU":      "Kuala Terengganu",
    "KELANTAN":        "Kota Bharu",
    "NEGERI_SEMBILAN": "Seremban",
}

def _haversine_km(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 6371.0
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat/2)**2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon/2)**2
    return 2 * R * asin(sqrt(a))

def _bearing_deg(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    """Compass bearing (0–360°, clockwise from North) from point 1 to point 2."""
    dlon = radians(lon2 - lon1)
    lat1r, lat2r = radians(lat1), radians(lat2)
    x = sin(dlon) * cos(lat2r)
    y = cos(lat1r) * sin(lat2r) - sin(lat1r) * cos(lat2r) * cos(dlon)
    return (degrees(atan2(x, y)) + 360) % 360

def _bearing_diff(b1: float, b2: float) -> float:
    """Smallest angular difference between two compass bearings (0–180°)."""
    diff = abs(b1 - b2) % 360
    return diff if diff <= 180 else 360 - diff

# Depot: Eng Sheng HQ — No 11 Persiaran Sabak Bernam, Section 26 (HICOM), 40400 Shah Alam
_DEPOT = (3.0340, 101.5563)

# Per-route GPS centroids derived from the LONGITUD column in the history file.
# Populated by LorryEngine._load_history(); takes priority over waypoint geocoding.
_HISTORY_CENTROIDS: dict[str, tuple] = {}

def _route_centroid(route: str) -> tuple | None:
    """Geographic centroid (lat, lng) of a route's destinations.

    Priority:
      1. GPS coordinates from the history LONGITUD column (most accurate).
      2. Waypoint geocoding via _MY_COORDS / Nominatim.
      3. Cluster capital city as last resort.
    """
    key = route.strip().upper()
    if key in _HISTORY_CENTROIDS:
        return _HISTORY_CENTROIDS[key]

    waypoints = list(_extract_waypoints(route))
    if not waypoints:
        intel = _extract_route_intelligence(route)
        city  = _CLUSTER_CITY.get(intel.get("cluster", "UNKNOWN"))
        return _geocode(city) if city else None

    coords = []
    for wp in waypoints[:6]:       # cap at 6 waypoints to limit network calls
        c = _geocode(wp)
        if c:
            coords.append(c)
    if not coords:
        intel = _extract_route_intelligence(route)
        city  = _CLUSTER_CITY.get(intel.get("cluster", "UNKNOWN"))
        return _geocode(city) if city else None
    return (
        sum(c[0] for c in coords) / len(coords),
        sum(c[1] for c in coords) / len(coords),
    )

def routes_distance_km(route1: str, route2: str) -> float | None:
    """Straight-line km between the geographic centroids of two routes."""
    c1 = _route_centroid(route1)
    c2 = _route_centroid(route2)
    if c1 is None or c2 is None:
        return None
    return _haversine_km(c1[0], c1[1], c2[0], c2[1])

def can_share_cross_cluster(route1: str, route2: str,
                             max_km: float = 300.0,
                             max_bearing_diff: float = 80.0,
                             min_depot_dist_km: float = 50.0) -> bool:
    """Two routes from different clusters can share a lorry when ALL of:
      1. Both route centroids are ≥ min_depot_dist_km from the KL depot —
         local suburban routes (Cheras, Serdang, Subang, <50 km) are never
         combined cross-cluster; only regional/long-haul routes qualify.
      2. Their geographic centroids are within max_km of each other.
      3. Their bearings from the KL depot differ by ≤ max_bearing_diff degrees —
         prevents directionally opposite routes being combined even if close
         (e.g. K.Selangor NW 304° vs Terengganu NE 34° = 90° diff).

    East Malaysia (Sabah/Sarawak ≈ 1 000–1 500 km from KL) is always rejected
    by the distance check alone.
    """
    c1 = _route_centroid(route1)
    c2 = _route_centroid(route2)
    if c1 is None or c2 is None:
        return False

    # Check 1: both routes must be regional (not local suburban)
    d1 = _haversine_km(_DEPOT[0], _DEPOT[1], c1[0], c1[1])
    d2 = _haversine_km(_DEPOT[0], _DEPOT[1], c2[0], c2[1])
    if d1 < min_depot_dist_km or d2 < min_depot_dist_km:
        return False

    # Check 2: routes must be geographically close to each other
    if _haversine_km(c1[0], c1[1], c2[0], c2[1]) > max_km:
        return False

    # Check 3: routes must head in roughly the same direction from the depot
    b1 = _bearing_deg(_DEPOT[0], _DEPOT[1], c1[0], c1[1])
    b2 = _bearing_deg(_DEPOT[0], _DEPOT[1], c2[0], c2[1])
    return _bearing_diff(b1, b2) <= max_bearing_diff


def _extract_route_intelligence(route: str) -> dict:
    """Derive cluster, corridor, route_code from a route string."""
    route_s = route.strip()
    prefix  = route_s[:2].upper()
    cluster = _CLUSTER_MAP.get(prefix, "UNKNOWN")

    m = re.match(r'^([A-Z]{2}\d+[A-Z]?)', route_s, re.IGNORECASE)
    route_code = m.group(1).upper() if m else prefix

    # Normalise separators so both "KV04A - N 4" and "JH09-->Kulai" work
    normalised = re.sub(r'-->|->|→', ' - ', route_s)
    corridor = "GENERAL"
    if " - " in normalised:
        parts      = [p.strip() for p in normalised.split(" - ")]
        suffix_raw = parts[-1].split()[0].upper() if parts else ""
        corridor   = _CORRIDOR_MAP.get(suffix_raw, "GENERAL")

    return {"cluster": cluster, "corridor": corridor, "route_code": route_code}


def _extract_waypoints(route: str) -> frozenset:
    """
    Extract intermediate place-name tokens from a route description string.
    Handles both "KV04A - PLACE1 - PLACE2 - N 4" and "JH09-->Kulai-Senai" formats.
    Returns a frozenset of uppercased tokens (empty if route is just a code).
    """
    normalised = re.sub(r'-->|->|→', ' - ', route.strip())
    parts = re.split(r'\s*-\s*', normalised)

    _dir_re  = re.compile(r'^(N|S|E|W|NE|NW|SE|SW|WN|ES|EN|WS|C|P)\s*\d*$', re.IGNORECASE)
    _code_re = re.compile(r'^[A-Z]{2}\d', re.IGNORECASE)

    waypoints = []
    for i, raw in enumerate(parts):
        tok = raw.strip().upper()
        if not tok:
            continue
        if i == 0 and _code_re.match(tok):
            continue             # skip leading route code (KV04A, JH09, …)
        if _dir_re.match(tok):
            continue             # skip direction suffix (N 4, WN 1, SE 3, …)
        tok = re.sub(r'\s+\d+\s*$', '', tok).strip()   # strip trailing numbers
        if tok and not tok.isdigit() and len(tok) > 1:
            waypoints.append(tok)
    return frozenset(waypoints)


def _routes_on_same_way(route1: str, route2: str) -> bool:
    """
    Return True when route1 and route2 can share a lorry because they travel
    in the same geographic direction.

    Two-path logic:
      Path A — named corridor (N / SE / C / …):
        Same cluster + same exact corridor string → same way.

      Path B — GENERAL corridor (route uses --> format or has no direction code):
        Same cluster + at least one shared waypoint (or one set is a
        subset of the other) → same way.

    Either path then applies a bearing sanity check: route centroids must
    point in a similar direction from the depot (≤80°).  This prevents
    "CENTRAL"-labelled routes that go in opposite directions (e.g. KV14A
    Puchong 64° ENE vs KV21A Puncak Alam 335° NNW) from being merged.

    Cross-cluster merges are never allowed (JH ≠ KV etc.).
    Routes whose cluster is UNKNOWN (bare codes like ZNA) are never merged.
    """
    ia = _extract_route_intelligence(route1)
    ib = _extract_route_intelligence(route2)

    if ia["cluster"] != ib["cluster"]:
        return False
    if ia["cluster"] == "UNKNOWN":
        return False

    # Path A: both have a named directional corridor — they must match
    if ia["corridor"] != "GENERAL" and ib["corridor"] != "GENERAL":
        if ia["corridor"] != ib["corridor"]:
            return False
    elif ia["corridor"] == "GENERAL" and ib["corridor"] == "GENERAL":
        # Path B: BOTH are GENERAL → need shared waypoints to confirm direction
        wp1 = _extract_waypoints(route1)
        wp2 = _extract_waypoints(route2)
        if not wp1 or not wp2:
            return False
        if not (bool(wp1 & wp2) or wp1 <= wp2 or wp2 <= wp1):
            return False
    # else: one has a named corridor and the other is GENERAL (same cluster).
    # The named corridor gives directional context; let the bearing check below
    # decide whether they actually point the same way.  This handles routes like
    # KV09A whose suffix contains extra text ("** START UKAY") that shifts
    # corridor parsing to GENERAL even though the route is in the same direction
    # as the adjacent NORTHEAST routes.

    # Bearing check: even same-corridor routes can go in opposite directions
    # (the "CENTRAL" code is reused for both E and W sub-routes in KV cluster).
    c1 = _route_centroid(route1)
    c2 = _route_centroid(route2)
    if c1 is not None and c2 is not None:
        d1 = _haversine_km(_DEPOT[0], _DEPOT[1], c1[0], c1[1])
        d2 = _haversine_km(_DEPOT[0], _DEPOT[1], c2[0], c2[1])
        if d1 >= 3.0 and d2 >= 3.0:  # skip if centroid is essentially at depot
            b1 = _bearing_deg(_DEPOT[0], _DEPOT[1], c1[0], c1[1])
            b2 = _bearing_deg(_DEPOT[0], _DEPOT[1], c2[0], c2[1])
            if _bearing_diff(b1, b2) > 80.0:
                return False

    return True


def _corridors_adjacent(c1: str, c2: str) -> bool:
    return c2 in _ADJACENT_CORRIDORS.get(c1, {c1})


def _distance_km(dist_str) -> Optional[float]:
    if not dist_str or str(dist_str).strip().lower() in ("nan", "", "-"):
        return None
    m = re.search(r"(\d+\.?\d*)", str(dist_str))
    return float(m.group(1)) if m else None


class LorryEngine:
    def __init__(self, master_path: str, history_path: str, owner_user: str):
        self.owner_user = owner_user.upper()
        self._load_master(master_path)
        self._load_history(history_path)
        self._build_route_frequency()
        self._build_daily_stop_counts()

    def _load_master(self, path):
        df = pd.read_excel(path)
        df.columns = [c.strip().upper() for c in df.columns]
        df["USER"] = df["USER"].str.strip().str.upper()
        self.eligible_lorries = df[df["USER"].isin({self.owner_user, "SPARE"})].copy()
        self.all_lorries = df.copy()

    @staticmethod
    def _parse_longitud_centroids(df: pd.DataFrame) -> dict:
        """Build {ROUTE_KEY: (lat, lon)} from the LONGITUD column.
        Each row may have one 'lat lon' pair; we average all pairs per route.
        """
        if "LONGITUD" not in df.columns or "ROUTE" not in df.columns:
            return {}
        out: dict[str, tuple] = {}
        for route, grp in df.groupby("ROUTE"):
            coords = []
            for val in grp["LONGITUD"].dropna():
                try:
                    parts = str(val).strip().split()
                    if len(parts) >= 2:
                        lat, lon = float(parts[0]), float(parts[1])
                        if -90 <= lat <= 90 and -180 <= lon <= 180:
                            coords.append((lat, lon))
                except (ValueError, IndexError):
                    continue
            if coords:
                out[str(route).strip().upper()] = (
                    sum(c[0] for c in coords) / len(coords),
                    sum(c[1] for c in coords) / len(coords),
                )
        return out

    def _load_history(self, path):
        import os, glob
        global _HISTORY_CENTROIDS
        paths = [path] if os.path.isfile(path) else (glob.glob(path + "*") or [path])
        frames = []
        for p in paths:
            try:
                eng = "xlrd" if str(p).lower().endswith(".xls") else "openpyxl"
                df  = pd.read_excel(p, engine=eng)
                df.columns = [str(c).strip().upper() for c in df.columns]
                # Some exports have a metadata row before the real header row.
                # Detect this: if "ROUTE" is missing but appears in the first data row.
                if "ROUTE" not in df.columns and len(df) > 0:
                    first_row = df.iloc[0].astype(str).str.strip().str.upper()
                    if "ROUTE" in first_row.values:
                        df = pd.read_excel(p, engine=eng, header=1)
                        df.columns = [str(c).strip().upper() for c in df.columns]
                if "GROSS WEIGHT" in df.columns and "WEIGHT(T)" not in df.columns:
                    df["WEIGHT(T)"] = pd.to_numeric(df["GROSS WEIGHT"], errors="coerce").fillna(0) / 1000.0
                if "LICENSE" in df.columns:
                    df["LICENSE"] = df["LICENSE"].fillna("").astype(str).str.strip().str.upper()
                    df = df[~df["LICENSE"].isin(["", "NAN", "NONE", "N/A", "-", "0", "0.0"])]
                if "ROUTE" in df.columns:
                    df["ROUTE"] = df["ROUTE"].fillna("").astype(str).str.strip()
                    df = df[df["ROUTE"] != ""]
                    intel = df["ROUTE"].apply(_extract_route_intelligence)
                    df["CLUSTER"]    = intel.apply(lambda x: x["cluster"])
                    df["CORRIDOR"]   = intel.apply(lambda x: x["corridor"])
                    df["ROUTE_CODE"] = intel.apply(lambda x: x["route_code"])
                    # Build GPS centroid cache from LONGITUD column
                    new_centroids = self._parse_longitud_centroids(df)
                    _HISTORY_CENTROIDS.update(new_centroids)
                if "DISTANCE" in df.columns:
                    df["DISTANCE_KM"] = df["DISTANCE"].apply(_distance_km)
                frames.append(df)
            except Exception as e:
                print(f"Warning: could not load history {p}: {e}")

        if not frames:
            self.history = pd.DataFrame(columns=["ROUTE", "LICENSE", "CUSTOMER NAME", "CLUSTER", "CORRIDOR"])
            return

        combined = pd.concat(frames, ignore_index=True)
        combined.columns = [c.strip().upper() for c in combined.columns]
        self.history = combined.dropna(subset=["ROUTE", "LICENSE"]).copy()

        for col in ["CLUSTER", "CORRIDOR", "ROUTE_CODE"]:
            if col not in self.history.columns:
                intel = self.history["ROUTE"].apply(_extract_route_intelligence)
                self.history["CLUSTER"]    = intel.apply(lambda x: x["cluster"])
                self.history["CORRIDOR"]   = intel.apply(lambda x: x["corridor"])
                self.history["ROUTE_CODE"] = intel.apply(lambda x: x["route_code"])
                break

        if "CUSTOMER NAME" not in self.history.columns:
            self.history["CUSTOMER NAME"] = ""
        self.history["CUSTOMER NAME"] = self.history["CUSTOMER NAME"].fillna("").astype(str).str.strip().str.upper()

        if "DATE" not in self.history.columns:
            self.history["DATE"] = pd.NaT
        else:
            self.history["DATE"] = pd.to_datetime(self.history["DATE"], errors="coerce")

    def _build_route_frequency(self):
        """Build 3 frequency tables: route, customer+route, cluster (Rules 4+5)."""
        self.route_freq = (
            self.history.groupby(["ROUTE", "LICENSE"]).size().reset_index(name="FREQ")
        )
        self.customer_route_freq = (
            self.history.groupby(["ROUTE", "CUSTOMER NAME", "LICENSE"]).size().reset_index(name="FREQ")
            if "CUSTOMER NAME" in self.history.columns
            else pd.DataFrame(columns=["ROUTE", "CUSTOMER NAME", "LICENSE", "FREQ"])
        )
        self.cluster_freq = (
            self.history.groupby(["CLUSTER", "LICENSE"]).size().reset_index(name="FREQ")
            if "CLUSTER" in self.history.columns
            else pd.DataFrame(columns=["CLUSTER", "LICENSE", "FREQ"])
        )

    def _build_daily_stop_counts(self):
        """Rule 6 — count stops (unique DOs) per lorry per date from history."""
        if "DATE" not in self.history.columns:
            self.daily_stop_counts: dict = {}
            return
        h = self.history.dropna(subset=["DATE", "LICENSE"]).copy()
        h["DATE_STR"] = h["DATE"].dt.strftime("%Y-%m-%d")
        counts = (
            h.groupby(["DATE_STR", "LICENSE"])["DO NUMBER"].nunique().reset_index(name="STOPS")
            if "DO NUMBER" in h.columns
            else h.groupby(["DATE_STR", "LICENSE"]).size().reset_index(name="STOPS")
        )
        self.daily_stop_counts = {
            (r["DATE_STR"], r["LICENSE"]): int(r["STOPS"])
            for _, r in counts.iterrows()
        }

    # ── Matching helpers ──────────────────────────────────────────────────────

    def _match_route(self, df, route, extra_filters=None):
        route_s = route.strip()
        for cmp in [
            lambda r: r == route_s,
            lambda r: r.upper() == route_s.upper(),
            lambda r: r.upper().startswith(route_s[:5].upper()) if len(route_s) >= 4 else False,
        ]:
            mask   = df["ROUTE"].str.strip().apply(cmp)
            subset = df[mask].copy()
            if not subset.empty:
                if extra_filters:
                    for col, val in extra_filters.items():
                        if col in subset.columns:
                            subset = subset[subset[col].str.upper() == val.upper()]
                if not subset.empty:
                    return subset.sort_values("FREQ", ascending=False)
        return pd.DataFrame(columns=df.columns)

    def get_route_frequencies(self, route):
        return self._match_route(self.route_freq, route)

    def get_customer_route_frequencies(self, route, customer_name):
        if not customer_name:
            return pd.DataFrame(columns=["ROUTE", "CUSTOMER NAME", "LICENSE", "FREQ"])
        return self._match_route(self.customer_route_freq, route,
                                  extra_filters={"CUSTOMER NAME": customer_name.strip().upper()})

    def get_cluster_frequencies(self, cluster):
        if self.cluster_freq.empty:
            return pd.DataFrame(columns=["CLUSTER", "LICENSE", "FREQ"])
        subset = self.cluster_freq[self.cluster_freq["CLUSTER"].str.upper() == cluster.upper()].copy()
        return subset.sort_values("FREQ", ascending=False)

    def get_stop_count_today(self, plate, date_str):
        return self.daily_stop_counts.get((date_str, plate.upper()), 0)

    # ── Rule 2 + Rule 7: merge check ─────────────────────────────────────────

    def can_merge_routes(self, route_a, route_b,
                         distance_a_km=None, distance_b_km=None):
        ia = _extract_route_intelligence(route_a)
        ib = _extract_route_intelligence(route_b)
        if ia["cluster"] != ib["cluster"]:
            return False
        if "GENERAL" in (ia["corridor"], ib["corridor"]):
            return False
        if not _corridors_adjacent(ia["corridor"], ib["corridor"]):
            return False
        if distance_a_km and distance_b_km:
            extra = abs(distance_b_km - distance_a_km)
            if distance_a_km > 0 and extra / distance_a_km > MERGE_DIST_THRESHOLD:
                return False
        return True

    def find_mergeable_routes(self, route, active_routes, distance_km=None):
        return [r for r in active_routes
                if r != route and self.can_merge_routes(route, r, distance_km)]

    # ── Core suggest ──────────────────────────────────────────────────────────

    def suggest(self, route, total_ton, unavailable=None, top_n=3,
                customer_name="", today_stop_counts=None, today_date_str=""):
        """
        Scoring (all PDF rules):
        1. Eligibility + Rule 6 (stop limit)
        2. Rule 4: CUST_FREQ    — customer+route history
        3. Rule 5: CLUSTER_FREQ — driver cluster familiarity
        4. Rule 3: UTIL_SCORE   — prefer >= 80% utilisation
        5.         SURPLUS ASC  — tightest fit
        6.         IS_OWNER     — owner before SPARE
        7.         ROUTE_FREQ   — general route history
        """
        if unavailable is None:
            unavailable = set()
        if today_stop_counts is None:
            today_stop_counts = {}

        intel   = _extract_route_intelligence(route)
        cluster = intel["cluster"]

        eligible = self.eligible_lorries[
            (self.eligible_lorries["TON"] >= total_ton) &
            (~self.eligible_lorries["LORRY"].isin(unavailable))
        ].copy()
        if eligible.empty:
            return []

        # Rule 6: exclude over-stop-limit lorries
        if today_stop_counts or today_date_str:
            over = {
                r["LORRY"] for _, r in eligible.iterrows()
                if (today_stop_counts.get(r["LORRY"], 0)
                    or self.get_stop_count_today(r["LORRY"], today_date_str)) >= MAX_STOPS_PER_LORRY
            }
            eligible = eligible[~eligible["LORRY"].isin(over)]
            if eligible.empty:
                return []

        eligible = eligible.copy()
        eligible["SURPLUS"] = eligible["TON"] - total_ton

        # Route freq
        freq_route = self.get_route_frequencies(route)
        merged = eligible.merge(
            freq_route[["LICENSE", "FREQ"]].rename(columns={"FREQ": "ROUTE_FREQ"}),
            left_on="LORRY", right_on="LICENSE", how="left")
        merged["ROUTE_FREQ"] = merged["ROUTE_FREQ"].fillna(0).astype(int)

        # Customer+route freq (Rule 4)
        if customer_name:
            cf = self.get_customer_route_frequencies(route, customer_name)
            merged = merged.merge(
                cf[["LICENSE", "FREQ"]].rename(columns={"FREQ": "CUST_FREQ"}),
                left_on="LORRY", right_on="LICENSE", how="left") if not cf.empty else merged
        if "CUST_FREQ" not in merged.columns:
            merged["CUST_FREQ"] = 0
        merged["CUST_FREQ"] = merged["CUST_FREQ"].fillna(0).astype(int)

        # Cluster freq (Rule 5)
        clf = self.get_cluster_frequencies(cluster)
        if not clf.empty:
            merged = merged.merge(
                clf[["LICENSE", "FREQ"]].rename(columns={"FREQ": "CLUSTER_FREQ"}),
                left_on="LORRY", right_on="LICENSE", how="left")
        if "CLUSTER_FREQ" not in merged.columns:
            merged["CLUSTER_FREQ"] = 0
        merged["CLUSTER_FREQ"] = merged["CLUSTER_FREQ"].fillna(0).astype(int)

        # Utilisation (Rule 3)
        merged["UTIL"] = total_ton / merged["TON"]
        merged["UTIL_SCORE"] = merged["UTIL"].apply(
            lambda u: 1.0 if u >= CAPACITY_TARGET else u / CAPACITY_TARGET)
        merged["IS_OWNER"] = (merged["USER"].str.upper() == self.owner_user).astype(int)

        # Per-tier sort — SURPLUS (tightest fit) is the primary key in every tier.
        # History (CUST_FREQ / CLUSTER_FREQ) breaks ties within the same capacity class
        # so a familiar driver is preferred when two lorries are equally efficient.
        # This ensures minimum capacity waste regardless of route history.
        _sort_cols = ["SURPLUS", "CUST_FREQ", "CLUSTER_FREQ", "UTIL_SCORE", "IS_OWNER", "ROUTE_FREQ"]
        _sort_asc  = [True,      False,       False,          False,        False,       False]
        UTIL_GOOD_THRESHOLD = 0.60
        UTIL_OK_THRESHOLD   = 0.40
        merged["UTIL_GOOD"] = (merged["UTIL"] >= UTIL_GOOD_THRESHOLD).astype(int)
        merged["UTIL_OK"]   = (merged["UTIL"] >= UTIL_OK_THRESHOLD).astype(int)

        tier2 = merged[merged["UTIL_GOOD"] == 1].sort_values(_sort_cols, ascending=_sort_asc)
        tier1 = merged[(merged["UTIL_GOOD"] == 0) & (merged["UTIL_OK"] == 1)].sort_values(_sort_cols, ascending=_sort_asc)
        tier0 = merged[merged["UTIL_OK"] == 0].sort_values(_sort_cols, ascending=_sort_asc)
        import pandas as _pd
        merged = _pd.concat([tier2, tier1, tier0]).reset_index(drop=True)

        results = []
        for _, row in merged.head(top_n).iterrows():
            surplus    = round(float(row["SURPLUS"]), 2)
            cust_freq  = int(row["CUST_FREQ"])
            clust_freq = int(row["CLUSTER_FREQ"])
            route_freq = int(row["ROUTE_FREQ"])
            util_pct   = round(float(row["UTIL"]) * 100, 1)

            if cust_freq > 0:
                reason = f"Served this customer {cust_freq}x ({util_pct}% utilised, {surplus}T spare)"
            elif clust_freq > 0:
                reason = f"Familiar with {cluster} region ({clust_freq}x) — {util_pct}% utilised"
            elif route_freq > 0:
                reason = f"{route_freq}x on this route — {util_pct}% utilised, {surplus}T spare"
            else:
                reason = f"Best fit — {util_pct}% utilised, {surplus}T spare"

            results.append({
                "LORRY":        row["LORRY"],
                "TON_CAPACITY": round(float(row["TON"]), 2),
                "SURPLUS":      surplus,
                "UTIL_PCT":     util_pct,
                "USER":         row["USER"],
                "FREQ":         route_freq,
                "CUST_FREQ":    cust_freq,
                "CLUSTER_FREQ": clust_freq,
                "CLUSTER":      cluster,
                "CORRIDOR":     intel["corridor"],
                "REASON":       reason,
            })
        return results

    def get_eligible_lorry_list(self):
        return self.eligible_lorries[["LORRY", "TON", "USER"]].copy()

    # ── Split suggestion ──────────────────────────────────────────────────────

    def suggest_split(self, route, total_ton, unavailable=None, max_lorries=6,
                      single_util_threshold=0.70, today_stop_counts=None, today_date_str=""):
        if unavailable is None:
            unavailable = set()
        if today_stop_counts is None:
            today_stop_counts = {}

        freq_route = self.get_route_frequencies(route)

        def _enrich(df, min_ton):
            d = df.copy()
            d["SURPLUS"] = d["TON"] - min_ton
            d = d.merge(freq_route[["LICENSE", "FREQ"]], left_on="LORRY", right_on="LICENSE", how="left")
            d["FREQ"] = d["FREQ"].fillna(0).astype(int)
            return d

        eligible = self.eligible_lorries[
            (self.eligible_lorries["TON"] >= total_ton) &
            (~self.eligible_lorries["LORRY"].isin(unavailable))
        ].copy()

        if eligible.empty:
            best_surplus = float("inf")
            best_cap     = float("inf")
        else:
            enriched = _enrich(eligible, total_ton)
            enriched["IS_OWNER"] = (enriched["USER"].str.upper() == self.owner_user).astype(int)
            enriched = enriched.sort_values(["SURPLUS", "IS_OWNER", "FREQ"], ascending=[True, False, False])
            best_row     = enriched.iloc[0]
            best_surplus = float(best_row["SURPLUS"])
            best_cap     = float(best_row["TON"])
            if total_ton / best_cap >= single_util_threshold:
                return None

        small_pool = self.eligible_lorries[
            (self.eligible_lorries["TON"] < best_cap) &
            (~self.eligible_lorries["LORRY"].isin(unavailable))
        ].copy()
        if small_pool.empty:
            return None

        small_pool = _enrich(small_pool, 0)
        small_pool["IS_OWNER"] = (small_pool["USER"].str.upper() == self.owner_user).astype(int)
        small_pool = small_pool.sort_values(["IS_OWNER", "FREQ", "TON"], ascending=[False, False, False])

        remain = total_ton
        used   = set(unavailable)
        chosen = []

        for _, row in small_pool.iterrows():
            if remain <= 0:
                break
            plate = row["LORRY"]
            if plate in used:
                continue
            stops = today_stop_counts.get(plate, 0) or self.get_stop_count_today(plate, today_date_str)
            if stops >= MAX_STOPS_PER_LORRY:
                continue
            cap     = float(row["TON"])
            portion = round(min(cap, remain), 6)
            surplus = round(cap - portion, 2)
            util_pct = round(portion / cap * 100, 1)
            chosen.append({
                "LORRY": plate, "TON_CAPACITY": round(cap, 2),
                "SURPLUS": surplus, "UTIL_PCT": util_pct,
                "USER": str(row["USER"]), "FREQ": int(row["FREQ"]),
                "REASON": f"Split {util_pct}% utilised, {surplus}T spare",
                "PORTION": portion,
            })
            used.add(plate)
            remain = round(remain - cap, 6)
            if len(chosen) >= max_lorries:
                break

        if remain > 0:
            return None
        if sum(c["SURPLUS"] for c in chosen) >= best_surplus:
            return None
        # Reject if any lorry in the split would be severely underutilised.
        # (e.g. VEA2818 1.07T carrying 58 kg overflow = 5% — better to use a
        # single slightly-large lorry at 59% than two lorries where one is empty)
        _MIN_SPLIT_UTIL = 40.0   # each split lorry must carry ≥ 40% of its capacity
        if any(c["UTIL_PCT"] < _MIN_SPLIT_UTIL for c in chosen):
            return None
        return chosen

    @staticmethod
    def route_intel(route: str) -> dict:
        return _extract_route_intelligence(route)

    def suggest_largest_available(self, route: str, unavailable=None,
                                   today_date_str: str = "",
                                   total_ton: float = 0.0) -> list:
        """
        Last-resort assignment: return the single largest lorry still available
        whose capacity is ≥ total_ton (so the lorry is NOT overloaded).

        If every remaining lorry would be overloaded (total_ton > all caps),
        returns [] — caller then assigns NO_LORRY.

        Returns a one-element list in the same format as suggest(), or [].
        """
        if unavailable is None:
            unavailable = set()

        eligible = self.eligible_lorries[
            ~self.eligible_lorries["LORRY"].isin(unavailable)
        ].copy()
        if eligible.empty:
            return []

        # Rule 6: honour the stop-count limit
        if today_date_str:
            over = {
                r["LORRY"] for _, r in eligible.iterrows()
                if self.get_stop_count_today(r["LORRY"], today_date_str) >= MAX_STOPS_PER_LORRY
            }
            eligible = eligible[~eligible["LORRY"].isin(over)]
        if eligible.empty:
            return []

        # Only assign lorries that can carry the load without exceeding capacity
        if total_ton > 0:
            eligible = eligible[eligible["TON"] >= total_ton].copy()
        if eligible.empty:
            return []   # every lorry would be overloaded → NO_LORRY

        intel      = _extract_route_intelligence(route)
        freq_route = self.get_route_frequencies(route)
        merged     = eligible.merge(
            freq_route[["LICENSE", "FREQ"]].rename(columns={"FREQ": "ROUTE_FREQ"}),
            left_on="LORRY", right_on="LICENSE", how="left")
        merged["ROUTE_FREQ"] = merged["ROUTE_FREQ"].fillna(0).astype(int)
        merged["IS_OWNER"]   = (merged["USER"].str.upper() == self.owner_user).astype(int)

        # Prefer tightest fit first (smallest surplus), then owner/history
        if total_ton > 0:
            merged["SURPLUS"] = merged["TON"] - total_ton
            merged = merged.sort_values(
                ["SURPLUS", "IS_OWNER", "ROUTE_FREQ"],
                ascending=[True, False, False])
        else:
            merged = merged.sort_values(
                ["TON", "IS_OWNER", "ROUTE_FREQ"],
                ascending=[False, False, False])

        row     = merged.iloc[0]
        cap     = round(float(row["TON"]), 2)
        surplus = round(cap - total_ton, 2) if total_ton > 0 else 0.0
        util_pct = round(total_ton / cap * 100, 1) if cap > 0 and total_ton > 0 else 0.0
        return [{
            "LORRY":        row["LORRY"],
            "TON_CAPACITY": cap,
            "SURPLUS":      surplus,
            "UTIL_PCT":     util_pct,
            "USER":         str(row["USER"]),
            "FREQ":         int(row["ROUTE_FREQ"]),
            "CLUSTER":      intel["cluster"],
            "CORRIDOR":     intel["corridor"],
            "REASON":       f"Best available lorry ({cap}T) — {util_pct}% utilised, {surplus}T spare",
        }]
