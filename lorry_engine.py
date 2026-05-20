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
from math import radians, cos, sin, asin, sqrt
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

# ── Geographic cross-cluster merging (Google Geocoding + Haversine) ───────────

_GEOCACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "data", "geocode_cache.json")
_geocode_cache: dict = {}
_geocache_dirty = False

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

def _geocode(place: str, api_key: str) -> tuple | None:
    """Geocode a place name to (lat, lng) using Google Geocoding API.
    Results are cached to data/geocode_cache.json to avoid repeated API calls.
    """
    key = place.strip().upper()
    if key in _geocode_cache:
        cached = _geocode_cache[key]
        return tuple(cached) if cached else None
    if not api_key:
        return None
    global _geocache_dirty
    try:
        resp = requests.get(
            "https://maps.googleapis.com/maps/api/geocode/json",
            params={"address": f"{place}, Malaysia", "key": api_key},
            timeout=5,
        )
        data = resp.json()
        if data.get("status") == "OK":
            loc = data["results"][0]["geometry"]["location"]
            coords = [loc["lat"], loc["lng"]]
            _geocode_cache[key] = coords
            _geocache_dirty = True
            _save_geocache()
            return tuple(coords)
    except Exception:
        pass
    # Cache the miss so we don't retry every time
    _geocode_cache[key] = None
    _geocache_dirty = True
    _save_geocache()
    return None

# Fallback city for each cluster when no waypoints can be extracted
_CLUSTER_CITY = {
    "KL_VALLEY":       "Petaling Jaya",
    "KL_CITY":         "Kuala Lumpur City Centre",
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

def _route_centroid(route: str, api_key: str) -> tuple | None:
    """Return the geographic centroid (lat, lng) of a route's destinations.
    Uses waypoints from the route description; falls back to cluster capital city.
    """
    waypoints = list(_extract_waypoints(route))
    if not waypoints:
        intel = _extract_route_intelligence(route)
        city  = _CLUSTER_CITY.get(intel.get("cluster", "UNKNOWN"))
        return _geocode(city, api_key) if city else None

    coords = []
    for wp in waypoints[:6]:  # cap at 6 to limit API calls
        c = _geocode(wp, api_key)
        if c:
            coords.append(c)
    if not coords:
        # Fallback to cluster city if no waypoints resolved
        intel = _extract_route_intelligence(route)
        city  = _CLUSTER_CITY.get(intel.get("cluster", "UNKNOWN"))
        return _geocode(city, api_key) if city else None
    return (
        sum(c[0] for c in coords) / len(coords),
        sum(c[1] for c in coords) / len(coords),
    )

def routes_distance_km(route1: str, route2: str, api_key: str) -> float | None:
    """Straight-line km between the geographic centroids of two routes.
    Returns None if either route cannot be geocoded.
    """
    c1 = _route_centroid(route1, api_key)
    c2 = _route_centroid(route2, api_key)
    if c1 is None or c2 is None:
        return None
    return _haversine_km(c1[0], c1[1], c2[0], c2[1])

def can_share_cross_cluster(route1: str, route2: str, api_key: str,
                             max_km: float = 300.0) -> bool:
    """Two routes from different clusters can share a lorry if their geographic
    centroids are within max_km straight-line distance of each other.
    300 km covers all intra-Peninsular combinations; East Malaysia (Sabah/Sarawak)
    is ~1 500 km from KL and will always be rejected.
    """
    if not api_key:
        return False
    dist = routes_distance_km(route1, route2, api_key)
    if dist is None:
        return False
    return dist <= max_km


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
        (The corridor code is already a geographic direction; no further
        waypoint check is required.  KV04A-NORTH and KV05A-NORTH both head
        north from the warehouse even if their specific stops differ.)

      Path B — GENERAL corridor (route uses --> format or has no direction code):
        Same cluster + at least one shared waypoint (or one set is a
        subset of the other) → same way.
        (All three Pahang routes go through Bentong; all Johor branches
        converge on JB — shared waypoints are the geographic signal.)

    Cross-cluster merges are never allowed (JH ≠ KV etc.).
    Routes whose cluster is UNKNOWN (bare codes like ZNA) are never merged.
    """
    ia = _extract_route_intelligence(route1)
    ib = _extract_route_intelligence(route2)

    if ia["cluster"] != ib["cluster"]:
        return False
    if ia["cluster"] == "UNKNOWN":
        return False

    # Path A: both have a named directional corridor
    if ia["corridor"] != "GENERAL" and ib["corridor"] != "GENERAL":
        return ia["corridor"] == ib["corridor"]

    # Path B: at least one is GENERAL → need shared waypoints
    wp1 = _extract_waypoints(route1)
    wp2 = _extract_waypoints(route2)
    if not wp1 or not wp2:
        return False
    return bool(wp1 & wp2) or wp1 <= wp2 or wp2 <= wp1


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

    def _load_history(self, path):
        import os, glob
        paths = [path] if os.path.isfile(path) else (glob.glob(path + "*") or [path])
        frames = []
        for p in paths:
            try:
                eng = "xlrd" if str(p).lower().endswith(".xls") else "openpyxl"
                df  = pd.read_excel(p, engine=eng)
                df.columns = [c.strip().upper() for c in df.columns]
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

        # Three-tier utilisation sort (Rules 3 + 4):
        #
        # Tier 2 (UTIL_GOOD=1): load uses ≥60% of capacity → "good fit".
        #   Within tier: history (CUST/CLUSTER/ROUTE_FREQ) breaks the tie,
        #   then SURPLUS (tightest fit), then owner preference.
        #
        # Tier 1 (UTIL_OK=1, UTIL_GOOD=0): 40–60% util → "acceptable".
        #   Within tier: SURPLUS first (prefer smaller lorry = less waste),
        #   then history as tiebreaker.  This prevents a 20T lorry with high
        #   history beating a 5T lorry just because it's driven that route before.
        #
        # Tier 0 (UTIL_OK=0): <40% util → "poor fit".
        #   Within tier: SURPLUS first, then history.
        #
        # Cross-tier: Tier 2 always beats Tier 1 always beats Tier 0.
        UTIL_GOOD_THRESHOLD = 0.60
        UTIL_OK_THRESHOLD   = 0.40
        merged["UTIL_GOOD"] = (merged["UTIL"] >= UTIL_GOOD_THRESHOLD).astype(int)
        merged["UTIL_OK"]   = (merged["UTIL"] >= UTIL_OK_THRESHOLD).astype(int)

        # Per-tier sort — each tier has its own key priority:
        # Tier 2 (≥60%): good fit already, prefer familiar driver → history first
        # Tier 1 (40-60%) and Tier 0 (<40%): prefer SMALLEST sufficient lorry
        #   (tightest fit / smallest SURPLUS) to avoid assigning a 10T truck for 1T load
        _t2_cols = ["CUST_FREQ", "CLUSTER_FREQ", "UTIL_SCORE", "SURPLUS", "IS_OWNER", "ROUTE_FREQ"]
        _t1_cols = ["SURPLUS", "CUST_FREQ", "CLUSTER_FREQ", "UTIL_SCORE", "IS_OWNER", "ROUTE_FREQ"]
        tier2 = merged[merged["UTIL_GOOD"] == 1].sort_values(
            _t2_cols, ascending=[False, False, False, True, False, False])
        tier1 = merged[(merged["UTIL_GOOD"] == 0) & (merged["UTIL_OK"] == 1)].sort_values(
            _t1_cols, ascending=[True, False, False, False, False, False])
        tier0 = merged[merged["UTIL_OK"] == 0].sort_values(
            _t1_cols, ascending=[True, False, False, False, False, False])
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
