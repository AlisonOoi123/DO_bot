"""Bridge between the DO engine and the CP-SAT optimizer.

Additive and OFF by default. When OPTIMIZER_ENABLED is set, `run_optimizer`
rebuilds the assignment from the engine's already-parsed items using the
optimizer, and overwrites each item's LORRY on success. On ANY problem
(solver missing, no solution, exception) it returns False and the caller keeps
the current greedy result — so the app can never be worse than today.

The optimizer's output is rule-valid by construction; the engine's audit gate
still runs as a second line of defense.
"""
from __future__ import annotations
import os
import re
import sys

_OPT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "optimizer")
if _OPT_DIR not in sys.path:
    sys.path.insert(0, _OPT_DIR)


def optimizer_enabled() -> bool:
    return str(os.environ.get("OPTIMIZER_ENABLED", "")).strip().lower() in ("1", "true", "yes", "on")


_SENTINELS = {"NO_LORRY", "NO_ELIGIBLE_LORRY", "OTHER_USER", "NOT_TODAY",
              "OUT_SOURCE", "REMARKS_SKIP", "SPLIT", "SKIPPED", "", None}


def _pfx(route: str) -> str:
    m = re.match(r'\s*([A-Z]{1,3}\d+[A-Z]?)', str(route))
    return m.group(1) if m else ""


def run_optimizer(sess, is_urban_fn, is_kuantan_fn) -> bool:
    """Rebuild the assignment via the optimizer. Returns True if it produced a
    complete, applied solution; False to fall back to the existing result.

    `is_urban_fn(item)` and `is_kuantan_fn(route)` are passed in from the engine
    so classification matches exactly.
    """
    try:
        from optimizer import Cluster, Lorry, optimize, MAX_DOS_PER_LORRY, OVERLOAD_CAP
    except Exception:
        return False

    eng = sess.get("engine")
    if eng is None:
        return False
    try:
        caps = {str(r["LORRY"]).strip().upper(): float(r["TON"])
                for _, r in eng.eligible_lorries.iterrows()}
    except Exception:
        return False
    # Fleet = eligible lorries minus only those genuinely unavailable: broken,
    # or already taken by ANOTHER user today. NOT sess["unavailable"] — that was
    # populated by the greedy pass we're replacing, so it would wrongly block
    # every lorry the greedy just used.
    blocked = set()
    try:
        import bot as _bot
        blocked = set(_bot.get_broken_lorries().keys())
        _me = str(sess.get("user_id", "")).strip().upper()
        _by = _bot.get_assigned_by()          # {plate: user} taken today
        blocked |= {p for p, u in _by.items() if u != _me}
    except Exception:
        pass
    fleet = [Lorry(p, c) for p, c in caps.items() if p not in blocked]
    if not fleet:
        return False
    maxcap = max(c for _, c in caps.items())

    items = sess.get("items", []) or []
    # Only optimize the DOs that are OURS to assign now (mine + capacity-stranded).
    # Skip other-user / not-today / out-source / van-remarks-skip — leave as-is.
    assignable = [it for it in items
                  if it.get("LORRY") not in ("OTHER_USER", "NOT_TODAY",
                                             "OUT_SOURCE", "REMARKS_SKIP")]
    if not assignable:
        return False

    # Build clusters: group by (route code, size-cap tier) so a VAN / MAX-N-TON
    # DO never caps the whole route to a tiny lorry (the engine keeps these
    # separate too). Then split into chunks <= maxcap / 15 stops.
    from collections import defaultdict
    by_code = defaultdict(list)
    for it in assignable:
        cap_tier = it.get("MAX_TON")            # None = uncapped
        by_code[(_pfx(it.get("ROUTE", "")), cap_tier)].append(it)

    clusters = []
    cid_items = {}
    for (code, _tier), group in by_code.items():
        group.sort(key=lambda it: -(it.get("WEIGHT", 0.0) or 0.0))
        lats = [it.get("GPS_LAT") for it in group if it.get("GPS_LAT") is not None]
        lons = [it.get("GPS_LON") for it in group if it.get("GPS_LON") is not None]
        clat = sum(lats) / len(lats) if lats else None
        clon = sum(lons) / len(lons) if lons else None
        iu = bool(is_urban_fn(group[0]))
        ik = bool(is_kuantan_fn(group[0].get("ROUTE", "")))
        chunk, cw, idx = [], 0.0, 0
        for it in group:
            w = it.get("WEIGHT", 0.0) or 0.0
            if chunk and (cw + w > maxcap or len(chunk) >= MAX_DOS_PER_LORRY):
                cid = f"{code}~{_tier}#{idx}"
                mt = min((x["MAX_TON"] for x in chunk if x.get("MAX_TON") is not None), default=None)
                fb = set()
                for x in chunk:
                    if x.get("FORBID_PLATES"):
                        fb |= x["FORBID_PLATES"]
                clusters.append(Cluster(cid, round(cw, 3), clat, clon, iu, ik, mt, len(chunk), fb))
                cid_items[cid] = list(chunk)
                idx += 1; chunk, cw = [], 0.0
            chunk.append(it); cw += w
        if chunk:
            cid = f"{code}~{_tier}#{idx}"
            mt = min((x["MAX_TON"] for x in chunk if x.get("MAX_TON") is not None), default=None)
            fb = set()
            for x in chunk:
                if x.get("FORBID_PLATES"):
                    fb |= x["FORBID_PLATES"]
            clusters.append(Cluster(cid, round(cw, 3), clat, clon, iu, ik, mt, len(chunk), fb))
            cid_items[cid] = list(chunk)

    if not clusters:
        return False

    # Warm-start hint: the greedy plate currently on each cluster's items.
    hint = {}
    for cid, its in cid_items.items():
        plates = [it.get("LORRY") for it in its if it.get("LORRY") not in _SENTINELS]
        if plates:
            hint[cid] = max(set(plates), key=plates.count)

    result = optimize(clusters, fleet, time_limit_s=5.0, hint=hint)
    if result.get("status") not in ("OPTIMAL", "FEASIBLE"):
        return False

    # Apply: write each cluster's plate to its items; unplaced clusters -> NO_LORRY.
    assign = result.get("assign", {})
    for cid, plate in assign.items():
        for it in cid_items.get(cid, []):
            if plate:
                it["LORRY"] = plate
                sess["assigned"][it["DO NUMBER"]] = plate
            else:
                it["LORRY"] = "NO_LORRY"
                sess["assigned"][it["DO NUMBER"]] = "NO_LORRY"
    import logging
    logging.info("[OPTIMIZER] applied CP-SAT plan: %d clusters, %d lorries.",
                 len(clusters), len({p for p in assign.values() if p}))
    return True
