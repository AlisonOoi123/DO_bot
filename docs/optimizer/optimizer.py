"""Phase-1 prototype: whole-day cluster->lorry optimizer (CP-SAT).

Standalone and side-effect-free. Given a day's clusters and the available fleet,
returns an assignment that honors the hard rules and maximizes delivered weight
then packs tight (fewest lorries, highest fill). NOT wired into the engine yet —
this is the shadow-mode model we validate before integration.

Middle-ground constants (option C) calibrated from the historical corpus:
  URBAN_MIX_DIST_DEG  0.45   (tuned: max full-delivery without over-mixing;
                              near the manual's p75-p90 span, below p99 0.9)
  OVERLOAD_CAP        1.20   (tuned single-trip cap)
  UTIL_FLOOR          0.35   (lorries below this are flagged "send tomorrow")
"""
from __future__ import annotations
from dataclasses import dataclass, field
import math

URBAN_MIX_DIST_DEG = 0.45
OVERLOAD_CAP = 1.20
UTIL_FLOOR = 0.35
MAX_DOS_PER_LORRY = 15


@dataclass
class Cluster:
    id: str
    weight: float
    lat: float | None
    lon: float | None
    is_urban: bool
    is_kuantan: bool
    max_ton: float | None            # per-DO size cap (min across the cluster)
    ndos: int = 1                    # number of delivery stops in this cluster
    forbid: set = field(default_factory=set)   # plates this cluster can't use


@dataclass
class Lorry:
    plate: str
    cap: float


def _far(a: Cluster, b: Cluster) -> bool:
    if a.lat is None or b.lat is None:
        return False
    return math.hypot(a.lat - b.lat, a.lon - b.lon) > URBAN_MIX_DIST_DEG


def _ffd(clusters: list[Cluster], fleet: list[Lorry]) -> dict:
    """First-fit-decreasing feasible assignment respecting ALL hard rules.
    Used as the CP-SAT warm-start AND as a delivery floor (so the final result
    is never worse than this even if the solver times out). {cid: plate|None}."""
    load = {l.plate: 0.0 for l in fleet}
    stops = {l.plate: 0 for l in fleet}
    is_urb = {l.plate: None for l in fleet}    # None=empty, True/False once set
    has_kt = {l.plate: False for l in fleet}
    pts = {l.plate: [] for l in fleet}         # urban (lat,lon) on the lorry
    capf = {l.plate: l.cap for l in fleet}
    out = {}
    for c in sorted(clusters, key=lambda z: -z.weight):
        best = None
        for l in fleet:
            p = l.plate
            if p in c.forbid:
                continue
            if c.max_ton is not None and capf[p] > c.max_ton + 1e-9:
                continue
            if load[p] + c.weight > capf[p] * OVERLOAD_CAP + 1e-9:
                continue
            if stops[p] + c.ndos > MAX_DOS_PER_LORRY:
                continue
            if has_kt[p]:                       # kuantan lorry takes nothing else
                continue
            if c.is_kuantan and load[p] > 1e-9:  # kuantan needs an empty lorry
                continue
            if is_urb[p] is not None and is_urb[p] != c.is_urban:
                continue                        # urban/outstation purity
            if c.is_urban and c.lat is not None and any(
                    math.hypot(c.lat - a[0], c.lon - a[1]) > URBAN_MIX_DIST_DEG
                    for a in pts[p]):
                continue                        # far urban mix
            # tightest fit that still leaves the lorry usable
            if best is None or capf[p] < capf[best.plate]:
                best = l
        if best is None:
            out[c.id] = None
            continue
        p = best.plate
        load[p] += c.weight; stops[p] += c.ndos
        is_urb[p] = c.is_urban
        if c.is_kuantan:
            has_kt[p] = True
        if c.is_urban and c.lat is not None:
            pts[p].append((c.lat, c.lon))
        out[c.id] = p
    return out


def optimize(clusters: list[Cluster], fleet: list[Lorry],
             time_limit_s: float = 5.0, hint: dict | None = None) -> dict:
    """Return {'assign': {cluster_id: plate|None}, 'status': str}.

    `hint` (optional) is {cluster_id: plate} — e.g. the greedy result — used to
    warm-start the solver so it converges fast within the time limit.
    """
    from ortools.sat.python import cp_model
    mdl = cp_model.CpModel()
    C, L = clusters, fleet
    capf = {l.plate: l.cap for l in L}

    # Allowed (cluster, lorry) pairs — hard per-pair filters.
    def allowed(c: Cluster, l: Lorry) -> bool:
        if l.plate in c.forbid:
            return False
        if c.max_ton is not None and l.cap > c.max_ton + 1e-9:
            return False
        if c.weight > l.cap * OVERLOAD_CAP + 1e-9:
            return False
        return True

    x = {}
    for ci, c in enumerate(C):
        for l in L:
            if allowed(c, l):
                x[(ci, l.plate)] = mdl.NewBoolVar(f"x_{ci}_{l.plate}")

    used = {l.plate: mdl.NewBoolVar(f"u_{l.plate}") for l in L}
    urban = {l.plate: mdl.NewBoolVar(f"urb_{l.plate}") for l in L}   # lorry serves urban

    # Each cluster to at most one lorry.
    assigned = []
    for ci, c in enumerate(C):
        xs = [x[(ci, l.plate)] for l in L if (ci, l.plate) in x]
        a = mdl.NewBoolVar(f"a_{ci}")
        if xs:
            mdl.Add(sum(xs) == a)
        else:
            mdl.Add(a == 0)
        assigned.append(a)

    for l in L:
        p = l.plate
        xs = [x[(ci, p)] for ci in range(len(C)) if (ci, p) in x]
        # capacity (with overload)
        mdl.Add(sum(int(round(C[ci].weight * 1000)) * x[(ci, p)]
                    for ci in range(len(C)) if (ci, p) in x)
                <= int(round(capf[p] * OVERLOAD_CAP * 1000)))
        # max stops (DOs) per lorry
        mdl.Add(sum(C[ci].ndos * x[(ci, p)]
                    for ci in range(len(C)) if (ci, p) in x) <= MAX_DOS_PER_LORRY)
        # used flag
        if xs:
            mdl.AddMaxEquality(used[p], xs)
        else:
            mdl.Add(used[p] == 0)
        # urban/outstation purity
        for ci, c in enumerate(C):
            if (ci, p) not in x:
                continue
            if c.is_urban:
                mdl.Add(urban[p] == 1).OnlyEnforceIf(x[(ci, p)])
            else:
                mdl.Add(urban[p] == 0).OnlyEnforceIf(x[(ci, p)])
        # kuantan alone: a kuantan cluster on l => l carries nothing else
        for ci, c in enumerate(C):
            if (ci, p) in x and c.is_kuantan:
                for cj, c2 in enumerate(C):
                    if cj != ci and (cj, p) in x:
                        mdl.Add(x[(cj, p)] == 0).OnlyEnforceIf(x[(ci, p)])

    # Far urban clusters can't share a lorry.
    for i in range(len(C)):
        for j in range(i + 1, len(C)):
            if C[i].is_urban and C[j].is_urban and _far(C[i], C[j]):
                for l in L:
                    if (i, l.plate) in x and (j, l.plate) in x:
                        mdl.Add(x[(i, l.plate)] + x[(j, l.plate)] <= 1)

    # Objective: deliver weight FIRST (dominant), fewer lorries as tiebreak.
    mdl.Maximize(
        sum(int(round(C[ci].weight * 1000)) * assigned[ci] for ci in range(len(C))) * 10000
        - sum(used[l.plate] for l in L)
    )

    # Feasible construction = warm-start AND delivery floor.
    floor = _ffd(C, L)
    id2i = {c.id: i for i, c in enumerate(C)}
    for cid, plate in floor.items():
        ci = id2i.get(cid)
        if ci is not None and plate and (ci, plate) in x:
            mdl.AddHint(x[(ci, plate)], 1)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_s
    solver.parameters.num_search_workers = 8
    st = solver.Solve(mdl)
    status = {cp_model.OPTIMAL: "OPTIMAL", cp_model.FEASIBLE: "FEASIBLE"}.get(st, "NONE")

    out = {}
    if st in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for ci, c in enumerate(C):
            plate = None
            for l in L:
                if (ci, l.plate) in x and solver.Value(x[(ci, l.plate)]) == 1:
                    plate = l.plate
                    break
            out[c.id] = plate
        # Delivery floor: if the solver (e.g. timed out) delivers LESS weight
        # than the feasible construction, use the construction instead.
        wid = {c.id: c.weight for c in C}
        w_solver = sum(wid[cid] for cid, p in out.items() if p)
        w_floor = sum(wid[cid] for cid, p in floor.items() if p)
        if w_floor > w_solver + 1e-6:
            return {"assign": floor, "status": status + "+FLOOR"}
        return {"assign": out, "status": status}
    # Solver found nothing usable — return the feasible construction.
    return {"assign": floor, "status": "FLOOR"}
