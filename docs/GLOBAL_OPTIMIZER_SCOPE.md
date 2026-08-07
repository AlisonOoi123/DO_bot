# Global Lorry-Assignment Optimizer — Scoping Document

**Status:** Draft for review · **Owner:** engineering · **Goal:** replace the
greedy per-cluster allocation with a whole-day optimizer that matches (or beats)
the manual planner, without regressing the rules already in production.

---

## 1. Why

The current engine assigns DOs greedily (cluster by cluster, then top-up/rescue
passes). It is **correct** — it never mis-routes and honors every hard rule —
but it decides locally, so on busy days it can:

- leave a few DOs unassigned when a cluster overflows its lorry (58_Test: 2 KV20A), or
- run a big lorry light because the small lorries filled up first (DO_9: VJN8929 14.8T at 16% for KV12A alone).

A human planner avoids both by balancing the **whole board at once** — choosing
which lorry serves which cluster/direction globally. That is an optimization
problem, and a greedy algorithm provably can't match it in general. This project
builds the global solver.

The two failed local attempts (documented in git history: the "urban reservation"
and "free-a-lorry post-pass") confirmed a local rule cannot fix this without
regressing other days. Hence a proper optimizer.

---

## 2. Objective

Minimize, in priority order (lexicographic, or weighted):

1. **Unassigned DOs** (deliver everything possible) — highest priority.
2. **Number of lorries used** (fewer trips / cost).
3. **Wasted capacity** on opened lorries (don't roll a 15T truck at 16%).

Subject to all the hard constraints in §3. "Match the manual" = produce plans
whose metrics are at least as good as the manual's on historical days.

---

## 3. Hard constraints (the rules the optimizer MUST satisfy)

These are already enforced piecemeal in `bot.py`; the optimizer must encode them
as constraints so a solution is never invalid.

| # | Rule | Current source of truth |
|---|------|------|
| C1 | A DO rides exactly one lorry (or is explicitly unassigned). | assignment core |
| C2 | Lorry load ≤ capacity (TON). Urban may slight-overload to **×1.20**; outstation **never** overloads. | `SLIGHT_OVERLOAD`, audit gate |
| C3 | **Same route code → one lorry** when it fits; a code heavier than any lorry may split. | split logic |
| C4 | **Different codes may share a lorry only if geographically near** — urban stops within `URBAN_MERGE_SPREAD_DEG` (≈28 km); single-linkage `MAX_GEO_GAP_DEG`. | `_lorry_geo_ok`, geo-cleanup |
| C5 | **Urban (KL/Selangor) never mixes with outstation** on one lorry. | `_is_urban_do`, rescue |
| C6 | **Kuantan (PH09) always independent** — Kuantan-only lorries. | `_is_kuantan` |
| C7 | Group by **actual STATE**, not route code, for mislabeled routes (e.g. NS04→Pahang). | `_state_corridor`, `_STATE_TO_CORRIDOR` |
| C8 | **Big lorries → outstation, medium/small → urban**, but only when urban demand needs it (else outstation may use a medium). | *(the unmet goal)* |
| C9 | Per-DO **size cap** (`MAX_TON`, incl. VAN=2T) and **forbidden plates** respected. | `_remarks_lorry_cap`, `FORBID_PLATES` |
| C10 | **VAN-remark DOs** ride a van (≤2T) separately, not the route's main lorry. | cap-split logic |
| C11 | **No cross-user lorry borrowing** (ABI uses ABI+SPARE; VIVIAN uses VIVIAN+SPARE). | fleet eligibility |
| C12 | Only lorries **Available** in the master file (not `Block`) are usable. | master upload |
| C13 | Per-lorry **max stops / max DOs** limits. | `MAX_STOPS_PER_LORRY`, `MAX_DOS_PER_LORRY` |
| C14 | Outstation minimum tonnage: lorries ≤5T can't run outstation. | `OUTSTATION_MIN_TON` |
| C15 | Off-schedule DOs handled by the existing day/schedule question (optimizer runs on the DOs the user chose to assign). | trip-day flow |

**Soft preferences** (objective terms, not hard):
- Prefer tightest-fitting lorry per cluster (minimize waste).
- Prefer the manual's observed lorry-per-route history where it doesn't conflict.

---

## 4. Approach options

| Option | Pros | Cons | Verdict |
|------|------|------|------|
| **A. OR-Tools CP-SAT** (constraint solver) | Handles all constraints natively; proven; finds optimal/near-optimal; time-boxable | New dependency (`ortools`); model complexity; geo constraints need pre-clustering | **Recommended** |
| B. MILP (PuLP/CBC) | Standard; explainable | Weaker on the combinatorial geo/clustering; slower | Fallback |
| C. Local-search / simulated annealing over the greedy result | No heavy dependency; incremental; easy fallback | No optimality guarantee; tuning-heavy | Good **Phase-2a** stepping stone |

**Recommendation:** pre-cluster DOs geographically (reuse the existing bucketing:
route code → GPS cluster → direction), then let **CP-SAT assign clusters→lorries**
with the constraints above and the objective in §2. Clustering shrinks the model
so it solves in well under the per-request time budget. Keep option C as a
warm-start / fallback and for environments where `ortools` can't be installed.

---

## 5. Architecture & safe integration

The optimizer is an **alternative allocator**, not a rewrite:

```
_handle_excel_upload
  → parse + bucket DOs           (UNCHANGED — reuse existing clustering)
  → allocator:
       if OPTIMIZER_ENABLED and solver available:
            solution = optimize(clusters, fleet, constraints)   # NEW
            if solution invalid / solver timeout:
                 fall back to current greedy allocator
       else:
            current greedy allocator                            # UNCHANGED
  → same audit gate, geo-cleanup, export                        (UNCHANGED)
```

Key safety properties:
- **Behind a flag** (`OPTIMIZER_ENABLED`, default OFF). Zero effect until switched on.
- **Fallback**: any solver failure/timeout/invalid result → current engine. The app can never be worse than today.
- **Same audit gate**: the optimizer's output passes through the existing rules-compliance audit, so an invalid assignment is impossible even if the model has a bug.
- **Shadow mode** (see §6): compute the optimizer plan, log the comparison, but keep applying the current plan — until we trust it.

---

## 6. Validation & rollout (this de-risks everything)

**Phase 0 — Evaluation harness** *(build first, useful regardless)*
- A script that runs any (master, DO, manual) triple and reports, per plan:
  unassigned count, lorries used, utilization histogram, far-mix count, and a
  hard-rule violation check.
- Collect a corpus of real days (the 58/DO_7/DO_9 + as many historical
  master+DO+manual sets as available).
- Baseline table: **current engine vs manual** on every day. This is the yardstick.

**Phase 1 — Optimizer in shadow mode**
- Implement the CP-SAT allocator behind the flag.
- Run it in shadow on the corpus; produce **optimizer vs current vs manual** metrics.
- Acceptance gate: optimizer ≥ current on unassigned AND utilization for **every**
  day in the corpus, with **zero** hard-rule violations, within the time budget.

**Phase 2 — Canary**
- Enable for one user (e.g. VIVIAN) on the live app; keep current as fallback.
- Watch for a week; compare against the planner's manual adjustments.

**Phase 3 — Default on**
- Enable for all; keep the flag and fallback indefinitely.

---

## 7. Testing (non-negotiable, given prior crashes)

- **Regression**: the existing isolated per-file harness (6/7/15/16/20/22/23/26 +
  58 + DO_7 + DO_9) must show unassigned ≤ current and no new far-mix on every file.
- **Crash/robustness**: the DO_7 recursion case; empty fleet; all-blocked master;
  single giant DO; solver timeout path.
- **Invariant check**: assert every hard constraint C1–C15 on the solver output
  before it's accepted (belt-and-suspenders with the audit gate).

---

## 8. Risks & mitigations

| Risk | Mitigation |
|------|------|
| `ortools` can't be installed on the Windows box | Ship option C (pure-Python local search) as the fallback allocator; CP-SAT optional. |
| Solver too slow for a live request | Time-box (e.g. 5 s); on timeout, return best-so-far or fall back to greedy. |
| Model bug produces an invalid plan | Audit gate + explicit C1–C15 invariant check reject it → fallback. |
| Over-fitting to a few sample days | Build the corpus first (Phase 0); gate on the whole corpus, not one file. |
| Regressing production | Flag OFF by default, shadow → canary → default; instant rollback via flag. |

---

## 9. Effort estimate (rough)

| Phase | Work | Est. |
|------|------|------|
| 0 | Evaluation harness + corpus + baseline table | ~1–2 days |
| 1 | CP-SAT model + integration + shadow logging | ~3–5 days |
| 2 | Canary + tuning | ~1 week elapsed (mostly observation) |
| 3 | Default-on + docs | ~1 day |

Prerequisite from you: **as many real (master + DO + manual) day-sets as you can
share** — the more historical days, the more trustworthy the gate. Two or three
isn't enough to prove we won't regress a fourth.

---

## 10. Open decisions for you

1. **`ortools` allowed** on the office PC? (If not, we go pure-Python local search — slightly lower quality, same safety.)
2. **Objective weights**: is "deliver everything" always #1, even if it means one extra lightly-loaded lorry? (Assumed yes.)
3. **How many historical day-sets** can you provide for the corpus?
4. **Canary user**: ABI or VIVIAN first?

Once these are answered, Phase 0 (the harness + baseline) starts — and it's
useful on its own even before any optimizer exists.
