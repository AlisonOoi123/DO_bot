# Phase 1 — CP-SAT optimizer prototype (progress)

`optimizer.py` — standalone CP-SAT model, solves a day in **seconds (optimal)**.
Not wired into the engine; validated in shadow against manual + current bot.

## Constraints encoded so far
- one lorry per cluster; capacity with overload cap (1.15); **max 15 stops/lorry**;
  urban/outstation purity; **Kuantan alone**; far-urban clusters can't share
  (dist ≤ 0.33°); per-DO size cap; forbidden plates.

## Middle-ground constants (calibrated, option C)
`URBAN_MIX_DIST_DEG=0.33`, `OVERLOAD_CAP=1.15`, `MAX_DOS_PER_LORRY=15`,
`UTIL_FLOOR=0.35` (see `calibration.json`).

## Head-to-head (ABI)
| Day | Manual | Optimizer | note |
|-----|--------|-----------|------|
| 2025-05-09 | 11 lorry @126% | 11 @65%, 0 unassigned | ✅ everything delivered |
| 2025-08-12 | 11 @110% | 11 @78%, ~2T unassigned | close |
| 2025-07-18 | 10 @133% | 9 @89%, **24T unassigned** | heavy day — can't fit |

## Scope decision (user): SINGLE-TRIP only — no morning/afternoon
Each upload is one assignment. **No 2nd-trip / trip-splitting modeling.** The
flow stays: user picks **today / tomorrow** → engine filters to that day's
scheduled routes → asks whether to include any off-schedule routes → optimizer
assigns the chosen DOs in a single trip. Genuine overflow (won't fit the day's
fleet) is flagged unassigned → "send tomorrow."

Note on the heavy-day 24 T: that test used `_ignore_schedule` (ALL DOs).
Under the real **schedule filter**, most of that overflow is off-schedule
routes the user wouldn't include for that day — so real single-trip overflow
should be small. The full shadow run below must use the schedule filter.

## Next steps
- **Full 227-day shadow (ABI)** using the **schedule filter** (real conditions):
  optimizer-vs-manual-vs-bot table + utilization histogram + a hard-rule
  invariant check on every day. This is the acceptance gate.
- Add the **"send tomorrow" flag**: after assigning, list any lorry below
  `UTIL_FLOOR` and any unassigned DOs → "consider sending tomorrow."
- Then wire behind the `OPTIMIZER_ENABLED` flag with automatic fallback to the
  current engine (shadow → canary → default).

## Status
Model proven on the hardest part (fast, optimal, rule-abiding, single-trip).
Remaining work is the full-corpus shadow gate + the tomorrow flag + guarded
integration. Engine still untouched.

## Shadow gate results (ABI, 57 days) + tuning
- **0 hard-rule violations** on every day. ✅
- Optimizer uses **fewer lorries than the manual** (9.0 vs 10.8).
- Tuning sweep (deliver-everything is priority #1):

  | dist / overload | fully-delivered (feasible days) | avg lorries | avg util |
  |---|---|---|---|
  | 0.33 / 1.15 | 14/17 | 9.9 | 77% |
  | **0.45 / 1.20** | **16/17** | **9.6** | **81%** |
  | 0.55 / 1.25 | 16/17 | 9.5 | 79% |

- **Locked: 0.45° / 1.20** — best delivery without over-mixing (0.45° sits
  between the manual's p75 and p90; below the p99 outliers). The ~6% of
  feasible days not fully delivered are genuine overflow → "send tomorrow."

## Remaining before integration
1. Full 227-day run at 0.45/1.20 for the final acceptance table (background).
2. The "send tomorrow" report (under-filled lorries + unassigned DOs).
3. Wire behind `OPTIMIZER_ENABLED` with fallback; shadow -> canary (ABI) -> default.
