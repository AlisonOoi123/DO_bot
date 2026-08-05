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

## The decisive finding: the manual uses **2nd trips**
On the heavy day (07-18) the manual ran lorries at **133%** — that's not just
overloading, it's a **second trip** (same lorry, morning + afternoon). A
single-trip fleet at a sane 115% cap physically cannot carry that day's weight,
so the optimizer correctly leaves 24 T unassigned rather than fake it.

Two consequences:
1. **Model 2nd trips.** Allow each lorry up to 2 runs/day (double stops/capacity,
   flagged Trip 1 / Trip 2). This is what closes the gap on heavy days. The
   engine already has a `TRIP` concept in the export — reuse it.
2. **This is exactly your "consider tomorrow" feature.** When even 2 trips can't
   fit the day, the optimizer delivers what fits and flags the overflow →
   "these DOs don't fit today's fleet — send tomorrow." Under-utilized lorries
   (< `UTIL_FLOOR`) get the same flag.

## Next steps
- Add **2nd-trip** decision to the model; re-run the 3 days (expect 07-18 to fit).
- Run the **full 227-day shadow** (ABI) → optimizer-vs-manual-vs-bot table +
  utilization histogram + hard-rule invariant check on every day.
- Then wire behind the `OPTIMIZER_ENABLED` flag with fallback (Phase 1 integration).

## Status
Model proven on the hardest part (fast, optimal, rule-abiding). The remaining
work is 2nd-trip modeling + the full-corpus shadow gate. Engine still untouched.
