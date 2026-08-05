# Phase 0 — Findings from the historical corpus

Data: `ZSDOROUTEWRH.xls`, manual assignments **24 Mar → 10 Dec 2025**, blank-plate
DOs ignored, capacities/owners from `master_lorry_1.xlsx`.

## Corpus
- **49,982** delivered DOs across **227 days**, **26** distinct lorries.
- ~**226 DOs/day** (median), ~**20 lorries/day** (both users combined).
- Enough data to gate a change without over-fitting. ✅

## Baseline: current bot vs manual (ABI, 11 sampled days)

| metric | MANUAL | BOT (current) |
|------|------|------|
| avg lorry utilization | **~110–133%** | **~57–81%** |
| lorries used | 9–13 | 9–13 (often more) |
| unassigned | 0 (by definition) | low (15 total; 14 on one bad day) |

### The two big learnings

**1. The bot's real weakness is UTILIZATION, not unassigned.**
The bot already delivers almost everything. But it runs lorries at **60–80%**
while the manual runs them at **110–130%**. Same DOs, but the bot picks **bigger
lorries loaded loosely**; the manual picks **smaller lorries packed tight** (and
overloads). This is the VJN8929-at-16% symptom, everywhere.

**2. The manual is LOOSER than the rules we hard-coded.**
- **Mixing distance:** manual mixes *different* urban route codes at a **median
  span of 0.277° (~31 km)** and **p90 0.515° (~57 km)**. Our guard is **0.25°** —
  i.e. **stricter than the manual**, which is why it stranded extra DOs.
- **Overload:** the manual routinely loads lorries to **110–130%** of the
  `master_lorry` TON — not the 20%-urban / 0%-outstation we enforce.

## Implications for the optimizer objective
Given "deliver everything is #1, then flag under-utilized lorries for tomorrow":
1. **Primary:** minimize unassigned. (Bot already good.)
2. **Strong secondary:** **maximize utilization / minimize lorries** — prefer the
   **smallest lorry that fits** a cluster, pack tight. This is where the win is.
3. Then flag any lorry left below a utilization floor → "consider sending these
   DOs tomorrow" (the feature you asked for).

## A reconciliation to decide (important)
The corpus shows the **manual mixes farther (~0.4–0.5°) and overloads more
(~120%)** than the rules you asked me to enforce (near-only ≤0.25°, urban-only
20% overload, outstation 0%). These directly conflict:

- If the optimizer follows **your stated rules** → it will look "cleaner" but
  will leave more DOs unassigned / run more lorries than the manual.
- If it follows **the manual's actual behavior** → it matches the manual's
  efficiency but mixes farther and overloads more than you said you wanted.

**We should pick the target before building the model.** Likely answer: a
middle ground — allow the manual's tighter packing (smaller lorries, higher
fill) but keep a *sane* distance cap (say the manual's p75, not p99) and a
capacity cap (say ~110–115%). We can tune these from the corpus.

## Next (Phase 0 remainder)
- Extend the baseline to **all 227 days** (batch run overnight) for both users,
  producing the full current-vs-manual table + a utilization histogram.
- From the corpus, **learn the real thresholds** (distance cap, overload cap,
  utilization floor) as tunable constants.
- Then Phase 1: CP-SAT model using those learned targets, in shadow mode.
