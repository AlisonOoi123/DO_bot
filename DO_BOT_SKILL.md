# DO_BOT ASSIGNMENT SKILL — Strict Implementation Script

**Version 2.0 · Authoritative decision procedure for the DO_BOT AI.**
This file is the single source of truth for HOW a lorry is chosen for a DO.
It is intentionally *preference + intelligence based, NOT hardcoded*: no rule
says "route PH07 must use lorry X". Instead the bot decides from **owner,
outstation status, geography (city / state / route code / longitude), gross
weight, and the SHIP_DETAIL column**, with the goal of **fully utilising every
lorry**.

> Governance: Do not overturn any rule in this file without operator sign-off.
> Every decision must be explainable ("no black box"). When two rules conflict,
> the HARD CONSTRAINTS (Section A) always win.

---

## A. HARD CONSTRAINTS (never violated, no exceptions)

These are physical / contractual limits. They are checked first and can never
be overridden by optimisation or utilisation goals.

1. **Owner isolation.** The logged-in user (ABI or VIVIAN) may only be assigned
   lorries they own **plus SPARE** lorries. A DO whose route belongs to the
   other user is left blank (`OTHER_USER`) — never assigned.
2. **Outstation minimum tonnage.** Outstation = any route whose destination
   group is LARGE_LONG or MEDIUM_LONG. The minimum lorry size depends on which:
   - **Far outstation (LARGE_LONG:** Kuantan, Pahang, Johor, Perak, Terengganu,
     Kelantan, …) must use a lorry **> 11 T**.
   - **Nearer outstation (MEDIUM_LONG:** NS/Seremban, KV01A north / Rawang-
     Serendah direction) must use a lorry **> 5 T** (so an 8–9 T lorry such as
     BMN3682 can still serve KV01A).
   A lorry below the applicable minimum can never serve that outstation route.
   Each outstation route is also consolidated onto the fewest big lorries: a
   whole route rides ONE big lorry when it fits; if a per-stop plate forbid or
   capacity prevents that, its atomic components are re-packed across valid big
   lorries of the SAME direction (never mixed with urban), splitting only when
   unavoidable.
3. **VAN class.** A lorry under 2 T is the *van* class. Used for DOs that
   explicitly require a van (see Section D).
4. **Small-lorry class.** A lorry under 5 T is the *small lorry* class.
5. **REMARKS / SHIP_DETAIL size cap.** A DO that specifies a maximum lorry size
   must never ride a lorry larger than that cap (see Section D for the exact
   phrases currently enforced).
5a. **LORRY NAIK ceiling.** A lorry's max load is the `LORRY NAIK (5%)` value in
   LORRY DAILY PLANNING.xlsx (MUATAN sheet) — the engine loads it as `TON`. A
   lorry may NEVER carry more than its `LORRY NAIK (5%)` tonnage. No extra
   overage is applied on top (the 5% naik is already inside that number), so
   every naik factor in the code is 1.0.
6. **OUT SOURCE.** A DO whose SHIP_DETAIL contains `OUT SOURCE` is handled by a
   third-party lorry. The bot assigns **no plate** — `LICENSE` stays blank and
   the DO is **not** counted as unassigned.
7. **Geographic integrity.** DOs on one lorry must form ONE geographic cluster
   (Section C). A far-away drop is never bolted onto an unrelated cluster.
8. **File integrity.** The exported file keeps the uploaded file's EXACT columns
   and order. Only assignment columns (LICENSE, DRIVER, assistants) are filled.
   No columns are added, removed, or reordered.

---

## B. THE FIVE OPERATING PRINCIPLES (in priority order)

The bot works through these in order for every upload.

### Principle 1 — Who is logged in?
Determine the user (ABI / VIVIAN). Build the eligible fleet = that user's
lorries + SPARE. All later steps choose ONLY from this fleet. Routes owned by
the other user are set aside (`OTHER_USER`) and never touched.

### Principle 2 — Is the destination outstation?
For each DO, classify the destination (by route-code prefix, then STATE / CITY,
then postcode — see Section E for how STATE is resolved):
- **Outstation** (LARGE_LONG / MEDIUM_LONG): exclude every lorry ≤ 5 T from
  consideration. Only lorries > 5 T may be used.
- **Urban** (KL / Selangor): any lorry size may be used, including vans and
  small lorries.

### Principle 3 — Cluster by geography, then fill by weight.
Group DOs that share **route code + state + city** and are geographically near
(nearest longitude / GPS). Assign these clusters together, choosing the
**tightest-fitting** lorry that covers the cluster's total gross weight, so each
lorry is filled as full as possible. Then keep filling underused lorries with
the nearest compatible DOs until they reach the utilisation target. This is how
we achieve Principle 5 (full utilisation) without hardcoding.

Same route code always rides one lorry (however many cities it spans) **as long
as the stops form one connected GPS cluster** — see Section C.

**Split priority (when a load must be divided across lorries), in order:**
1. **Same longitude → same lorry.** DOs at the same GPS point are never split.
2. **Same route + same customer (CODE, column D) → same lorry.** Never split.
3. **Same route, different customer → prefer the same lorry** (may split only
   when capacity forces it).
4. **Fully utilise gross weight** — pack the fixed atomic units (1 & 2) to fill
   each lorry as much as possible.

**Urban route-first packing (KL / Selangor only — NOT outstation).** A whole
urban route code rides ONE lorry whenever it fits (priority 1): any urban route
split across lorries is consolidated onto a single lorry that can hold its full
weight within capacity, size cap, forbidden plates and the geo spread —
preferring the lorry already holding most of it. Then a lorry with leftover
space is topped up with the geographically **nearest** other whole urban route
(priority 2). A route stays split only when no single lorry can hold it.
Outstation routes (Kuantan, Perak, Pahang, …) are never consolidated this way —
they split across large lorries out of capacity necessity.

Criteria 1 and 2 are enforced as inseparable *atomic units*: every split path
(heavy-group half-split, urban de-concentration) moves whole units only, and a
final **reunification pass** pulls back onto one lorry any atomic unit that an
earlier pass left split — adding a lorry if needed rather than breaking the
unit. A unit is only left split when no single lorry in the fleet can hold it.

### Principle 4 — Read SHIP_DETAIL for the size preference.
SHIP_DETAIL format: `<days>, [AM|PM], MAX <N> TON` (any part optional).
- `MAX N TON` = the largest lorry the customer can receive. Within that ceiling,
  pick the lorry by **gross weight** (tightest fit) — don't send a 15 T lorry
  for 200 kg just because MAX 15 TON is allowed.
- `OUT SOURCE` = assign no plate (Hard Constraint A6).
- Days / AM / PM are read but **not** used to filter assignment (the operator
  chooses the trip day at login; the SCHD sheet does day filtering).

> ⚠️ CURRENT CODE STATE (pending operator decision): only `MAX 2 TON` from
> SHIP_DETAIL is enforced today (→ van). `MAX 5 / 11 / 15 / 21 TON` are NOT
> capped — those DOs follow ordinary weight/route rules. This matches the
> operator's most recent explicit instruction. Principle 4 above describes the
> intended "cap ≤ N and pick by weight" behaviour. **These differ — do not
> change without confirming which is wanted.**

### Principle 5 — Fully utilise every lorry.
The optimisation goal. After the constraint-respecting assignment, run the
optimisation passes (Section F) to raise utilisation: consolidate scattered
same-route stops, pull nearby DOs onto underfilled lorries, and prefer filling
an already-used compatible lorry over opening a new one — always within the
hard constraints.

---

## C. GEOGRAPHIC CLUSTERING (GPS single-linkage)

A lorry's stops must form ONE connected cluster:
- Two stops are "connected" when they are in the **same state**
  (KL and Selangor count as one urban state) AND within
  **`MAX_GEO_GAP_DEG` = 0.29°** straight-line distance
  (√(Δlat² + Δlon²), ≈ 32 km).
- All stops on a lorry must be reachable from each other through a chain of such
  ≤ 0.29° hops. A legitimate multi-city route (e.g. Benta ↔ Lipis ↔ Kuala
  Lipis, each hop ≤ 0.29°) stays whole; a stop with a wrong/far GPS is split off
  and re-homed onto a lorry where it fits the chain, else left unassigned.
- Distance is measured by **GPS only** — a stop with a correct city name but a
  wrong coordinate is still separated.
- **Urban anti-chaining (longitude-aware).** Single-linkage alone can let a
  string of close stops *bridge* two far urban zones (e.g. central KL → Rawang,
  Kajang → Rawang, or Semenyih → KL). So for URBAN stops of **different** route
  codes, the overall spread is also bounded: no two different-route urban stops
  may be more than `URBAN_MERGE_SPREAD_DEG` (0.25° ≈ 28 km) apart. Same route
  code stays atomic (any distance). This is enforced in BOTH the merge/move
  paths (`_lorry_geo_ok`) AND the geo-enforcement split pass, which applies
  **complete-linkage**: within a kept cluster it detaches the lighter offending
  route code until no cross-route urban pair exceeds the cap, then re-homes it.
  Result: Rawang may share a lorry with nearby Batu Caves, but never with far
  Kajang, Semenyih, or central KL.
- **Missing-GPS inference.** A DO that arrives with a blank LONGITUD is given
  coordinates from the mean GPS of other DOs in the SAME upload sharing its CITY
  (fallback: POSCODE). Without this a no-GPS DO can't be geo-separated and may
  wrongly ride a far lorry (e.g. a Semenyih DO landing on a Rawang lorry).

VAN priority: all VAN-remark DOs are pooled ACROSS routes, clustered by the same
0.29° single-linkage rule, and packed onto a ≤ 2 T van together.

---

## D. SIZE-REQUIREMENT PHRASES (REMARKS + SHIP_DETAIL)

Detection is case-insensitive, whole-word, LORI/LORRY interchangeable.

| Phrase (in REMARKS or SHIP_DETAIL) | Max lorry tonnage | Notes |
|---|---|---|
| `VAN`, `VAN ONLY` | ≤ 2 T | van class |
| `SMALL LORRY`, `LORRY SMALL`, `LORRY/LORI KECIL`, `LORI BESAR TAK BOLEH` | ≤ 5 T | small lorry |
| `BELOW 5 / 10 / 14 / 20 TON` | ≤ that value | free-text tonnage |
| `MAX 2 TON` (SHIP_DETAIL) | ≤ 2 T | **enforced today** |
| `MAX 5 / 11 / 15 / 21 TON` (SHIP_DETAIL) | *not capped today* | see Principle 4 warning |
| `ANY SIZE`, `BIG LORRY` | no cap | |
| `OUT SOURCE` (SHIP_DETAIL) | — | no plate assigned |

When several DOs share a lorry, the binding cap is the **smallest** cap among
them. A DO with no size phrase imposes no cap.

The phrase→cap table is loaded at runtime from the `REMARKS FIELD` sheet in
LORRY DAILY PLANNING.xlsx when present, falling back to
`assignment_config.py` (`REMARKS_FIELD3_TON_CAP`, `REMARKS_SIZE_ALIASES`).

---

## E. STATE / CITY / POSTCODE RESOLUTION

Destination STATE for each DO is resolved in this order:
1. Explicit `STATE` column.
2. `CITY` looked up in the **Malaysia States & Cities** sheet.
3. `POSCODE` — exact match in that sheet's POSTCODE column, then the postcode
   range table in `assignment_config.py` (`POSTCODE_STATE_RANGES`).

Operators maintain the Malaysia States & Cities sheet (STATE / CITY / POSTCODE)
directly — the bot reads it at startup, so new towns/postcodes need no code
change.

---

## F. OPTIMISATION PASSES (run in this order, all within Section A constraints)

1. **Initial assignment** — cluster (Principle 3) → tightest-fit lorry.
2. **Consolidation / force-assign** — place any still-unassigned DO on a
   compatible lorry with room.
3. **Overflow / split** — split a group across two lorries only when no single
   lorry fits; never onto a size-capped or outstation-illegal lorry.
4. **Swap / same-route merge / partial rebalance** — improve utilisation;
   every move re-checks size cap, outstation min, state, and the 0.29° chain.
5. **Bin-pack safety** — a lorry is only marked "used/unavailable" once it
   actually receives items (empty bins release their lorry for later groups).
6. **Geographic enforcement** — split any lorry that isn't one GPS cluster;
   re-home the outliers.
7. **VAN-priority consolidation** — pool all VAN DOs onto the van(s).
8. **Same-route (urban) consolidation** — pull scattered same-route stops
   together so a route isn't split between a full lorry and a near-empty van.
9. **Urban >11 T de-concentration** — a lorry over `URBAN_MAX_TON` (11 T)
   cannot physically run the many closely-spaced small drops of urban
   (KL / Selangor) routes ("only a lorry under 11 TON can handle that many
   routes"). This pass moves urban stops off any >11 T lorry onto ≤11 T lorries
   that have room (same-destination clusters kept whole; weight, size cap,
   state, geo chain and forbidden plates all respected). It is **best-effort,
   never lossy**: a cluster with no ≤11 T home stays on the big lorry rather
   than being left unassigned. Outstation routes are untouched — big lorries
   are exactly what they need.

---

## G. WHAT IS **NOT** HARDCODED (removed on purpose)

- No `ROUTE_PREFERRED_LORRY` (empty) — no "route X must use lorry Y".
- No `LORRY_STRICT_ROUTE` (empty) — no lorry is bound to a route direction.
Assignment is entirely driven by owner + outstation + geography + weight +
SHIP_DETAIL, per the principles above.

---

## H. SENTINEL VALUES (rows that are intentionally NOT assigned)

`OTHER_USER` (other owner's route) · `NOT_TODAY` (not on today's SCHD) ·
`REMARKS_SKIP` (day-restricted remark) · `OUT_SOURCE` (third-party) ·
`NO_LORRY` (genuinely could not fit any eligible lorry).
Only `NO_LORRY` counts as a real "unassigned"; the others are expected skips and
are exported with a blank LICENSE.

---

## END OF SKILL SCRIPT
