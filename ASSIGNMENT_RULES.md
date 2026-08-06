# Lorry Assignment Rules
# Version: 1.0 — Read by bot on every run. All rules below are MANDATORY.
# Last updated: 2026-06-15

---

## SECTION 1 — DATE PRIORITY

**RULE 1.1 — Earliest Date First**
Always assign DOs with the earliest date first. Within any group sharing the same route + city + state, sort items by DATE ascending before assigning lorries. Older orders must never be displaced by newer orders for the same lorry slot.

**RULE 1.2 — Capacity Split by Date**
If a group's total weight exceeds any single lorry's capacity and must be split, the earliest-date items fill the first lorry. Newer items go to the second lorry or remain pending (NO_LORRY).

---

## SECTION 2 — GEOGRAPHIC GROUPING (Mandatory before assignment)

**RULE 2.1 — Group by Route + City + State First**
Before assigning any lorry, categorize all items by:
  1. ROUTE (exact route code, e.g. KV11A)
  2. STATE (e.g. KUALA LUMPUR, SELANGOR, PAHANG)
  3. CITY (e.g. AMPANG, CHERAS, KUANTAN)
Items sharing the same Route + State + City must be assigned to the SAME lorry.

**RULE 2.2 — Nearest Longitude Within Same Bucket**
Within the same Route + State + City bucket, sort items by nearest longitude (GPS proximity to lorry's route centroid). Geographically closest items are assigned together to avoid mixed-direction delivery runs.

**RULE 2.3 — No Route Mixing Across Different Geography**
Do NOT assign items from different cities, different states, or geographically distant longitudes to the same lorry unless they are in the same corridor group (see Rule 8.1). Items far apart in longitude or from different cities/states must use separate lorries.

**RULE 2.4 — Urban Sub-Bucketing (KL / Selangor)**
Urban routes (KV codes, KL, Selangor): split buckets by STATE + CITY.
  - Example: KV11A items in KUALA LUMPUR city and KV11A items in AMPANG (SELANGOR) are separate buckets — different lorries.

**RULE 2.5 — Outstation Sub-Bucketing (GPS Bearing)**
Outstation routes (PH, NS, TR, PK, etc.): split buckets by STATE + GPS bearing octant from depot (N/NE/E/SE/S/SW/W/NW). Prevents same route code items pointing in genuinely opposite directions from sharing a lorry.

---

## SECTION 3 — LORRY SIZE RULES

**RULE 3.1 — Small Lorry (≤5T) — Urban Only**
Lorries with tonnage ≤ 5T (including 4.2T vans) may ONLY be assigned to KL or Selangor (urban) routes. They are STRICTLY FORBIDDEN on any outstation route (Kuantan, Pahang, Seremban, Negeri Sembilan, Tanjung Malim, Rawang outstation, Kemaman, Port Dickson, or any route > ~50 km from depot).

**RULE 3.2 — Outstation Minimum Tonnage: 5.001T**
Any route classified as outstation (LARGE_LONG or MEDIUM_LONG) requires a lorry with TON > 5T. The minimum effective tonnage is 5.001T. This applies to:
  - LARGE_LONG: PH, TR, KB, JH, PK, KD, PN, MC, SB, SR (Pahang, Terengganu, Kelantan, Johor, Perak, Kedah, Penang, Melaka, Sabah, Sarawak)
  - MEDIUM_LONG: NS (Negeri Sembilan / Seremban), KV01A (Rawang / Tanjung Malim direction toward Perak)

**RULE 3.3 — Urban Routes Accept Any Lorry Size**
KL and Selangor urban routes (KV codes) accept lorries of ANY size (van, small, medium, large). Lorry selection is weight-optimized (best-fit first). No upper tonnage cap for urban routes.

**RULE 3.4 — Lorry Size Categories**
  - SMALL: ≤ 5T (vans and small lorries)
  - MEDIUM: 5T – 11T
  - LARGE: ≥ 11T

---

## SECTION 4 — OUTSTATION ROUTE RULES

**RULE 4.1 — Outstation Lorries Cannot Share Routes**
Lorries assigned to outstation destinations (Kuantan, Pahang, Seremban, Tanjung Malim, Rawang outstation, Kemaman, Port Dickson, and all LARGE_LONG/MEDIUM_LONG clusters) must NOT share their lorry with items from a different outstation destination UNLESS the routes travel in the same geographic direction (bearing difference ≤ 80°) AND are within 60 km of each other.

**RULE 4.2 — Direction Compatibility**
Two outstation routes may share a lorry only if:
  - Both are in the same route corridor group (see Rule 8.1), OR
  - GPS centroids are within 60 km of each other AND bearing from depot differs by ≤ 80°.

**RULE 4.3 — Outstation Hard Cap (1 Trip)**
Outstation lorries make exactly ONE trip. Effective capacity = lorry's rated tonnage (no double-trip multiplier). LORRY NAIK (5% tolerance) still applies for same-route loads.

**RULE 4.4 — Strict Lorry Reservations for Outstation**
  - BQU3875: Pahang (PH) routes ONLY.
  - WA6899M: Pahang (PH) routes ONLY — spare lorry, cannot serve urban.
  - These lorries are forbidden from ALL other route types.

---

## SECTION 5 — DAILY ROUTE SCHEDULE (SCHD Sheet)

**RULE 5.1 — Check SCHD Sheet Before Assignment**
Before every assignment run, load the SCHD sheet for the current user from the LORRY DAILY PLANNING file:
  - Sheet name: SCHD(abi) or SCHD(vivian) depending on logged-in user.
  - The sheet maps each weekday (Monday–Saturday) to which route codes are scheduled.

**RULE 5.2 — Only Assign Today's Scheduled Routes**
Items whose route is NOT scheduled for today (or tomorrow if planning tomorrow's trip) are marked NOT_TODAY and left blank. The user is prompted to confirm if they want to assign NOT_TODAY routes anyway.

**RULE 5.3 — Working Days: Monday to Saturday**
The bot treats Monday through Saturday as working days. Sunday is the only rest day. "Tomorrow" always skips Sunday and lands on the next weekday (Mon–Sat). Saturday is a valid working day and must NOT be skipped.

---

## SECTION 6 — REMARKS-BASED RULES

**RULE 6.1 — Day-Restricted Delivery (REMARKS_SKIP)**
If a DO's REMARKS specify delivery days (e.g., "SELASA DAN JUMAAT", "TUESDAY & FRIDAY DELIVERY", "WEDNESDAY & SATURDAY") and today's weekday is NOT in that list, the item is marked REMARKS_SKIP and NOT assigned. It will not appear in the assignment output.

**RULE 6.2 — Remarks Day Parsing**
Day names accepted in remarks (Malay and English):
  - ISNIN / MONDAY = 0, SELASA / TUESDAY = 1, RABU / WEDNESDAY = 2
  - KHAMIS / THURSDAY = 3, JUMAAT / FRIDAY = 4, SABTU / SATURDAY = 5, AHAD / SUNDAY = 6
Negative patterns (day is OFF/CLOSED) exclude that day from valid delivery days.

**RULE 6.3 — Remarks That Are NOT Day Restrictions (Ignore)**
Do not parse these as delivery-day restrictions:
  - "SETIAP HARI", "DAILY", "EVERY DAY"
  - "NEXT DAY MUST DELIVER", "SAME DAY DELIVERY"
  - "LUNCH TIME", "MORNING TRIP", "AFTERNOON TRIP"
  - "SMALL LORRY", "BIG LORRY", "LARGE LORRY", "LORRY KECIL"
  - "WAKTU OPERASI", "OPERATION HOURS"
  - Time ranges like "9AM - 4PM"

**RULE 6.4 — Special Delivery Requirements in Remarks**
  - "LORRY KECIL SAJA" / "LORRY KECIL" → assign only small lorry (≤5T)
  - "VAN" / "VAN ONLY" → assign VAN-class vehicle only
  - "LORI BESAR TIDAK BOLEH MASUK" → assign small lorry or van, NOT large
  - These remarks override standard weight-based lorry selection.

**RULE 6.5 — REMARKS FIELD Sheet (FIELD 3 — Lorry Tonnage Requirement)**
The "REMARKS FIELD" sheet in LORRY DAILY PLANNING.xlsx defines the canonical
remark phrases planners use, in three columns:
  - **FIELD 1** — delivery days (MON,WED,FRI / EVERY MON / EVERY DAY …) → Rule 6.1/6.2
  - **FIELD 2** — time windows (AM ONLY, PM ONLY, 8AM-5PM, 24 HOURS …) → informational
  - **FIELD 3** — required lorry size → hard tonnage cap on lorry selection

FIELD 3 phrase → maximum lorry tonnage allowed:
  - "VAN" → ≤ 2T (van class only)
  - "BELOW 5 TON" → ≤ 5T
  - "BELOW 10 TON" → ≤ 10T
  - "BELOW 14 TON" → ≤ 14T
  - "BELOW 20 TON" → ≤ 20T
  - "ANY SIZE" → no cap

**RULE 6.6 — REMARKS-Driven Size vs Geographic Grouping**
  - If a DO's REMARKS specify a lorry size (FIELD 3 phrase or a free-text alias
    like "LORRY KECIL", "VAN", "LORI BESAR TIDAK BOLEH"), that DO MUST be assigned
    to a lorry at or below the cap. The cap is enforced in every assignment pass
    (initial, consolidation, force-assign, overflow, fill-to-80%).
  - When a bucket of DOs shares a lorry, the binding cap is the SMALLEST cap among
    the bucket's DOs (DOs without a size remark impose no cap).
  - If REMARKS are EMPTY, no size cap applies and the DO follows normal grouping
    by city + state + nearest longitude (Section 2).
  - The phrase→cap table is loaded from the REMARKS FIELD sheet at runtime when
    present, falling back to the defaults in assignment_config.py.

---

## SECTION 7 — LORRY UTILIZATION RULES

**RULE 7.1 — 80% Utilization Target**
Every lorry that is assigned at least one DO must reach ≥ 80% utilization (based on gross weight vs lorry rated tonnage). After initial assignment, run a fill-to-80% pass: find unassigned items from the same route + city + state bucket, sorted by GPS proximity, and add them to underloaded lorries until ≥ 80% is reached or no more matching items exist.

**RULE 7.2 — LORRY NAIK (5% Overload Tolerance)**
Lorries may carry up to 5% over their rated tonnage (LORRY NAIK). Effective maximum capacity = rated TON × 1.05. Same-route loads may use up to 10% overage (× 1.10).
  - Use NAIK capacity in split threshold checks: if a group fits within TON × 1.05, keep it on ONE lorry — do NOT split.

**RULE 7.3 — Double-Trip for Urban Small/Medium Lorries**
Small and medium lorries (< 11T) on urban (KL/Selangor) routes can make 2 trips per day (morning + afternoon). Their effective capacity = rated TON × 2 for urban route scheduling. Large lorries (≥ 11T) and any lorry on outstation routes make exactly 1 trip (hard cap, no double-trip).

**RULE 7.4 — Minimum Utilization Threshold (Outstation)**
Do not dispatch an outstation lorry with < 10% utilization. Exception: urban routes (always ship regardless of load) and tiny-item routes (e.g. KV11A).

---

## SECTION 8 — ROUTE CORRIDOR GROUPS

**RULE 8.1 — Corridor Group Definition**
Routes within the same corridor group may share a lorry even if they don't perfectly overlap geographically:
  - NS: [NS04, NS05, NS06, NS07, NS08] — Negeri Sembilan corridor
  - PH_INT: [PH01, PH02, PH03, PH04, PH05, PH06, PH07, PH08] — Pahang interior
  - KV_NORTH: [KV01A, KV02A, KV04A] — North KL / Damansara / Rawang
  - KV_EAST: [KV10A, KV11A, KV12A] — East KL / Chow Kit / Pudu / Ampang

**RULE 8.2 — Urban Routes: Corridor Group Only for Cross-Route Sharing**
For KL/Selangor urban routes, a lorry may only accept items from a DIFFERENT route if both routes are in the SAME corridor group. GPS bearing alone is NOT sufficient to mix urban routes. Example: KV08A (Gombak/Setapak) and KV11A (Pudu/Cheras) both point "east" from depot but are NOT in the same corridor group → must use separate lorries.

**RULE 8.3 — Outstation Routes: Corridor Group OR Same-Way Bearing**
For outstation routes, cross-route sharing is allowed if:
  - Both routes are in the same corridor group, OR
  - GPS bearing from depot differs by ≤ 80° AND centroids are within 60 km.

---

## SECTION 9 — LORRY SELECTION (PREFERENCE, NOT HARDCODE)

**RULE 9.1 — No route→lorry hardcode.**
There is NO fixed "route X must use lorry Y" table. `ROUTE_PREFERRED_LORRY` and
`LORRY_STRICT_ROUTE` in assignment_config.py are intentionally EMPTY. A lorry is
chosen purely from:
  1. **Owner** — only the logged-in user's lorries + SPARE are eligible.
  2. **Outstation status** — lorries ≤ 5T are excluded from outstation routes.
  3. **Size cap** — REMARKS / SHIP_DETAIL caps (Section 6, Section 6A).
  4. **Gross weight** — the tightest-fitting eligible lorry that covers the
     cluster's weight wins (maximise utilisation).
  5. **Geography** — same route code + state + city + nearest longitude cluster
     together (Section 2, and the 0.29° GPS single-linkage rule).
The full decision procedure lives in **DO_BOT_SKILL.md** (authoritative).

**RULE 9.2 — Tiny-Item Routes (Small Lorry Mandatory)**
Routes with average DO weight ≤ 150 kg (e.g., KV11A ~46 kg/DO) MUST use small lorries or vans. Lorries ≥ 4.5T are excluded from tiny-item routes because large trucks cannot maneuver in narrow shophouse streets.

**RULE 9A — FIT IN LORRY Default-Lorry List (ABI only, PREFERENCE not restriction)**
The optional "FIT IN LORRY" sheet in LORRY DAILY PLANNING.xlsx lists, per
route, the ordered set of DEFAULT plates that route should try first. This
populates RULE 9.1's `_preferred_lorries_for_route()` hint mechanism from
the planners' sheet instead of a code constant — it is data-driven, not a
hardcoded route→lorry table, and it does NOT exclude any other lorry.
  - **Scope**: cell A1 of the sheet names the owning user (currently "ABI").
    The preference applies only to that user's routes; other users' routes
    are unaffected and keep the RULE 9.1 open selection.
  - **Preference, not restriction — full fallback to the master fleet**: the
    listed plates are tried first (tightest-fitting available one wins,
    same as any other `_preferred_lorries_for_route` hint). If none of a
    route's listed plates are available/free/big-enough that day, ANY other
    lorry owned by that user and marked `Available` in the MUATAN (master
    lorry) sheet remains fully eligible — the existing open weight-based
    selection (RULE 9.1) takes over exactly as if the route had no FIT IN
    LORRY entry at all. A route is never left `NO_LORRY` just because its
    default plates happen to be busy.
  - **Contention priority (closer-to-full-utilisation wins)**: if two
    different routes' lists overlap and, on a given day, both would
    naturally claim the SAME plate as their tightest-fitting preferred
    choice, the route whose gross weight pushes that plate CLOSER to full
    capacity is processed first and claims it; the other route falls
    through to its own next preferred plate, or the open master-fleet
    fallback above. Example: NS05 totalling 14T and PH09 totalling 13T both
    preferring VJN9910 (15.389T) — NS05 wins (91% vs 84% utilisation). This
    priority OVERRIDES RULE 1.1's earliest-date-first ordering for the
    specific pair of groups in contention; all other groups keep their
    normal date order.
  - **Does not affect mixing**: routes that the corridor/geography rules
    (Section 2, Section 8) already merge onto one shared lorry — e.g.
    KV19A + KV20A — are combined into a single group before this rule ever
    runs, so they are never treated as competitors under RULE 9A.

---

## SECTION 6A — SHIP_DETAIL COLUMN

The uploaded DO file may include a `SHIP_DETAIL` column:
`<days>, [AM|PM], MAX <N> TON`.

- **`MAX N TON`** — the largest lorry the customer can receive. **Current code
  enforces only `MAX 2 TON` (→ van, ≤2T).** `MAX 5 / 11 / 15 / 21 TON` are NOT
  capped and follow ordinary weight/route rules. (The intended long-term rule is
  "cap at ≤N and pick the lorry by gross weight" — pending operator sign-off; do
  not change silently.)
- **`OUT SOURCE`** — the DO is delivered by a third-party lorry. Assign NO plate:
  `LICENSE` stays blank and the DO is NOT counted as unassigned (sentinel
  `OUT_SOURCE`).
- **days / AM / PM** — read but not used for assignment filtering.

---

## SECTION 10 — STATE BOUNDARY RULES

**RULE 10.1 — No Cross-State Lorry Assignment**
Once a lorry is committed to a destination state (e.g., PAHANG), it cannot serve items destined for a different state (e.g., TERENGGANU) in the same session.

**RULE 10.2 — KL / Selangor Exception**
KUALA LUMPUR and SELANGOR are treated as compatible states. A lorry serving KL items may also serve Selangor items (and vice versa) within the same session. This applies only to urban KV routes.

---

## SECTION 11 — SPLIT RULES

**RULE 11.1 — Do Not Split If It Fits**
Before splitting a group across multiple lorries, check if the total weight fits within any available lorry's NAIK capacity (TON × 1.05). If yes, keep the group on ONE lorry. Never create two underloaded lorries when one lorry can handle the load.

**RULE 11.2 — Split Threshold: 15 DOs Per Lorry**
If a bucket group has more than 15 DOs AND total weight exceeds all available lorry NAIK capacities, split the group. Always split by DATE (earliest DOs fill the first lorry).

**RULE 11.3 — Bin-Pack for Overflow**
For groups of > 1 unassigned item sharing the same city + state: attempt a bin-pack split across 2 eligible lorries (heaviest items first) before resorting to individual overflow. If bin-pack fails, process individually with ≤ 1T overage tolerance.

---

## SECTION 12 — CROSS-USER AND SCHEDULE FILTERING

**RULE 12.1 — Cross-User Routes (OTHER_USER)**
Each user (ABI, VIVIAN, TRANSPORT) owns specific route prefixes (defined in their ROUTE sheet). Items with a route prefix NOT belonging to the logged-in user are marked OTHER_USER and left blank. Cross-user assignment is NOT allowed.

**RULE 12.2 — Pre-Filled LICENSE Column**
  - Case A (all rows filled): Register plates as assigned today. Show summary. Do NOT re-assign.
  - Case B (some rows blank): Respect pre-filled plates as capacity seeds. Auto-assign only blank rows.

---

## SECTION 13 — SENTINEL VALUES (Skip List)

Items with any of the following statuses are NEVER auto-assigned:
  - OTHER_USER — route belongs to a different user
  - NOT_TODAY — route not on today's SCHD schedule
  - REMARKS_SKIP — REMARKS restrict delivery to days that don't include today
  - NO_LORRY — no eligible lorry found (manual intervention required)
  - SKIPPED — manually skipped by user
  - Blank / None / "nan" / "n/a" / "-"

---

## SECTION 14 — OPTIMIZATION PASSES (Run in Order)

After initial group assignment, run the following passes in sequence:

1. **Lorry-Swap Pass** — Swap large underloaded lorry with small overloaded lorry if: large load fits on small, small is heavier, and large load fills small to ≥ 70%.
2. **Same-Route Merge Pass** — If lorry A's entire load fits on lorry B (same or compatible route), merge A onto B to free a lorry.
3. **Partial-Transfer Rebalance Pass** — Move individual items from well-loaded lorries (≥ 50%) onto underloaded lorries (< 50%) when routes are compatible and source stays above threshold.
4. **Fill-to-80% Pass** — Absorb unassigned items from same route + city + state bucket (sorted by GPS proximity) onto lorries below 80% utilization, up to NAIK capacity (× 1.05).
5. **Hard Capacity Guard** — Trim any lorry exceeding physical capacity: keep nearest-to-depot items, mark overflow as NO_LORRY.

---

## SECTION 15 — GEOGRAPHIC CONSTANTS

  - Depot / HQ: HICOM Shah Alam (3.0340°N, 101.5563°E)
  - Outstation distance threshold: ≥ 50 km from depot
  - Max outstation city merge radius: 60 km
  - Max bearing difference for same-direction merge: 80°
  - LOCAL zone radius: 8 km from depot (fallback to city sub-key if within zone)

### Distance / routing

  - **Customer location**: real GPS pin from the `LONGITUD` column of the upload
    when present; otherwise the place name is geocoded via Nominatim
    (OpenStreetMap, free) and cached in `data/geocode_cache.json`.
  - **Distance between two points**:
      - Grouping / proximity decisions use **haversine** (straight-line) km.
      - The driver trip-manifest **stop ordering** (greedy nearest-neighbour
        from depot) and the **return-time estimate** use **real driving
        distance via OSRM** (free OpenStreetMap routing) when reachable, and
        fall back automatically to haversine if OSRM is unavailable.
  - **Enabling OSRM** (optional — improves stop ordering accuracy):
      - Self-host (recommended, no limits, offline): run an OSRM backend with a
        Malaysia OSM extract, then set `OSRM_URL=http://localhost:5000`.
      - Public demo: leave `OSRM_URL` unset to use
        `https://router.project-osrm.org`. NOTE: in a managed/web environment
        this host must be added to the network egress allowlist, otherwise
        OSRM calls are blocked (403) and the bot silently falls back to
        haversine.
      - Env vars: `OSRM_URL` (base URL), `OSRM_TIMEOUT` (seconds, default 8),
        `OSRM_USE_DEMO=0` to disable the demo fallback entirely.

---

## SECTION 16 — SPECIFIC ROUTE CLASSIFICATIONS

  - KV02A (Batu Beruntung / Serendah / Rawang area): classified as SELANGOR (urban) — any lorry size eligible.
  - KV01A (Rawang / Tanjung Malim toward Perak): classified as MEDIUM_LONG (outstation) — minimum 5.001T.
  - KV05A (Selayang / Batu Caves): SELANGOR urban.
  - KV04A (Sungai Buloh / Kota Damansara): KV_NORTH corridor, SELANGOR urban.
  - PH routes: all outstation LARGE_LONG — Pahang minimum 5.001T, preferred BPE9788 / WA6899M.
  - NS routes: outstation MEDIUM_LONG — Negeri Sembilan minimum 5.001T, preferred BQX9983 / BMN3682.
  - TR02 (Kemaman): outstation LARGE_LONG.
  - NS04, NS06 (Port Dickson, Seremban): outstation MEDIUM_LONG — minimum 5.001T.

---

## END OF RULES FILE
