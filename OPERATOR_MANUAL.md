# DO Bot — Operator & Maintenance Manual

**System:** WhatsApp-based automatic lorry assignment bot for daily delivery operations (Malaysia)
**Repository:** `AlisonOoi123/DO_bot`
**Last updated:** 2026-06-18

---

## TABLE OF CONTENTS

1. [What This Bot Does](#1-what-this-bot-does)
2. [System Architecture](#2-system-architecture)
3. [User Menu — Complete Flow](#3-user-menu--complete-flow)
4. [All WhatsApp Commands](#4-all-whatsapp-commands)
5. [Assignment Rules (Full)](#5-assignment-rules-full)
6. [How Assignment Works Internally](#6-how-assignment-works-internally)
7. [Master Data Files](#7-master-data-files)
8. [External Services](#8-external-services)
9. [Environment Variables & Config](#9-environment-variables--config)
10. [Restart & Service Commands](#10-restart--service-commands)
11. [How to Maintain & Update](#11-how-to-maintain--update)
12. [Troubleshooting](#12-troubleshooting)
13. [Quick Reference Cards](#13-quick-reference-cards)

---

## 1. WHAT THIS BOT DOES

A planner uploads the day's delivery orders (DOs) as an Excel file via WhatsApp. The bot:

1. Reads every DO (route, weight, customer, remarks, GPS)
2. Auto-assigns each DO to the best available lorry using 7 scoring rules
3. Respects day restrictions in REMARKS (e.g. "SELASA & JUMAAT"), lorry size caps, strict route ownership, preferred lorries, destination state boundaries, and load capacity
4. Lets the planner review, change individual assignments, block broken lorries, and re-export
5. Sends back two Excel files: the filled DO sheet + a stop-ordered trip manifest per lorry

The bot maintains a **daily log** (`data/daily_assignments.json`) that persists across conversations so broken lorries and today's assignments are remembered even if the planner reconnects.

---

## 2. SYSTEM ARCHITECTURE

```
WhatsApp (Meta Cloud API)
        │  HTTPS POST /webhook
        ▼
  app.py  ─────────────────────────────────────────────────────────────────
  Flask webhook server                                                      │
  • Receives messages (text / interactive buttons / Excel files)            │
  • Calls bot.py handler for all logic                                      │
  • Sends replies back to WhatsApp via Meta REST API                        │
        │                                                                   │
        ▼                                                                   │
  bot.py                                                                    │
  State machine per phone number (dict "sessions")                          │
  • IDLE → AWAIT_USER_ID → AWAIT_TRIP_DAY → AWAIT_EXCEL                   │
  • → REVIEWING → CONFIRMING → DONE                                         │
  • Calls lorry_engine.py for suggestions                                   │
  • Calls Anthropic / Local LLM for ambiguous remarks parsing              │
        │                                                                   │
        ▼                                                                   │
  lorry_engine.py                                                           │
  LorryEngine class                                                         │
  • Loads lorry master data from LORRY DAILY PLANNING.xlsx                 │
  • Loads history from ZSDOROUTEWRH.xlsx                                   │
  • suggest() / suggest_split() — scored lorry recommendation              │
  • Geocoding via built-in _MY_COORDS → Nominatim (OSM)                   │
  • Road distances via OSRM → haversine fallback                           │
        │
        ▼
  assignment_config.py
  Single source of truth for ALL constants, thresholds, route maps,
  preferred lorry lists, corridor groups, strict reservations.
```

**Key files:**

| File | Role |
|------|------|
| `app.py` | Flask server, webhook, WhatsApp send helpers |
| `bot.py` | Full session state machine, all user flows |
| `lorry_engine.py` | Assignment scoring engine, geocoding, OSRM |
| `assignment_config.py` | All constants & business config |
| `ASSIGNMENT_RULES.md` | Business rules reference (human-readable) |
| `data/LORRY DAILY PLANNING.xlsx` | Lorry capacities, schedules, routes |
| `data/ZSDOROUTEWRH.xlsx` | Assignment history (GPS, customer, weights) |
| `data/daily_assignments.json` | Today's assigned & broken lorries (auto-created) |
| `data/geocode_cache.json` | OSM geocoding cache (auto-created) |
| `data/remarks_day_cache.json` | LLM remarks parsing cache (auto-created) |
| `config.txt` | Dev credentials (never commit; gitignored) |

---

## 3. USER MENU — COMPLETE FLOW

### Step 1 — Start Session

User sends any of: `hi`, `hello`, `start`

Bot replies with the login menu:
```
👋 Lorry Assignment Bot
Please tap your name below or type it to continue.

[ABI]  [VIVIAN]  [Other...]
```

### Step 2 — Pick Trip Day

After selecting user:
```
📅 Which day's DOs are you planning now?

[Today (Mon)]  [Tomorrow (Tue)]
[Clear daily log]
```

- **Today / Tomorrow** — Sets the planning date; routes scheduled for that weekday are expected
- **Clear daily log** — Wipes all assigned lorries for today (useful if starting fresh after midnight)

### Step 3 — Upload DO File

```
📎 Please upload your DO Excel file (.xlsx) now.

(Or upload a lorry status file to bulk block/release lorries)
```

The bot accepts:
- **DO file** — must have columns: `DO NUMBER`, `ROUTE`, `CUSTOMER NAME`, `GROSS WEIGHT` (kg) or `WEIGHT(T)` (tonnes)
- **Lorry status file** — must have columns: `LORRY` + `STATUS` (Blocked / Available) to bulk set availability

### Step 4 — Review Assignments

Bot auto-assigns all DOs and shows a summary:
```
✅ Assignment Complete — 47 DOs assigned

🚛 BQY7823 (14.5T) — 12.3T / 85% util
  KV01A × 4 | KV11A × 3 | NS04 × 2

🚛 VEA2818 (5T) — 4.1T / 82% util
  KV11A × 6

⚠️ NO_LORRY — 2 DOs (use "change [DO#]" to reassign)
  1DO6030948, 1DO6030951
```

If any DOs are still `NO_LORRY` after the automatic pass, the bot also asks
right away, without waiting for a `change`/`force` command:
```
🚚 2 DO(s) totalling 8.4T are still unassigned.
Reply with available lorry plate(s) (e.g. VJN9910 BQX9983) to assign them
now — no need to re-upload — or reply SKIP to leave them as is.
```
Reply with one or more plates and the bot bin-packs the still-unassigned
DOs across them (heaviest DO first, tightest-fitting plate wins), still
enforcing the same hard rules (owner fleet, outstation minimum tonnage,
REMARKS/SHIP_DETAIL size cap, state compatibility) — nothing gets forced
onto a plate that genuinely can't take it. Whatever doesn't fit stays
`NO_LORRY` and the bot asks again for a different plate, or `SKIP` to stop.
This re-uses the same DOs already in the session, so no re-upload is
needed.

Planner can also still handle it manually:
- `change 1DO6030948` → Pick a different lorry
- `block WUD4927` → Mark lorry unavailable
- `force BQY7823 1DO6030948` → Override assignment
- `yes` → Confirm & export

### Step 5 — Export

Bot sends two files:
1. **DO_Assigned_NewRows.xlsx** — The original DO sheet with `LICENSE` column filled in
2. **Trip_Manifest_DDMMYYYY.xlsx** — One sheet per lorry, stops sorted by nearest-neighbour GPS route from depot, with estimated return time

After export, the planner can still:
- `block [PLATE]` → Blocks lorry and re-exports automatically
- `change [DO#]` → Reassigns one DO and re-exports
- `release [PLATE]` → Unblocks a lorry

Send `hi` to start a new session. Daily log persists.

---

## 4. ALL WHATSAPP COMMANDS

### Global (any state)

| Command | Example | What it does |
|---------|---------|--------------|
| `hi` / `hello` / `start` | `hi` | Start or restart session |
| `reset` / `restart` / `start over` | `reset` | Clear current session state |
| `manage lorry` | `manage lorry` | Open lorry maintenance submenu |
| `manage [PLATE]` | `manage BQY7823` | Jump to actions for one lorry |
| `show assigned today` | `show assigned` | List all lorries assigned today |
| `show blocked` | `blocked list` | List all blocked/unavailable lorries |
| `broken list` | `broken list` | List all broken lorries + replacements |

### Lorry Management Submenu

Enter via `manage lorry`, then tap a button:

| Button | What it does |
|--------|-------------|
| **Block a Lorry** | Mark one lorry unavailable for today (not broken, just not using) |
| **Log Breakdown** | Mark as broken + pick replacement lorry |
| **Release Blocked** | Multi-select lorries to unblock |
| **Mark Fixed** | Multi-select broken lorries to restore |

Lorry picker shows:
- ✅ Available
- ⛔ Assigned / blocked today
- 🔧 Broken → [Replacement Plate]
- Paginated at 9 per page (WhatsApp limit)

### Broken Lorry Commands (shortcut)

| Command | Example | What it does |
|---------|---------|--------------|
| `broken [PLATE]` | `broken VJN9910` | Log breakdown, bot asks for replacement |
| `broken [PLATE] [REPLACE]` | `broken VJN9910 BQU3875` | Log + set replacement in one step |
| `fixed [PLATE]` | `fixed VJN9910` | Remove from broken list |
| `release [PLATE]...` | `release VJN9910 BQU3875` | Unblock multiple lorries |

### During Review / After Export

| Command | Example | What it does |
|---------|---------|--------------|
| `change [DO#]` | `change 1DO6030948` | Pick different lorry for this DO |
| `block [PLATE]` | `block WUD4927` | Mark lorry unavailable, re-assign its DOs |
| `force [PLATE] [DO#]` | `force BQY7823 1DO6030948` | Force-assign without validation |
| `custom [DO#]` | `custom 1DO6030948` | Manually specify weight split across lorries |
| `yes` | `yes` | Confirm & export Excel files |

---

## 5. ASSIGNMENT RULES (FULL)

All rules are enforced in order. Earlier rules are never overridden by later ones.

### Rule 0 — Owner Isolation

When ABI logs in, only ABI + SPARE lorries are eligible. When VIVIAN logs in, only VIVIAN + SPARE. Cross-owner borrowing is architecturally impossible — the engine's `eligible_lorries` DataFrame is filtered at login time.

### Rule 1 — Preferred Lorry (Owner-First)

Each route has a preferred lorry list defined in `assignment_config.py → ROUTE_PREFERRED_LORRY`. The matching uses the **longest route prefix** (e.g. `KV24A` wins over `KV24`).

**Behaviour:** If ANY preferred lorry has capacity and is not excluded, it is used **unconditionally**. Only if ALL preferred lorries are full, unavailable, wrong state, or size-capped by REMARKS does the bot fall through to open fleet assignment.

Key examples:

| Route(s) | Preferred Lorries (in order) |
|----------|------------------------------|
| KV01A | BQY7823, BMN3682, VER2872 |
| KV06A, KV07A (tight streets) | W3826C, BQX7228, W3618U |
| KV11A (Pudu/Ampang shophouses) | VEA2818, VKN8836, W3826C |
| PH01–PH07 (Pahang) | BPE9788, BQX9983 |
| NS04–NS08 (N. Sembilan) | BQX9983, BMN3682 |
| JH01 (Johor) | VJN8929, VJA7981, VNL6819 |

### Rule 2 — Strict Lorry Reservations

Some lorries are **permanently forbidden** from all routes except their designated state:

| Plate | Restriction | Reason |
|-------|-------------|--------|
| BQU3875 (21T) | Pahang (PH) routes ONLY | Contractual |
| WA6899M (13T) | Pahang (PH) routes ONLY | Spare lorry, Pahang-only |

If a strict lorry would be assigned to any other route, it is hard-excluded before any suggestion is made.

### Rule 3 — Destination Size Minimums

| Destination Type | Routes | Min Lorry Size |
|-----------------|--------|----------------|
| LARGE_LONG (outstation) | PH, TR, KB, JH, PK, KD, PN, MC, SB, SR | > 5T (outstation min) |
| MEDIUM_LONG (outstation) | NS, KV01A | > 5T |
| KL / SELANGOR (urban) | All KV routes, etc. | Any size |

Preferred lorries **bypass** this minimum because they are operationally designated (e.g. BMN3682 at 8.66T handles NS04/NS05 even though the general MEDIUM_LONG min is 5T).

### Rule 4 — REMARKS Day Restrictions

The REMARKS column is parsed for delivery-day restrictions:

| Remark (Malay/English) | Parsed as (weekdays, 0=Mon) |
|------------------------|----------------------------|
| `SELASA & JUMAAT` | Tuesday (1), Friday (4) |
| `SETIAP HARI` / `EVERY DAY` | No restriction |
| `TUTUP ISNIN` | Closed Monday → Tue–Sat only |
| `ISNIN SAHAJA` | Monday (0) only |
| `NEXT DAY DELIVERY` | No day restriction (delivery timing only) |

DOs with day restrictions incompatible with the planning date are marked `NOT_TODAY` and excluded from assignment automatically.

**Lorry size caps from REMARKS (FIELD 3):**

| Remark | Max lorry (T) |
|--------|--------------|
| `VAN` | 2.0 |
| `LORRY KECIL` / `BELOW 5 TON` / `LORI BESAR TIDAK BOLEH` | 5.0 |
| `BELOW 10 TON` | 10.0 |
| `BELOW 14 TON` | 14.0 |
| `BELOW 20 TON` | 20.0 |
| `ANY SIZE` | No cap |

Parsing priority: canonical `REMARKS FIELD` sheet → built-in aliases → regex patterns (`BELOW \d+ TON`, `MAX \d+T`, `\d+T KE BAWAH`) → LLM fallback for ambiguous text.

### Rule 5 — Capacity Utilisation

- **Target:** ≥ 80% utilisation per lorry (`CAPACITY_TARGET = 0.80`)
- **Minimum:** ≥ 10% for outstation loads (`MIN_UTIL_TO_ASSIGN = 0.10`); below this the DO is marked `NO_LORRY` rather than waste a large truck
- **Overload tolerance:** +5% for multi-route loads; +10% for same-route loads (`NAIK_FACTOR`)
- **Max stops per lorry per day:** 8 (`MAX_STOPS_PER_LORRY`)

### Rule 6 — Tiny-Item Routes

Routes with average DO weight ≤ 0.15T (150 kg) — typically narrow shophouse streets (e.g. KV11A) — must use small lorries (≤ `LORRY_TINY_EXCL_TON` = 4.5T). Large trucks cannot navigate those streets.

### Rule 7 — Same Route / Same City / Same Direction

Items are grouped **route-first** before lorry assignment:

1. Items with the same ROUTE code + same STATE + same CITY are collapsed into one group for one lorry
2. Outstation items with the same ROUTE code are further split by **GPS bearing octant** (N / NE / E / SE / S / SW / W / NW from depot) — this prevents items pointing in opposite directions from being forced onto the same lorry
3. Corridor groups allow compatible routes to share a lorry (e.g. NS04 + NS05 can both ride BQX9983 if weight fits)

### Rule 8 — Scoring for Open Fleet Assignment

When no preferred lorry is available, `engine.suggest()` ranks eligible lorries by this cascade (each tiebreaker applied in order):

| Priority | Factor | Meaning |
|----------|--------|---------|
| 1 | Customer+Route history | How often this lorry served this exact customer on this route |
| 2 | Cluster familiarity | How often this lorry served this geographic cluster |
| 3 | Utilisation ≥ 80% | Prefer lorries that will be well-loaded |
| 4 | Smallest surplus | Tightest fit (least wasted capacity) |
| 5 | Owner before SPARE | Owner lorries preferred over shared pool |
| 6 | Route history | General frequency on this route |

### Rule 9 — State Boundary

A lorry already committed to `SELANGOR` deliveries in this session will not be assigned `PAHANG` deliveries (and vice versa). KL and Selangor are treated as compatible (allowed to mix).

### Rule 10 — Stop Ordering in Trip Manifest

Stops within each lorry's manifest are sorted using a **greedy nearest-neighbour chain** starting from the depot (HICOM Shah Alam). When OSRM (road routing) is reachable, real driving distances are used; otherwise straight-line haversine. This ensures drivers travel without zig-zagging.

---

## 6. HOW ASSIGNMENT WORKS INTERNALLY

### Phase 1 — Parse & Normalise

For each row in the uploaded Excel:

1. Extract DO NUMBER, ROUTE, CUSTOMER NAME, WEIGHT
   - New format: `GROSS WEIGHT` column (in kg) → divide by 1000 → tonnes
   - Old format: `WEIGHT(T)` column (already in tonnes)
2. Resolve STATE (tries: STATE column → CITY lookup → POSTCODE range)
3. Parse REMARKS for day restrictions and lorry size cap
4. Mark as `NOT_TODAY` if day restriction excludes planning date
5. Mark as `OTHER_USER` if route belongs to a different owner
6. If `LICENSE` column already filled → register as pre-assigned, skip

### Phase 2 — Group by Route + City + Direction

```
Items
  → bucket by ROUTE + STATE + CITY
  → outstation items: further split by GPS bearing octant
  → same-route prefix: collapse sub-buckets
  → corridor merge: combine compatible routes onto one lorry if weight fits
  → capacity sub-divide: if combined weight > max lorry, bin-pack into smaller groups
  → cross-cluster merge: pull geographically close groups from different corridors
```

### Phase 3 — Assign Each Group

For each group (in priority order — heavier groups first):

```
1. Preferred lorry available?
   YES → assign unconditionally (owner-first rule)
   NO  ↓
2. Lorry already assigned in this session with room left?
   YES → assign (reuses partially-loaded lorry)
   NO  ↓
3. engine.suggest() — open fleet, scored by history + fit
   Got suggestion → check utilisation ≥ 10%
   YES → assign
   NO  → mark NO_LORRY
```

### Phase 4 — Post-Assignment Passes

After all groups are assigned, 6 cleanup passes run in order:

| Pass | Purpose |
|------|---------|
| **Consolidation** | Give NO_LORRY items a 2nd chance on partially-full lorries |
| **Force-assign** | Must-ship items: relax daily-log exclusions, try any eligible lorry |
| **Same-route merge** | Move an entire small lorry's load onto a larger lorry if it fits |
| **Partial-transfer rebalance** | Move individual items from >50% full lorries to <50% lorries |
| **Fill-to-80%** | Absorb remaining unassigned items into lorries below 80% utilisation |
| **Hard capacity guard** | Trim any lorry over physical capacity; keep nearest-to-depot items |

### Phase 5 — Export

1. Build output DataFrame with `LICENSE` column filled
2. Generate trip manifest Excel: one sheet per lorry, stops in NN-sorted GPS order
3. Return-time estimate per lorry (OSRM road km when available, else haversine × 1.3 road factor)
4. Send both files back via WhatsApp

---

## 7. MASTER DATA FILES

### 7.1 `data/LORRY DAILY PLANNING.xlsx`

**Required sheets:**

#### Sheet: `MUATAN` — Lorry Capacities

| Column | Meaning | Example |
|--------|---------|---------|
| NAME / Col 0 | Section label (ABI / VIVIAN / SPARE) or plate | ABI |
| LORRY / Col 1 | Lorry plate | BQY7823 |
| BDM / Col 2 | Tare weight (kg) | 4200 |
| BTM / Col 3 | Body+tare (kg) | 5100 |
| MUATAN / Col 4 | Net payload (kg) | 14500 |
| LORRY NAIK (5%) / Col 5 | **This is the operative capacity used by bot** | 14500 |

Bot reads `LORRY NAIK (5%)` as the lorry's capacity in kg, converts to tonnes.

#### Sheet: `SCHD(abi)` / `SCHD(vivian)` — Route Schedules

```
Header row: MON  TUES  WED  THUS  FRI  SAT
Data rows:  KV01A  KV11A  PH05  ...
```

Bot checks planned routes against today's schedule to flag unexpected routes.

#### Sheet: `ABI ROUTE` / `VIVIAN ROUTE` — Route Ownership

List of route codes owned by each user. DOs on the wrong user's routes are tagged `OTHER_USER`.

#### Sheet: `Malaysia States & Cities` — Geographic Reference

```
Col 0: STATE    (e.g. SELANGOR)
Col 1: CITY     (e.g. KAJANG)
```

Used to map CITY → STATE when the STATE column is blank.

#### Sheet: `REMARKS FIELD` — Canonical Remarks (optional)

```
FIELD 1: day restriction   (SELASA & JUMAAT, SETIAP HARI, ...)
FIELD 2: time window       (AM ONLY, 8AM-5PM, ...)
FIELD 3: lorry size cap    (VAN, BELOW 5 TON, ANY SIZE, ...)
```

Adding rows here teaches the bot new canonical remark phrases without code changes.

### 7.2 `data/ZSDOROUTEWRH.xlsx` — Assignment History

**Key columns used for scoring:**

| Column | Used for |
|--------|---------|
| DO NUMBER | Tracking |
| ROUTE | Route frequency learning |
| CUSTOMER NAME | Customer frequency learning |
| LICENSE | Which lorry served this DO |
| LONGITUD | GPS `"lat lon"` → route centroid per route code |
| STATE / CITY | Geographic reference |
| GROSS WEIGHT | Weight learning |

The bot rebuilds frequency tables every session from this file. The more history it has, the smarter the suggestions.

### 7.3 Auto-Created Data Files

#### `data/daily_assignments.json`

```json
{
  "date": "2026-06-18",
  "assigned": ["BQY7823", "VJN9910", "WUD4927"],
  "broken": {
    "VEA2818": "BQX7228",
    "VJN8929": "NONE"
  }
}
```

- **assigned**: Plates that have been used today; excluded from new sessions until released or midnight
- **broken**: Maps broken plate → replacement plate (`"NONE"` = blocked without replacement)
- Resets automatically at midnight (UTC). **Do not delete manually during the working day.**

#### `data/geocode_cache.json`

Stores `{ "PLACE NAME": [lat, lon] }` for every OSM geocoding lookup. Grows over time — safe to delete if corrupted (will rebuild on next run, one API call per unknown place name).

#### `data/remarks_day_cache.json`

Stores `{ "REMARK TEXT": [weekday_ints] }` for every LLM parsing result. Safe to delete if results are wrong (will re-parse on next run).

---

## 8. EXTERNAL SERVICES

| Service | URL | Auth | Cost | Fallback |
|---------|-----|------|------|----------|
| **WhatsApp Meta Cloud API** | `https://graph.facebook.com/v19.0/...` | Bearer token (`META_ACCESS_TOKEN`) | Pay-per-message (very low) | None — required |
| **Anthropic Claude** | `https://api.anthropic.com` | `ANTHROPIC_API_KEY` | Pay-per-token (low usage) | Local LLM → regex |
| **Local LLM (LM Studio)** | `http://localhost:1234/v1` | None | Free | Regex fallback |
| **Nominatim (OSM geocoding)** | `https://nominatim.openstreetmap.org` | None (rate-limited 1 req/sec) | Free | Built-in `_MY_COORDS` dict |
| **OSRM (road routing)** | `https://router.project-osrm.org` (demo) or self-hosted | None | Free | Haversine straight-line |

### Setting Up OSRM Self-Hosted (Recommended)

The demo OSRM server may be blocked by network egress policies. Self-hosting gives unlimited, offline-capable road routing:

```bash
# 1. Download Malaysia OSM extract (~300 MB)
wget https://download.geofabrik.de/asia/malaysia-singapore-brunei-latest.osm.pbf

# 2. Run OSRM backend in Docker
docker run -t -v "${PWD}:/data" ghcr.io/project-osrm/osrm-backend \
    osrm-extract -p /opt/car.lua /data/malaysia-singapore-brunei-latest.osm.pbf
docker run -t -v "${PWD}:/data" ghcr.io/project-osrm/osrm-backend \
    osrm-partition /data/malaysia-singapore-brunei-latest.osrm
docker run -t -v "${PWD}:/data" ghcr.io/project-osrm/osrm-backend \
    osrm-customize /data/malaysia-singapore-brunei-latest.osrm

# 3. Start the routing server
docker run -d -p 5000:5000 -v "${PWD}:/data" ghcr.io/project-osrm/osrm-backend \
    osrm-routed --algorithm mld /data/malaysia-singapore-brunei-latest.osrm

# 4. Set environment variable for the bot
set OSRM_URL=http://localhost:5000
```

---

## 9. ENVIRONMENT VARIABLES & CONFIG

### Production Environment Variables

Set in Windows System Environment or NSSM service config:

```
META_ACCESS_TOKEN        (required) WhatsApp API Bearer token
META_PHONE_NUMBER_ID     (required) WhatsApp Business phone number ID
META_VERIFY_TOKEN        (optional, default: eslorrybot2026) Webhook verification token
PUBLIC_BASE_URL          (required) Your public HTTPS URL (e.g. https://yourdomain.com)
PORT                     (optional, default: 5000) Server listen port

ANTHROPIC_API_KEY        (optional) Claude API key for remarks parsing
LOCAL_LLM_URL            (optional, default: http://localhost:1234/v1) LM Studio endpoint
LOCAL_LLM_MODEL          (optional, default: local-model) Local model name

OSRM_URL                 (optional) OSRM routing server URL
OSRM_TIMEOUT             (optional, default: 8) Seconds before OSRM request times out
OSRM_USE_DEMO            (optional) Set to "0" to disable public OSRM demo entirely

PYTHONUTF8               Set to "1" to avoid encoding errors on Windows
```

### Development (`config.txt`)

For local development, create `config.txt` in the project root (this file is gitignored — never commit it):

```ini
META_ACCESS_TOKEN=EAABsQ7XX...
META_PHONE_NUMBER_ID=123456789012345
META_VERIFY_TOKEN=eslorrybot2026
PUBLIC_BASE_URL=https://xxxx.ngrok.io
PORT=5000
ANTHROPIC_API_KEY=sk-ant-...
```

---

## 10. RESTART & SERVICE COMMANDS

### ⚡ Standard Restart (after any code update)

Run this in PowerShell **as Administrator** on the production Windows server:

```powershell
cd "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
Stop-Service ngrok_do_bot; Restart-Service do_bot; Start-Service ngrok_do_bot
```

**What this does:**
1. `Stop-Service ngrok_do_bot` — Stops the ngrok tunnel first (prevents stale tunnel)
2. `Restart-Service do_bot` — Restarts the Flask/Python bot server with new code
3. `Start-Service ngrok_do_bot` — Restarts the ngrok tunnel to re-expose the server

**Why this order matters:** If you restart `do_bot` while ngrok is running, the tunnel briefly disconnects, which can cause Meta to stop delivering webhooks. Stopping ngrok first prevents this.

### Service Status Check

```powershell
Get-Service do_bot, ngrok_do_bot
```

Expected output:
```
Status   Name            DisplayName
------   ----            -----------
Running  do_bot          DO Bot Lorry Assignment
Running  ngrok_do_bot    ngrok Tunnel
```

### View Live Logs

```powershell
# Bot server logs (adjust path to your NSSM log config):
Get-Content "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot\logs\stdout.log" -Tail 50 -Wait

# Or if using NSSM:
nssm status do_bot
```

### Emergency — Stop Everything

```powershell
Stop-Service ngrok_do_bot
Stop-Service do_bot
```

### Pull Latest Code and Restart

```powershell
cd "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
git pull origin main
Stop-Service ngrok_do_bot; Restart-Service do_bot; Start-Service ngrok_do_bot
```

---

## 11. HOW TO MAINTAIN & UPDATE

### 11.1 Adding a New Lorry

1. Open `data/LORRY DAILY PLANNING.xlsx`
2. In the `MUATAN` sheet, add a row under the correct owner section (ABI / VIVIAN / SPARE):
   ```
   [NAME]  [PLATE]  [BDM kg]  [BTM kg]  [MUATAN kg]  [LORRY NAIK (5%) kg]
   ABI     VXX1234  4000       5000       14000         14000
   ```
3. The `LORRY NAIK (5%)` value is what the bot uses as capacity — set it carefully
4. Save & restart the service

**No code changes needed.** The bot reads lorry list from the Excel file at every session start.

### 11.2 Adding a New Route

1. Open `assignment_config.py`

2. If the new route belongs to an existing **cluster**, add it to `CLUSTER_MAP`:
   ```python
   CLUSTER_MAP = {
       ...
       "KV99A": "KL_VALLEY",   # ← add your route here
       ...
   }
   ```

3. If it belongs to an existing **corridor**, add it to `CORRIDOR_MAP`:
   ```python
   CORRIDOR_MAP = {
       ...
       "KV99A": "KV_EAST",
       ...
   }
   ```

4. If it should share a lorry with other routes (e.g. KV99A + KV99B same lorry), add a corridor group in `ROUTE_CORRIDOR_GROUPS`:
   ```python
   ROUTE_CORRIDOR_GROUPS = [
       ...
       {"KV99A", "KV99B"},    # ← these can share a lorry
   ]
   ```

5. If the route has a **preferred lorry**, add to `ROUTE_PREFERRED_LORRY`:
   ```python
   ROUTE_PREFERRED_LORRY = {
       ...
       "KV99A": ["BQY7823", "VEA2818"],  # longest-prefix match, first tried first
   }
   ```

6. If the route must be classified as outstation, verify it appears in `_classify_dest_group()` in `lorry_engine.py` or update `assignment_config.py → OUTSTATION_ROUTES`.

7. Save & restart the service.

### 11.3 Changing a Preferred Lorry Assignment

Edit `assignment_config.py → ROUTE_PREFERRED_LORRY`:

```python
"KV11A": ["VEA2818", "VKN8836", "W3826C"],
# Change to:
"KV11A": ["NEW_PLATE", "VEA2818", "VKN8836"],
```

Save & restart.

### 11.4 Blocking a Route to a Specific Lorry (Strict Reservation)

To permanently ban a lorry from a route type, edit `assignment_config.py → LORRY_STRICT_ROUTE`:

```python
LORRY_STRICT_ROUTE = {
    "BQU3875": {"PH"},      # Pahang only
    "WA6899M": {"PH"},      # Pahang only
    "NEW_PLATE": {"NS"},    # ← add: Negeri Sembilan only
}
```

Save & restart.

### 11.5 Adding/Changing a REMARKS Lorry Size Cap

**Option A (no code change):** Add a row to the `REMARKS FIELD` sheet in `LORRY DAILY PLANNING.xlsx`:
```
FIELD 3: BELOW 8 TON
```
The bot loads canonical phrases from this sheet at startup.

**Option B (code):** Edit `assignment_config.py → REMARKS_FIELD3_TON_CAP`:
```python
REMARKS_FIELD3_TON_CAP = {
    ...
    "BELOW 8 TON": 8.0,    # ← add new phrase
    "8T KE BAWAH": 8.0,    # ← Malay alias
}
```

Save & restart.

### 11.6 Updating Lorry Capacity in the Bot

Edit the `LORRY NAIK (5%)` value in `MUATAN` sheet of `LORRY DAILY PLANNING.xlsx`. Bot re-reads at each session start. No restart needed.

### 11.7 Adding a New User (Owner)

1. Add a new section to `MUATAN` sheet with the user's lorries (section header row = user name)
2. Create a new schedule sheet `SCHD(username)` following the same day-column format
3. Create a new route ownership sheet `USERNAME ROUTE`
4. In `lorry_engine.py`, the `LorryEngine` class reads `owner_user` and filters `eligible_lorries` by `USER` column — ensure your new section header matches exactly
5. Add the new user to the login menu in `bot.py` (search for `AWAIT_USER_ID` state)

Save & restart.

### 11.8 Updating WhatsApp Token

Meta access tokens expire periodically.

1. Get new token from Meta Business Suite → WhatsApp → API Setup
2. Update `META_ACCESS_TOKEN` in Windows System Environment (or NSSM service env)
3. Restart the service:
   ```powershell
   Stop-Service ngrok_do_bot; Restart-Service do_bot; Start-Service ngrok_do_bot
   ```

### 11.9 After Any Code Change

```powershell
cd "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot"
git pull origin main
Stop-Service ngrok_do_bot; Restart-Service do_bot; Start-Service ngrok_do_bot
```

Wait 5–10 seconds, then send `hi` to the bot from WhatsApp to verify it responds.

### 11.10 Backing Up Data

Back up these files regularly (daily recommended):

```
data/LORRY DAILY PLANNING.xlsx     ← lorry master (update here for permanent changes)
data/ZSDOROUTEWRH.xlsx             ← history (grows over time, improves accuracy)
data/daily_assignments.json        ← today's state (can regenerate, but useful for audit)
data/geocode_cache.json            ← geocoding cache (safe to delete, rebuilds slowly)
data/remarks_day_cache.json        ← LLM remarks cache (safe to delete, rebuilds on demand)
assignment_config.py               ← all business rules (version controlled in git)
```

---

## 12. TROUBLESHOOTING

### Bot not responding to WhatsApp

1. Check services are running:
   ```powershell
   Get-Service do_bot, ngrok_do_bot
   ```
2. Check bot logs for errors
3. Verify Meta access token hasn't expired (tokens typically expire every 60 days)
4. Verify ngrok tunnel URL matches `PUBLIC_BASE_URL` in environment
5. Restart:
   ```powershell
   Stop-Service ngrok_do_bot; Restart-Service do_bot; Start-Service ngrok_do_bot
   ```

### "Upload failed" or Excel not accepted

- File must be `.xlsx` (not `.xls` or `.csv`)
- File must be under ~20 MB (Meta API limit)
- Required columns: `DO NUMBER`, `ROUTE`, `CUSTOMER NAME`, and either `GROSS WEIGHT` (kg) or `WEIGHT(T)` (tonnes)
- Do not password-protect the Excel file

### Wrong lorry assigned

1. Check if a preferred lorry was full/unavailable — run `show assigned today`
2. Check if the route has a preferred lorry in `assignment_config.py → ROUTE_PREFERRED_LORRY`
3. Check if history file has outdated data pointing to a different lorry
4. Use `change [DO#]` during review to manually correct
5. For permanent fix: update `ROUTE_PREFERRED_LORRY` in `assignment_config.py` and restart

### DO marked NOT_TODAY

REMARKS contains a day restriction that excludes today. Check the REMARKS column — it may say "SELASA SAHAJA" (Tuesday only) and today is Monday. This is correct behaviour. If the remark is wrong, edit it in the source file.

### DO marked NO_LORRY

- Weight too heavy for all available lorries → check if a lorry is broken/blocked that shouldn't be
- REMARKS size cap too restrictive (e.g. "VAN" on a 3T item)
- No lorry available with correct state compatibility
- Use `change [DO#]` to manually assign, or `force [PLATE] [DO#]` to override

### Remarks parsed wrong / wrong day restriction

1. Check `data/remarks_day_cache.json` — find the remark text, delete its entry
2. Restart the bot (cache reloads)
3. Re-upload the file — the bot will re-parse via LLM
4. If still wrong, add the phrase to `REMARKS FIELD` sheet in the Excel master file (canonical override)

### OSRM stop ordering not using real roads

The bot falls back to straight-line haversine if OSRM is unreachable. Check:
1. Is `OSRM_URL` set in the environment?
2. Is the OSRM server running? (`docker ps` on the OSRM host)
3. Is the OSRM host accessible from the bot server? (check network/firewall)
4. Check logs for OSRM timeout messages

### Session stuck in wrong state

Send `reset` or `start over` from WhatsApp. If that doesn't work, restart the service (all sessions are in-memory and reset on restart).

### Daily log not resetting

The daily log resets at midnight UTC. Malaysia time is UTC+8, so 8:00 AM Malaysia time. If the log is stuck:
```powershell
# Manually clear: rename the file and restart
cd "C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot\data"
del daily_assignments.json
```
Then restart the service.

---

## 13. QUICK REFERENCE CARDS

### Business Constants (from `assignment_config.py`)

| Constant | Value | What it controls |
|----------|-------|-----------------|
| `CAPACITY_TARGET` | 0.80 | Target 80% lorry utilisation |
| `MIN_UTIL_TO_ASSIGN` | 0.10 | Skip assignment if load < 10% of lorry capacity |
| `MAX_STOPS_PER_LORRY` | 8 | Max deliveries per lorry per day |
| `NAIK_FACTOR` | 1.05 | Allow 5% overload (multi-route) |
| `SAME_ROUTE_NAIK` | 1.10 | Allow 10% overload (same-route) |
| `FILL_TARGET` | 0.80 | Fill-to-80% pass target |
| `TINY_ITEM_AVG_WEIGHT_T` | 0.15 | Avg DO weight threshold for narrow-street restriction |
| `OUTSTATION_MIN_TON` | 5.001 | Minimum lorry size for outstation routes |
| `OUTSTATION_DIST_KM` | 50.0 | Distance from depot to classify as outstation |
| `DEPOT_LAT / LON` | 3.0340, 101.5563 | HICOM Shah Alam depot coordinates |
| `MAX_CROSS_CLUSTER_KM` | 180.0 | Max GPS distance for route merging |
| `MAX_CROSS_CLUSTER_BEARING` | 80.0 | Max compass bearing difference for merging |
| `LOCAL_ZONE_KM` | 8.0 | Radius around depot treated as "local" |
| `OSRM_MAX_FAILS` | 3 | Stop trying OSRM after 3 consecutive failures |

### Port & Protocol Summary

| Service | Port | Notes |
|---------|------|-------|
| DO Bot Flask server | 5000 | Listens on 0.0.0.0 |
| ngrok tunnel | 443 (public) | Exposes port 5000 to internet |
| OSRM (self-hosted) | 5000 | Set via `OSRM_URL` |
| Local LLM (LM Studio) | 1234 | Set via `LOCAL_LLM_URL` |
| WhatsApp Meta API | 443 | `https://graph.facebook.com/v19.0/` |
| Nominatim | 443 | `https://nominatim.openstreetmap.org` |

### File Locations

```
C:\Users\Administrator\OneDrive\Documents\GitHub\DO_bot\
├── bot.py                        Main logic & session state machine
├── app.py                        Flask server & WhatsApp API helpers
├── lorry_engine.py               Assignment scoring engine
├── assignment_config.py          ⭐ ALL RULES & CONSTANTS — edit here first
├── ASSIGNMENT_RULES.md           Human-readable business rules
├── OPERATOR_MANUAL.md            This file
├── config.txt                    Dev credentials (gitignored — never commit)
└── data\
    ├── LORRY DAILY PLANNING.xlsx  ⭐ Master lorry data — edit here for lorry/route changes
    ├── ZSDOROUTEWRH.xlsx          Assignment history (improves suggestions over time)
    ├── daily_assignments.json     Today's state (auto-reset midnight UTC)
    ├── geocode_cache.json         GPS cache (auto-built, safe to delete)
    └── remarks_day_cache.json     LLM remarks cache (safe to delete)
```

### After Every Code Change — Checklist

```
[ ] git pull origin main
[ ] Review assignment_config.py if rules changed
[ ] PowerShell (Admin):
    Stop-Service ngrok_do_bot; Restart-Service do_bot; Start-Service ngrok_do_bot
[ ] Wait 10 seconds
[ ] Send "hi" to bot from WhatsApp — confirm it replies
[ ] Test with a small DO file upload
[ ] Check logs for errors
```

### Adding Something New — Which File to Edit

| What to add/change | File to edit |
|--------------------|-------------|
| New lorry / capacity | `data/LORRY DAILY PLANNING.xlsx` → `MUATAN` sheet |
| New route preference | `assignment_config.py` → `ROUTE_PREFERRED_LORRY` |
| New corridor grouping | `assignment_config.py` → `ROUTE_CORRIDOR_GROUPS` |
| New cluster mapping | `assignment_config.py` → `CLUSTER_MAP` |
| Strict lorry restriction | `assignment_config.py` → `LORRY_STRICT_ROUTE` |
| New REMARKS size cap | `LORRY DAILY PLANNING.xlsx` → `REMARKS FIELD` sheet OR `assignment_config.py` → `REMARKS_FIELD3_TON_CAP` |
| New city→state mapping | `data/LORRY DAILY PLANNING.xlsx` → `Malaysia States & Cities` sheet |
| New user (owner) | `MUATAN` sheet + `SCHD(user)` sheet + `USER ROUTE` sheet + `bot.py` login menu |
| Threshold change (util %, max stops) | `assignment_config.py` (change the constant) |

---

*Document maintained alongside the codebase. Update this file whenever rules or procedures change.*
