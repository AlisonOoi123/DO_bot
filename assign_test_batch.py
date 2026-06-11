"""
Batch DO Assignment for Test File
==================================
Rules:
  - Long-distance areas (Kuantan, Pahang, Seremban, T.Malim, Rawang, Kemaman, Port Dickson)
    → large lorries (>10.8T); one lorry per area group where possible.
  - KL/Selangor → medium/small lorries (≤10.8T).
  - BQU3875 (21T) is Pahang-route-only.
  - BQY7823 (14.5T) is Rawang/T.Malim direction only.
  - All DOs must be assigned.
"""

import pandas as pd

INPUT_FILE  = '/root/.claude/uploads/77106d3c-bdbc-414f-8cbe-1ec22cc5ce2c/0d4e7a05-4626_test_ABI.xlsx'
OUTPUT_FILE = '/home/user/DO_bot/DO_Assigned_test_ABI.xlsx'

# ── Lorry fleet (MUATAN sheet, ABI + SPARE, LORRY NAIK 5% values) ─────────────
LORRY_FLEET = [
    {'plate': 'BQU3875', 'cap_kg': 21000.0,  'class': 'large'},
    {'plate': 'BQY7823', 'cap_kg': 14542.5,  'class': 'large'},
    {'plate': 'VJN9910', 'cap_kg': 14689.5,  'class': 'large'},
    {'plate': 'VER2872', 'cap_kg': 14311.5,  'class': 'large'},
    {'plate': 'BPE9788', 'cap_kg': 13251.0,  'class': 'large'},
    {'plate': 'WA6899M', 'cap_kg': 13314.0,  'class': 'large'},
    {'plate': 'BQX9983', 'cap_kg': 10762.5,  'class': 'medium'},
    {'plate': 'BMN3682', 'cap_kg':  8662.5,  'class': 'medium'},
    {'plate': 'BQX7228', 'cap_kg':  4200.0,  'class': 'small'},
    {'plate': 'W3618U',  'cap_kg':  4200.0,  'class': 'small'},
    {'plate': 'W3826C',  'cap_kg':  4200.0,  'class': 'small'},
    {'plate': 'WLD8738', 'cap_kg':  4200.0,  'class': 'small'},
    {'plate': 'VEA2818', 'cap_kg':  1113.0,  'class': 'van'},
]
FLEET_MAP = {L['plate']: L for L in LORRY_FLEET}

# Lorries that must NOT appear in last-resort fallback for other areas
LORRY_RESTRICT = {
    'BQU3875': {'KUANTAN', 'PAHANG'},         # Pahang routes only
    'BQY7823': {'RAWANG'},                     # Rawang/T.Malim direction only
}


def classify_route(route: str) -> str:
    r = str(route).upper()
    if 'PH09' in r or 'KUANTAN' in r:
        return 'KUANTAN'
    if r[:4] == 'PH03' or r[:4] == 'PH06' or r[:2] == 'PH':
        return 'PAHANG'
    if 'KV02A' in r or 'RAWANG' in r or 'KV01A' in r or 'T.MALIM' in r:
        return 'RAWANG'
    if 'NS' in r[:3]:
        return 'NS_SOUTH'
    return 'KL_SELANGOR'


# Per-area lorry priority — ordered by preference
AREA_LORRY_PRIORITY = {
    'KUANTAN':     ['BQU3875', 'VJN9910', 'VER2872', 'WA6899M', 'BPE9788'],
    'PAHANG':      ['BPE9788', 'BQU3875', 'VJN9910', 'VER2872', 'WA6899M'],
    'RAWANG':      ['BQY7823', 'BMN3682', 'BQX9983'],
    'NS_SOUTH':    ['BMN3682', 'BQX9983', 'WA6899M', 'VER2872', 'BPE9788'],
    'KL_SELANGOR': ['BQX9983', 'BMN3682', 'BQX7228', 'W3618U', 'W3826C', 'WLD8738', 'VEA2818'],
}

# Preferred lorry per KL sub-route code (best-fit by weight)
KL_ROUTE_LORRY_PREF = {
    'KV20A': 'BQX9983',   # Semenyih/Bangi/Mantin 14.3T (needs 2 trips)
    'KV11A': 'BQX9983',   # Pudu/Cheras 5.2T
    'KV19A': 'BMN3682',   # Kajang 7.4T
    'KV04A': 'BMN3682',   # Sungai Buloh 5.1T
    'KV07A': 'BQX7228',   # Segambut 4.3T
    'KV03A': 'BQX7228',   # K.Selangor 2.5T
    'KV18A': 'W3618U',    # Sungai Besi 2.2T
    'KV06A': 'W3618U',    # Kepong 1.6T
    'KV13A': 'W3826C',    # Serdang 0.8T
    'KV12A': 'W3826C',    # Bangsar 0.7T
    'KV08A': 'WLD8738',   # Gombak 1.0T
    'KV10A': 'WLD8738',   # Chow Kit 0.6T
    'KV09A': 'VEA2818',   # Wangsa Maju 0.14T (van)
}


def assign():
    df = pd.read_excel(INPUT_FILE)
    df['AREA_GROUP'] = df['ROUTE'].apply(classify_route)
    df['LICENSE']    = ''
    df['TRIP']       = pd.NA

    # Track remaining capacity and trip count per lorry
    lorry_remain = {L['plate']: L['cap_kg'] for L in LORRY_FLEET}
    lorry_trips  = {L['plate']: 1           for L in LORRY_FLEET}

    # ── Long-distance areas ────────────────────────────────────────────────────
    for area in ['KUANTAN', 'PAHANG', 'RAWANG', 'NS_SOUTH']:
        mask = (df['AREA_GROUP'] == area) & (df['LICENSE'] == '')
        if not mask.any():
            continue
        allowed_fallback = {
            p for p, areas in LORRY_RESTRICT.items()
            if area in areas
        } | {  # also allow unrestricted lorries
            p for p in FLEET_MAP if p not in LORRY_RESTRICT
        }
        # For long-distance, don't use medium/small lorries unless no alternative
        priority = AREA_LORRY_PRIORITY[area]
        _assign_dos(df, mask, priority, lorry_remain, lorry_trips,
                    fallback_allowed=allowed_fallback)

    # ── KL/Selangor ────────────────────────────────────────────────────────────
    kl_mask = (df['AREA_GROUP'] == 'KL_SELANGOR') & (df['LICENSE'] == '')
    kl_route_kg = (
        df[kl_mask].groupby('ROUTE')['GROSS WEIGHT']
        .sum().sort_values(ascending=False)
    )
    for route in kl_route_kg.index:
        code = route.strip().upper()[:5]
        pref = next((v for k, v in KL_ROUTE_LORRY_PREF.items() if code.startswith(k)), None)
        base = AREA_LORRY_PRIORITY['KL_SELANGOR']
        queue = ([pref] + [p for p in base if p != pref]) if pref else base
        route_mask = (
            (df['AREA_GROUP'] == 'KL_SELANGOR') &
            (df['ROUTE'] == route) &
            (df['LICENSE'] == '')
        )
        kl_fallback = set(FLEET_MAP.keys()) - {'BQU3875', 'BQY7823'}
        _assign_dos(df, route_mask, queue, lorry_remain, lorry_trips,
                    double_trip=True, fallback_allowed=kl_fallback)

    # ── Remaining unassigned (safety net) ──────────────────────────────────────
    unassigned = df['LICENSE'] == ''
    if unassigned.any():
        all_plates = [L['plate'] for L in LORRY_FLEET]
        _assign_dos(df, unassigned, all_plates, lorry_remain, lorry_trips,
                    double_trip=True, ignore_cap=True,
                    fallback_allowed=set(FLEET_MAP.keys()))

    # ── Output ─────────────────────────────────────────────────────────────────
    out_cols = [c for c in df.columns if c not in ('AREA_GROUP',)]
    df[out_cols].to_excel(OUTPUT_FILE, index=False)
    print(f'Saved → {OUTPUT_FILE}')
    _print_summary(df, lorry_remain)


def _assign_dos(df, mask, priority_plates, lorry_remain, lorry_trips,
                double_trip=False, ignore_cap=False, fallback_allowed=None):
    """Assign each DO to the first lorry in priority_plates that has capacity.

    For each DO we scan the entire priority list (not a persistent cursor) so
    every available lorry is considered.  A lorry is skipped only when it truly
    has insufficient remaining capacity for that specific DO.

    double_trip: if True, a KV-route lorry may start a second trip (capacity reset).
    fallback_allowed: set of plates eligible when priority list is exhausted.
    """
    if fallback_allowed is None:
        fallback_allowed = set(FLEET_MAP.keys())

    for idx in df.index[mask].tolist():
        weight = float(df.at[idx, 'GROSS WEIGHT'])
        assigned = False

        # 1. Try priority list first
        for plate in priority_plates:
            remain = lorry_remain[plate]
            if ignore_cap or remain >= weight:
                _do_assign(df, idx, plate, weight, lorry_remain, lorry_trips)
                assigned = True
                break
            elif double_trip and lorry_trips[plate] == 1:
                cap = FLEET_MAP[plate]['cap_kg']
                if ignore_cap or cap >= weight:
                    lorry_trips[plate]  = 2
                    lorry_remain[plate] = cap
                    _do_assign(df, idx, plate, weight, lorry_remain, lorry_trips)
                    assigned = True
                    break

        if assigned:
            continue

        # 2. Fallback: any allowed lorry sorted by remaining capacity desc
        candidates = sorted(
            [p for p in fallback_allowed if p not in priority_plates],
            key=lambda p: lorry_remain[p], reverse=True
        )
        for plate in candidates:
            remain = lorry_remain[plate]
            if ignore_cap or remain >= weight:
                _do_assign(df, idx, plate, weight, lorry_remain, lorry_trips)
                assigned = True
                break

        if not assigned:
            # Absolute last resort — ignore capacity
            best = max(FLEET_MAP.keys(), key=lambda p: lorry_remain[p])
            _do_assign(df, idx, best, weight, lorry_remain, lorry_trips)


def _do_assign(df, idx, plate, weight, lorry_remain, lorry_trips):
    df.at[idx, 'LICENSE'] = plate
    df.at[idx, 'TRIP']    = lorry_trips[plate]
    lorry_remain[plate]   = max(0.0, lorry_remain[plate] - weight)


def _print_summary(df, lorry_remain):
    print('\n=== Assignment Summary ===')
    summary = (
        df.groupby(['LICENSE', 'AREA_GROUP'])
        .agg(DOs=('DO NUMBER', 'count'), KG=('GROSS WEIGHT', 'sum'))
        .reset_index()
    )
    print(summary.to_string(index=False))
    print(f'\nUnassigned DOs: {(df["LICENSE"] == "").sum()}')
    print('\nLorry utilisation:')
    for L in LORRY_FLEET:
        plate = L['plate']
        cap   = L['cap_kg']
        used  = cap - lorry_remain[plate]
        if used > 0:
            pct = round(used / cap * 100, 1)
            print(f'  {plate:>8}: {used:>8.0f} / {cap:>8.0f} kg  ({pct:>5.1f}%)')


if __name__ == '__main__':
    assign()
