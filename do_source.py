"""
Live DO (Delivery Order) source — pulls today's delivery data directly from
the ERP database (X3), for the web portal's "fetch DOs from system" button.

Query verified directly against the production SQL (shared 2026-08-26,
see sql_query.docx) — REMARKS and DISTANCE actually live on BPDLVCUST,
not SDELIVERY, and INVOICE DATE needs a join to SINVOICE.ACCDAT_0; the
first version of this file selected none of those correctly and always
sent REMARKS/DISTANCE/INVOICE DATE blank, silently skipping the
REMARKS-driven size-cap/day-restriction logic on every live fetch.

Mirrors the standalone generate_script.py report (same ROUTE_MAP, same
SHIP_DETAIL rules, same filters) with differences, all deliberate:

  1. Two output columns are renamed to match what bot.py's upload parser
     actually looks for (generate_script.py's names didn't line up, which
     would have silently disabled the affected features on any file from
     this source): 'SHIP DETAIL' -> 'SHIP_DETAIL', 'LONGITUDE' -> 'LONGITUD'.
  2. UVYDAY1_0-UVYDAY7_0 and ZUVYDAY8_0-ZUVYDAY16_0 are pulled in (matching
     the production query) but not yet interpreted into anything — their
     meaning hasn't been confirmed. Only ZUVYDAY17_0-21_0 (the size-cap
     flags) are used, as before.

Rows are NOT filtered on ETD (2026-08-26 fix — see fetch_delivery_report's
docstring): an earlier version of this file dropped every DO whose ETD was
the DB's NULL-sentinel default (1753-01-01, meaning "no ETD set"), which
silently removed a large share of real DOs from every live fetch, since
most DOs never get an ETD populated in the ERP at all. bot.py never reads
ETD, so there was no correctness reason to filter on it — DATE (DLVDAT_0)
is what drives scheduling/priority.

Query cross-checked against the actual production SQL text (shared
2026-08-26, sql_query.txt) and a raw multi-year result dump used to derive
the correct row count for a known day (164 DOs), NOT hardcoded to that
number — several real discrepancies were found and fixed:
  1. Missing a date floor. This file returned every un-invoiced DO ever
     recorded, including un-invoiced orphan rows years old (2021-2025)
     that were clearly dead/abandoned records, not real pending work —
     confirmed by cross-checking the reference dump: every one of those
     stale rows had a delivery date before the current year, every
     legitimate row didn't. Originally added as `DLVDAT_0 >= <start of
     the current calendar year>`; tightened 2026-09-01 per explicit
     request to a rolling 30-day floor instead (`DLVDAT_0 >= DATEADD(day,
     -30, CAST(GETDATE() AS DATE))`; pandas mirror: `today - 30 days`,
     computed at call time) — a DO more than 30 days old is never useful
     to fetch, matching the same 30-day backdate-priority cutoff bot.py
     already applies after the fetch (see _past_date_cutoff). No upper
     bound — today onward is unrestricted.
  2. CUSTOMER NAME was pulled from SDELIVERY.BPDNAM_0; the real query
     reads BPDLVCUST.BPDNAM_0. POSTCODE/CITY/STATE were pulled from the
     BPADDRESS join (the customer's current master address), but the real
     query reads them straight off SDELIVERY (BPDPOSCOD_0/BPDCTY_0/
     BPDSAT_0 — the delivery's own snapshotted address, which can differ
     from the master address if it changed since, or if the delivery used
     a one-off address).
  3. CODE (customer code) was aliased from SDELIVERY.BPCORD_0; the real
     query reads BPDLVCUST.BPCNUM_0 directly. Equal for any row that
     survives the join (that's the join predicate), so not a behavioural
     bug, but changed to match the source of truth.
  4. SALESREP was joined in and REPNAM_0 added as a new 'SALES REP' output
     column — present in the real query, wasn't pulled in at all before.
     Informational only; bot.py's parser doesn't read it.
  5. Added SELECT DISTINCT, matching the real query — belt-and-braces
     against a join fan-out producing duplicate rows for a customer with
     more than one matching BPDLVCUST/BPADDRESS record; didn't change the
     row count on the reference dump (no duplicates existed there), but
     matches the source of truth exactly.

2026-08-27: briefly removed the SIHNUM_0 IS NULL ("not yet invoiced")
WHERE-clause filter to match the raw production query text exactly,
on the theory it was silently under-fetching. Wrong call, reverted
immediately — without it, the date floor (start of the CURRENT YEAR,
~8 months of history by August) let through every already-invoiced,
already-delivered DO for the whole year as well, ballooning one fetch
to 25,000+ rows instead of the handful of still-pending ones. The raw
query text is the report's base SELECT; the "not yet invoiced" filter
is this file's own addition on top of it to scope that base query down
to "still needs a lorry today", and DOES need to stay. The original
under-fetch complaint that prompted removing it was never actually
tied to this filter — it needs separate investigation.

Configuration: a JSON file with {server, database, username, password,
driver} — see configrd.json. NEVER commit this file; it holds a real
production DB password (already gitignored). Path is read from the
DO_DB_CONFIG_PATH env var; if unset, defaults to configrd.json sitting
right next to this file (i.e. in the DO_bot folder itself).
"""
from __future__ import annotations

import io
import json
import os
import urllib.parse
from datetime import datetime, timedelta

import pandas as pd
from sqlalchemy import create_engine, text

CONFIG_PATH = os.environ.get(
    "DO_DB_CONFIG_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "configrd.json"))
SCHEMA_NAME = "ENGSHENG"

# Route Mapping Dictionary — identical to generate_script.py.
ROUTE_MAP = {
    1: 'NA',
    2: 'JH01-->Tangkak-Pagoh-Muar',
    3: 'JH05-->Yong Peng- Parit Raja -Batu Pahat',
    4: 'JH06-->Ayer hitam-Kluang',
    5: 'JH09-->Kulai-Senai-Skudai-Johor Bharu',
    6: 'JH10-->Pdn-Plentong-Masai-Pasir gudang-Nusajaya',
    7: 'NA', 8: 'NA', 9: 'NA', 10: 'NA', 11: 'NA', 12: 'NA', 13: 'NA',
    14: 'NA', 15: 'NA', 16: 'NA', 17: 'NA', 18: 'NA', 19: 'NA', 20: 'NA',
    21: 'NA', 22: 'NA',
    23: 'MC01-->Tampin-Melaka',
    24: 'NS01-->Nilai-Mantin',
    25: 'NS04-->Port Dickson-Lukut',
    26: 'NS05-->Seremban',
    27: 'NS06-->Kuala Pilah-Bahau-Gemas',
    28: 'PH01-->Karak',
    29: 'PH03-->Raub',
    30: 'PH05-->Jerantut & Damak',
    31: 'PH07-->Bera',
    32: 'PH09-->Kuantan',
    33: 'PK01-->Sabak Bernam-Hutan Melintang',
    34: 'PK02-->Teluk Intan - Langkap - Slim River',
    35: 'PK04-->Ipoh - Town',
    36: 'PK05-->Batu Gajah - Menglembu - Lahat - Tapah',
    37: 'TR02-->Kemaman',
    38: 'ZNA',
    39: 'PH02-->Bentong',
    40: 'PH04-->Benta & Lipis',
    41: 'PH06-->Lanchang-K.Krau-Mentakab-Temerloh',
    42: 'NA', 43: 'NA', 44: 'NA',
    45: 'KV24-->SEMENYIH',
    46: 'NA', 47: 'NA', 48: 'NA', 49: 'NA', 50: 'NA', 51: 'NA', 52: 'NA', 53: 'NA',
    54: 'SBH01-->KOTA KINABALU',
    55: 'SRW04-->KUCHING',
    56: 'SRW03-->SIBU',
    57: 'SRW01-->BINTULU',
    58: 'KLT01-->KOTA BHARU',
    59: 'SRW02-->MIRI',
    60: 'PH10-->MENTAKAB',
    61: 'PH11-->TEMERLOH',
    62: 'PG01-->PENANG ISLAND',
    63: 'PG02-->PENANG MAINLAND',
    64: 'KDH01-->KEDAH',
    65: 'TR01-->TERENGGANU',
    66: 'KV01A - T.MALIM - K.KUBU  - B.KALI  -N 1',
    67: 'KV02A - B.BERUNTUNG - SERENDAH - RAWANG - N 2',
    68: 'KV03A- K.SELANGOR - SEKINCHAN - T.KARANG - S.BESAR - WN 1',
    69: 'KV04A - SUNGAI BULOH - U5 - KOTA DAMANSARA - N 4',
    70: 'KV05A - SELAYANG - BATU CAVES - N 5',
    71: 'KV06A - KEPONG - SRI DAMANSARA - MONT KIARA - N 6',
    72: 'KV07A - SEGAMBUT - SENTUL - NE 1',
    73: 'KV08A - GOMBAK - SETAPAK - NE 2',
    74: 'KV09A - WANGSAMAJU - AMPANG SOURTH -NE 3  ** START UKAY',
    75: 'KV10A - CHOW KIT - RAJA LAUT - KL - E 1',
    76: 'KV11A - PUDU - AMPANG NORTH - CHERAS NORTH - E 2',
    77: 'KV12A - BANGSAR - BRICKFIELD - TAMAN DESA - KUCHAI LAMA - OUG - KINRARA - ES 1',
    78: 'KV13A - S.KEMBANGAN - SERDANG - B.JALIL- P.JALIL - ES 2',
    79: 'KV14A - PUTRA PERDANA - B.B.PUCHONG - USJ - SUBANG 1 - C 1',
    80: 'KV15A - SUNWAY - GLENMARIE - PJ OLD TOWN - C 2',
    81: 'KV16A - SUBANG GOLF CLUB - SEA PARK - UM - C 3',
    82: 'KV17A - SUBANG AIRPORT - ARA DAMANSARA - BU - BUKIT DAMAN SARA -C 4',
    83: 'KV18A - SUNGAI BESI - BALAKONG - SERDANG - SE 3',
    84: 'KV19A - HULU LANGAT - CHERAS SOUTH - KAJANG - SE 4 ** START BATU 9',
    85: 'KV20A - SEMENYIH - BANGI - BERANANG - MANTIN - SE 5',
    86: 'KV21A - PUNCAK ALAM - KAPAR - SETIA ALAM - HOSPITAL SHAH ALAM - C 5',
    87: 'KV22A - S.ALAM - S.MUDA - K.KEMUNING - C 6',
    88: 'KV23A - PANDAMARAN - P.KLANG - B.TINGGI - T.GONG - P.INDAH - W 1',
    89: 'KV24A - DENGKIL - CYBER - PUTRAJAYA - KLIA - S 1',
    90: 'KV25A - JENJAROM - BANTING - T.SEPAT - T.P.GARANG - S 2',
    91: 'PK03 --> CAMERON HIGHLAND',
    92: 'PK06 --> Manjung - Pantai Remis',
    93: 'PK07 --> Tanjung Rambutan - Sungai Siput',
    94: 'PK08 --> Taiping',
}

# DRN_0 codes to fetch from the DB — restricted to routes either planner
# actually handles (matches the two ROUTE sheets in LORRY DAILY PLANNING.xlsx),
# not every DRN_0 value that exists (e.g. East Malaysia/other-state codes
# neither ABI nor VIVIAN runs stay excluded). The query only ever fetched
# ABI's codes (2026-08-27 fix) — VIVIAN's route sheet exists but her DOs were
# never even pulled from the DB, so no amount of downstream route-ownership
# filtering could produce anything for her. Widening this list to include
# her known routes too is the actual fix; ABI's codes are unchanged.
_ABI_DRN = [25, 26, 27, 28, 29, 30, 31, 32, 37, 39, 40, 41, 45, 60, 61, 66, 67,
            69, 70, 71, 72, 73, 74, 75, 76, 77, 84, 85]
_VIVIAN_DRN = [2, 3, 4, 5, 6, 23, 24, 33, 34, 35, 36, 68, 78, 79, 80, 81, 82,
               83, 86, 87, 88, 89, 90, 92, 93, 94]
# DRN_0 codes whose ROUTE_MAP value is the literal "NA" — a customer whose
# route has been cleared in the Warehouse Route & Remarks Update portal
# (same DB, same DRN_0 column) uses one of these. Included so those DOs are
# still fetched at all instead of silently dropped by the WHERE clause —
# bot.py routes anything with ROUTE == "NA" into the board's Manual Assign
# Only section rather than letting AI Assign touch it.
_NA_DRN = sorted(k for k, v in ROUTE_MAP.items() if v == 'NA')
DRN_LIST = sorted(set(_ABI_DRN) | set(_VIVIAN_DRN) | set(_NA_DRN))

# Set by fetch_delivery_report() on every call — a stage-by-stage row count
# breakdown, so a 0-row result can be explained (which filter zeroed it out)
# instead of just reported.
LAST_FETCH_DIAGNOSTICS: dict = {}

_QUERY_TEMPLATE = """
SELECT DISTINCT
    d.DLVDAT_0,
    d.SDHNUM_0,
    c.BPCNUM_0,
    c.BPDNAM_0,
    a.BPADES_0,
    d.GROWEI_0,
    d.SIHNUM_0,
    d.STOFCY_0,
    d.SDHTYP_0,
    d.CFMFLG_0,
    c.DRN_0,
    ISNULL(c.ZDOREMARKS_0, '') AS ZDOREMARKS_0,
    ISNULL(c.ZDISTANCE_0, '') AS ZDISTANCE_0,
    ISNULL(c.ZLONGITUD_0, '') AS ZLONGITUD_0,
    ISNULL(c.ZDROPPOINT_0, '') AS ZDROPPOINT_0,
    i.ACCDAT_0 AS INVOICE_DATE,
    d.BPDPOSCOD_0,
    d.BPDCTY_0,
    d.BPDSAT_0,
    ISNULL(r.REPNAM_0, '') AS REPNAM_0,
    ISNULL(d.ZLICENSE_0, '') AS ZLICENSE_0,
    ISNULL(d.ZDRIVER_0, '') AS ZDRIVER_0,
    ISNULL(d.ZFOLLOWER1_0, '') AS ZFOLLOWER1_0,
    ISNULL(d.ZFOLLOWER2_0, '') AS ZFOLLOWER2_0,
    ISNULL(d.ZETD_0, '1753-01-01') AS ZETD_0,
    ISNULL(c.UVYDAY1_0, 0) AS UVYDAY1_0,
    ISNULL(c.UVYDAY2_0, 0) AS UVYDAY2_0,
    ISNULL(c.UVYDAY3_0, 0) AS UVYDAY3_0,
    ISNULL(c.UVYDAY4_0, 0) AS UVYDAY4_0,
    ISNULL(c.UVYDAY5_0, 0) AS UVYDAY5_0,
    ISNULL(c.UVYDAY6_0, 0) AS UVYDAY6_0,
    ISNULL(c.UVYDAY7_0, 0) AS UVYDAY7_0,
    ISNULL(c.ZUVYDAY8_0, 0) AS ZUVYDAY8_0,
    ISNULL(c.ZUVYDAY9_0, 0) AS ZUVYDAY9_0,
    ISNULL(c.ZUVYDAY10_0, 0) AS ZUVYDAY10_0,
    ISNULL(c.ZUVYDAY11_0, 0) AS ZUVYDAY11_0,
    ISNULL(c.ZUVYDAY12_0, 0) AS ZUVYDAY12_0,
    ISNULL(c.ZUVYDAY13_0, 0) AS ZUVYDAY13_0,
    ISNULL(c.ZUVYDAY14_0, 0) AS ZUVYDAY14_0,
    ISNULL(c.ZUVYDAY15_0, 0) AS ZUVYDAY15_0,
    ISNULL(c.ZUVYDAY16_0, 0) AS ZUVYDAY16_0,
    ISNULL(c.ZUVYDAY17_0, 0) AS ZUVYDAY17_0,
    ISNULL(c.ZUVYDAY18_0, 0) AS ZUVYDAY18_0,
    ISNULL(c.ZUVYDAY19_0, 0) AS ZUVYDAY19_0,
    ISNULL(c.ZUVYDAY20_0, 0) AS ZUVYDAY20_0,
    ISNULL(c.ZUVYDAY21_0, 0) AS ZUVYDAY21_0,
    ISNULL(p.PRODUCTS_0, '') AS PRODUCTS_0
FROM {schema}.SDELIVERY d
LEFT JOIN {schema}.BPDLVCUST c
    ON d.BPAADD_0 = c.BPAADD_0
   AND d.BPCORD_0 = c.BPCNUM_0
LEFT JOIN {schema}.SINVOICE i
    ON d.SIHNUM_0 = i.NUM_0
LEFT JOIN {schema}.SALESREP r
    ON d.REP_0 = r.REPNUM_0
LEFT JOIN {schema}.BPADDRESS a
    ON c.BPCNUM_0 = a.BPANUM_0
   AND c.BPAADD_0 = a.BPAADD_0
-- One row per DO with every SDELIVERYD line item ("<description> x<qty>")
-- concatenated together — SDELIVERYD is a one-to-many detail table (each DO
-- can have several product lines), so it's pre-aggregated here rather than
-- joined directly into the main SELECT DISTINCT, which would fan out one
-- SDELIVERY row per line item and break every weight/route aggregate above.
-- STRING_AGG requires SQL Server 2017+; if the real server predates that,
-- this fails loudly with a clear "not a recognized function" error rather
-- than silently returning wrong data.
LEFT JOIN (
    SELECT SDHNUM_0,
           -- QTY_0 is a decimal column with a long fixed scale (e.g.
           -- 10.0000000000000000) — rounded to a whole number here since
           -- delivery quantities are always whole units; the web app splits
           -- on '; ' and renders each product on its own bulleted line.
           -- ITMREF_0 is prefixed (before a '|') so the web app can bold
           -- only the item codes on its highlight list without a second
           -- round trip — stripped back out before display.
           STRING_AGG(CONCAT(ITMREF_0, '|', ITMDES_0, ' x', CAST(ROUND(QTY_0, 0) AS BIGINT)), '; ') AS PRODUCTS_0
    FROM {schema}.SDELIVERYD
    GROUP BY SDHNUM_0
) p ON d.SDHNUM_0 = p.SDHNUM_0
WHERE d.SDHTYP_0 <> 'LOAN'
  AND d.STOFCY_0 = '1SA'
  AND (d.SIHNUM_0 IS NULL OR LTRIM(RTRIM(d.SIHNUM_0)) = '')
  AND d.DLVDAT_0 >= DATEADD(day, -30, CAST(GETDATE() AS DATE))
  AND c.DRN_0 IN ({drn_list})
  {etd_clause}
ORDER BY d.DLVDAT_0
"""


_ACTIVE = ('2', '2.0')  # BPDLVCUST's UVYDAY-style flags: '2' = active, confirmed
                        # via the ZUVYDAY17-21 size-cap fields (matches
                        # generate_script.py) and the Crystal Report's REMARKS
                        # formula field, which checks each named sub-formula for
                        # a non-blank/non-"0" result the same way.

# UVYDAY1_0..UVYDAY7_0 -> the single-weekday tags in the Crystal Report's
# REMARKS formula (MONDAY..SUNDAY), in field-number order — confirmed from
# the formula's own MONDAY-first, SUNDAY-last evaluation order.
_WEEKDAY_FIELDS = [
    ('UVYDAY1_0', 'MONDAY'), ('UVYDAY2_0', 'TUESDAY'), ('UVYDAY3_0', 'WEDNESDAY'),
    ('UVYDAY4_0', 'THURSDAY'), ('UVYDAY5_0', 'FRIDAY'), ('UVYDAY6_0', 'SATURDAY'),
    ('UVYDAY7_0', 'SUNDAY'),
]

# NOT YET IMPLEMENTED: the Crystal Report's REMARKS formula also has 9 more
# tags — EVERYDAY, MON TO FRIDAY, MON,THUR, MON,WED,FRI, TUES,FRI,
# TUES,THUR,SAT, WED,FRI, AM, PM — presumably each backed by one of
# ZUVYDAY8_0..ZUVYDAY16_0 (9 fields, 9 tags), but the .docx only showed the
# combining formula, not each sub-formula's own field mapping, so which
# ZUVYDAY8-16 field maps to which tag is NOT confirmed. Guessing here would
# risk mislabeling a real day-restriction (e.g. showing "WED,FRI" when the
# DB flag actually means "MON,THUR"), so these are left out until confirmed
# — ask for each @-formula's definition (or just the field-to-tag mapping)
# to fill this in correctly.

def _calculate_ship_detail(row) -> str:
    tags = []
    for field, label in _WEEKDAY_FIELDS:
        if str(row.get(field, '')).strip() in _ACTIVE:
            tags.append(label)
    if str(row.get('ZUVYDAY17_0', '')).strip() in _ACTIVE: tags.append("MAX 2 TON")
    if str(row.get('ZUVYDAY18_0', '')).strip() in _ACTIVE: tags.append("MAX 5 TON")
    if str(row.get('ZUVYDAY19_0', '')).strip() in _ACTIVE: tags.append("MAX 11 TON")
    if str(row.get('ZUVYDAY20_0', '')).strip() in _ACTIVE: tags.append("MAX 15 TON")
    if str(row.get('ZUVYDAY21_0', '')).strip() in _ACTIVE: tags.append("MAX 21 TON")
    return ", ".join(tags)


def _load_config(path: str = None) -> dict:
    path = path or CONFIG_PATH
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"DB config not found at {path}. Set the DO_DB_CONFIG_PATH "
            f"environment variable to its location, or place configrd.json there."
        )
    with open(path) as f:
        return json.load(f)


def _build_engine(config: dict):
    encoded_password = urllib.parse.quote_plus(config['password'])
    encoded_driver = urllib.parse.quote_plus(config['driver'])
    connection_string = (
        f"mssql+pyodbc://{config['username']}:{encoded_password}@"
        f"{config['server']}/{config['database']}?driver={encoded_driver}"
    )
    return create_engine(connection_string)


def fetch_delivery_report(config_path: str = None, etd_days: int = None) -> pd.DataFrame:
    """Query the live ERP DB and return today's Delivery Report rows in the
    shape bot.py's DO-file parser expects — see module docstring for the two
    column-name fixes and the ETD-sentinel exclusion versus
    generate_script.py's original output. Raises on any DB/config failure;
    the caller decides how to surface that to the user.

    etd_days: when given, keep only rows whose ETD falls within
    [today - etd_days, today + etd_days] inclusive — a symmetric window
    the user picks in the portal (e.g. etd_days=2 with today=26/8 keeps
    ETD 24/8 through 28/8, catching DOs whose ETD already slipped a
    couple of days as well as ones coming up). None (default) applies no
    ETD-range filter at all — every still-pending DO from this year is
    fetched, matching this file's default behaviour before this window
    was made user-configurable."""
    config = _load_config(config_path)
    engine = _build_engine(config)

    etd_clause = ""
    if etd_days is not None:
        etd_days = int(etd_days)
        if etd_days < 0:
            raise ValueError("etd_days can't be negative")
        etd_clause = (
            f"AND d.ZETD_0 >= DATEADD(day, -{etd_days}, CAST(GETDATE() AS DATE)) "
            f"AND d.ZETD_0 <= DATEADD(day, {etd_days}, CAST(GETDATE() AS DATE))"
        )
    query = _QUERY_TEMPLATE.format(
        schema=SCHEMA_NAME,
        drn_list=", ".join(str(n) for n in DRN_LIST),
        etd_clause=etd_clause,
    )
    df = pd.read_sql(query, engine)

    df['ROUTE'] = df['DRN_0'].map(ROUTE_MAP).fillna('NA')
    df['SHIP_DETAIL'] = df.apply(_calculate_ship_detail, axis=1)
    df['VALIDATED'] = df['SIHNUM_0'].apply(lambda x: 'YES' if pd.notnull(x) and str(x).strip() != '' else 'NO')

    # Diagnostics: when the final result is empty, this shows which single
    # filter step actually zeroed it out, instead of leaving "0 DOs" to guess
    # at. Overwritten on every call — read it right after this function
    # returns if you need to know why.
    global LAST_FETCH_DIAGNOSTICS
    _cutoff_30d = pd.Timestamp(datetime.now().date()) - timedelta(days=30)
    _not_loan = df['SDHTYP_0'] != 'LOAN'
    _right_site = df['STOFCY_0'] == '1SA'
    _not_validated = df['VALIDATED'] == 'NO'
    _known_route = df['DRN_0'].isin(DRN_LIST)
    _within_30_days = pd.to_datetime(df['DLVDAT_0'], errors='coerce') >= _cutoff_30d
    LAST_FETCH_DIAGNOSTICS = {
        "raw_rows_from_sql": int(len(df)),
        "after_not_loan": int(_not_loan.sum()),
        "after_site_1SA": int((_not_loan & _right_site).sum()),
        "after_not_validated": int((_not_loan & _right_site & _not_validated).sum()),
        "after_known_route": int((_not_loan & _right_site & _not_validated & _known_route).sum()),
        "after_30day_floor": int((_not_loan & _right_site & _not_validated & _known_route & _within_30_days).sum()),
        "distinct_stofcy_seen": sorted(str(v) for v in df['STOFCY_0'].dropna().unique())[:10] if len(df) else [],
        "distinct_sdhtyp_seen": sorted(str(v) for v in df['SDHTYP_0'].dropna().unique())[:10] if len(df) else [],
    }

    filtered_df = df[_not_loan & _right_site & _not_validated & _known_route & _within_30_days].copy()

    if etd_days is not None:
        _etd_dt = pd.to_datetime(filtered_df['ZETD_0'], errors='coerce')
        _today = pd.Timestamp(datetime.now().date())
        _lo = _today - timedelta(days=etd_days)
        _hi = _today + timedelta(days=etd_days)
        filtered_df = filtered_df[(_etd_dt >= _lo) & (_etd_dt <= _hi)].copy()
        LAST_FETCH_DIAGNOSTICS["after_etd_window"] = int(len(filtered_df))

    filtered_df['DATE_FORMATTED'] = pd.to_datetime(
        filtered_df['DLVDAT_0'], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')
    filtered_df['ETD_FORMATTED'] = pd.to_datetime(
        filtered_df['ZETD_0'], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')
    filtered_df['INVOICE_DATE_FORMATTED'] = pd.to_datetime(
        filtered_df['INVOICE_DATE'], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')

    for col in ['ZDOREMARKS_0', 'ZLONGITUD_0', 'ZDISTANCE_0', 'ZDROPPOINT_0', 'PRODUCTS_0']:
        if col not in filtered_df.columns:
            filtered_df[col] = ''

    # Matches the Crystal Report's own Distance formula:
    # {BPDLVCUST.ZDISTANCE_0}+"KM" — bot.py's distance parser (_distance_km)
    # extracts the leading number regardless of a trailing unit, so this is
    # purely for display fidelity with what the report already shows.
    filtered_df['DISTANCE_FORMATTED'] = filtered_df['ZDISTANCE_0'].apply(
        lambda v: f"{v}KM" if str(v).strip() not in ('', 'nan', 'None') else '')

    filtered_df['no'] = range(1, len(filtered_df) + 1)

    report_df = pd.DataFrame({
        'NO': filtered_df['no'],
        'DATE': filtered_df['DATE_FORMATTED'],
        'DO NUMBER': filtered_df['SDHNUM_0'],
        'CODE': filtered_df['BPCNUM_0'],
        'ROUTE': filtered_df['ROUTE'],
        'CUSTOMER NAME': filtered_df['BPDNAM_0'],
        'BRANCH': filtered_df['BPADES_0'],
        'GROSS WEIGHT': filtered_df['GROWEI_0'],
        'REMARKS': filtered_df['ZDOREMARKS_0'],
        'VALIDATED': filtered_df['VALIDATED'],
        'INVOICE NO': filtered_df['SIHNUM_0'],
        'INVOICE DATE': filtered_df['INVOICE_DATE_FORMATTED'],
        'SITE': filtered_df['STOFCY_0'],
        'LICENSE': filtered_df['ZLICENSE_0'],
        'DRIVER': filtered_df['ZDRIVER_0'],
        'LORRY ASST 1': filtered_df['ZFOLLOWER1_0'],
        'LORRY ASST 2': filtered_df['ZFOLLOWER2_0'],
        'DISTANCE': filtered_df['DISTANCE_FORMATTED'],
        'LONGITUD': filtered_df['ZLONGITUD_0'],          # was 'LONGITUDE' — see docstring
        'POSTCODE': filtered_df['BPDPOSCOD_0'],
        'CITY': filtered_df['BPDCTY_0'],
        'STATE': filtered_df['BPDSAT_0'],
        'SHIP_DETAIL': filtered_df['SHIP_DETAIL'],        # was 'SHIP DETAIL' — see docstring
        'ETD': filtered_df['ETD_FORMATTED'],
        'SALES REP': filtered_df['REPNAM_0'],
        'DROPPOINT': filtered_df['ZDROPPOINT_0'],
        'PRODUCTS': filtered_df['PRODUCTS_0'],
    })

    # NOTE: rows are intentionally NOT dropped just because ETD is unset (the
    # DB's NULL-sentinel default, 1753-01-01). ETD isn't used anywhere in the
    # assignment logic (bot.py never reads it) and most real DOs never get an
    # ETD populated in the ERP at all, so filtering on it was silently
    # dropping legitimate DOs from every live fetch. ETD is kept as a display
    # column only; DATE (DLVDAT_0) drives scheduling/priority.
    report_df['NO'] = range(1, len(report_df) + 1)

    return report_df


def report_to_xlsx_bytes(report_df: pd.DataFrame) -> bytes:
    """Render report_df to plain .xlsx bytes — no styling, no TOTAL row, just
    the data with a header row and default column widths. Matches the plain
    look of the file this replaces (no custom formatting applied)."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        report_df.to_excel(writer, index=False, sheet_name="Delivery Report")
    return buf.getvalue()


# ── Write path: save the board's assignments back to the ERP ────────────────
# By explicit request. This is the only place in the whole app that writes
# to production Sage X3 tables rather than just reading from them — treat
# any change here with the same care as the read query's own history (see
# the module docstring's cross-check notes).
# Confirmed via INFORMATION_SCHEMA.COLUMNS: the real column name is
# literally "ACTUAL_D.DATE" (a period inside the identifier, not an
# underscore) — needs bracket-quoting in T-SQL or the dot parses as a
# schema/table separator.
_ZLORRY_DATE_COL = "[ACTUAL_D.DATE]"


def save_board_to_erp(rows: list[dict], add_date, trip: str, config_path: str = None) -> dict:
    """Write one planner's board back to the ERP — SDELIVERY per DO, ZLORRY
    per (add_date, trip, plate).

    rows: one dict per DO in the board's own scope (assigned AND
    unassigned — unassigned rows are only used for the ZLORRY totals
    below; they are never written to SDELIVERY):
        {"do_number": str, "plate": str | None, "weight_kg": float,
         "ai_assigned": bool}
    add_date: the "Assign for" date (a datetime.date) — ZLORRY.<date col>
        and SDELIVERY.ARVDAT_0.
    trip: the Trip number as a string ("1".."4") — ZLORRY.TRIP.

    SDELIVERY.ZLICENSE_0/ARVDAT_0/ZAIASSIGN_0 are written ONLY for assigned
    DOs (by explicit request — unassigned DOs must never be touched, even
    to clear a stale prior plate). A DO with no plate this save is simply
    skipped; whatever SDELIVERY already holds for it is left alone.

    ZAIASSIGN_0 (confirmed via SSMS — the real column, not the earlier
    guessed "AI_Assign" name) is written 'Yes'/'No' per the caller's own
    ai_assigned flag — never NULL, matching this DB's usual blank-string
    convention for char columns (same reasoning as ZLICENSE_0).

    ZLORRY.TON (converted to KG here — callers pass weight_kg already) is
    the SUMMED weight of everything assigned to that plate for this exact
    (add_date, trip) — the natural key the user confirmed is unique per
    date+trip+vehicle. A plate that had DOs before but has none now still
    gets its row (never deleted, DRIVER/ASSIT_1/ASSIT_2/OVERNIGHT
    untouched) — just its TON drops to 0, so ZLORRY always mirrors what's
    actually assigned rather than accumulating stale totals. A brand-new
    (add_date, trip, plate) combination gets INSERTed.

    Everything happens in one transaction — a failure partway rolls back
    the whole save rather than leaving Sage X3 half-updated. Raises on any
    DB/config failure; the caller decides how to surface that.
    """
    config = _load_config(config_path)
    engine = _build_engine(config)

    # The legacy "[Microsoft][ODBC SQL Server Driver]" driver this DB uses
    # can't bind a bare Python `datetime.date` (raises HYC00 "Optional
    # feature not implemented" from SQLBindParameter) — only a full
    # `datetime.datetime`. Normalize once, up front, for every :add_date
    # param used below.
    if isinstance(add_date, datetime):
        pass
    elif hasattr(add_date, "year"):
        add_date = datetime(add_date.year, add_date.month, add_date.day)

    plate_totals: dict[str, float] = {}
    for r in rows:
        plate = r.get("plate")
        if plate:
            plate_totals[plate] = plate_totals.get(plate, 0.0) + float(r.get("weight_kg", 0) or 0)

    dos_written = 0
    lorries_written = 0
    with engine.begin() as conn:
        # 1) SDELIVERY — one row per DO, matched by SDHNUM_0 (the DO number,
        # same column the read query already keys off). No OUTPUT clause —
        # BPDLVCUST is documented to have an update trigger that breaks
        # OUTPUT, and SDELIVERY may well have one too; a plain UPDATE is
        # safe either way.
        for r in rows:
            plate = r.get("plate")
            if not plate:
                # Unassigned DO — leave whatever SDELIVERY already has for
                # it untouched, by explicit request.
                continue
            do_number = str(r.get("do_number", "")).strip()
            if not do_number:
                continue
            conn.execute(
                text(f"""
                    UPDATE {SCHEMA_NAME}.SDELIVERY
                    SET ZLICENSE_0 = :plate,
                        ARVDAT_0 = :add_date,
                        ZAIASSIGN_0 = :ai_assign
                    WHERE SDHNUM_0 = :do_number
                """),
                {
                    "plate": plate,
                    "add_date": add_date,
                    "ai_assign": "Yes" if r.get("ai_assigned") else "No",
                    "do_number": do_number,
                },
            )
            dos_written += 1

        # 2) ZLORRY — union of plates assigned today and plates that already
        # have a row for this exact (add_date, trip), so a plate dropped to
        # zero DOs still gets its TON reset instead of being left stale.
        existing = conn.execute(
            text(f"""
                SELECT VEHICLE FROM {SCHEMA_NAME}.ZLORRY
                WHERE {_ZLORRY_DATE_COL} = :add_date AND TRIP = :trip
            """),
            {"add_date": add_date, "trip": trip},
        ).fetchall()
        existing_plates = {row[0] for row in existing}

        for plate in sorted(set(plate_totals) | existing_plates):
            ton_kg = round(plate_totals.get(plate, 0.0), 3)
            if plate in existing_plates:
                conn.execute(
                    text(f"""
                        UPDATE {SCHEMA_NAME}.ZLORRY
                        SET TON = :ton
                        WHERE {_ZLORRY_DATE_COL} = :add_date AND TRIP = :trip AND VEHICLE = :plate
                    """),
                    {"ton": ton_kg, "add_date": add_date, "trip": trip, "plate": plate},
                )
            else:
                conn.execute(
                    text(f"""
                        INSERT INTO {SCHEMA_NAME}.ZLORRY ({_ZLORRY_DATE_COL}, TRIP, VEHICLE, TON)
                        VALUES (:add_date, :trip, :plate, :ton)
                    """),
                    {"add_date": add_date, "trip": trip, "plate": plate, "ton": ton_kg},
                )
            lorries_written += 1

    return {"dos_written": dos_written, "lorries_written": lorries_written}
