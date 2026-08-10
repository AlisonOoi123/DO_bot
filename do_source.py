"""
Live DO (Delivery Order) source — pulls today's delivery data directly from
the ERP database (X3), for the web portal's "fetch DOs from system" button.

Mirrors the standalone generate_script.py report (same query, same
ROUTE_MAP, same SHIP_DETAIL rules, same filters) with two differences,
both deliberate:

  1. Rows whose ETD is the database's NULL sentinel (1753-01-01, from
     `ISNULL(d.ZETD_0, '1753-01-01')` in the SQL) are dropped — those DOs
     have no real ETD set and shouldn't appear in the working file.
  2. Two output columns are renamed to match what bot.py's upload parser
     actually looks for (generate_script.py's names didn't line up, which
     would have silently disabled the affected features on any file from
     this source): 'SHIP DETAIL' -> 'SHIP_DETAIL', 'LONGITUDE' -> 'LONGITUD'.

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
from sqlalchemy import create_engine

CONFIG_PATH = os.environ.get(
    "DO_DB_CONFIG_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "configrd.json"))
SCHEMA_NAME = "ENGSHENG"

# SQL's ISNULL(d.ZETD_0, '1753-01-01') fallback for a DO with no real ETD
# set, formatted the same way as DATE ('%d-%m-%Y') further down.
_ETD_NULL_SENTINEL = "01-01-1753"

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

DRN_LIST = [25, 26, 27, 28, 29, 30, 31, 32, 37, 39, 40, 41, 45, 60, 61, 66, 67,
            69, 70, 71, 72, 73, 74, 75, 76, 77, 84, 85]

_QUERY_TEMPLATE = """
SELECT
    d.DLVDAT_0,
    d.SDHNUM_0,
    d.BPCORD_0 AS BPCNUM_0,
    d.BPDNAM_0,
    a.BPADES_0,
    d.GROWEI_0,
    d.SIHNUM_0,
    d.STOFCY_0,
    d.SDHTYP_0,
    c.DRN_0,
    a.POSCOD_0 AS BPDPOSCOD_0,
    a.CTY_0 AS BPDCTY_0,
    a.SAT_0 AS BPDSAT_0,
    ISNULL(d.ZLICENSE_0, '') AS ZLICENSE_0,
    ISNULL(d.ZDRIVER_0, '') AS ZDRIVER_0,
    ISNULL(d.ZFOLLOWER1_0, '') AS ZFOLLOWER1_0,
    ISNULL(d.ZFOLLOWER2_0, '') AS ZFOLLOWER2_0,
    ISNULL(d.ZETD_0, '1753-01-01') AS ZETD_0,
    ISNULL(c.ZUVYDAY17_0, 0) AS ZUVYDAY17_0,
    ISNULL(c.ZUVYDAY18_0, 0) AS ZUVYDAY18_0,
    ISNULL(c.ZUVYDAY19_0, 0) AS ZUVYDAY19_0,
    ISNULL(c.ZUVYDAY20_0, 0) AS ZUVYDAY20_0,
    ISNULL(c.ZUVYDAY21_0, 0) AS ZUVYDAY21_0
FROM {schema}.SDELIVERY d
LEFT JOIN {schema}.BPDLVCUST c
    ON d.BPCORD_0 = c.BPCNUM_0
   AND d.BPAADD_0 = c.BPAADD_0
LEFT JOIN {schema}.BPADDRESS a
    ON d.BPCORD_0 = a.BPANUM_0
   AND d.BPAADD_0 = a.BPAADD_0
"""


def _calculate_ship_detail(row) -> str:
    tags = []
    if str(row.get('ZUVYDAY17_0', '')).strip() in ['2', '2.0']: tags.append("MAX 2 TON")
    if str(row.get('ZUVYDAY18_0', '')).strip() in ['2', '2.0']: tags.append("MAX 5 TON")
    if str(row.get('ZUVYDAY19_0', '')).strip() in ['2', '2.0']: tags.append("MAX 11 TON")
    if str(row.get('ZUVYDAY20_0', '')).strip() in ['2', '2.0']: tags.append("MAX 15 TON")
    if str(row.get('ZUVYDAY21_0', '')).strip() in ['2', '2.0']: tags.append("MAX 21 TON")
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
    [today, today + etd_days] inclusive (e.g. etd_days=2 with today=10/8
    keeps ETD 10/8 through 12/8). None (default) applies no ETD-range
    filter — only the NULL-sentinel exclusion below still applies."""
    config = _load_config(config_path)
    engine = _build_engine(config)
    query = _QUERY_TEMPLATE.format(schema=SCHEMA_NAME)
    df = pd.read_sql(query, engine)

    df['ROUTE'] = df['DRN_0'].map(ROUTE_MAP).fillna('NA')
    df['SHIP_DETAIL'] = df.apply(_calculate_ship_detail, axis=1)
    df['VALIDATED'] = df['SIHNUM_0'].apply(lambda x: 'YES' if pd.notnull(x) and str(x).strip() != '' else 'NO')

    filtered_df = df[
        (df['SDHTYP_0'] != 'LOAN') &
        (df['STOFCY_0'] == '1SA') &
        (df['VALIDATED'] == 'NO') &
        (df['DRN_0'].isin(DRN_LIST))
    ].copy()

    if etd_days is not None:
        _etd_dt = pd.to_datetime(filtered_df['ZETD_0'], errors='coerce')
        _today = pd.Timestamp(datetime.now().date())
        _cutoff = _today + timedelta(days=etd_days)
        filtered_df = filtered_df[(_etd_dt >= _today) & (_etd_dt <= _cutoff)].copy()

    filtered_df['DATE_FORMATTED'] = pd.to_datetime(
        filtered_df['DLVDAT_0'], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')
    filtered_df['ETD_FORMATTED'] = pd.to_datetime(
        filtered_df['ZETD_0'], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')

    for col in ['ZDOREMARKS_0', 'ZLONGITUD_0', 'INVOICE_DATE', 'DISTANCE']:
        if col not in filtered_df.columns:
            filtered_df[col] = ''

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
        'INVOICE DATE': filtered_df['INVOICE_DATE'],
        'SITE': filtered_df['STOFCY_0'],
        'LICENSE': filtered_df['ZLICENSE_0'],
        'DRIVER': filtered_df['ZDRIVER_0'],
        'LORRY ASST 1': filtered_df['ZFOLLOWER1_0'],
        'LORRY ASST 2': filtered_df['ZFOLLOWER2_0'],
        'DISTANCE': filtered_df['DISTANCE'],
        'LONGITUD': filtered_df['ZLONGITUD_0'],          # was 'LONGITUDE' — see docstring
        'POSTCODE': filtered_df['BPDPOSCOD_0'],
        'CITY': filtered_df['BPDCTY_0'],
        'STATE': filtered_df['BPDSAT_0'],
        'SHIP_DETAIL': filtered_df['SHIP_DETAIL'],        # was 'SHIP DETAIL' — see docstring
        'ETD': filtered_df['ETD_FORMATTED'],
    })

    # Drop rows with no real ETD set (the DB's NULL-sentinel default).
    report_df = report_df[report_df['ETD'] != _ETD_NULL_SENTINEL].reset_index(drop=True)
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
