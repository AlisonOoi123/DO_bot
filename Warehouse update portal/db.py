import urllib.parse
from sqlalchemy import create_engine
from config import get_db_config

_engine = None


def get_engine():
    """Same connection-string pattern as Statement_Version19.py's get_engine()."""
    global _engine
    if _engine is None:
        db = get_db_config()
        params = urllib.parse.quote_plus(
            f"DRIVER={{{db['driver']}}};"
            f"SERVER={db['server']};"
            f"DATABASE={db['database']};"
            f"UID={db['uid']};"
            f"PWD={db['pwd']}"
        )
        _engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
    return _engine
