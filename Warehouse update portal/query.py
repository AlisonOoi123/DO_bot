"""
Quick ad-hoc SQL query tool for environments without SSMS access.
Uses the exact same connection setup as main.py (same config.json).

Usage (from inside the activated venv, in the project folder):
    python query.py "SELECT * FROM ENGSHENG.UVYDAY_AUDIT_LOG"
    python query.py "SELECT BPCNUM_0, BPAADD_0, ZUVYDAY15_0, ZUVYDAY16_0 FROM ENGSHENG.BPDLVCUST WHERE BPCNUM_0 = 'G018-E' AND BPAADD_0 = 'A01'"

If your query itself contains double quotes, wrap the whole thing in single
quotes instead, or save it in a .sql file and use --file:
    python query.py --file myquery.sql
"""
import sys
from sqlalchemy import text
from db import get_engine


def run_query(sql_text: str):
    engine = get_engine()
    with engine.begin() as conn:  # begin() commits automatically at the end; connect() does NOT
        result = conn.execute(text(sql_text))
        if not result.returns_rows:
            print(f"OK - {result.rowcount} row(s) affected (committed).")
            return
        rows = result.mappings().all()

    if not rows:
        print("(0 rows)")
        return

    columns = list(rows[0].keys())
    widths = [max(len(str(c)), *(len(str(r[c])) for r in rows)) for c in columns]

    def fmt_row(values):
        return " | ".join(str(v).ljust(w) for v, w in zip(values, widths))

    print(fmt_row(columns))
    print("-+-".join("-" * w for w in widths))
    for r in rows:
        print(fmt_row([r[c] for c in columns]))
    print(f"\n({len(rows)} row(s))")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python query.py \"SELECT ...\"")
        print("   or: python query.py --file myquery.sql")
        sys.exit(1)

    if sys.argv[1] == "--file":
        with open(sys.argv[2]) as f:
            sql_text = f.read()
    else:
        sql_text = sys.argv[1]

    try:
        run_query(sql_text)
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
