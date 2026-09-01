"""
Configuration for the UVYDAY maintenance tool.

Follows the same pattern as Statement_Version19.py: DB connection details
live in a JSON file sitting next to this script, not in the source code.

config.json (copy config.sample.json and rename it):
{
    "database": {
        "driver": "ODBC Driver 17 for SQL Server",
        "server": "localhost",
        "database": "ENGSHENG",
        "uid": "your_sql_login",
        "pwd": "your_password"
    },
    "app_users": {
        "alison": "choose-a-strong-password",
        "ops": "another-password"
    }
}

"app_users" controls who can log in to THIS tool (separate from the DB
login above). Give the DB login in "database" UPDATE + SELECT rights
limited to ENGSHENG.BPDLVCUST and ENGSHENG.UVYDAY_AUDIT_LOG only - this
app is reachable from phones/laptops, so don't point it at a full-access
Sage X3 service account.
"""
import json
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.json")

with open(CONFIG_PATH) as f:
    config = json.load(f)


def get_db_config() -> dict:
    return config["database"]


def get_app_users() -> dict:
    return config.get("app_users", {})


def get_secret_key() -> str:
    """
    Used to sign the session cookie (same HMAC pattern as web_aging_app.py's
    SECRET_KEY). Stored in config.json instead of hard-coded in source, so
    it's not accidentally committed/shared with the code.
    """
    key = config.get("secret_key")
    if not key:
        raise RuntimeError(
            "config.json is missing \"secret_key\". Add a long random string, "
            "e.g. \"secret_key\": \"<any random 32+ character string>\""
        )
    return key
