import hmac
import hashlib
import base64
import logging
import traceback
from typing import Dict, List, Optional

from fastapi import FastAPI, Depends, HTTPException, Request, Form
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel, Field
from sqlalchemy import text

from config import get_app_users, get_secret_key
from categories import CATEGORIES
from db import get_engine
from routes import get_route_list, get_route_message

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("uvyday")

app = FastAPI(title="Warehouse Route & Remarks Update")
templates = Jinja2Templates(directory="templates")
engine = get_engine()
SECRET_KEY = get_secret_key()


@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    """
    Catches anything not already an HTTPException (e.g. a SQL error) and
    returns JSON instead of a plain-text 500 page. Full traceback goes to
    the server console/log; the browser gets a generic message (no DB
    connection details leaked to the client).
    """
    logger.error("Unhandled error on %s:\n%s", request.url.path, traceback.format_exc())
    return JSONResponse(
        status_code=500,
        content={"detail": f"Server error: {exc.__class__.__name__}: {exc}"},
    )


# ---------------------------------------------------------------------------
# Signed cookie session helpers (same HMAC pattern as web_aging_app.py)
# ---------------------------------------------------------------------------
def sign_value(value: str) -> str:
    sig = hmac.new(SECRET_KEY.encode(), value.encode(), hashlib.sha256).digest()
    return value + "." + base64.urlsafe_b64encode(sig).decode()


def verify_signed_value(signed_value: str) -> Optional[str]:
    try:
        value, sig_b64 = signed_value.rsplit(".", 1)
        expected_sig = hmac.new(SECRET_KEY.encode(), value.encode(), hashlib.sha256).digest()
        if hmac.compare_digest(expected_sig, base64.urlsafe_b64decode(sig_b64.encode())):
            return value
    except Exception:
        return None
    return None


def get_current_user(request: Request) -> Optional[str]:
    cookie = request.cookies.get("session")
    if not cookie:
        return None
    return verify_signed_value(cookie)


def require_login_page(request: Request) -> Optional[str]:
    """For page routes: returns username, or None (caller should redirect)."""
    return get_current_user(request)


def require_login_api(request: Request) -> str:
    """For API routes: returns username, or raises 401 JSON."""
    user = get_current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return user


# ---------------------------------------------------------------------------
# Login / logout
# ---------------------------------------------------------------------------
@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request, error: Optional[str] = None):
    return templates.TemplateResponse(
        "login.html", {"request": request, "error": error}
    )


@app.post("/login")
def login_submit(username: str = Form(...), password: str = Form(...)):
    users = get_app_users()
    stored_password = users.get(username.strip())

    if stored_password is None or not hmac.compare_digest(password, stored_password):
        return RedirectResponse(url="/login?error=1", status_code=303)

    response = RedirectResponse(url="/", status_code=303)
    response.set_cookie(
        key="session",
        value=sign_value(username.strip()),
        httponly=True,
        samesite="lax",
        max_age=60 * 60 * 12,  # 12 hours
    )
    return response


@app.get("/logout")
def logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie("session")
    return response


# ---------------------------------------------------------------------------
# Main page
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    user = require_login_page(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    return templates.TemplateResponse(
        "index.html",
        {"request": request, "categories": CATEGORIES, "username": user},
    )


# ---------------------------------------------------------------------------
# API: search customers
# ---------------------------------------------------------------------------
@app.get("/api/search")
def search_customers(q: str, user: str = Depends(require_login_api)):
    q = (q or "").strip()
    if len(q) < 2:
        return {"results": []}

    like = f"%{q}%"
    sql = text("""
        SELECT TOP 25 BPCNUM_0, BPAADD_0, BPDNAM_0, DRN_0
        FROM ENGSHENG.BPDLVCUST
        WHERE BPCNUM_0 LIKE :like OR BPDNAM_0 LIKE :like
        ORDER BY BPCNUM_0, BPAADD_0
    """)
    with engine.connect() as conn:
        rows = conn.execute(sql, {"like": like}).mappings().all()

    return {
        "results": [
            {
                "bpcnum": r["BPCNUM_0"],
                "bpaadd": r["BPAADD_0"],
                "name": (r["BPDNAM_0"] or "").strip(),
                "route": r["DRN_0"],
                "route_message": get_route_message(r["DRN_0"]),
            }
            for r in rows
        ]
    }


# ---------------------------------------------------------------------------
# API: get current values for a customer/address, for all categories
# ---------------------------------------------------------------------------
@app.get("/api/customer")
def get_customer(bpcnum: str, bpaadd: str, user: str = Depends(require_login_api)):
    col_list = [col for cat in CATEGORIES.values() for col, _ in cat.get("columns", [])]
    select_cols = ", ".join(col_list)
    sql = text(f"""
        SELECT BPCNUM_0, BPAADD_0, BPDNAM_0, DRN_0, {select_cols}
        FROM ENGSHENG.BPDLVCUST
        WHERE BPCNUM_0 = :bpcnum AND BPAADD_0 = :bpaadd
    """)
    with engine.connect() as conn:
        row = conn.execute(sql, {"bpcnum": bpcnum, "bpaadd": bpaadd}).mappings().first()

    if row is None:
        raise HTTPException(status_code=404, detail="Customer/address not found")

    result = {
        "bpcnum": row["BPCNUM_0"],
        "bpaadd": row["BPAADD_0"],
        "name": (row["BPDNAM_0"] or "").strip(),
        "categories": {},
        "route": {
            "drn": row["DRN_0"],
            "message": get_route_message(row["DRN_0"]),
        },
    }

    for cat_key, cat in CATEGORIES.items():
        if cat.get("type") == "route":
            continue
        result["categories"][cat_key] = [
            {"column": col, "label": label, "checked": row[col] == 2}
            for col, label in cat["columns"]
        ]

    return result


# ---------------------------------------------------------------------------
# API: update one category for a customer/address
# ---------------------------------------------------------------------------
class UpdateRequest(BaseModel):
    bpcnum: str
    bpaadd: str
    category: str
    values: Dict[str, bool]  # {column_name: checked}


def _apply_update(conn, bpcnum: str, bpaadd: str, category: str,
                   cols_in_order: list, update_params: dict, user: str) -> dict:
    """
    Applies one category update to one customer/address inside an
    already-open connection/transaction. Returns a small result dict
    rather than raising, so callers (single or bulk) can decide how to
    handle a miss without unwinding the whole request.

    update_params is expected to be {"val_0": 2|1, "val_1": 2|1, ...} in the
    same order as cols_in_order (this is exactly what
    _validate_category_and_columns produces).
    """
    col_to_param = {col: f"val_{i}" for i, col in enumerate(cols_in_order)}

    select_cols = ", ".join(cols_in_order)
    select_sql = text(f"""
        SELECT {select_cols}
        FROM ENGSHENG.BPDLVCUST
        WHERE BPCNUM_0 = :bpcnum AND BPAADD_0 = :bpaadd
    """)
    before_row = conn.execute(select_sql, {"bpcnum": bpcnum, "bpaadd": bpaadd}).mappings().first()
    if before_row is None:
        return {"bpcnum": bpcnum, "bpaadd": bpaadd, "status": "not_found"}

    set_clause = ", ".join(f"{col} = :{col_to_param[col]}" for col in cols_in_order)
    update_sql = text(f"""
        UPDATE ENGSHENG.BPDLVCUST
        SET {set_clause}
        WHERE BPCNUM_0 = :bpcnum AND BPAADD_0 = :bpaadd
    """)
    full_params = dict(update_params)
    full_params["bpcnum"] = bpcnum
    full_params["bpaadd"] = bpaadd
    result = conn.execute(update_sql, full_params)

    if result.rowcount == 0:
        return {"bpcnum": bpcnum, "bpaadd": bpaadd, "status": "not_found"}

    audit_sql = text("""
        INSERT INTO ENGSHENG.UVYDAY_AUDIT_LOG
            (CHANGED_BY, BPCNUM_0, BPAADD_0, CATEGORY, COLUMN_NAME, OLD_VALUE, NEW_VALUE)
        VALUES (:changed_by, :bpcnum, :bpaadd, :category, :column_name, :old_value, :new_value)
    """)
    for col in cols_in_order:
        conn.execute(
            audit_sql,
            {
                "changed_by": user,
                "bpcnum": bpcnum,
                "bpaadd": bpaadd,
                "category": category,
                "column_name": col,
                "old_value": before_row[col],
                "new_value": update_params[col_to_param[col]],
            },
        )

    return {"bpcnum": bpcnum, "bpaadd": bpaadd, "status": "ok"}


def _validate_category_and_columns(category: str, values: Dict[str, bool]):
    if category not in CATEGORIES:
        raise HTTPException(status_code=400, detail="Unknown category")

    if CATEGORIES[category].get("type") == "route":
        raise HTTPException(
            status_code=400,
            detail="Route code updates must use /api/update-route or /api/bulk-update-route",
        )

    if CATEGORIES[category].get("locked"):
        raise HTTPException(
            status_code=403,
            detail=CATEGORIES[category].get(
                "locked_message",
                "This category cannot be changed here. Please contact IT.",
            ),
        )

    allowed_cols = {col for col, _ in CATEGORIES[category]["columns"]}
    submitted_cols = set(values.keys())

    if not submitted_cols.issubset(allowed_cols):
        raise HTTPException(
            status_code=400,
            detail="Request contains columns not permitted for this category",
        )
    if not submitted_cols:
        raise HTTPException(status_code=400, detail="No columns submitted")

    cols_in_order = [c for c in allowed_cols if c in values]
    update_params = {f"val_{i}": (2 if values[col] else 1) for i, col in enumerate(cols_in_order)}
    return cols_in_order, update_params


@app.post("/api/update")
def update_customer(payload: UpdateRequest, user: str = Depends(require_login_api)):
    cols_in_order, update_params = _validate_category_and_columns(payload.category, payload.values)

    with engine.begin() as conn:  # commits on success, rolls back on any exception
        result = _apply_update(conn, payload.bpcnum, payload.bpaadd, payload.category,
                                cols_in_order, update_params, user)

    if result["status"] == "not_found":
        raise HTTPException(status_code=404, detail="Customer/address not found")

    return {"status": "ok", "bpcnum": payload.bpcnum, "bpaadd": payload.bpaadd}


class CustomerRef(BaseModel):
    bpcnum: str
    bpaadd: str


class BulkUpdateRequest(BaseModel):
    customers: List[CustomerRef]
    category: str
    values: Dict[str, bool]


@app.post("/api/bulk-update")
def bulk_update_customers(payload: BulkUpdateRequest, user: str = Depends(require_login_api)):
    if not payload.customers:
        raise HTTPException(status_code=400, detail="No customers selected")
    if len(payload.customers) > 500:
        raise HTTPException(status_code=400, detail="Too many customers selected at once (max 500)")

    cols_in_order, update_params = _validate_category_and_columns(payload.category, payload.values)

    results = []
    # Each customer gets its own small transaction, so one missing/bad row
    # doesn't roll back everyone else's successful update.
    for cust in payload.customers:
        try:
            with engine.begin() as conn:
                r = _apply_update(conn, cust.bpcnum, cust.bpaadd, payload.category,
                                   cols_in_order, update_params, user)
            results.append(r)
        except Exception as e:
            logger.error("Bulk update failed for %s/%s: %s", cust.bpcnum, cust.bpaadd, e)
            results.append({"bpcnum": cust.bpcnum, "bpaadd": cust.bpaadd, "status": "error", "detail": str(e)})

    ok_count = sum(1 for r in results if r["status"] == "ok")
    return {
        "status": "ok",
        "total": len(results),
        "succeeded": ok_count,
        "failed": len(results) - ok_count,
        "results": results,
    }


# ---------------------------------------------------------------------------
# Route code (DRN_0): the route NAME lives only in data/route_codes.csv,
# never in the database. These endpoints only ever write the raw numeric
# code to DRN_0 - the column is hard-coded here, never taken from the
# client, same guarantee as the whitelisted category columns above.
# ---------------------------------------------------------------------------
@app.get("/api/routes")
def list_routes(user: str = Depends(require_login_api)):
    return {"routes": get_route_list()}


def _apply_route_update(conn, bpcnum: str, bpaadd: str, new_drn: int, user: str) -> dict:
    select_sql = text("""
        SELECT DRN_0 FROM ENGSHENG.BPDLVCUST
        WHERE BPCNUM_0 = :bpcnum AND BPAADD_0 = :bpaadd
    """)
    before_row = conn.execute(select_sql, {"bpcnum": bpcnum, "bpaadd": bpaadd}).mappings().first()
    if before_row is None:
        return {"bpcnum": bpcnum, "bpaadd": bpaadd, "status": "not_found"}

    update_sql = text("""
        UPDATE ENGSHENG.BPDLVCUST
        SET DRN_0 = :drn
        WHERE BPCNUM_0 = :bpcnum AND BPAADD_0 = :bpaadd
    """)
    result = conn.execute(update_sql, {"drn": new_drn, "bpcnum": bpcnum, "bpaadd": bpaadd})
    if result.rowcount == 0:
        return {"bpcnum": bpcnum, "bpaadd": bpaadd, "status": "not_found"}

    audit_sql = text("""
        INSERT INTO ENGSHENG.UVYDAY_AUDIT_LOG
            (CHANGED_BY, BPCNUM_0, BPAADD_0, CATEGORY, COLUMN_NAME, OLD_VALUE, NEW_VALUE)
        VALUES (:changed_by, :bpcnum, :bpaadd, 'ROUTE', 'DRN_0', :old_value, :new_value)
    """)
    conn.execute(
        audit_sql,
        {
            "changed_by": user,
            "bpcnum": bpcnum,
            "bpaadd": bpaadd,
            "old_value": before_row["DRN_0"],
            "new_value": new_drn,
        },
    )

    return {"bpcnum": bpcnum, "bpaadd": bpaadd, "status": "ok"}


class RouteUpdateRequest(BaseModel):
    bpcnum: str
    bpaadd: str
    drn: int = Field(ge=0, le=255)


@app.post("/api/update-route")
def update_route(payload: RouteUpdateRequest, user: str = Depends(require_login_api)):
    with engine.begin() as conn:
        result = _apply_route_update(conn, payload.bpcnum, payload.bpaadd, payload.drn, user)

    if result["status"] == "not_found":
        raise HTTPException(status_code=404, detail="Customer/address not found")

    return {
        "status": "ok",
        "bpcnum": payload.bpcnum,
        "bpaadd": payload.bpaadd,
        "drn": payload.drn,
        "message": get_route_message(payload.drn),
    }


class BulkRouteUpdateRequest(BaseModel):
    customers: List[CustomerRef]
    drn: int = Field(ge=0, le=255)


@app.post("/api/bulk-update-route")
def bulk_update_route(payload: BulkRouteUpdateRequest, user: str = Depends(require_login_api)):
    if not payload.customers:
        raise HTTPException(status_code=400, detail="No customers selected")
    if len(payload.customers) > 500:
        raise HTTPException(status_code=400, detail="Too many customers selected at once (max 500)")

    results = []
    for cust in payload.customers:
        try:
            with engine.begin() as conn:
                r = _apply_route_update(conn, cust.bpcnum, cust.bpaadd, payload.drn, user)
            results.append(r)
        except Exception as e:
            logger.error("Bulk route update failed for %s/%s: %s", cust.bpcnum, cust.bpaadd, e)
            results.append({"bpcnum": cust.bpcnum, "bpaadd": cust.bpaadd, "status": "error", "detail": str(e)})

    ok_count = sum(1 for r in results if r["status"] == "ok")
    return {
        "status": "ok",
        "total": len(results),
        "succeeded": ok_count,
        "failed": len(results) - ok_count,
        "results": results,
    }
