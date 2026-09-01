# Warehouse Route & Remarks Update

A small internal web app so warehouse/office staff can update
`ENGSHENG.BPDLVCUST` delivery settings (Day / Time / Lorry Size / Out
Source / **Route Code**) from a phone or laptop browser, without opening
Sage X3 or running SQL by hand.

Built to match the patterns already used in `Statement_Version19.py` and
`web_aging_app.py`: same `config.json` + SQLAlchemy connection style, same
signed-cookie login session (instead of a browser Basic Auth popup).

## What it touches

| Category    | Columns updated                                       |
|-------------|--------------------------------------------------------|
| Day         | `UVYDAY1_0` ... `UVYDAY7_0` (Mon-Sun)                   |
| Time        | `ZUVYDAY15_0` (AM), `ZUVYDAY16_0` (PM)                  |
| Lorry Size  | `ZUVYDAY17_0`...`ZUVYDAY21_0` (MAX 2/5/11/15/21 TON)     |
| Out Source  | `ZUVYDAY14_0`                                           |
| Route Code  | `DRN_0`                                                 |

The app **only** ever writes to these 16 columns - never to any other field
on `BPDLVCUST`, and never to any other table, because the column list is
hard-coded server-side in `categories.py` and validated on every request.

Every change is written to `ENGSHENG.UVYDAY_AUDIT_LOG` (who / when / customer /
column / old value / new value).

### Route Code

`DRN_0` is just a number in the database - Sage X3/`BPDLVCUST` does **not**
store the route *name*. The route names shown in the portal come only from
[`data/route_codes.csv`](data/route_codes.csv) (Number -> route name),
which is exported from IT's route master list. This file is read fresh on
every request, so IT can update it on the server without restarting the
app, but the app itself never writes to it.

Because of this, the Route Code screen:

- Lets the user search/pick from the known route list (shows the code
  and its name side by side), or enter a raw numeric code that isn't on
  the list yet.
- Shows a clear warning when entering a code that isn't on the list,
  asking the user to confirm with the IT department first - an unlisted
  code will show a customer's route as a bare number with no name until
  `data/route_codes.csv` is updated to match, which is a good way to end
  up with inconsistent routes if it's done without IT's sign-off.
- Does **not** try to guess which of several same-named customers (e.g.
  two `BPCNUM_0`s for the same company) is "the real one" - search
  results always show `BPCNUM_0` + `BPAADD_0` separately so the user can
  update, say, the duplicate entry's route to an "NA" code without
  touching the correct one.

**Note:** `BPDLVCUST` has an enabled Sage X3 trigger (it auto-populates
`UPDUSR_0`/`UPDDAT_0` on update), so the update logic reads the "before"
values with a plain `SELECT` inside the same transaction rather than using
SQL Server's `OUTPUT` clause - `OUTPUT` isn't allowed on tables with
triggers enabled.

## 1. One-time database setup

Run this in SSMS (or via `query.py --file ...`, see below, if SSMS isn't
available on the machine you're working from):

```
sql/create_audit_table.sql
```

If you're using a scoped login rather than `sa`, also run:

```
sql/create_scoped_login.sql
```

(edit the password in that script first)

## 2. Install

```bat
cd uvyday_maintenance
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## 3. Configure

Copy the sample file and fill in real values (`config.json` is **not**
committed to source control / should not be zipped up and shared):

```bat
copy config.sample.json config.json
notepad config.json
```

```json
{
    "database": {
        "driver": "ODBC Driver 17 for SQL Server",
        "server": "10.0.0.19\\X3V12",
        "database": "x3v12",
        "uid": "<sql login for this app>",
        "pwd": "<its password>"
    },
    "secret_key": "<any long random string>",
    "app_users": {
        "alison": "choose-a-strong-password",
        "ops": "another-password"
    }
}
```

- `database` - same server/database as `config2.json`. `ENGSHENG` is a
  **schema** inside `x3v12`, not the database name itself - all the SQL in
  this app already references it as `ENGSHENG.BPDLVCUST`.
- `secret_key` - signs the login session cookie. Any long random string;
  changing it later logs everyone out (harmless, just log back in).
- `app_users` - controls who can log in to *this tool* (separate from the
  database login above). Add one entry per person who should have access.

## 4. Run it (test manually first)

```bat
uvicorn main:app --reload
```

Watch this console while testing - any error will print its full traceback
here. Once you've confirmed login + search + save all work, stop this
(`Ctrl+C`) before moving to the service install below.

## 5. Install as a Windows service via NSSM

```bat
C:\nssm\win64\nssm.exe install UvydayMaintenance "C:\py01\uvyday_maintenance\venv\Scripts\uvicorn.exe" "main:app --host 0.0.0.0 --port 8007"
C:\nssm\win64\nssm.exe set UvydayMaintenance AppDirectory "C:\py01\uvyday_maintenance"
C:\nssm\win64\nssm.exe start UvydayMaintenance
```

(adjust the `nssm.exe` path and port to match your setup)

## 6. Access from phone / laptop

- **On the office network / VPN:** browse to `http://<server-ip>:8007`
- **From outside (phone on mobile data):** needs HTTPS - either through
  your existing reverse proxy/Cloudflare Tunnel setup (same as
  `web_aging_app.py` uses), or a VPN into the office network. The login
  form posts a plaintext password, which is only safe over HTTPS.

The page is a single mobile-responsive HTML page (no app install needed).
Visiting it redirects to `/login` if you don't have a valid session; after
logging in you get a signed cookie valid for 12 hours, with a "Log out"
link in the header.

## How the update flow works

1. User searches by customer code or name -> picks the exact
   `BPCNUM_0` + `BPAADD_0` (customer + delivery address/route), since a
   customer can have multiple routes/addresses with different settings.
2. Picks one of the 5 categories. For Day/Time/Lorry Size/Out Source, sees
   toggle switches pre-filled with the current DB values for just that
   category. For Route Code, sees the current route (code + name, if
   known) and can search the route reference list or enter a raw code.
3. Saves -> reads current values, then runs one `UPDATE` for only that
   category's column(s), then writes one audit-log row per changed
   column, all in a single transaction (rolls back entirely if anything
   fails).

## query.py - ad-hoc SQL from the command line

If SSMS isn't available on the machine you're working from, `query.py`
reuses the same `config.json` connection to run SQL directly:

```bat
python query.py "SELECT * FROM ENGSHENG.UVYDAY_AUDIT_LOG"
python query.py --file sql\create_audit_table.sql
```

## Known simplifications (fine for an internal tool, flag if this grows)

- Login is a simple username/password list in `config.json`, not tied to
  Windows/AD. Fine for a handful of internal users; worth revisiting past
  that.
- No optimistic locking - if two people edit the exact same customer/route
  within seconds of each other, the second save wins.
