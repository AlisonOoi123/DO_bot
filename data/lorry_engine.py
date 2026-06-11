"""
Lorry Assignment Engine
- Loads master lorry list and historical trip data
- Supports multiple history sources (old format + new ZSDOROUTEWRH format)
- Suggests best lorry per DO based on:
    1. Eligibility (owner/SPARE) and capacity >= load
    2. Closest TON fit (least wasted capacity)
    3. Tiebreaker A: ROUTE + CUSTOMER familiarity (most specific match)
    4. Tiebreaker B: ROUTE-only familiarity (broader match)
"""

import os
import pandas as pd


class LorryEngine:
    def __init__(self, master_path: str, history_path: str, owner_user: str,
                 ref_history_path: str = None):
        """
        master_path      — master lorry list (.xlsx)
        history_path     — primary working history (ZSDOROUTEWRH-bot.xlsx)
        owner_user       — user ID (e.g. "ABI") — only their lorries + SPARE eligible
        ref_history_path — read-only reference history (ZSDOROUTEWRH.xls from user)
                           Merged with history_path for richer frequency analysis.
        """
        self.owner_user = owner_user.upper()
        self._load_master(master_path)
        self._load_history(history_path, ref_history_path)
        self._build_frequency()

    # ── Data loading ──────────────────────────────────────────────────────────

    def _load_master(self, path):
        df = pd.read_excel(path)
        df.columns = [c.strip().upper() for c in df.columns]
        df['USER'] = df['USER'].str.strip().str.upper()
        allowed_users = {self.owner_user, 'SPARE'}
        self.eligible_lorries = df[df['USER'].isin(allowed_users)].copy()
        self.all_lorries = df.copy()

    def _load_one_history(self, path: str) -> pd.DataFrame | None:
        """Load one history file; normalise columns; return None on failure."""
        if not path or not os.path.exists(path):
            return None
        try:
            df = pd.read_excel(path)
        except Exception as e:
            print(f"LorryEngine: could not read {path}: {e}")
            return None

        df.columns = [c.strip().upper() for c in df.columns]

        # GROSS WEIGHT (kg) -> WEIGHT(T) for new ZSDOROUTEWRH format
        if 'GROSS WEIGHT' in df.columns and 'WEIGHT(T)' not in df.columns:
            df['WEIGHT(T)'] = (
                pd.to_numeric(df['GROSS WEIGHT'], errors='coerce').fillna(0) / 1000.0
            )

        # Normalise customer column name across formats
        for cand in ('CUSTOMER NAME', 'CUSTOMER', 'CUST NAME', 'CUST'):
            if cand in df.columns and 'CUSTOMER NAME' not in df.columns:
                df['CUSTOMER NAME'] = df[cand]
                break
        if 'CUSTOMER NAME' not in df.columns:
            df['CUSTOMER NAME'] = ''

        df = df.dropna(subset=['ROUTE', 'LICENSE']).copy()
        df['ROUTE']         = df['ROUTE'].astype(str).str.strip()
        df['LICENSE']       = df['LICENSE'].astype(str).str.strip().str.upper()
        df['CUSTOMER NAME'] = df['CUSTOMER NAME'].astype(str).str.strip()
        return df

    def _load_history(self, primary_path: str, ref_path: str = None):
        """Merge primary working history + optional read-only reference history."""
        parts = []
        for p in [primary_path, ref_path]:
            df = self._load_one_history(p)
            if df is not None and not df.empty:
                parts.append(df)

        if not parts:
            self.history = pd.DataFrame(
                columns=['ROUTE', 'LICENSE', 'CUSTOMER NAME'])
        else:
            self.history = pd.concat(parts, ignore_index=True)

    # ── Frequency tables ──────────────────────────────────────────────────────

    def _build_frequency(self):
        """
        Two frequency tables:
          route_customer_freq — ROUTE + CUSTOMER NAME + LICENSE: most specific
          route_freq          — ROUTE + LICENSE: broader route familiarity
        """
        if self.history.empty:
            self.route_customer_freq = pd.DataFrame(
                columns=['ROUTE', 'CUSTOMER NAME', 'LICENSE', 'FREQ'])
            self.route_freq = pd.DataFrame(
                columns=['ROUTE', 'LICENSE', 'FREQ'])
            return

        self.route_customer_freq = (
            self.history
            .groupby(['ROUTE', 'CUSTOMER NAME', 'LICENSE'])
            .size()
            .reset_index(name='FREQ')
        )
        self.route_freq = (
            self.history
            .groupby(['ROUTE', 'LICENSE'])
            .size()
            .reset_index(name='FREQ')
        )

    def get_route_frequencies(self, route: str) -> pd.DataFrame:
        subset = self.route_freq[self.route_freq['ROUTE'] == route].copy()
        return subset.sort_values('FREQ', ascending=False)

    def get_route_customer_frequencies(self, route: str,
                                       customer: str = '') -> pd.DataFrame:
        mask = self.route_customer_freq['ROUTE'] == route
        if customer:
            mask &= (self.route_customer_freq['CUSTOMER NAME']
                     .str.upper() == customer.strip().upper())
        return self.route_customer_freq[mask].copy().sort_values(
            'FREQ', ascending=False)

    # ── Suggestion ────────────────────────────────────────────────────────────

    def suggest(self, route: str, total_ton: float,
                customer: str = '',
                unavailable: set = None, top_n: int = 3) -> list[dict]:
        """
        Sort key:  SURPLUS ASC  →  CUST_FREQ DESC  →  ROUTE_FREQ DESC

        CUST_FREQ  = times this lorry served THIS route + THIS customer
        ROUTE_FREQ = times this lorry served THIS route (any customer)

        Tightest capacity fit wins first. Among equal-surplus lorries, prefer
        the one most familiar with this specific customer, then the route.
        """
        if unavailable is None:
            unavailable = set()

        eligible = self.eligible_lorries[
            (self.eligible_lorries['TON'] >= total_ton) &
            (~self.eligible_lorries['LORRY'].isin(unavailable))
        ].copy()

        if eligible.empty:
            return []

        eligible['SURPLUS'] = eligible['TON'] - total_ton

        # Route+customer frequency
        rc_freq = self.get_route_customer_frequencies(route, customer)
        merged = eligible.merge(
            rc_freq[['LICENSE', 'FREQ']].rename(columns={'FREQ': 'CUST_FREQ'}),
            left_on='LORRY', right_on='LICENSE', how='left'
        ) if not rc_freq.empty else eligible.assign(CUST_FREQ=0)
        merged['CUST_FREQ'] = merged['CUST_FREQ'].fillna(0).astype(int)

        # Route-only frequency
        r_freq = self.get_route_frequencies(route)
        merged = merged.merge(
            r_freq[['LICENSE', 'FREQ']].rename(columns={'FREQ': 'ROUTE_FREQ'}),
            left_on='LORRY', right_on='LICENSE', how='left'
        ) if not r_freq.empty else merged.assign(ROUTE_FREQ=0)
        merged['ROUTE_FREQ'] = merged['ROUTE_FREQ'].fillna(0).astype(int)

        merged = merged.sort_values(
            ['SURPLUS', 'CUST_FREQ', 'ROUTE_FREQ'],
            ascending=[True, False, False]
        )

        results = []
        for _, row in merged.head(top_n).iterrows():
            surplus    = round(float(row['SURPLUS']), 2)
            cust_freq  = int(row['CUST_FREQ'])
            route_freq = int(row['ROUTE_FREQ'])

            if cust_freq > 0:
                reason = (f"Closest fit ({surplus}T spare) — "
                          f"served this customer on this route {cust_freq}\u00d7")
            elif route_freq > 0:
                reason = (f"Closest fit ({surplus}T spare) — "
                          f"used on this route {route_freq}\u00d7 (other customers)")
            else:
                reason = f"Closest fit ({surplus}T spare) — no prior route history"

            results.append({
                'LORRY':        row['LORRY'],
                'TON_CAPACITY': round(float(row['TON']), 2),
                'SURPLUS':      surplus,
                'USER':         row['USER'],
                'CUST_FREQ':    cust_freq,
                'ROUTE_FREQ':   route_freq,
                'FREQ':         max(cust_freq, route_freq),  # backward-compat
                'REASON':       reason,
            })
        return results

    def get_eligible_lorry_list(self) -> pd.DataFrame:
        return self.eligible_lorries[['LORRY', 'TON', 'USER']].copy()

    def suggest_split(self, route: str, total_ton: float,
                      customer: str = '',
                      unavailable: set = None, max_lorries: int = 6,
                      single_util_threshold: float = 0.70) -> list[dict] | None:
        """Split across smaller lorries if it gives a tighter fit."""
        if unavailable is None:
            unavailable = set()

        rc_freq = self.get_route_customer_frequencies(route, customer)
        r_freq  = self.get_route_frequencies(route)

        def _enrich(df: pd.DataFrame) -> pd.DataFrame:
            d = df.copy()
            if not rc_freq.empty:
                d = d.merge(rc_freq[['LICENSE', 'FREQ']].rename(
                    columns={'FREQ': 'CUST_FREQ'}),
                    left_on='LORRY', right_on='LICENSE', how='left')
            else:
                d['CUST_FREQ'] = 0
            d['CUST_FREQ'] = d['CUST_FREQ'].fillna(0).astype(int)
            if not r_freq.empty:
                d = d.merge(r_freq[['LICENSE', 'FREQ']].rename(
                    columns={'FREQ': 'ROUTE_FREQ'}),
                    left_on='LORRY', right_on='LICENSE', how='left')
            else:
                d['ROUTE_FREQ'] = 0
            d['ROUTE_FREQ'] = d['ROUTE_FREQ'].fillna(0).astype(int)
            return d

        # Best single-lorry baseline
        eligible = self.eligible_lorries[
            (self.eligible_lorries['TON'] >= total_ton) &
            (~self.eligible_lorries['LORRY'].isin(unavailable))
        ].copy()

        if eligible.empty:
            best_single_surplus = float('inf')
            best_single_cap     = float('inf')
        else:
            enriched = _enrich(eligible)
            enriched['SURPLUS'] = enriched['TON'] - total_ton
            enriched = enriched.sort_values(
                ['SURPLUS', 'CUST_FREQ', 'ROUTE_FREQ'],
                ascending=[True, False, False])
            best_row            = enriched.iloc[0]
            best_single_surplus = float(best_row['SURPLUS'])
            best_single_cap     = float(best_row['TON'])
            if (total_ton / best_single_cap) >= single_util_threshold:
                return None

        # Small-lorry pool
        small_pool = self.eligible_lorries[
            (self.eligible_lorries['TON'] < best_single_cap) &
            (~self.eligible_lorries['LORRY'].isin(unavailable))
        ].copy()
        if small_pool.empty:
            return None

        small_pool = _enrich(small_pool)
        small_pool = small_pool.sort_values(
            ['CUST_FREQ', 'ROUTE_FREQ', 'TON'],
            ascending=[False, False, False])

        remain = total_ton
        used   = set(unavailable)
        chosen = []

        for _, row in small_pool.iterrows():
            if remain <= 0:
                break
            if row['LORRY'] in used:
                continue
            cap       = float(row['TON'])
            portion   = round(min(cap, remain), 6)
            surplus   = round(cap - portion, 2)
            cust_freq = int(row['CUST_FREQ'])
            r_freq_v  = int(row['ROUTE_FREQ'])
            if cust_freq > 0:
                reason = f"Split — {surplus}T spare, {cust_freq}\u00d7 this customer+route"
            elif r_freq_v > 0:
                reason = f"Split — {surplus}T spare, {r_freq_v}\u00d7 this route"
            else:
                reason = f"Split — {surplus}T spare, no route history"
            chosen.append({
                'LORRY':        row['LORRY'],
                'TON_CAPACITY': round(cap, 2),
                'SURPLUS':      surplus,
                'USER':         str(row['USER']),
                'CUST_FREQ':    cust_freq,
                'ROUTE_FREQ':   r_freq_v,
                'FREQ':         max(cust_freq, r_freq_v),
                'REASON':       reason,
                'PORTION':      portion,
            })
            used.add(row['LORRY'])
            remain = round(remain - cap, 6)
            if len(chosen) >= max_lorries:
                break

        if remain > 0:
            return None
        if round(sum(c['SURPLUS'] for c in chosen), 2) >= best_single_surplus:
            return None
        return chosen
