"""
Lorry Assignment Engine
- Loads master lorry list and historical trip data
- Suggests best lorry per DO based on: user ownership, route frequency, closest TON fit
"""

import pandas as pd


class LorryEngine:
    def __init__(self, master_path: str, history_path: str, owner_user: str):
        self.owner_user = owner_user.upper()
        self._load_master(master_path)
        self._load_history(history_path)
        self._build_route_frequency()

    def _load_master(self, path):
        df = pd.read_excel(path)
        df.columns = [c.strip().upper() for c in df.columns]
        df['USER'] = df['USER'].str.strip().str.upper()
        allowed_users = {self.owner_user, 'SPARE'}
        self.eligible_lorries = df[df['USER'].isin(allowed_users)].copy()
        self.all_lorries = df.copy()

    def _load_history(self, path):
        df = pd.read_excel(path)
        df.columns = [c.strip().upper() for c in df.columns]
        # Support both old format (WEIGHT(T), ITMREF_0) and new ZSDOROUTEWRH format
        # New format has GROSS WEIGHT (kg) instead of WEIGHT(T)
        if 'GROSS WEIGHT' in df.columns and 'WEIGHT(T)' not in df.columns:
            df['WEIGHT(T)'] = pd.to_numeric(df['GROSS WEIGHT'], errors='coerce').fillna(0) / 1000.0
        self.history = df.dropna(subset=['ROUTE', 'LICENSE']).copy()

    def _build_route_frequency(self):
        freq = (
            self.history.groupby(['ROUTE', 'LICENSE'])
            .size()
            .reset_index(name='FREQ')
        )
        self.route_freq = freq

    def get_route_frequencies(self, route: str) -> pd.DataFrame:
        """Exact → case-insensitive → 5-char prefix match for new format compatibility."""
        route_s = route.strip()
        subset = self.route_freq[self.route_freq['ROUTE'].str.strip() == route_s].copy()
        if not subset.empty:
            return subset.sort_values('FREQ', ascending=False)
        subset = self.route_freq[
            self.route_freq['ROUTE'].str.strip().str.upper() == route_s.upper()
        ].copy()
        if not subset.empty:
            return subset.sort_values('FREQ', ascending=False)
        prefix = route_s[:5].upper()
        if len(prefix) >= 4:
            subset = self.route_freq[
                self.route_freq['ROUTE'].str.strip().str.upper().str.startswith(prefix)
            ].copy()
            if not subset.empty:
                return subset.sort_values('FREQ', ascending=False)
        return pd.DataFrame(columns=['ROUTE', 'LICENSE', 'FREQ'])

    def suggest(self, route: str, total_ton: float,
                unavailable: set = None, top_n: int = 3) -> list[dict]:
        """
        Ranking logic — Option B: Tight fit first, route frequency as tiebreaker.
        Priority order:
        1. Must be eligible (owner/SPARE) and TON >= total_ton
        2. Must not be in unavailable set
        3. Sort by: TON_SURPLUS ASC (closest fit — least wasted capacity)
        4. Tiebreak: FREQ DESC (prefer lorry that knows the route better)
           e.g. for 7.5T load: 8T lorry ranked above 13T lorry even if 13T
           has more route history, unless both have identical surplus.
        """
        if unavailable is None:
            unavailable = set()

        eligible = self.eligible_lorries[
            (self.eligible_lorries['TON'] >= total_ton) &
            (~self.eligible_lorries['LORRY'].isin(unavailable))
        ].copy()

        if eligible.empty:
            return []

        # Add surplus column: how much extra capacity (smaller = tighter fit)
        eligible = eligible.copy()
        eligible['SURPLUS'] = eligible['TON'] - total_ton

        # Merge with route frequency
        freq_on_route = self.get_route_frequencies(route)
        merged = eligible.merge(freq_on_route[['LICENSE', 'FREQ']],
                                left_on='LORRY', right_on='LICENSE', how='left')
        merged['FREQ'] = merged['FREQ'].fillna(0).astype(int)

        # Sort: 1) tightest fit  2) owner before SPARE  3) route familiarity
        merged['IS_OWNER'] = (merged['USER'].str.upper() == self.owner_user).astype(int)
        merged = merged.sort_values(
            ['SURPLUS', 'IS_OWNER', 'FREQ'], ascending=[True, False, False])

        results = []
        for _, row in merged.head(top_n).iterrows():
            surplus = round(float(row['SURPLUS']), 2)
            if row['FREQ'] > 0:
                reason = (f"Closest fit ({surplus}T spare) — "
                          f"used on this route {int(row['FREQ'])} time(s)")
            else:
                reason = f"Closest fit ({surplus}T spare) — no prior route history"
            results.append({
                'LORRY': row['LORRY'],
                'TON_CAPACITY': round(float(row['TON']), 2),
                'SURPLUS': surplus,
                'USER': row['USER'],
                'FREQ': int(row['FREQ']),
                'REASON': reason,
            })
        return results

    def get_eligible_lorry_list(self) -> pd.DataFrame:
        return self.eligible_lorries[['LORRY', 'TON', 'USER']].copy()

    def suggest_split(self, route: str, total_ton: float,
                      unavailable: set = None, max_lorries: int = 6,
                      single_util_threshold: float = 0.70) -> list[dict] | None:
        """
        Checks whether splitting across smaller lorries gives a tighter fit
        than the best single lorry.

        Strategy:
          1. Find the best single-lorry surplus (baseline) and its utilization.
          2. If single lorry utilization >= single_util_threshold (default 70%),
             do NOT split — a well-utilized single lorry is always preferred.
          3. Try packing using ONLY lorries smaller than the best single option
             (sorted by route frequency DESC so familiar lorries go first).
          4. If the combined surplus of the small-lorry split is strictly less
             than the single-lorry surplus, return the split.
          5. Otherwise return None so caller uses single lorry.
        """
        if unavailable is None:
            unavailable = set()

        freq_on_route = self.get_route_frequencies(route)

        def _enrich(df: pd.DataFrame, min_ton: float) -> pd.DataFrame:
            d = df.copy()
            d["SURPLUS"] = d["TON"] - min_ton
            d = d.merge(freq_on_route[["LICENSE", "FREQ"]],
                        left_on="LORRY", right_on="LICENSE", how="left")
            d["FREQ"] = d["FREQ"].fillna(0).astype(int)
            return d

        # Step 1 — best single lorry
        eligible = self.eligible_lorries[
            (self.eligible_lorries["TON"] >= total_ton) &
            (~self.eligible_lorries["LORRY"].isin(unavailable))
        ].copy()

        if eligible.empty:
            best_single_surplus = float("inf")
            best_single_cap     = float("inf")
        else:
            enriched = _enrich(eligible, total_ton)
            enriched["IS_OWNER"] = (enriched["USER"].str.upper() == self.owner_user).astype(int)
            enriched = enriched.sort_values(
                ["SURPLUS", "IS_OWNER", "FREQ"], ascending=[True, False, False])
            best_row            = enriched.iloc[0]
            best_single_surplus = float(best_row["SURPLUS"])
            best_single_cap     = float(best_row["TON"])
            # If single lorry is already well-utilized, never split
            single_util = total_ton / best_single_cap if best_single_cap > 0 else 0
            if single_util >= single_util_threshold:
                return None

        # Step 2 — small-lorry pool: only lorries SMALLER than best single option
        # (so we never accidentally pick the same big lorry inside the split)
        small_pool = self.eligible_lorries[
            (self.eligible_lorries["TON"] < best_single_cap) &
            (~self.eligible_lorries["LORRY"].isin(unavailable))
        ].copy()

        if small_pool.empty:
            return None  # nothing smaller to split with

        small_pool = _enrich(small_pool, 0)
        # Sort: owner first, then highest frequency, then largest lorry
        small_pool["IS_OWNER"] = (small_pool["USER"].str.upper() == self.owner_user).astype(int)
        small_pool = small_pool.sort_values(
            ["IS_OWNER", "FREQ", "TON"], ascending=[False, False, False])

        # Greedy fill: keep taking lorries until load is covered
        remain = total_ton
        used   = set(unavailable)
        chosen = []

        for _, row in small_pool.iterrows():
            if remain <= 0:
                break
            if row["LORRY"] in used:
                continue
            lorry   = row["LORRY"]
            cap     = float(row["TON"])
            portion = round(min(cap, remain), 6)
            surplus = round(cap - portion, 2)
            freq    = int(row["FREQ"])
            reason  = (f"Split load — {surplus}T spare"
                       + (f", {freq} trips on route" if freq > 0 else ", no route history"))
            chosen.append({
                "LORRY":        lorry,
                "TON_CAPACITY": round(cap, 2),
                "SURPLUS":      surplus,
                "USER":         str(row["USER"]),
                "FREQ":         freq,
                "REASON":       reason,
                "PORTION":      portion,
            })
            used.add(lorry)
            remain = round(remain - cap, 6)

            if len(chosen) >= max_lorries:
                break

        # If small lorries can not cover the full load, abort
        if remain > 0:
            return None

        # Step 3 — only return split if it's strictly tighter than single lorry
        combined_surplus = round(sum(c["SURPLUS"] for c in chosen), 2)
        if combined_surplus >= best_single_surplus:
            return None

        return chosen
