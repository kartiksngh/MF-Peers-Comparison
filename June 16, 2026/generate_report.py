"""
generate_report.py
==================
Monthly automation for the Peer Performance Monitor.

Reads:
    MFI Data/HistoricalNav_*.xlsx       — NAV history per scheme
    MFI Data/Scheme wise AUM Report*.xlsx — AUM history
    MFI Data/Map.xlsx                    — MFI category mapping
    Value Research Data/Value Research Benchmarks NAV*.xlsx
                                         — VR exact-peer mapping + benchmark NAVs

Emits:
    out/percent_AUM_in_Q1_to_Q4 - <DATE>.xlsx        (internal sharing)
    out/Scheme_Scoring_on_Exact_Peer_Set - <DATE>.xlsx (internal sharing)
    out/dashboard_data.json                          (data powering the dashboard)
    out/dashboard.html                               (interactive dashboard — host on intranet)
    out/dashboard_offline.html                       (single self-contained file — email / double-click)

Run:
    python generate_report.py
    python generate_report.py --mfi "MFI Data" --vr "Value Research Data" --out out

One command does everything: ingest → analyse → Excel → JSON → both dashboards.
The dashboard template is embedded in this script (see bottom), so nothing else
is needed. To use a customised template instead, pass --dashboard-html my.html.

The analysis logic is a clean refactor of the notebook
"All_Peer_Exact_Peer_based_Q1_to_Q4_and_Scores".
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
from datetime import datetime
from functools import reduce
from pathlib import Path

import numpy as np
import pandas as pd
from joblib import Parallel, delayed


# ═════════════════════════════════════════════════════════════════════
# 1. DATA LOADING
# ═════════════════════════════════════════════════════════════════════

ENGINE = "calamine"  # falls back to openpyxl if not installed; swap in CLI if needed


def _read_nav(path: str | Path) -> pd.DataFrame:
    """NAV xlsx → DataFrame (dates × schemes, '--' and 0 → NaN, ffilled)."""
    df = pd.read_excel(path, skiprows=5, header=0, index_col=0, engine=ENGINE)
    df = df.iloc[3:-3]
    df.index = pd.to_datetime(df.index, format="%m/%d/%Y", errors="coerce")
    df = df.loc[df.index.notna()]
    df.index.name = "Date"
    return df.apply(pd.to_numeric, errors="coerce")


def _read_aum(path: str | Path) -> pd.DataFrame:
    """AUM xlsx → long DataFrame [Fund Name, Scheme Name, Date, Value]."""
    df = pd.read_excel(path, skiprows=5, header=0, engine=ENGINE)
    df = df.iloc[:-3][["Fund Name", "Scheme Name", "Date", "Value"]].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    return df.dropna(subset=["Date"])


def load_mfi(mfi_folder: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Load MFI NAV + AUM + Map. Returns (nav, aum_wide, category_map)."""
    files = [f for f in os.listdir(mfi_folder)
             if f.endswith(".xlsx") and not f.startswith("~")]
    nav_paths = sorted(mfi_folder / f for f in files if "nav" in f.lower())
    aum_paths = sorted(mfi_folder / f for f in files if "aum" in f.lower())

    print(f"  NAV files: {[p.name for p in nav_paths]}")
    print(f"  AUM files: {[p.name for p in aum_paths]}")

    print("  Reading NAV...")
    nav_dfs = [_read_nav(p) for p in nav_paths]
    nav = pd.concat(nav_dfs, axis=1).replace(0, np.nan).sort_index().ffill()
    print(f"    {nav.shape[1]} schemes × {len(nav)} dates")

    print("  Reading AUM...")
    aum_long = pd.concat([_read_aum(p) for p in aum_paths]).drop_duplicates()
    aum = aum_long.pivot_table(values="Value", index="Scheme Name", columns="Date")
    aum = aum.drop([i for i in aum.index if "Adjusted" in i or "Segregated" in i])
    print(f"    {len(aum)} schemes × {aum.shape[1]} months")

    cmap = pd.read_excel(mfi_folder / "Map.xlsx", engine=ENGINE)
    print(f"  Map: {len(cmap)} schemes")
    return nav, aum, cmap


# ═════════════════════════════════════════════════════════════════════
# 2. CATEGORY MAPPING
# ═════════════════════════════════════════════════════════════════════

BREAKDOWN_RULES = {
    "Bottoms Up": ["Large & Mid Cap Fund", "Small cap Fund", "Mid Cap Fund", "Multi Cap Fund"],
    "Top Down": ["Large Cap Fund", "Flexi Cap Fund", "Focused Fund", "ELSS"],
    "Asset Allocation": ["Aggressive Hybrid Fund", "Balanced Hybrid Fund",
                         "Dynamic Asset Allocation or Balanced Advantage",
                         "Conservative Hybrid Fund", "Multi Asset Allocation"],
    "Arbitrage+": ["Equity Savings", "Arbitrage Fund"],
}
MULTI_ASSET_DROP = ["WhiteOak", "Edelweiss", "Mahindra"]  # debt taxation — exclude


def build_category_map(raw_map: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Clean raw MFI Map → (category_map with Breakdown/fund house, ordered categories)."""
    cm = raw_map.copy()
    cm.index = cm.pop("Scheme Name")
    cm = cm[["Scheme Sub Nature", "Category"]].copy()
    cm["Category"] = cm["Category"].fillna("Unknown").astype(str)

    # Merge Value + Contra → Value/Contra
    cm.loc[cm["Category"].isin(["Contra Fund", "Value Fund"]), "Category"] = "Value/Contra"

    categories = sorted(set(cm["Category"]))
    thematic_cats = set(cm.loc[cm["Scheme Sub Nature"] == "Thematic", "Category"])
    sectoral_cats = set(cm.loc[cm["Scheme Sub Nature"] == "Sectoral", "Category"])
    thematic_cats -= sectoral_cats
    # Sectoral and thematic go at the end
    categories = ([c for c in categories if c not in thematic_cats | sectoral_cats]
                  + sorted(sectoral_cats) + sorted(thematic_cats))

    cm["fund house"] = cm.index.to_series().str.split(" ").str[0]

    # Breakdown
    assigned = set()
    for breakdown, subnatures in BREAKDOWN_RULES.items():
        mask = cm["Scheme Sub Nature"].isin(subnatures)
        cm.loc[mask, "Breakdown"] = breakdown
        assigned |= set(subnatures)
    # Everything else is Thematic
    cm.loc[~cm["Scheme Sub Nature"].isin(assigned), "Breakdown"] = "Thematic"

    # Drop multi-asset schemes with debt-like taxation
    drop = cm.loc[(cm["Scheme Sub Nature"] == "Multi Asset Allocation")
                  & (cm["fund house"].isin(MULTI_ASSET_DROP))].index
    cm = cm.drop(drop)

    # Keep only categories with ≥4 fund houses
    counts = cm.reset_index().groupby("Category")["fund house"].count()
    valid_cats = set(counts[counts >= 4].index)
    categories = [c for c in categories if c in valid_cats]
    return cm, categories


def load_vr(vr_folder: Path):
    """Load Value Research exact-peer mapping + benchmark NAVs."""
    # Find the latest VR file
    vr_files = sorted([f for f in os.listdir(vr_folder)
                       if "Value Research Benchmarks NAV" in f and f.endswith(".xlsx")])
    if not vr_files:
        raise FileNotFoundError(f"No Value Research Benchmarks file in {vr_folder}")
    vr_path = vr_folder / vr_files[-1]
    print(f"  Using VR file: {vr_path.name}")

    cme = pd.read_excel(vr_path, sheet_name="Ret. Compr.(Equity) - Dir", engine=ENGINE)
    cme.index = cme.pop("Scheme")
    missing_in_mfi = cme.loc[pd.isna(cme["Scheme Name From MFI"])
                             & (cme["AMFI Code"] != "bench")].index.tolist()
    if missing_in_mfi:
        print(f"    Dropping {len(missing_in_mfi)} VR schemes with no MFI NAV match")
        cme = cme.drop(missing_in_mfi)
    cme["fund house"] = cme.index.to_series().str.split(" ").str[0]

    # Benchmarks
    bench = cme.loc[cme["AMFI Code"] == "bench", ["Category"]]
    bench = bench.reset_index().set_index("Category")
    bench.columns = ["Benchmark"]
    cme = cme.drop(bench["Benchmark"])
    cme = cme.reset_index().set_index("Scheme Name From MFI")

    # Benchmark NAVs
    sheets = ["Bloomberg Bench NAV", "Crisil Bench NAV", "VR Bench NAV", "Silver"]
    bench_dfs = []
    for sh in sheets:
        try:
            bdf = pd.read_excel(vr_path, sheet_name=sh, engine=ENGINE)
            date_col = next(c for c in ["Dates", "Date", "NAV Date"] if c in bdf.columns)
            bdf = bdf.set_index(date_col)
            bench_dfs.append(bdf)
        except Exception as e:
            print(f"    Warn: sheet '{sh}' skipped — {e}")
    bench_nav = reduce(lambda l, r: pd.merge(l, r, left_index=True, right_index=True, how="left"),
                       bench_dfs) if bench_dfs else pd.DataFrame()

    try:
        name_map = pd.read_excel(vr_path, sheet_name="Benchmarks and sources", engine=ENGINE)
        name_map = name_map.set_index("Identifier")
        bench_nav.columns = [name_map.loc[c, "Benchmarks"] if c in name_map.index else c
                             for c in bench_nav.columns]
    except Exception:
        pass

    return cme, bench, bench_nav


# ═════════════════════════════════════════════════════════════════════
# 3. AUM SHARE per scheme (daily, ffilled from monthly)
# ═════════════════════════════════════════════════════════════════════

def percent_aum_daily(aum: pd.DataFrame, nav_index: pd.Index,
                       cmap: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Daily ffilled scheme AUM, and each scheme's share of its fund house's AUM.
    Returns (aum_filled_long, percent_aum_with_house) where rows are schemes."""
    aum1 = aum.copy().T.sort_index().replace("--", np.nan).ffill().T
    aum2 = aum1.iloc[:, 1:].T   # dates × schemes
    aum1.insert(0, "fund house", aum1.index.to_series().str.split().str[0])
    fund_aum = aum1.groupby("fund house").sum()
    houses = list(fund_aum.index)

    pct_df = pd.DataFrame()
    for fund in houses:
        sub = aum1.loc[aum1["fund house"] == fund].iloc[:, 1:]
        denom = fund_aum.loc[fund]
        pct = (sub / denom).T  # dates × schemes
        # Re-index to daily nav index (ffill)
        pct_daily = pd.DataFrame([1.0] * len(nav_index), index=nav_index)
        pct_daily = pct_daily.merge(pct, left_index=True, right_index=True, how="outer") \
                              .ffill().drop(0, axis=1)
        pct_daily = pct_daily.loc[nav_index]
        pct_df = pct_df.merge(pct_daily, left_index=True, right_index=True, how="outer")

    pct_with_house = cmap[["fund house"]].merge(pct_df.T, left_index=True, right_index=True, how="right")
    return aum2, pct_with_house


# ═════════════════════════════════════════════════════════════════════
# 4. QUARTILE ANALYSIS (parallel per category)
# ═════════════════════════════════════════════════════════════════════

def _process_category(cat: str, fs: list[str], nav: pd.DataFrame,
                       rets_1y: pd.DataFrame, rets_3y: pd.DataFrame,
                       aum2: pd.DataFrame):
    """One category at a time. Returns dict with quartile assignments + distances."""
    # Defensive: drop scheme names that don't actually exist as NAV columns. This
    # happens when the VR exact-peer mapping references schemes that are absent
    # from the final NAV universe (whitespace mismatch, dropped by MULTI_ASSET_DROP,
    # not in AUM, etc.).
    available = set(rets_1y.columns)
    fs = [f for f in fs if f in available]
    if not fs:
        # Nothing usable for this category — return empty frames
        return dict(cat=cat, fs=[],
                    qy_1y=pd.DataFrame(index=rets_1y.index),
                    qy_3y=pd.DataFrame(index=rets_3y.index),
                    dist_1y=pd.DataFrame(), dist_3y=pd.DataFrame(),
                    pct_1y=pd.DataFrame(), pct_3y=pd.DataFrame())

    f1, f3 = rets_1y[fs], rets_3y[fs]

    dates = f3.loc[pd.isna(f3).sum(axis=1) != len(f3.columns)].index
    dates = [d for d in dates if d > aum2.index.min()]

    qy_1y = pd.DataFrame(index=f1.index, columns=fs)
    qy_3y = pd.DataFrame(index=f3.index, columns=fs)
    dist_1y = pd.DataFrame()
    dist_3y = pd.DataFrame()

    def split_quartiles(sorted_series):
        n = len(sorted_series)
        base, extra = divmod(n, 4)
        sizes = [base + (1 if i < extra else 0) for i in range(4)]
        out, start = [], 0
        for sz in sizes:
            out.append(list(sorted_series.iloc[start:start+sz].index))
            start += sz
        return out  # [q1, q2, q3, q4]

    for d in dates:
        for label, f_rets, qy, dist in [("1y", f1, qy_1y, dist_1y), ("3y", f3, qy_3y, dist_3y)]:
            avail = f_rets.loc[d].dropna()
            if avail.empty:
                continue
            srt = avail.sort_values(ascending=False)
            q1, q2, q3, q4 = split_quartiles(srt)
            qy.loc[d, q1] = "q1"; qy.loc[d, q2] = "q2"
            qy.loc[d, q3] = "q3"; qy.loc[d, q4] = "q4"

            abs_fn = [i for i in srt.index if "Aditya" in i]
            dist.loc[d, "Category"] = cat
            dist.loc[d, "category top return"] = srt.iloc[0]
            dist.loc[d, "category bottom return"] = srt.iloc[-1]
            if len(srt) > 2: dist.loc[d, "category mean return"] = srt.iloc[len(srt)//2]
            if len(srt) > 3 and q1: dist.loc[d, "category Q1 return"] = srt[q1].iloc[-1]
            if abs_fn:
                ar = srt[abs_fn].iloc[0]
                dist.loc[d, "category Birla scheme return"] = ar
                behind = srt.loc[ar >= srt].index
                dist.loc[d, "% of AUM in category Birla is outperforming"] = (
                    aum2.loc[:d, behind].iloc[-1].sum()
                    / aum2.loc[:d, srt.index].iloc[-1].sum()
                )
                dist.loc[d, "% of funds in category Birla is outperforming"] = len(behind) / len(srt)
            dist.loc[d, "category top scheme"] = srt.index[0]
            dist.loc[d, "category bottom scheme"] = srt.index[-1]
            if len(srt) > 2:
                dist.loc[d, "category mean scheme"] = srt.index[len(srt)//2 + (len(srt) % 2)]
            if q1:
                dist.loc[d, "category Q1 scheme"] = srt[q1].index[-1]
            if abs_fn:
                dist.loc[d, "category Birla scheme"] = abs_fn[0]

    # Rolling percentile within category
    f1_per = f1.rank(axis=1).div(f1.notna().sum(axis=1), axis=0)
    f3_per = f3.rank(axis=1).div(f3.notna().sum(axis=1), axis=0)

    return dict(cat=cat, fs=fs,
                qy_1y=qy_1y, qy_3y=qy_3y,
                dist_1y=dist_1y, dist_3y=dist_3y,
                pct_1y=f1_per, pct_3y=f3_per)


def run_quartile_analysis(nav, cmap, categories, aum2, member_lookup=None):
    """Run the parallel quartile analysis. member_lookup: cat → list of schemes
    (default: use cmap['Category']). Returns combined dataframes."""
    rets_1y = nav / nav.shift(250) - 1
    rets_3y = (nav / nav.shift(250 * 3)) ** (1/3) - 1

    if member_lookup is None:
        member_lookup = {c: list(cmap[cmap["Category"] == c].index) for c in categories}

    results = Parallel(n_jobs=-1, verbose=1)(
        delayed(_process_category)(c, member_lookup[c], nav, rets_1y, rets_3y, aum2)
        for c in categories if member_lookup.get(c)
    )

    qy_1y = pd.DataFrame(index=nav.index, columns=nav.columns)
    qy_3y = pd.DataFrame(index=nav.index, columns=nav.columns)
    dist_1y = pd.DataFrame(); dist_3y = pd.DataFrame()
    pct_1y_all = pd.DataFrame(); pct_3y_all = pd.DataFrame()
    for r in results:
        qy_1y.loc[:, r["fs"]] = r["qy_1y"]
        qy_3y.loc[:, r["fs"]] = r["qy_3y"]
        dist_1y = pd.concat([dist_1y, r["dist_1y"]])
        dist_3y = pd.concat([dist_3y, r["dist_3y"]])
        pct_1y_all = pct_1y_all.merge(r["pct_1y"], left_index=True, right_index=True, how="outer")
        pct_3y_all = pct_3y_all.merge(r["pct_3y"], left_index=True, right_index=True, how="outer")
    return qy_1y, qy_3y, dist_1y, dist_3y, pct_1y_all, pct_3y_all


# ═════════════════════════════════════════════════════════════════════
# 5. AGGREGATION → SLEEVE & AMC DATAFRAMES
# ═════════════════════════════════════════════════════════════════════

def build_quartile_aum_tables(qy_1y, qy_3y, pct_with_house, cmap, top15):
    """Build the sleevewise and AMCwise %AUM in Q1-Q4 tables (same shape as the
    Excel sheets the user currently presents)."""

    def slice_q(qy, label):
        out = {}
        for q in ["q1", "q2", "q3", "q4"]:
            mask = qy.T == q
            df = mask * pct_with_house.iloc[:, 1:]
            df = df.copy()
            df["Quartile"] = f"% in Q{q[-1]}"
            df["Rolling Window"] = label
            out[q] = df
        return out

    p1, p3 = slice_q(qy_1y, "1 Year"), slice_q(qy_3y, "3 Year")
    pct_q_1y = pd.concat(p1.values())
    pct_q_3y = pd.concat(p3.values())

    # Filter empty rows
    def nonempty(df):
        date_cols = [c for c in df.columns if isinstance(c, pd.Timestamp)]
        return df.loc[df[date_cols].notna().any(axis=1)]

    pct_q_1y = nonempty(pct_q_1y)
    pct_q_3y = nonempty(pct_q_3y)

    # Merge with cmap
    cmeta = cmap[["Breakdown", "Category", "Scheme Sub Nature", "fund house"]]
    merged_1y = pct_q_1y.merge(cmeta, left_index=True, right_index=True)
    merged_3y = pct_q_3y.merge(cmeta, left_index=True, right_index=True)
    merged_1y["Scheme"] = merged_1y.index
    merged_3y["Scheme"] = merged_3y.index

    group_keys = ["Rolling Window", "Breakdown", "Category", "Scheme Sub Nature",
                  "fund house", "Scheme", "Quartile"]
    g1 = merged_1y.groupby(group_keys).sum(numeric_only=True)
    g3 = merged_3y.groupby(group_keys).sum(numeric_only=True)
    merged_quartile_df = pd.concat([g3, g1])
    truncate = pd.to_datetime("2013-12-31")
    merged_quartile_df = merged_quartile_df.loc[:,
        [c for c in merged_quartile_df.columns if isinstance(c, pd.Timestamp) and c >= truncate]
    ]

    # ── Aggregated sleeve & AMC tables ─────────────
    df = merged_quartile_df.copy().reset_index()
    # Latest AUM% per scheme
    latest_pct = pd.DataFrame(pct_with_house.iloc[:, -1]).reset_index()
    latest_pct.columns = ["Scheme", "% of Total AUM"]
    df = df.merge(latest_pct, on="Scheme")

    # Sleevewise = (Rolling Window, fund house, Breakdown, Quartile)
    visual_df = df.groupby(["Rolling Window", "fund house", "Breakdown", "Quartile"]).sum(numeric_only=True)
    visual_df = visual_df.drop(columns=["% of Total AUM"], errors="ignore")
    date_cols = [c for c in visual_df.columns if isinstance(c, pd.Timestamp)]
    visual_df = visual_df[date_cols]

    # Add the % of AMC AUM row
    aum_df = cmeta.merge(pct_with_house.iloc[:, 1:][visual_df.columns],
                         left_index=True, right_index=True, how="right")
    aum_df = aum_df.loc[aum_df["fund house"].isin(top15)]
    aum_df["Quartile"] = "% of AMC AUM"
    aum1 = aum_df.copy(); aum1["Rolling Window"] = "1 Year"
    aum3 = aum_df.copy(); aum3["Rolling Window"] = "3 Year"
    aum_combined = pd.concat([aum1, aum3])
    aum_total = aum_combined.groupby(["Rolling Window", "fund house",
                                       "Breakdown", "Quartile"]).sum(numeric_only=True)
    aum_total = aum_total[visual_df.columns]

    sleeve_df = pd.concat([visual_df, aum_total]).sort_index(
        level=["Rolling Window", "fund house", "Breakdown", "Quartile"])
    sleeve_df = sleeve_df.loc[sleeve_df.index.get_level_values("fund house").isin(top15)]

    # AMCwise = collapse Breakdown
    amc_df = sleeve_df.reset_index().groupby(
        ["fund house", "Rolling Window", "Quartile"]).sum(numeric_only=True)

    return merged_quartile_df, sleeve_df, amc_df


def build_distance_table(dist_1y, dist_3y, categories):
    """Build merged distance table with alpha columns."""
    dist_1y = dist_1y.copy(); dist_3y = dist_3y.copy()
    dist_1y["Rolling window"] = "1 Year"; dist_3y["Rolling window"] = "3 Year"
    for d in (dist_1y, dist_3y):
        d["Birla - category Average"] = d["category Birla scheme return"] - d.get("category mean return", np.nan)
        d["Birla - Category Top"]      = d["category Birla scheme return"] - d.get("category top return", np.nan)
        d["Birla - Category Bottom"]   = d["category Birla scheme return"] - d.get("category bottom return", np.nan)
        d["Birla - Q1"]                = d["category Birla scheme return"] - d.get("category Q1 return", np.nan)

    for c in categories:
        m1 = dist_1y["Category"] == c
        m3 = dist_3y["Category"] == c
        dist_1y.loc[m1, "% times Birla with +ve alpha over category average"] = (
            (dist_1y.loc[m1, "Birla - category Average"] > 0).rolling(250).sum() / 250
        )
        dist_3y.loc[m3, "% times Birla with +ve alpha over category average"] = (
            (dist_3y.loc[m3, "Birla - category Average"] > 0).rolling(250).sum() / 250
        )

    cols = ["Rolling window", "Category",
            "category top return", "category bottom return", "category mean return",
            "category Q1 return", "category Birla scheme return",
            "% of AUM in category Birla is outperforming",
            "% of funds in category Birla is outperforming",
            "Birla - category Average", "Birla - Category Top",
            "Birla - Category Bottom", "Birla - Q1",
            "% times Birla with +ve alpha over category average",
            "category top scheme", "category bottom scheme",
            "category mean scheme", "category Q1 scheme", "category Birla scheme"]
    merged = pd.concat([dist_3y, dist_1y])
    merged = merged[[c for c in cols if c in merged.columns]]
    return merged.sort_values(["Rolling window", "Category"])


# ═════════════════════════════════════════════════════════════════════
# 6. SCHEME-LEVEL SCORES (quartile assignment 4=Q1, 1=Q4)
# ═════════════════════════════════════════════════════════════════════

def quartile_to_score(qy_df: pd.DataFrame) -> pd.DataFrame:
    """Map q1→4, q2→3, q3→2, q4→1 (higher score = better). NaN passes through."""
    # Build a float result by stacking conditionals — avoids pandas' deprecated
    # silent dtype-downcasting in .replace() on object DataFrames.
    out = pd.DataFrame(np.nan, index=qy_df.index, columns=qy_df.columns, dtype=float)
    for label, score in {"q1": 4.0, "q2": 3.0, "q3": 2.0, "q4": 1.0}.items():
        out = out.where(qy_df != label, score)
    return out


def composite_score(score_1y: pd.DataFrame, score_3y: pd.DataFrame,
                    w_1y: float = 0.4, w_3y: float = 0.6) -> pd.DataFrame:
    """Weighted composite. Default 40/60 split (1Y / 3Y)."""
    return score_1y * w_1y + score_3y * w_3y


# ═════════════════════════════════════════════════════════════════════
# 7. EXPORTS
# ═════════════════════════════════════════════════════════════════════

def write_aum_workbook(path: Path, sleeve_df, amc_all_df, amc_vr_df):
    """Mirror the structure of the existing 'percent AUM in Q1 to Q4' file."""
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        sleeve_df.reset_index().to_excel(w, sheet_name="Sleevewise %AUM in Q1-Q4 - All", index=False)
        amc_all_df.reset_index().to_excel(w, sheet_name="AMCwise %AUM in Q1-Q4 - All", index=False)
        if amc_vr_df is not None:
            amc_vr_df.reset_index().to_excel(w, sheet_name="AMCwise %AUM in Q1-Q4 - VR peer", index=False)
    print(f"  Wrote {path}")


def write_scoring_workbook(path: Path, score_1y, score_3y, composite, peer_map):
    """Mirror the scheme scoring file. Each score sheet is dates × schemes."""
    def add_category_header(df, cat_lookup):
        """Add category as a top header row (cat_lookup: scheme → category)."""
        cats = [cat_lookup.get(c, "") for c in df.columns]
        out = pd.DataFrame([cats], columns=df.columns)
        return pd.concat([out, df])

    # peer_map already has scheme name, category; build cat_lookup from it
    cat_lookup = dict(zip(peer_map.index, peer_map["Category"])) if "Category" in peer_map.columns else {}

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, df in [("1Y score", score_1y), ("3Y score", score_3y),
                          ("Composite daily score", composite)]:
            add_category_header(df, cat_lookup).to_excel(w, sheet_name=name)
        peer_map.reset_index().to_excel(w, sheet_name="Peer & Bench Mapping", index=False)
    print(f"  Wrote {path}")


def write_dashboard_json(out_path: Path, sleeve_df, amc_all_df, amc_vr_df,
                          composite, score_1y, score_3y, peer_map):
    """Emit dashboard_data.json. Sample to month-ends."""
    def month_end_of(values):
        """Given any iterable of pandas Timestamps, return one per (year, month) — the latest."""
        by_ym = {}
        for v in values:
            if isinstance(v, pd.Timestamp):
                key = (v.year, v.month)
                if key not in by_ym or v > by_ym[key]:
                    by_ym[key] = v
        return sorted(by_ym.values())

    def df_to_records(df, idx_names, value_name, me_cols):
        recs = []
        for ix, row in df[me_cols].iterrows():
            r = {short: ix[i] for i, short in enumerate(idx_names)}
            r["v"] = [None if pd.isna(v) else round(float(v), 6) for v in row]
            recs.append(r)
        return recs

    # sleeve_df, amc_*_df: dates are in COLUMNS (after the groupby pivots)
    me_cols = month_end_of(sleeve_df.columns)
    date_strs = [d.date().isoformat() for d in me_cols]
    sleeve_records = df_to_records(sleeve_df, ["rw", "fh", "br", "qt"], "v", me_cols)
    amc_all_records = df_to_records(amc_all_df, ["fh", "rw", "qt"], "v", me_cols)
    amc_vr_records = df_to_records(amc_vr_df, ["fh", "rw", "qt"], "v", me_cols) if amc_vr_df is not None else []

    # composite / score_1y / score_3y: dates are in the INDEX, schemes in columns
    score_me_dates = month_end_of(composite.index)
    score_dates = [d.date().isoformat() for d in score_me_dates]

    def scheme_records(df, cat_lookup):
        out = []
        # Slice once for performance, then iterate scheme columns
        df_me = df.loc[score_me_dates] if score_me_dates else df.iloc[0:0]
        for sch in df.columns:
            vals = df_me[sch].values
            v = [None if pd.isna(x) else round(float(x), 3) for x in vals]
            if all(x is None for x in v):
                continue
            out.append({"cat": cat_lookup.get(sch, ""), "sch": sch, "v": v})
        return out

    cat_lookup = dict(zip(peer_map.index, peer_map["Category"])) if "Category" in peer_map.columns else {}
    comp_recs = scheme_records(composite, cat_lookup)
    y1_recs = scheme_records(score_1y, cat_lookup)
    y3_recs = scheme_records(score_3y, cat_lookup)

    peer_records = []
    for ix, row in peer_map.reset_index().iterrows():
        peer_records.append({
            "scheme": row.get("Scheme") or row.get("index"),
            "amfi": str(row.get("AMFI Code", "")),
            "cat": row.get("Category", ""),
            "mfi_name": row.get("Scheme Name From MFI", ""),
        })

    payload = {
        "meta": {
            "generated": datetime.now().isoformat(timespec="seconds"),
            "latest_aum_date": date_strs[-1] if date_strs else None,
            "latest_score_date": score_dates[-1] if score_dates else None,
            "n_schemes": len(comp_recs),
            "n_categories": len({s["cat"] for s in comp_recs}),
        },
        "aum_dates": date_strs,
        "score_dates": score_dates,
        "sleeve": sleeve_records,
        "amc_all": amc_all_records,
        "amc_vr": amc_vr_records,
        "composite": comp_recs,
        "score_1y": y1_recs,
        "score_3y": y3_recs,
        "peer_map": peer_records,
    }
    with open(out_path, "w") as f:
        json.dump(payload, f, separators=(",", ":"))
    size = out_path.stat().st_size / 1e6
    print(f"  Wrote {out_path}  ({size:.2f} MB)")


def _dashboard_template(override: str | None = None) -> str:
    """Return the dashboard HTML template — from an external file if supplied
    and present, otherwise the copy embedded at the bottom of this script."""
    if override:
        p = Path(override)
        if p.exists():
            print(f"  Using custom dashboard template: {p}")
            return p.read_text(encoding="utf-8")
        print(f"  (custom template {override} not found — using built-in template)")
    return base64.b64decode(DASHBOARD_TEMPLATE_B64).decode("utf-8")


def write_dashboards(out: Path, json_path: Path, override: str | None = None):
    """Write two ready-to-use dashboards into `out`:
        dashboard.html         — fetches dashboard_data.json at runtime (host on intranet)
        dashboard_offline.html — JSON inlined; a single file you can email / double-click
    """
    html = _dashboard_template(override)

    # 1) Live version — ships next to dashboard_data.json
    (out / "dashboard.html").write_text(html, encoding="utf-8")
    print(f"  Wrote {out / 'dashboard.html'}  (loads dashboard_data.json)")

    # 2) Offline single-file version — inline the JSON into the placeholder tag
    payload = json_path.read_text(encoding="utf-8")
    marker = re.compile(
        r'(<script id="embedded-data" type="application/json">)(.*?)(</script>)',
        re.DOTALL,
    )
    if not marker.search(html):
        print("  WARNING: embedded-data placeholder not found; skipping offline build")
        return
    # Guard against an accidental </script> inside the data (escape the slash).
    safe = payload.replace("</", "<\\/")
    offline = marker.sub(lambda m: m.group(1) + safe + m.group(3), html, count=1)
    off_path = out / "dashboard_offline.html"
    off_path.write_text(offline, encoding="utf-8")
    print(f"  Wrote {off_path}  ({off_path.stat().st_size / 1e6:.2f} MB, self-contained)")


# ═════════════════════════════════════════════════════════════════════
# 8. MAIN PIPELINE
# ═════════════════════════════════════════════════════════════════════

TOP15 = ["Axis", "Franklin", "Kotak", "HDFC", "Aditya", "DSP", "UTI", "Invesco",
         "Canara", "SBI", "Mirae", "Nippon", "Tata", "HSBC", "ICICI"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mfi", default="MFI Data", help="Path to MFI Data folder")
    ap.add_argument("--vr", default="Value Research Data", help="Path to Value Research Data folder")
    ap.add_argument("--out", default="out", help="Output folder")
    ap.add_argument("--dashboard-html", default=None,
                    help="Optional path to a custom dashboard template; "
                         "defaults to the template embedded in this script")
    args = ap.parse_args()

    mfi = Path(args.mfi); vr = Path(args.vr); out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    print("\n[1/7] Loading MFI data...")
    nav, aum_wide, raw_map = load_mfi(mfi)

    print("\n[2/7] Building category map...")
    cmap, categories = build_category_map(raw_map)
    nav = nav[aum_wide.index.intersection(cmap.index)]
    cmap = cmap.loc[nav.columns]
    print(f"  Final universe: {len(cmap)} schemes, {len(categories)} categories")

    print("\n[3/7] Loading Value Research mapping (optional)...")
    try:
        cme, bench_map, bench_nav = load_vr(vr)
        # Align benchmark NAV dates with NAV
        common = nav.index.intersection(bench_nav.index)
        bench_nav = bench_nav.loc[common]
        nav = nav.loc[common]
        # Trim VR mapping to schemes that actually exist in our NAV universe.
        # VR's "Scheme Name From MFI" may reference names with whitespace
        # mismatches, or schemes we dropped (MULTI_ASSET_DROP, no AUM, etc.).
        before = len(cme)
        cme = cme.loc[cme.index.intersection(nav.columns)]
        dropped = before - len(cme)
        if dropped:
            print(f"  Trimmed {dropped} VR rows whose MFI scheme name is not in our NAV universe")
    except Exception as e:
        print(f"  Skipped — {e}")
        cme = None

    print("\n[4/7] Building daily AUM share...")
    aum2, pct_with_house = percent_aum_daily(aum_wide, nav.index, cmap)

    print("\n[5/7] Running quartile analysis (All Peers — MFI categories)...")
    qy_1y, qy_3y, dist_1y, dist_3y, pct_1y, pct_3y = run_quartile_analysis(
        nav, cmap, categories, aum2
    )

    merged_quartile_df, sleeve_df, amc_all_df = build_quartile_aum_tables(
        qy_1y, qy_3y, pct_with_house, cmap, TOP15
    )
    dist_all = build_distance_table(dist_1y, dist_3y, categories)

    # Exact peer pass (Value Research)
    amc_vr_df = None
    score_1y = score_3y = composite = None
    if cme is not None:
        print("\n[6/7] Running exact-peer analysis (Value Research)...")
        cats_exact = sorted({c for c in cme["Category"]
                             if c not in ["Domestic + International", "Multi Index FoF",
                                          "Other Competitor Thematic/Contra Funds"]})
        member_lookup = {c: list(cme.loc[cme["Category"] == c].index) for c in cats_exact}
        qy_1y_e, qy_3y_e, _, _, _, _ = run_quartile_analysis(
            nav, cmap, cats_exact, aum2, member_lookup=member_lookup
        )
        _, _, amc_vr_df = build_quartile_aum_tables(
            qy_1y_e, qy_3y_e, pct_with_house, cmap, TOP15
        )
        # Scheme scores
        score_1y = quartile_to_score(qy_1y_e)
        score_3y = quartile_to_score(qy_3y_e)
        composite = composite_score(score_1y, score_3y)
        # Trim to schemes that actually have any score
        keep = composite.notna().any(axis=0)
        score_1y = score_1y.loc[:, keep]
        score_3y = score_3y.loc[:, keep]
        composite = composite.loc[:, keep]
    else:
        # Use the all-peer quartiles as fallback
        score_1y = quartile_to_score(qy_1y)
        score_3y = quartile_to_score(qy_3y)
        composite = composite_score(score_1y, score_3y)

    print("\n[7/7] Writing outputs...")
    asof = datetime.now().strftime("%B %d, %Y")
    aum_path = out / f"percent AUM in Q1 to Q4 - {asof}.xlsx"
    score_path = out / f"Scheme Scoring on Exact Peer Set - Calendar 1Y 3Y - {asof}.xlsx"

    write_aum_workbook(aum_path, sleeve_df, amc_all_df, amc_vr_df)

    peer_map_df = cme[["Scheme", "AMFI Code", "Category"]].copy() if cme is not None else pd.DataFrame()
    if cme is not None:
        peer_map_df.index.name = "Scheme Name From MFI"
        peer_map_df = peer_map_df.reset_index().set_index("Scheme")
        peer_map_df = peer_map_df.rename(columns={"index": "Scheme Name From MFI"})

    write_scoring_workbook(score_path, score_1y, score_3y, composite,
                            peer_map_df if not peer_map_df.empty else cmap[["Category"]])

    json_path = out / "dashboard_data.json"
    write_dashboard_json(json_path,
                          sleeve_df, amc_all_df, amc_vr_df,
                          composite, score_1y, score_3y,
                          peer_map_df if not peer_map_df.empty else cmap[["Category"]])

    write_dashboards(out, json_path, args.dashboard_html)

    print("\n✓ Done.")
    print(f"  Outputs: {out.resolve()}")
    print("  • Excel (internal sharing):")
    print(f"      {aum_path.name}")
    print(f"      {score_path.name}")
    print("  • Dashboard (host on intranet):  dashboard.html  +  dashboard_data.json")
    print("  • Dashboard (email / portable):  dashboard_offline.html")
    print(f"\n  Quick preview — open: {(out / 'dashboard_offline.html').resolve()}")


# ═════════════════════════════════════════════════════════════════════
# 9. EMBEDDED DASHBOARD TEMPLATE
# ═════════════════════════════════════════════════════════════════════
# The full interactive dashboard (HTML/CSS/JS) is stored below as a base64
# blob so this script is fully self-contained — running it on the data folders
# produces ready-to-use dashboards with no other files required.
# To edit the dashboard, decode this to dashboard.html, change it, and either
# re-embed it (base64) or pass it via --dashboard-html.
DASHBOARD_TEMPLATE_B64 = (
    "PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04Ij4KPG1ldGEgbmFtZT0i"
    "dmlld3BvcnQiIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xLjAiPgo8dGl0bGU+UGVlciBQZXJm"
    "b3JtYW5jZSBNb25pdG9yIOKAlCBFcXVpdHkgJmFtcDsgSHlicmlkPC90aXRsZT4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVm"
    "PSJodHRwczovL2ZvbnRzLmdvb2dsZWFwaXMuY29tIj4KPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2ZvbnRz"
    "LmdzdGF0aWMuY29tIiBjcm9zc29yaWdpbj4KPGxpbmsgaHJlZj0iaHR0cHM6Ly9mb250cy5nb29nbGVhcGlzLmNvbS9jc3MyP2Zh"
    "bWlseT1GcmF1bmNlczppdGFsLG9wc3osd2dodEAwLDkuLjE0NCw0MDA7MCw5Li4xNDQsNTAwOzAsOS4uMTQ0LDYwMDswLDkuLjE0"
    "NCw3MDA7MSw5Li4xNDQsNDAwOzEsOS4uMTQ0LDUwMCZmYW1pbHk9SW50ZXI6d2dodEA0MDA7NTAwOzYwMDs3MDAmZmFtaWx5PUpl"
    "dEJyYWlucytNb25vOndnaHRANDAwOzUwMDs2MDA7NzAwJmRpc3BsYXk9c3dhcCIgcmVsPSJzdHlsZXNoZWV0Ij4KPHNjcmlwdCBz"
    "cmM9Imh0dHBzOi8vY2RuLmpzZGVsaXZyLm5ldC9ucG0vY2hhcnQuanNANC40LjAvZGlzdC9jaGFydC51bWQubWluLmpzIj48L3Nj"
    "cmlwdD4KPHNjcmlwdCBzcmM9Imh0dHBzOi8vY2RuLmpzZGVsaXZyLm5ldC9ucG0vY2hhcnRqcy1hZGFwdGVyLWRhdGUtZm5zQDMu"
    "MC4wL2Rpc3QvY2hhcnRqcy1hZGFwdGVyLWRhdGUtZm5zLmJ1bmRsZS5taW4uanMiPjwvc2NyaXB0Pgo8c3R5bGU+CiAgOnJvb3Qg"
    "ewogICAgLS1wYXBlcjogICAgICAjZjRlZWUyOwogICAgLS1wYXBlci0yOiAgICAjZWZlN2Q1OwogICAgLS1wYXBlci0zOiAgICAj"
    "ZTdkY2MzOwogICAgLS1jYXJkOiAgICAgICAjZmJmN2VlOwogICAgLS1pbms6ICAgICAgICAjMTYxMzBmOwogICAgLS1pbmstMjog"
    "ICAgICAjM2IzNDJhOwogICAgLS1pbmstMzogICAgICAjNmE1ZjUwOwogICAgLS1pbmstNDogICAgICAjOWM4ZjdhOwogICAgLS1s"
    "aW5lOiAgICAgICAjZGRkMWI4OwogICAgLS1saW5lLTI6ICAgICAjZWJlMWNjOwogICAgLS1hYnNsOiAgICAgICAjYzgzOTJiOwog"
    "ICAgLS1hYnNsLTI6ICAgICAjYTcyYzIwOwogICAgLS1hYnNsLTM6ICAgICAjOGYyNjFjOwogICAgLS1hYnNsLWdsb3c6ICByZ2Jh"
    "KDIwMCw1Nyw0MywwLjEwKTsKICAgIC0tYWJzbC1iZzogICAgI2Y2ZGJkNjsKICAgIC0tZ29sZDogICAgICAgI2IzODUyZjsKICAg"
    "IC0tbmF2eTogICAgICAgIzIwNDM2YjsKICAgIC0tcTE6ICAgICAgICAgIzFmNWMzZDsKICAgIC0tcTI6ICAgICAgICAgIzZmYTM3"
    "YzsKICAgIC0tcTM6ICAgICAgICAgI2Q2YTIzZjsKICAgIC0tcTQ6ICAgICAgICAgI2MwNTIzZTsKICAgIC0tcTEtdDogICAgICAg"
    "cmdiYSgzMSw5Miw2MSwwLjg1KTsKICAgIC0tcTItdDogICAgICAgcmdiYSgxMTEsMTYzLDEyNCwwLjgwKTsKICAgIC0tcTMtdDog"
    "ICAgICAgcmdiYSgyMTQsMTYyLDYzLDAuODApOwogICAgLS1xNC10OiAgICAgICByZ2JhKDE5Miw4Miw2MiwwLjgyKTsKICAgIC0t"
    "c2hhZG93LXNtOiAgMCAxcHggMnB4IHJnYmEoMjIsMTksMTUsMC4wNCksIDAgMnB4IDhweCAtNHB4IHJnYmEoMjIsMTksMTUsMC4x"
    "MCk7CiAgICAtLXNoYWRvdy1tZDogIDAgMXB4IDAgcmdiYSgyMiwxOSwxNSwwLjA0KSwgMCAxMnB4IDMycHggLTE2cHggcmdiYSgy"
    "MiwxOSwxNSwwLjIyKTsKICAgIC0tc2hhZG93LWxnOiAgMCAycHggMCByZ2JhKDIyLDE5LDE1LDAuMDQpLCAwIDI4cHggNjBweCAt"
    "MjhweCByZ2JhKDIyLDE5LDE1LDAuMzApOwogICAgLS1yOiAgICAgICAgICA2cHg7CiAgICAtLXItbGc6ICAgICAgIDEwcHg7CiAg"
    "fQoKICAqIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgbWFyZ2luOiAwOyBwYWRkaW5nOiAwOyB9CgogIGh0bWwgeyAtd2Via2l0"
    "LWZvbnQtc21vb3RoaW5nOiBhbnRpYWxpYXNlZDsgdGV4dC1yZW5kZXJpbmc6IG9wdGltaXplTGVnaWJpbGl0eTsgfQogIGJvZHkg"
    "ewogICAgYmFja2dyb3VuZC1jb2xvcjogdmFyKC0tcGFwZXIpOwogICAgYmFja2dyb3VuZC1pbWFnZToKICAgICAgcmFkaWFsLWdy"
    "YWRpZW50KDEyMDBweCA2MDBweCBhdCA3OCUgLTglLCByZ2JhKDIwMCw1Nyw0MywwLjA1KSwgdHJhbnNwYXJlbnQgNjAlKSwKICAg"
    "ICAgcmFkaWFsLWdyYWRpZW50KDkwMHB4IDUwMHB4IGF0IC01JSA4JSwgcmdiYSgzMiw2NywxMDcsMC4wNDUpLCB0cmFuc3BhcmVu"
    "dCA1NSUpOwogICAgYmFja2dyb3VuZC1hdHRhY2htZW50OiBmaXhlZDsKICAgIGNvbG9yOiB2YXIoLS1pbmspOwogICAgZm9udC1m"
    "YW1pbHk6ICdJbnRlcicsIHNhbnMtc2VyaWY7CiAgICBmb250LXNpemU6IDE0cHg7CiAgICBsaW5lLWhlaWdodDogMS41OwogICAg"
    "cG9zaXRpb246IHJlbGF0aXZlOwogIH0KICAvKiBzdWJ0bGUgcGFwZXIgZ3JhaW4gKi8KICBib2R5OjpiZWZvcmUgewogICAgY29u"
    "dGVudDogIiI7CiAgICBwb3NpdGlvbjogZml4ZWQ7IGluc2V0OiAwOyB6LWluZGV4OiAwOyBwb2ludGVyLWV2ZW50czogbm9uZTsg"
    "b3BhY2l0eTogMC40MDsKICAgIGJhY2tncm91bmQtaW1hZ2U6IHVybCgiZGF0YTppbWFnZS9zdmcreG1sLCUzQ3N2ZyB4bWxucz0n"
    "aHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnIHdpZHRoPScxNDAnIGhlaWdodD0nMTQwJyUzRSUzQ2ZpbHRlciBpZD0nbiclM0Ul"
    "M0NmZVR1cmJ1bGVuY2UgdHlwZT0nZnJhY3RhbE5vaXNlJyBiYXNlRnJlcXVlbmN5PScwLjg1JyBudW1PY3RhdmVzPScyJyBzdGl0"
    "Y2hUaWxlcz0nc3RpdGNoJy8lM0UlM0MvZmlsdGVyJTNFJTNDcmVjdCB3aWR0aD0nMTAwJTI1JyBoZWlnaHQ9JzEwMCUyNScgZmls"
    "dGVyPSd1cmwoJTIzbiknIG9wYWNpdHk9JzAuMDM1Jy8lM0UlM0Mvc3ZnJTNFIik7CiAgfQogICNhcHAgeyBwb3NpdGlvbjogcmVs"
    "YXRpdmU7IHotaW5kZXg6IDE7IH0KCiAgLyog4paR4paRIE1BU1RIRUFEIOKWkeKWkSAqLwogIC5tYXN0aGVhZCB7CiAgICBiYWNr"
    "Z3JvdW5kOiBsaW5lYXItZ3JhZGllbnQoMTE4ZGVnLCAjMWExNjExIDAlLCAjMjQxZDE1IDU1JSwgIzJhMjAxNyAxMDAlKTsKICAg"
    "IGNvbG9yOiAjZjNlYWQ4OwogICAgcGFkZGluZzogMzBweCA0NHB4IDI2cHg7CiAgICBkaXNwbGF5OiBncmlkOwogICAgZ3JpZC10"
    "ZW1wbGF0ZS1jb2x1bW5zOiAxZnIgYXV0bzsKICAgIGFsaWduLWl0ZW1zOiBjZW50ZXI7CiAgICBnYXA6IDI0cHg7CiAgICBwb3Np"
    "dGlvbjogcmVsYXRpdmU7CiAgICBvdmVyZmxvdzogaGlkZGVuOwogICAgYm9yZGVyLWJvdHRvbTogM3B4IHNvbGlkIHZhcigtLWFi"
    "c2wpOwogIH0KICAubWFzdGhlYWQ6OmFmdGVyIHsKICAgIGNvbnRlbnQ6ICIiOwogICAgcG9zaXRpb246IGFic29sdXRlOyByaWdo"
    "dDogLTgwcHg7IHRvcDogLTEyMHB4OyB3aWR0aDogNDIwcHg7IGhlaWdodDogNDIwcHg7CiAgICBiYWNrZ3JvdW5kOiByYWRpYWwt"
    "Z3JhZGllbnQoY2lyY2xlLCByZ2JhKDIwMCw1Nyw0MywwLjIyKSwgdHJhbnNwYXJlbnQgNjUlKTsKICAgIHBvaW50ZXItZXZlbnRz"
    "OiBub25lOwogIH0KICAubWFzdC1raWNrZXIgewogICAgZm9udC1mYW1pbHk6ICdKZXRCcmFpbnMgTW9ubycsIG1vbm9zcGFjZTsK"
    "ICAgIGZvbnQtc2l6ZTogMTBweDsgbGV0dGVyLXNwYWNpbmc6IDAuMzJlbTsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsKICAg"
    "IGNvbG9yOiB2YXIoLS1hYnNsKTsgZm9udC13ZWlnaHQ6IDYwMDsgbWFyZ2luLWJvdHRvbTogMTBweDsKICAgIGRpc3BsYXk6IGZs"
    "ZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTBweDsKICB9CiAgLm1hc3Qta2lja2VyOjpiZWZvcmUgeyBjb250ZW50OiIi"
    "OyB3aWR0aDogMjZweDsgaGVpZ2h0OiAxcHg7IGJhY2tncm91bmQ6IHZhcigtLWFic2wpOyBkaXNwbGF5OmlubGluZS1ibG9jazsg"
    "fQogIC5tYXN0aGVhZCBoMSB7CiAgICBmb250LWZhbWlseTogJ0ZyYXVuY2VzJywgc2VyaWY7IGZvbnQtd2VpZ2h0OiA2MDA7IGZv"
    "bnQtc2l6ZTogNDBweDsKICAgIGxldHRlci1zcGFjaW5nOiAtMC4wMmVtOyBsaW5lLWhlaWdodDogMS4wOyBjb2xvcjogI2Y3ZjBl"
    "MjsKICB9CiAgLm1hc3RoZWFkIGgxIGVtIHsgZm9udC1zdHlsZTogaXRhbGljOyBmb250LXdlaWdodDogNTAwOyBjb2xvcjogI2U5"
    "YzljMjsgfQogIC5tYXN0LXN1YiB7CiAgICBmb250LXNpemU6IDEzLjVweDsgY29sb3I6ICNiM2E2OTA7IG1hcmdpbi10b3A6IDlw"
    "eDsgbWF4LXdpZHRoOiA1NjBweDsgZm9udC13ZWlnaHQ6IDQwMDsKICB9CiAgLm1hc3QtbWV0YSB7CiAgICBkaXNwbGF5OiBncmlk"
    "OyBnYXA6IDEycHg7IHRleHQtYWxpZ246IHJpZ2h0OwogICAgZm9udC1mYW1pbHk6ICdKZXRCcmFpbnMgTW9ubycsIG1vbm9zcGFj"
    "ZTsKICB9CiAgLm1hc3QtbWV0YSAucm93IC5sYmwgewogICAgZm9udC1zaXplOiA5cHg7IGxldHRlci1zcGFjaW5nOiAwLjE2ZW07"
    "IHRleHQtdHJhbnNmb3JtOiB1cHBlcmNhc2U7IGNvbG9yOiAjODk3YzY2OyBkaXNwbGF5OmJsb2NrOyBtYXJnaW4tYm90dG9tOiAz"
    "cHg7CiAgfQogIC5tYXN0LW1ldGEgLnJvdyAudmFsIHsgZm9udC1zaXplOiAxM3B4OyBjb2xvcjogI2YwZTZkMjsgZm9udC13ZWln"
    "aHQ6IDUwMDsgfQogIC5tYXN0LWJhZGdlIHsKICAgIGRpc3BsYXk6aW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOmNlbnRlcjsgZ2Fw"
    "OjdweDsgcGFkZGluZzogN3B4IDEzcHg7CiAgICBiYWNrZ3JvdW5kOiByZ2JhKDI0MywyMzQsMjE2LDAuMDYpOyBib3JkZXI6IDFw"
    "eCBzb2xpZCByZ2JhKDI0MywyMzQsMjE2LDAuMTQpOwogICAgYm9yZGVyLXJhZGl1czogMTAwcHg7IGZvbnQtc2l6ZTogMTFweDsg"
    "Y29sb3I6I2U3ZGNjNjsKICB9CiAgLm1hc3QtYmFkZ2UgLmRvdCB7IHdpZHRoOjdweDtoZWlnaHQ6N3B4O2JvcmRlci1yYWRpdXM6"
    "NTAlO2JhY2tncm91bmQ6IzVmYjg3ZjsgYm94LXNoYWRvdzowIDAgMCAzcHggcmdiYSg5NSwxODQsMTI3LDAuMjApOyB9CgogIC8q"
    "IOKWkeKWkSBOQVYgVEFCUyDilpHilpEgKi8KICAudGFicyB7CiAgICBkaXNwbGF5OiBmbGV4OyBwYWRkaW5nOiAwIDQ0cHg7IGdh"
    "cDogNHB4OwogICAgYmFja2dyb3VuZDogcmdiYSgyNDQsMjM4LDIyNiwwLjg2KTsKICAgIGJhY2tkcm9wLWZpbHRlcjogc2F0dXJh"
    "dGUoMTQwJSkgYmx1cig4cHgpOwogICAgLXdlYmtpdC1iYWNrZHJvcC1maWx0ZXI6IHNhdHVyYXRlKDE0MCUpIGJsdXIoOHB4KTsK"
    "ICAgIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1saW5lKTsKICAgIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogMDsgei1p"
    "bmRleDogNjA7CiAgfQogIC50YWIgewogICAgYXBwZWFyYW5jZTogbm9uZTsgYm9yZGVyOiBub25lOyBiYWNrZ3JvdW5kOiBub25l"
    "OyBjdXJzb3I6IHBvaW50ZXI7CiAgICBwYWRkaW5nOiAxNnB4IDIwcHggMTRweDsgZm9udC1mYW1pbHk6ICdGcmF1bmNlcycsIHNl"
    "cmlmOyBmb250LXNpemU6IDE1cHg7IGZvbnQtd2VpZ2h0OiA1MDA7CiAgICBjb2xvcjogdmFyKC0taW5rLTQpOyBib3JkZXItYm90"
    "dG9tOiAyLjVweCBzb2xpZCB0cmFuc3BhcmVudDsgbWFyZ2luLWJvdHRvbTogLTFweDsKICAgIHRyYW5zaXRpb246IGNvbG9yIC4y"
    "cywgYm9yZGVyLWNvbG9yIC4yczsgbGV0dGVyLXNwYWNpbmc6IC4wMDVlbTsgZGlzcGxheTpmbGV4OyBhbGlnbi1pdGVtczpjZW50"
    "ZXI7IGdhcDo5cHg7CiAgfQogIC50YWIgLnRudW0geyBmb250LWZhbWlseTonSmV0QnJhaW5zIE1vbm8nLCBtb25vc3BhY2U7IGZv"
    "bnQtc2l6ZToxMHB4OyBmb250LXdlaWdodDo2MDA7IGNvbG9yOnZhcigtLWluay00KTsgb3BhY2l0eTouNzsgfQogIC50YWI6aG92"
    "ZXIgeyBjb2xvcjogdmFyKC0taW5rLTIpOyB9CiAgLnRhYi5hY3RpdmUgeyBjb2xvcjogdmFyKC0taW5rKTsgYm9yZGVyLWJvdHRv"
    "bS1jb2xvcjogdmFyKC0tYWJzbCk7IH0KICAudGFiLmFjdGl2ZSAudG51bSB7IGNvbG9yOiB2YXIoLS1hYnNsKTsgb3BhY2l0eTox"
    "OyB9CgogIC8qIOKWkeKWkSBGSUxURVIgQkFSIOKWkeKWkSAqLwogIC5maWx0ZXJzIHsKICAgIHBhZGRpbmc6IDE2cHggNDRweDsg"
    "ZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGZsZXgtZW5kOyBnYXA6IDIwcHg7IGZsZXgtd3JhcDogd3JhcDsKICAgIGJhY2tn"
    "cm91bmQ6IGxpbmVhci1ncmFkaWVudCgxODBkZWcsIHZhcigtLXBhcGVyLTIpLCB2YXIoLS1wYXBlcikpOwogICAgYm9yZGVyLWJv"
    "dHRvbTogMXB4IHNvbGlkIHZhcigtLWxpbmUpOwogIH0KICAuZmcgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29s"
    "dW1uOyBnYXA6IDZweDsgfQogIC5mZyBsYWJlbCB7CiAgICBmb250LXNpemU6IDkuNXB4OyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJj"
    "YXNlOyBsZXR0ZXItc3BhY2luZzogLjE0ZW07IGNvbG9yOiB2YXIoLS1pbmstMyk7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgfQogIC5m"
    "ZyBzZWxlY3QsIC5mZyBpbnB1dCB7CiAgICBmb250LWZhbWlseTogJ0ludGVyJywgc2Fucy1zZXJpZjsgZm9udC1zaXplOiAxM3B4"
    "OyBwYWRkaW5nOiA5cHggMTNweDsKICAgIGJhY2tncm91bmQ6IHZhcigtLWNhcmQpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1s"
    "aW5lKTsgY29sb3I6IHZhcigtLWluayk7CiAgICBib3JkZXItcmFkaXVzOiB2YXIoLS1yKTsgbWluLXdpZHRoOiAxNTBweDsgY3Vy"
    "c29yOiBwb2ludGVyOyB0cmFuc2l0aW9uOiBib3JkZXItY29sb3IgLjE1cywgYm94LXNoYWRvdyAuMTVzOwogIH0KICAuZmcgc2Vs"
    "ZWN0OmhvdmVyIHsgYm9yZGVyLWNvbG9yOiB2YXIoLS1pbmstNCk7IH0KICAuZmcgc2VsZWN0OmZvY3VzLCAuZmcgaW5wdXQ6Zm9j"
    "dXMgeyBvdXRsaW5lOiBub25lOyBib3JkZXItY29sb3I6IHZhcigtLWFic2wpOyBib3gtc2hhZG93OiAwIDAgMCAzcHggdmFyKC0t"
    "YWJzbC1nbG93KTsgfQogIC5mZyBpbnB1dCB7IGN1cnNvcjogdGV4dDsgbWluLXdpZHRoOiAyMzBweDsgfQogIC5mZy53aWRlIHNl"
    "bGVjdCB7IG1pbi13aWR0aDogMzIwcHg7IH0KCiAgLnNlZyB7CiAgICBkaXNwbGF5OiBpbmxpbmUtZmxleDsgYm9yZGVyOiAxcHgg"
    "c29saWQgdmFyKC0tbGluZSk7IGJvcmRlci1yYWRpdXM6IHZhcigtLXIpOyBvdmVyZmxvdzogaGlkZGVuOyBiYWNrZ3JvdW5kOiB2"
    "YXIoLS1jYXJkKTsKICB9CiAgLnNlZyBidXR0b24gewogICAgYXBwZWFyYW5jZTpub25lOyBib3JkZXI6bm9uZTsgYmFja2dyb3Vu"
    "ZDpub25lOyBjdXJzb3I6cG9pbnRlcjsKICAgIHBhZGRpbmc6IDlweCAxNXB4OyBmb250LWZhbWlseTogJ0pldEJyYWlucyBNb25v"
    "JywgbW9ub3NwYWNlOyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS1pbmstMyk7CiAgICBib3JkZXItcmlnaHQ6IDFweCBz"
    "b2xpZCB2YXIoLS1saW5lKTsgdHJhbnNpdGlvbjogYmFja2dyb3VuZCAuMTVzLCBjb2xvciAuMTVzOwogIH0KICAuc2VnIGJ1dHRv"
    "bjpsYXN0LWNoaWxkIHsgYm9yZGVyLXJpZ2h0OiBub25lOyB9CiAgLnNlZyBidXR0b24uYWN0aXZlIHsgYmFja2dyb3VuZDogdmFy"
    "KC0taW5rKTsgY29sb3I6IHZhcigtLXBhcGVyKTsgfQogIC5zZWcgYnV0dG9uOm5vdCguYWN0aXZlKTpob3ZlciB7IGJhY2tncm91"
    "bmQ6IHZhcigtLXBhcGVyLTMpOyBjb2xvcjogdmFyKC0taW5rKTsgfQoKICAuc3BhY2VyIHsgbWFyZ2luLWxlZnQ6IGF1dG87IH0K"
    "ICAuYnRuIHsKICAgIGFwcGVhcmFuY2U6bm9uZTsgY3Vyc29yOnBvaW50ZXI7IHBhZGRpbmc6IDlweCAxOHB4OyBib3JkZXItcmFk"
    "aXVzOiB2YXIoLS1yKTsgZm9udC1zaXplOiAxMi41cHg7IGZvbnQtd2VpZ2h0OiA2MDA7CiAgICBmb250LWZhbWlseTogJ0ludGVy"
    "Jywgc2Fucy1zZXJpZjsgbGV0dGVyLXNwYWNpbmc6IC4wMmVtOyBib3JkZXI6IDFweCBzb2xpZCB0cmFuc3BhcmVudDsgdHJhbnNp"
    "dGlvbjogYWxsIC4xNXM7CiAgICBkaXNwbGF5OmlubGluZS1mbGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGdhcDo4cHg7CiAgfQog"
    "IC5idG4tZGFyayB7IGJhY2tncm91bmQ6IHZhcigtLWluayk7IGNvbG9yOiB2YXIoLS1wYXBlcik7IH0KICAuYnRuLWRhcms6aG92"
    "ZXIgeyBiYWNrZ3JvdW5kOiB2YXIoLS1hYnNsLTIpOyB9CiAgLmJ0biBzdmcgeyB3aWR0aDogMTRweDsgaGVpZ2h0OiAxNHB4OyB9"
    "CgogIC8qIOKWkeKWkSBDT05URU5UIOKWkeKWkSAqLwogIC5jb250ZW50IHsgcGFkZGluZzogMzRweCA0NHB4IDkwcHg7IG1heC13"
    "aWR0aDogMTY0MHB4OyBtYXJnaW46IDAgYXV0bzsgfQogIC52aWV3IHsgZGlzcGxheTogbm9uZTsgfQogIC52aWV3LmFjdGl2ZSB7"
    "IGRpc3BsYXk6IGJsb2NrOyBhbmltYXRpb246IHZpZXdJbiAuNDVzIGN1YmljLWJlemllciguMjIsLjYxLC4zNiwxKTsgfQogIEBr"
    "ZXlmcmFtZXMgdmlld0luIHsgZnJvbSB7IG9wYWNpdHk6IDA7IHRyYW5zZm9ybTogdHJhbnNsYXRlWSg4cHgpOyB9IHRvIHsgb3Bh"
    "Y2l0eTogMTsgdHJhbnNmb3JtOiBub25lOyB9IH0KCiAgLnNoZWFkIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGZsZXgt"
    "ZW5kOyBqdXN0aWZ5LWNvbnRlbnQ6IHNwYWNlLWJldHdlZW47IGdhcDogMjRweDsgbWFyZ2luLWJvdHRvbTogMjJweDsgfQogIC5z"
    "aGVhZCBoMiB7IGZvbnQtZmFtaWx5OiAnRnJhdW5jZXMnLCBzZXJpZjsgZm9udC13ZWlnaHQ6IDUwMDsgZm9udC1zaXplOiAyNnB4"
    "OyBsZXR0ZXItc3BhY2luZzogLTAuMDE1ZW07IGxpbmUtaGVpZ2h0OiAxLjE7IH0KICAuc2hlYWQgaDIgZW0geyBmb250LXN0eWxl"
    "OiBpdGFsaWM7IGNvbG9yOiB2YXIoLS1pbmstMyk7IGZvbnQtd2VpZ2h0OiA0MDA7IH0KICAuc2hlYWQgLmN0eCB7IGZvbnQtZmFt"
    "aWx5OidKZXRCcmFpbnMgTW9ubycsIG1vbm9zcGFjZTsgZm9udC1zaXplOiAxMXB4OyBjb2xvcjogdmFyKC0taW5rLTMpOyB0ZXh0"
    "LWFsaWduOnJpZ2h0OyBsaW5lLWhlaWdodDoxLjY7IH0KCiAgLyogS1BJIGNhcmRzICovCiAgLmtwaXMgeyBkaXNwbGF5OiBncmlk"
    "OyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCg0LCAxZnIpOyBnYXA6IDE0cHg7IG1hcmdpbi1ib3R0b206IDI2cHg7IH0K"
    "ICAua3BpIHsKICAgIGJhY2tncm91bmQ6IHZhcigtLWNhcmQpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1saW5lLTIpOyBib3Jk"
    "ZXItcmFkaXVzOiB2YXIoLS1yLWxnKTsgcGFkZGluZzogMThweCAyMHB4OwogICAgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LXNt"
    "KTsgcG9zaXRpb246IHJlbGF0aXZlOyBvdmVyZmxvdzogaGlkZGVuOwogICAgdHJhbnNpdGlvbjogdHJhbnNmb3JtIC4ycywgYm94"
    "LXNoYWRvdyAuMnM7CiAgfQogIC5rcGk6OmJlZm9yZSB7IGNvbnRlbnQ6IiI7IHBvc2l0aW9uOmFic29sdXRlOyBsZWZ0OjA7IHRv"
    "cDowOyBib3R0b206MDsgd2lkdGg6M3B4OyBiYWNrZ3JvdW5kOiB2YXIoLS1pbmstNCk7IH0KICAua3BpLmFjY2VudDo6YmVmb3Jl"
    "IHsgYmFja2dyb3VuZDogdmFyKC0tYWJzbCk7IH0KICAua3BpLmdvb2Q6OmJlZm9yZSB7IGJhY2tncm91bmQ6IHZhcigtLXExKTsg"
    "fQogIC5rcGkud2Fybjo6YmVmb3JlIHsgYmFja2dyb3VuZDogdmFyKC0tcTQpOyB9CiAgLmtwaTpob3ZlciB7IHRyYW5zZm9ybTog"
    "dHJhbnNsYXRlWSgtMnB4KTsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1kKTsgfQogIC5rcGkgLmstbGJsIHsgZm9udC1zaXpl"
    "OiAxMHB4OyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogLjExZW07IGNvbG9yOiB2YXIoLS1pbmst"
    "Myk7IGZvbnQtd2VpZ2h0OiA2MDA7IG1hcmdpbi1ib3R0b206IDlweDsgfQogIC5rcGkgLmstdmFsIHsgZm9udC1mYW1pbHk6ICdG"
    "cmF1bmNlcycsIHNlcmlmOyBmb250LXNpemU6IDMwcHg7IGZvbnQtd2VpZ2h0OiA1MDA7IGxldHRlci1zcGFjaW5nOiAtMC4wMmVt"
    "OyBsaW5lLWhlaWdodDogMTsgfQogIC5rcGkgLmstdmFsLmFjY2VudCB7IGNvbG9yOiB2YXIoLS1hYnNsKTsgfQogIC5rcGkgLmst"
    "dmFsLmdvb2QgeyBjb2xvcjogdmFyKC0tcTEpOyB9CiAgLmtwaSAuay12YWwud2FybiB7IGNvbG9yOiB2YXIoLS1xNCk7IH0KICAu"
    "a3BpIC5rLXN1YiB7IGZvbnQtZmFtaWx5OidKZXRCcmFpbnMgTW9ubycsIG1vbm9zcGFjZTsgZm9udC1zaXplOiAxMC41cHg7IGNv"
    "bG9yOiB2YXIoLS1pbmstNCk7IG1hcmdpbi10b3A6IDdweDsgfQoKICAvKiBDYXJkcyAqLwogIC5jYXJkIHsKICAgIGJhY2tncm91"
    "bmQ6IHZhcigtLWNhcmQpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1saW5lLTIpOyBib3JkZXItcmFkaXVzOiB2YXIoLS1yLWxn"
    "KTsKICAgIHBhZGRpbmc6IDI0cHggMjZweDsgbWFyZ2luLWJvdHRvbTogMjJweDsgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LW1k"
    "KTsKICB9CiAgLmNhcmQtaCB7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBmbGV4LXN0YXJ0OyBqdXN0aWZ5LWNvbnRlbnQ6"
    "IHNwYWNlLWJldHdlZW47IGdhcDogMjRweDsgbWFyZ2luLWJvdHRvbTogMThweDsgfQogIC5jYXJkLWggaDMgeyBmb250LWZhbWls"
    "eTogJ0ZyYXVuY2VzJywgc2VyaWY7IGZvbnQtd2VpZ2h0OiA1MDA7IGZvbnQtc2l6ZTogMTlweDsgbGV0dGVyLXNwYWNpbmc6IC0w"
    "LjAxZW07IH0KICAuY2FyZC1oIC5zdWIgeyBmb250LXNpemU6IDEycHg7IGNvbG9yOiB2YXIoLS1pbmstMyk7IG1hcmdpbi10b3A6"
    "IDNweDsgfQogIC5sZWdlbmQgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDE1cHg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGZvbnQtZmFt"
    "aWx5OidKZXRCcmFpbnMgTW9ubycsIG1vbm9zcGFjZTsgZm9udC1zaXplOiAxMXB4OyBjb2xvcjogdmFyKC0taW5rLTIpOyBmbGV4"
    "LXdyYXA6IHdyYXA7IH0KICAubGRvdCB7IHdpZHRoOiAxMXB4OyBoZWlnaHQ6IDExcHg7IGJvcmRlci1yYWRpdXM6IDNweDsgZGlz"
    "cGxheTogaW5saW5lLWJsb2NrOyBtYXJnaW4tcmlnaHQ6IDZweDsgdmVydGljYWwtYWxpZ246IC0xcHg7IH0KICAubGxpbmUgeyB3"
    "aWR0aDogMTZweDsgaGVpZ2h0OiAyLjVweDsgYm9yZGVyLXJhZGl1czogMnB4OyBkaXNwbGF5OmlubGluZS1ibG9jazsgbWFyZ2lu"
    "LXJpZ2h0OjZweDsgdmVydGljYWwtYWxpZ246IDNweDsgfQoKICAuY2hhcnRib3ggeyBwb3NpdGlvbjogcmVsYXRpdmU7IGhlaWdo"
    "dDogMzgwcHg7IH0KICAuY2hhcnRib3gudGFsbCB7IGhlaWdodDogNDYwcHg7IH0KICAuY2hhcnRib3guc2hvcnQgeyBoZWlnaHQ6"
    "IDMwMHB4OyB9CgogIC5ncmlkLTIgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDEuNDVmciAxZnI7IGdh"
    "cDogMjJweDsgfQogIC5ncmlkLTIuZXZlbiB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyIDFmcjsgfQoKICAvKiBNZXRob2Rv"
    "bG9neSBub3RlICovCiAgLm1ldGhvZCB7CiAgICBtYXJnaW4tdG9wOiAxNnB4OyBwYWRkaW5nOiAxM3B4IDE2cHggMTNweCAxNXB4"
    "OyBiYWNrZ3JvdW5kOiBsaW5lYXItZ3JhZGllbnQoMTgwZGVnLCB2YXIoLS1wYXBlci0yKSwgdmFyKC0tcGFwZXIpKTsKICAgIGJv"
    "cmRlcjogMXB4IHNvbGlkIHZhcigtLWxpbmUtMik7IGJvcmRlci1sZWZ0OiAzcHggc29saWQgdmFyKC0tZ29sZCk7IGJvcmRlci1y"
    "YWRpdXM6IHZhcigtLXIpOwogICAgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0taW5rLTIpOyBsaW5lLWhlaWdodDogMS42"
    "OwogIH0KICAubWV0aG9kIC5tLWxibCB7CiAgICBmb250LWZhbWlseTonSmV0QnJhaW5zIE1vbm8nLCBtb25vc3BhY2U7IGZvbnQt"
    "c2l6ZTogOS41cHg7IGxldHRlci1zcGFjaW5nOiAuMTRlbTsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsKICAgIGNvbG9yOiB2"
    "YXIoLS1nb2xkKTsgZm9udC13ZWlnaHQ6IDcwMDsgbWFyZ2luLXJpZ2h0OiA4cHg7CiAgfQogIC5tZXRob2QgYiB7IGNvbG9yOiB2"
    "YXIoLS1pbmspOyBmb250LXdlaWdodDogNjAwOyB9CgogIC8qIFRhYmxlcyAqLwogIC50Ymwtd3JhcCB7IG92ZXJmbG93LXg6IGF1"
    "dG87IH0KICB0YWJsZS5ydCB7IHdpZHRoOiAxMDAlOyBib3JkZXItY29sbGFwc2U6IGNvbGxhcHNlOyBmb250LXNpemU6IDEyLjVw"
    "eDsgfQogIHRhYmxlLnJ0IHRoIHsKICAgIHRleHQtYWxpZ246IGxlZnQ7IHBhZGRpbmc6IDExcHggMTBweDsgYm9yZGVyLWJvdHRv"
    "bTogMS41cHggc29saWQgdmFyKC0taW5rKTsKICAgIGZvbnQtZmFtaWx5OiAnSW50ZXInLCBzYW5zLXNlcmlmOyBmb250LXNpemU6"
    "IDkuNXB4OyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogLjA5ZW07CiAgICBjb2xvcjogdmFyKC0t"
    "aW5rLTIpOyBmb250LXdlaWdodDogNzAwOyB3aGl0ZS1zcGFjZTogbm93cmFwOyBwb3NpdGlvbjogc3RpY2t5OyB0b3A6IDA7IGJh"
    "Y2tncm91bmQ6IHZhcigtLWNhcmQpOwogIH0KICB0YWJsZS5ydCB0aC5udW0sIHRhYmxlLnJ0IHRkLm51bSB7IHRleHQtYWxpZ246"
    "IHJpZ2h0OyBmb250LXZhcmlhbnQtbnVtZXJpYzogdGFidWxhci1udW1zOyB9CiAgdGFibGUucnQgdGQgeyBwYWRkaW5nOiAxMHB4"
    "IDEwcHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1saW5lLTIpOyBmb250LWZhbWlseTonSmV0QnJhaW5zIE1vbm8n"
    "LCBtb25vc3BhY2U7IGNvbG9yOiB2YXIoLS1pbmstMik7IH0KICB0YWJsZS5ydCB0ZC5sYmwgeyBmb250LWZhbWlseTogJ0ludGVy"
    "Jywgc2Fucy1zZXJpZjsgY29sb3I6IHZhcigtLWluayk7IGZvbnQtd2VpZ2h0OiA1MDA7IH0KICB0YWJsZS5ydCB0cjpob3ZlciB0"
    "ZCB7IGJhY2tncm91bmQ6IHZhcigtLXBhcGVyLTIpOyB9CiAgdGFibGUucnQgdHIuaGwgdGQgeyBiYWNrZ3JvdW5kOiB2YXIoLS1h"
    "YnNsLWJnKTsgfQogIHRhYmxlLnJ0IHRyLmhsIHRkLmxibCB7IGNvbG9yOiB2YXIoLS1hYnNsLTIpOyBmb250LXdlaWdodDogNzAw"
    "OyB9CiAgLnNjcm9sbC10YWxsIHsgbWF4LWhlaWdodDogNTgwcHg7IG92ZXJmbG93LXk6IGF1dG87IGJvcmRlcjogMXB4IHNvbGlk"
    "IHZhcigtLWxpbmUtMik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXIpOyB9CgogIC5waWxsIHsgZGlzcGxheTppbmxpbmUtYmxvY2s7"
    "IG1pbi13aWR0aDogMjRweDsgcGFkZGluZzogM3B4IDhweDsgZm9udC1mYW1pbHk6J0pldEJyYWlucyBNb25vJywgbW9ub3NwYWNl"
    "OyBmb250LXNpemU6IDEwLjVweDsgZm9udC13ZWlnaHQ6IDcwMDsgdGV4dC1hbGlnbjpjZW50ZXI7IGJvcmRlci1yYWRpdXM6IDEw"
    "MHB4OyBiYWNrZ3JvdW5kOiB2YXIoLS1wYXBlci0zKTsgY29sb3I6IHZhcigtLWluay0yKTsgfQogIC5waWxsLnIxIHsgYmFja2dy"
    "b3VuZDogdmFyKC0tcTEpOyBjb2xvcjogI2ZmZjsgfQogIC5waWxsLnVwIHsgYmFja2dyb3VuZDogcmdiYSgzMSw5Miw2MSwwLjE2"
    "KTsgY29sb3I6IHZhcigtLXExKTsgfQogIC5waWxsLmRuIHsgYmFja2dyb3VuZDogcmdiYSgxOTIsODIsNjIsMC4xNik7IGNvbG9y"
    "OiB2YXIoLS1xNCk7IH0KICAucGlsbC5sYXN0IHsgYmFja2dyb3VuZDogdmFyKC0tcTQpOyBjb2xvcjogI2ZmZjsgfQoKICAuZGVs"
    "dGEtcG9zIHsgY29sb3I6IHZhcigtLXExKTsgfQogIC5kZWx0YS1uZWcgeyBjb2xvcjogdmFyKC0tcTQpOyB9CgogIC5zZWFyY2gg"
    "ewogICAgd2lkdGg6IDEwMCU7IHBhZGRpbmc6IDEycHggMTZweDsgZm9udC1zaXplOiAxMy41cHg7IGJvcmRlcjogMXB4IHNvbGlk"
    "IHZhcigtLWxpbmUpOyBiYWNrZ3JvdW5kOiB2YXIoLS1jYXJkKTsKICAgIGJvcmRlci1yYWRpdXM6IHZhcigtLXIpOyBmb250LWZh"
    "bWlseTogJ0ludGVyJywgc2Fucy1zZXJpZjsgbWFyZ2luLWJvdHRvbTogMTRweDsKICB9CiAgLnNlYXJjaDpmb2N1cyB7IG91dGxp"
    "bmU6IG5vbmU7IGJvcmRlci1jb2xvcjogdmFyKC0tYWJzbCk7IGJveC1zaGFkb3c6IDAgMCAwIDNweCB2YXIoLS1hYnNsLWdsb3cp"
    "OyB9CgogIC5sb2FkaW5nIHsgZGlzcGxheTpmbGV4OyBmbGV4LWRpcmVjdGlvbjpjb2x1bW47IGFsaWduLWl0ZW1zOmNlbnRlcjsg"
    "anVzdGlmeS1jb250ZW50OmNlbnRlcjsgbWluLWhlaWdodDogNzB2aDsgZ2FwOiAxNnB4OyBjb2xvcjogdmFyKC0taW5rLTMpOyB9"
    "CiAgLmxvYWRpbmcgLnNwaW4geyB3aWR0aDogMzBweDsgaGVpZ2h0OiAzMHB4OyBib3JkZXI6IDNweCBzb2xpZCB2YXIoLS1saW5l"
    "KTsgYm9yZGVyLXRvcC1jb2xvcjogdmFyKC0tYWJzbCk7IGJvcmRlci1yYWRpdXM6IDUwJTsgYW5pbWF0aW9uOiBzcCAwLjhzIGxp"
    "bmVhciBpbmZpbml0ZTsgfQogIEBrZXlmcmFtZXMgc3AgeyB0byB7IHRyYW5zZm9ybTogcm90YXRlKDM2MGRlZyk7IH0gfQogIC5s"
    "b2FkaW5nIC5sdCB7IGZvbnQtZmFtaWx5OiAnRnJhdW5jZXMnLCBzZXJpZjsgZm9udC1zaXplOiAxN3B4OyBmb250LXN0eWxlOiBp"
    "dGFsaWM7IH0KCiAgLnRhZyB7IGRpc3BsYXk6aW5saW5lLWJsb2NrOyBwYWRkaW5nOiAzcHggOXB4OyBib3JkZXItcmFkaXVzOiAx"
    "MDBweDsgZm9udC1zaXplOiAxMC41cHg7IGZvbnQtZmFtaWx5OidKZXRCcmFpbnMgTW9ubycsbW9ub3NwYWNlOyBmb250LXdlaWdo"
    "dDo2MDA7IGJhY2tncm91bmQ6IHZhcigtLXBhcGVyLTMpOyBjb2xvcjogdmFyKC0taW5rLTIpOyBib3JkZXI6MXB4IHNvbGlkIHZh"
    "cigtLWxpbmUpOyB9CgogIEBtZWRpYSAobWF4LXdpZHRoOiAxMTgwcHgpIHsKICAgIC5ncmlkLTIsIC5ncmlkLTIuZXZlbiB7IGdy"
    "aWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9CiAgICAua3BpcyB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogcmVwZWF0KDIsIDFm"
    "cik7IH0KICB9CiAgQG1lZGlhIChtYXgtd2lkdGg6IDcyMHB4KSB7CiAgICAubWFzdGhlYWQgeyBncmlkLXRlbXBsYXRlLWNvbHVt"
    "bnM6IDFmcjsgcGFkZGluZzogMjJweDsgfQogICAgLm1hc3QtbWV0YSB7IHRleHQtYWxpZ246IGxlZnQ7IH0KICAgIC50YWJzLCAu"
    "ZmlsdGVycywgLmNvbnRlbnQgeyBwYWRkaW5nLWxlZnQ6IDE4cHg7IHBhZGRpbmctcmlnaHQ6IDE4cHg7IH0KICAgIC5tYXN0aGVh"
    "ZCBoMSB7IGZvbnQtc2l6ZTogMzBweDsgfQogICAgLmtwaXMgeyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IH0KICAg"
    "IC50YWJzIHsgb3ZlcmZsb3cteDogYXV0bzsgfQogIH0KPC9zdHlsZT4KPC9oZWFkPgo8Ym9keT4KPGRpdiBpZD0iYXBwIj4KCjwh"
    "LS0gT3B0aW9uYWwgaW5saW5lIGRhdGEgYmxvY2sg4oCUIHBvcHVsYXRlZCBieSBlbWJlZF9kYXRhLnB5IGZvciBvZmZsaW5lIGRp"
    "c3RyaWJ1dGlvbi4gLS0+CjxzY3JpcHQgaWQ9ImVtYmVkZGVkLWRhdGEiIHR5cGU9ImFwcGxpY2F0aW9uL2pzb24iPjwvc2NyaXB0"
    "PgoKPGhlYWRlciBjbGFzcz0ibWFzdGhlYWQiPgogIDxkaXY+CiAgICA8ZGl2IGNsYXNzPSJtYXN0LWtpY2tlciI+QWRpdHlhIEJp"
    "cmxhIFN1biBMaWZlIMK3IEludmVzdG1lbnQgUmVzZWFyY2g8L2Rpdj4KICAgIDxoMT5QZWVyIFBlcmZvcm1hbmNlIDxlbT5Nb25p"
    "dG9yPC9lbT48L2gxPgogICAgPGRpdiBjbGFzcz0ibWFzdC1zdWIiPlF1YXJ0aWxlIHBvc2l0aW9uaW5nLCBBVU0gY29uY2VudHJh"
    "dGlvbiBhbmQgc2NoZW1lIHNjb3JpbmcgYWNyb3NzIGVxdWl0eSAmYW1wOyBoeWJyaWQgc2xlZXZlcyDigJQgYmVuY2htYXJrZWQg"
    "YWdhaW5zdCB0aGUgVG9wLTE1IGluZHVzdHJ5IEFNQ3MuPC9kaXY+CiAgPC9kaXY+CiAgPGRpdiBjbGFzcz0ibWFzdC1tZXRhIj4K"
    "ICAgIDxkaXYgY2xhc3M9Im1hc3QtYmFkZ2UiPjxzcGFuIGNsYXNzPSJkb3QiPjwvc3Bhbj48c3BhbiBpZD0ibWItYXNvZiI+4oCU"
    "PC9zcGFuPjwvZGl2PgogICAgPGRpdiBjbGFzcz0icm93Ij48c3BhbiBjbGFzcz0ibGJsIj5Vbml2ZXJzZTwvc3Bhbj48c3BhbiBj"
    "bGFzcz0idmFsIiBpZD0ibWItdW5pdiI+4oCUPC9zcGFuPjwvZGl2PgogICAgPGRpdiBjbGFzcz0icm93Ij48c3BhbiBjbGFzcz0i"
    "bGJsIj5HZW5lcmF0ZWQ8L3NwYW4+PHNwYW4gY2xhc3M9InZhbCIgaWQ9Im1iLWdlbiI+4oCUPC9zcGFuPjwvZGl2PgogIDwvZGl2"
    "Pgo8L2hlYWRlcj4KCjxuYXYgY2xhc3M9InRhYnMiIGlkPSJ0YWJzIj4KICA8YnV0dG9uIGNsYXNzPSJ0YWIgYWN0aXZlIiBkYXRh"
    "LXZpZXc9InNsZWV2ZSI+PHNwYW4gY2xhc3M9InRudW0iPjAxPC9zcGFuPlNsZWV2ZSBWaWV3PC9idXR0b24+CiAgPGJ1dHRvbiBj"
    "bGFzcz0idGFiIiBkYXRhLXZpZXc9ImFtYyI+PHNwYW4gY2xhc3M9InRudW0iPjAyPC9zcGFuPkFNQyBSb2xsLXVwPC9idXR0b24+"
    "CiAgPGJ1dHRvbiBjbGFzcz0idGFiIiBkYXRhLXZpZXc9InNjaGVtZSI+PHNwYW4gY2xhc3M9InRudW0iPjAzPC9zcGFuPlNjaGVt"
    "ZSBEZXRhaWw8L2J1dHRvbj4KICA8YnV0dG9uIGNsYXNzPSJ0YWIiIGRhdGEtdmlldz0ibWF0cml4Ij48c3BhbiBjbGFzcz0idG51"
    "bSI+MDQ8L3NwYW4+UXVhcnRpbGUgTWF0cml4PC9idXR0b24+CiAgPGJ1dHRvbiBjbGFzcz0idGFiIiBkYXRhLXZpZXc9InBlZXIi"
    "PjxzcGFuIGNsYXNzPSJ0bnVtIj4wNTwvc3Bhbj5QZWVyIE1hcHBpbmc8L2J1dHRvbj4KPC9uYXY+Cgo8bWFpbiBpZD0idmlld3Mi"
    "PjwvbWFpbj4KCjwvZGl2PgoKPHNjcmlwdD4KLy/ilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDi"
    "lZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDi"
    "lZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDi"
    "lZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZAKLy8gIFBFRVIgUEVSRk9STUFOQ0UgTU9OSVRPUiDigJQgYXBwbGljYXRp"
    "b24gbG9naWMKLy/ilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDi"
    "lZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDi"
    "lZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDilZDi"
    "lZDilZDilZDilZDilZAKY29uc3QgQyA9IHsKICBxMTonIzFmNWMzZCcsIHEyOicjNmZhMzdjJywgcTM6JyNkNmEyM2YnLCBxNDon"
    "I2MwNTIzZScsCiAgcTF0OidyZ2JhKDMxLDkyLDYxLDAuODIpJywgcTJ0OidyZ2JhKDExMSwxNjMsMTI0LDAuNzgpJywgcTN0Oidy"
    "Z2JhKDIxNCwxNjIsNjMsMC44MCknLCBxNHQ6J3JnYmEoMTkyLDgyLDYyLDAuODIpJywKICBhYnNsOicjYzgzOTJiJywgbmF2eTon"
    "IzIwNDM2YicsIGdvbGQ6JyNiMzg1MmYnLCBpbms6JyMxNjEzMGYnLCBpbmszOicjNmE1ZjUwJywgaW5rNDonIzljOGY3YScsCiAg"
    "bGluZTonI2RkZDFiOCcsIGxpbmUyOicjZWJlMWNjJywgcGVlcjonI2M5YmNhMycKfTsKY29uc3QgUU9SREVSID0gWyclIGluIFEx"
    "JywnJSBpbiBRMicsJyUgaW4gUTMnLCclIGluIFE0J107CmNvbnN0IFFDT0wgICA9IHsgJyUgaW4gUTEnOkMucTEsICclIGluIFEy"
    "JzpDLnEyLCAnJSBpbiBRMyc6Qy5xMywgJyUgaW4gUTQnOkMucTQgfTsKY29uc3QgUUZJTEwgID0geyAnJSBpbiBRMSc6Qy5xMXQs"
    "JyUgaW4gUTInOkMucTJ0LCclIGluIFEzJzpDLnEzdCwnJSBpbiBRNCc6Qy5xNHQgfTsKCmNvbnN0IFMgPSB7CiAgZGF0YTpudWxs"
    "LAogIHNsZWV2ZTp7IGZoOidBZGl0eWEnLCBicjonQm90dG9tcyBVcCcsIHJ3OicxIFllYXInLCBtb2RlOid3aXRoaW4nLCByYW5r"
    "RGF0ZTpudWxsIH0sCiAgYW1jOnsgZmg6J0FkaXR5YScsIHJ3OicxIFllYXInLCB1bml2OidhbGwnLCByYW5rRGF0ZTpudWxsIH0s"
    "CiAgc2NoZW1lOnsgY2F0Om51bGwsIG5hbWU6bnVsbCwgcGVlckRhdGU6bnVsbCB9LAogIG1hdHJpeDp7IHJ3OicxIFllYXInLCBt"
    "ZXRyaWM6J3RvcGhhbGYnLCBkYXRlOm51bGwgfSwKICBjaGFydHM6e30KfTsKCmNvbnN0IGZtdCA9IHsKICBwY3Q6KHYsZD0xKT0+"
    "IHY9PW51bGx8fGlzTmFOKHYpID8gJ+KAlCcgOiAoMTAwKnYpLnRvRml4ZWQoZCkrJyUnLAogIHBwOih2LGQ9MSk9PiB2PT1udWxs"
    "fHxpc05hTih2KSA/ICfigJQnIDogKHY+PTA/JysnOicnKSsoMTAwKnYpLnRvRml4ZWQoZCkrJ3BwJywKICBudW06KHYsZD0yKT0+"
    "IHY9PW51bGx8fGlzTmFOKHYpID8gJ+KAlCcgOiB2LnRvRml4ZWQoZCksCiAgZGF0ZTppc289PnsgaWYoIWlzbykgcmV0dXJuICfi"
    "gJQnOyBjb25zdCBkPW5ldyBEYXRlKGlzbysnVDAwOjAwOjAwJyk7IHJldHVybiBkLnRvTG9jYWxlRGF0ZVN0cmluZygnZW4tR0In"
    "LHtkYXk6JzItZGlnaXQnLG1vbnRoOidzaG9ydCcseWVhcjonbnVtZXJpYyd9KTsgfSwKICBtb255cjppc289PnsgaWYoIWlzbykg"
    "cmV0dXJuICfigJQnOyBjb25zdCBkPW5ldyBEYXRlKGlzbysnVDAwOjAwOjAwJyk7IHJldHVybiBkLnRvTG9jYWxlRGF0ZVN0cmlu"
    "ZygnZW4tR0InLHttb250aDonc2hvcnQnLHllYXI6J251bWVyaWMnfSk7IH0KfTsKCi8vIENoYXJ0LmpzIGdsb2JhbCB0aGVtZQpD"
    "aGFydC5kZWZhdWx0cy5mb250LmZhbWlseSA9ICInSW50ZXInLCBzYW5zLXNlcmlmIjsKQ2hhcnQuZGVmYXVsdHMuZm9udC5zaXpl"
    "ID0gMTE7CkNoYXJ0LmRlZmF1bHRzLmNvbG9yID0gQy5pbmszOwpDaGFydC5kZWZhdWx0cy5ib3JkZXJDb2xvciA9IEMubGluZTI7"
    "CgpmdW5jdGlvbiBncmFkaWVudChjdHgsIGFyZWEsIGhleCwgdG9wQT0wLjQyLCBib3RBPTAuMDIpewogIGlmKCFhcmVhKSByZXR1"
    "cm4gaGV4OwogIGNvbnN0IGcgPSBjdHguY3JlYXRlTGluZWFyR3JhZGllbnQoMCwgYXJlYS50b3AsIDAsIGFyZWEuYm90dG9tKTsK"
    "ICBjb25zdCBoID0gaGV4LnJlcGxhY2UoJyMnLCcnKTsKICBjb25zdCByPXBhcnNlSW50KGguc2xpY2UoMCwyKSwxNiksIGdnPXBh"
    "cnNlSW50KGguc2xpY2UoMiw0KSwxNiksIGI9cGFyc2VJbnQoaC5zbGljZSg0LDYpLDE2KTsKICBnLmFkZENvbG9yU3RvcCgwLCBg"
    "cmdiYSgke3J9LCR7Z2d9LCR7Yn0sJHt0b3BBfSlgKTsKICBnLmFkZENvbG9yU3RvcCgxLCBgcmdiYSgke3J9LCR7Z2d9LCR7Yn0s"
    "JHtib3RBfSlgKTsKICByZXR1cm4gZzsKfQoKLy/ilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDi"
    "lIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAgREFUQSBMT0FESU5HIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKU"
    "gOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKU"
    "gApmdW5jdGlvbiBsb2FkRGF0YSgpewogIGNvbnN0IGVtYiA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdlbWJlZGRlZC1kYXRh"
    "Jyk7CiAgaWYgKGVtYiAmJiBlbWIudGV4dENvbnRlbnQudHJpbSgpLmxlbmd0aCA+IDEwMCl7CiAgICB0cnkgeyBib290KEpTT04u"
    "cGFyc2UoZW1iLnRleHRDb250ZW50KSk7IHJldHVybjsgfSBjYXRjaChlKXsgY29uc29sZS53YXJuKCdlbWJlZGRlZCBwYXJzZSBm"
    "YWlsZWQnLCBlKTsgfQogIH0KICBmZXRjaCgnZGFzaGJvYXJkX2RhdGEuanNvbicpCiAgICAudGhlbihyPT57IGlmKCFyLm9rKSB0"
    "aHJvdyBuZXcgRXJyb3IoJ0hUVFAgJytyLnN0YXR1cyk7IHJldHVybiByLmpzb24oKTsgfSkKICAgIC50aGVuKGJvb3QpLmNhdGNo"
    "KHNob3dFcnIpOwp9CgpmdW5jdGlvbiBzaG93RXJyKGVycil7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXdzJykuaW5u"
    "ZXJIVE1MID0KICAgIGA8ZGl2IGNsYXNzPSJsb2FkaW5nIj48ZGl2IGNsYXNzPSJsdCI+Q291bGRuJ3QgbG9hZCBkYXNoYm9hcmRf"
    "ZGF0YS5qc29uPC9kaXY+CiAgICAgPGRpdiBzdHlsZT0iZm9udC1zaXplOjEzcHg7bWF4LXdpZHRoOjUyMHB4O3RleHQtYWxpZ246"
    "Y2VudGVyO2xpbmUtaGVpZ2h0OjEuNyI+CiAgICAgU2VydmUgdGhpcyBwYWdlIG92ZXIgSFRUUCAoPGNvZGU+cHl0aG9uIC1tIGh0"
    "dHAuc2VydmVyPC9jb2RlPikgd2l0aCB0aGUgSlNPTiBpbiB0aGUgc2FtZSBmb2xkZXIsCiAgICAgb3IgcnVuIDxjb2RlPmVtYmVk"
    "X2RhdGEucHk8L2NvZGU+IHRvIGJ1bmRsZSB0aGUgZGF0YSBpbnRvIHRoaXMgZmlsZS48YnI+PGJyPgogICAgIDxzbWFsbCBzdHls"
    "ZT0iY29sb3I6dmFyKC0taW5rLTQpIj4ke2Vycn08L3NtYWxsPjwvZGl2PjwvZGl2PmA7Cn0KCmZ1bmN0aW9uIGJvb3QoZGF0YSl7"
    "CiAgUy5kYXRhID0gZGF0YTsKICAvLyBkZWZhdWx0IGRhdGVzID0gbGF0ZXN0CiAgUy5zbGVldmUucmFua0RhdGUgPSBkYXRhLmF1"
    "bV9kYXRlcy5sZW5ndGgtMTsKICBTLmFtYy5yYW5rRGF0ZSAgICA9IGRhdGEuYXVtX2RhdGVzLmxlbmd0aC0xOwogIC8vIHNjb3Jl"
    "cyBjYW4gYmUgbnVsbCBvbiB0aGUgdmVyeSBsYXN0IGVkZ2UgZGF0ZSAoM1kgcm9sbGluZyDihpIgTmFOKTsgZGVmYXVsdCB0byBs"
    "YXRlc3QgcG9wdWxhdGVkIGRhdGUKICBTLnNjaGVtZS5wZWVyRGF0ZSA9IGxhdGVzdFNjb3JlRGF0YUluZGV4KGRhdGEpOwogIFMu"
    "bWF0cml4LmRhdGUgICAgID0gZGF0YS5hdW1fZGF0ZXMubGVuZ3RoLTE7CgogIC8vIG1ldGEKICBkb2N1bWVudC5nZXRFbGVtZW50"
    "QnlJZCgnbWItYXNvZicpLnRleHRDb250ZW50ID0gJ0FzIG9mICcgKyBmbXQuZGF0ZShkYXRhLm1ldGEubGF0ZXN0X2F1bV9kYXRl"
    "KTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbWItdW5pdicpLnRleHRDb250ZW50ID0gYCR7ZGF0YS5tZXRhLm5fc2NoZW1l"
    "c30gc2NoZW1lcyDCtyAke2RhdGEubWV0YS5uX2NhdGVnb3JpZXN9IGNhdGVnb3JpZXNgOwogIGRvY3VtZW50LmdldEVsZW1lbnRC"
    "eUlkKCdtYi1nZW4nKS50ZXh0Q29udGVudCAgPSAoZGF0YS5tZXRhLmdlbmVyYXRlZHx8JycpLnJlcGxhY2UoJ1QnLCcgJyk7Cgog"
    "IGJ1aWxkVmlld3MoKTsKICB3aXJlVGFicygpOwogIHJlbmRlclNsZWV2ZSgpOyByZW5kZXJBTUMoKTsgcmVuZGVyU2NoZW1lKCk7"
    "IHJlbmRlck1hdHJpeCgpOyByZW5kZXJQZWVyKCk7Cn0KCi8v4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSAIEhFTFBFUlMg4pSA4pSA4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSA4pSA4pSA4pSA4pSACmNvbnN0IHVuaXEgPSAoYXJyLGspPT4gWy4uLm5ldyBTZXQoYXJyLm1hcChyPT5yW2tdKSldOwpj"
    "b25zdCBmaFNvcnQgPSAoYSxiKT0+IGE9PT0nQWRpdHlhJz8tMSA6IGI9PT0nQWRpdHlhJz8xIDogYS5sb2NhbGVDb21wYXJlKGIp"
    "OwoKLy8gbGF0ZXN0IHNjb3JlLWRhdGUgaW5kZXggdGhhdCBoYXMgYXQgbGVhc3Qgb25lIG5vbi1udWxsIGNvbXBvc2l0ZSAoZWRn"
    "ZSBkYXRlcyBjYW4gYmUgYWxsLW51bGwpCmZ1bmN0aW9uIGxhdGVzdFNjb3JlRGF0YUluZGV4KGRhdGEpewogIGZvcihsZXQgaT1k"
    "YXRhLnNjb3JlX2RhdGVzLmxlbmd0aC0xO2k+PTA7aS0tKXsKICAgIGlmKGRhdGEuY29tcG9zaXRlLnNvbWUocz0+cy52W2ldIT1u"
    "dWxsKSkgcmV0dXJuIGk7CiAgfQogIHJldHVybiBkYXRhLnNjb3JlX2RhdGVzLmxlbmd0aC0xOwp9CgpmdW5jdGlvbiBzbGVldmVT"
    "bGljZShmaCxicixydyl7CiAgY29uc3QgcmVjcyA9IFMuZGF0YS5zbGVldmUuZmlsdGVyKHI9PnIuZmg9PT1maCAmJiByLmJyPT09"
    "YnIgJiYgci5ydz09PXJ3KTsKICBjb25zdCBieVEgPSB7fTsgcmVjcy5mb3JFYWNoKHI9PiBieVFbci5xdF09ci52KTsgcmV0dXJu"
    "IGJ5UTsKfQpmdW5jdGlvbiBhbWNTbGljZShmaCxydyx1bml2KXsKICBjb25zdCBVID0gdW5pdj09PSdhbGwnPyBTLmRhdGEuYW1j"
    "X2FsbCA6IFMuZGF0YS5hbWNfdnI7CiAgY29uc3QgcmVjcyA9IFUuZmlsdGVyKHI9PnIuZmg9PT1maCAmJiByLnJ3PT09cncpOwog"
    "IGNvbnN0IGJ5USA9IHt9OyByZWNzLmZvckVhY2gocj0+IGJ5UVtyLnF0XT1yLnYpOyByZXR1cm4gYnlROwp9Ci8vIEFVTS13ZWln"
    "aHRlZCBxdWFydGlsZSBzY29yZSBhdCBpbmRleCBpOiBRMT00IOKApiBRND0xLCB3ZWlnaHRlZCBieSBxdWFydGlsZSBBVU0uIDHi"
    "gJM0IHNjYWxlLgpmdW5jdGlvbiBxdWFsaXR5U2NvcmUoYnlRLCBpKXsKICBsZXQgbnVtPTAsIGRlbj0wOwogIGNvbnN0IHcgPSB7"
    "JyUgaW4gUTEnOjQsJyUgaW4gUTInOjMsJyUgaW4gUTMnOjIsJyUgaW4gUTQnOjF9OwogIFFPUkRFUi5mb3JFYWNoKHE9PnsgY29u"
    "c3Qgdj1ieVFbcV0/YnlRW3FdW2ldOjA7IGlmKHYhPW51bGwpeyBudW0rPXdbcV0qdjsgZGVuKz12OyB9IH0pOwogIHJldHVybiBk"
    "ZW4+MCA/IG51bS9kZW4gOiBudWxsOwp9Ci8vIGRyb3Bkb3duIG9mIG1vbnRoLWVuZCBkYXRlcwpmdW5jdGlvbiBkYXRlT3B0aW9u"
    "cyhkYXRlcywgc2VsKXsKICByZXR1cm4gZGF0ZXMubWFwKChkLGkpPT5gPG9wdGlvbiB2YWx1ZT0iJHtpfSIgJHtpPT09c2VsPydz"
    "ZWxlY3RlZCc6Jyd9PiR7Zm10LmRhdGUoZCl9PC9vcHRpb24+YCkuam9pbignJyk7Cn0KZnVuY3Rpb24gbWV0aG9kTm90ZShodG1s"
    "KXsgcmV0dXJuIGA8ZGl2IGNsYXNzPSJtZXRob2QiPjxzcGFuIGNsYXNzPSJtLWxibCI+TWV0aG9kb2xvZ3k8L3NwYW4+JHtodG1s"
    "fTwvZGl2PmA7IH0KCmZ1bmN0aW9uIGRvd25sb2FkQ1NWKGNvbnRlbnQsIGZpbGVuYW1lKXsKICBjb25zdCBibG9iID0gbmV3IEJs"
    "b2IoW2NvbnRlbnRdLHt0eXBlOid0ZXh0L2NzdjtjaGFyc2V0PXV0Zi04Oyd9KTsKICBjb25zdCB1cmwgPSBVUkwuY3JlYXRlT2Jq"
    "ZWN0VVJMKGJsb2IpOyBjb25zdCBhPWRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoJ2EnKTsKICBhLmhyZWY9dXJsOyBhLmRvd25sb2Fk"
    "PWZpbGVuYW1lOyBhLmNsaWNrKCk7IFVSTC5yZXZva2VPYmplY3RVUkwodXJsKTsKfQoKZnVuY3Rpb24gd2lyZVRhYnMoKXsKICBk"
    "b2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcudGFiJykuZm9yRWFjaChidG49PnsKICAgIGJ0bi5hZGRFdmVudExpc3RlbmVyKCdj"
    "bGljaycsKCk9PnsKICAgICAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLnRhYicpLmZvckVhY2goYj0+Yi5jbGFzc0xpc3Qu"
    "cmVtb3ZlKCdhY3RpdmUnKSk7CiAgICAgIGRvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy52aWV3JykuZm9yRWFjaCh2PT52LmNs"
    "YXNzTGlzdC5yZW1vdmUoJ2FjdGl2ZScpKTsKICAgICAgYnRuLmNsYXNzTGlzdC5hZGQoJ2FjdGl2ZScpOwogICAgICBkb2N1bWVu"
    "dC5nZXRFbGVtZW50QnlJZCgndmlldy0nK2J0bi5kYXRhc2V0LnZpZXcpLmNsYXNzTGlzdC5hZGQoJ2FjdGl2ZScpOwogICAgICAv"
    "LyBjaGFydHMgc29tZXRpbWVzIG5lZWQgcmVzaXplIGFmdGVyIGJlaW5nIHNob3duCiAgICAgIHNldFRpbWVvdXQoKCk9Pk9iamVj"
    "dC52YWx1ZXMoUy5jaGFydHMpLmZvckVhY2goYz0+eyB0cnl7Yy5yZXNpemUoKTt9Y2F0Y2goZSl7fSB9KSwgNjApOwogICAgfSk7"
    "CiAgfSk7Cn0KCi8v4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSA4pSA4pSA4pSA4pSA4pSAIEJVSUxEIFZJRVcgU0NBRkZPTERTIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKU"
    "gOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiBidWlsZFZpZXdzKCl7CiAgZG9jdW1l"
    "bnQuZ2V0RWxlbWVudEJ5SWQoJ3ZpZXdzJykuaW5uZXJIVE1MID0gYAogIDwhLS0g4paR4paRIFNMRUVWRSDilpHilpEgLS0+CiAg"
    "PHNlY3Rpb24gY2xhc3M9InZpZXcgYWN0aXZlIiBpZD0idmlldy1zbGVldmUiPgogICAgPGRpdiBjbGFzcz0iZmlsdGVycyI+CiAg"
    "ICAgIDxkaXYgY2xhc3M9ImZnIj48bGFiZWw+RnVuZCBIb3VzZTwvbGFiZWw+PHNlbGVjdCBpZD0ic2wtZmgiPjwvc2VsZWN0Pjwv"
    "ZGl2PgogICAgICA8ZGl2IGNsYXNzPSJmZyI+PGxhYmVsPkJyZWFrZG93biAvIFNsZWV2ZTwvbGFiZWw+PHNlbGVjdCBpZD0ic2wt"
    "YnIiPjwvc2VsZWN0PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJmZyI+PGxhYmVsPlJvbGxpbmcgV2luZG93PC9sYWJlbD4KICAg"
    "ICAgICA8ZGl2IGNsYXNzPSJzZWciIGlkPSJzbC1ydyI+PGJ1dHRvbiBkYXRhLXY9IjEgWWVhciIgY2xhc3M9ImFjdGl2ZSI+MVk8"
    "L2J1dHRvbj48YnV0dG9uIGRhdGEtdj0iMyBZZWFyIj4zWTwvYnV0dG9uPjwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJm"
    "ZyI+PGxhYmVsPlF1YXJ0aWxlIEJhc2lzPC9sYWJlbD4KICAgICAgICA8ZGl2IGNsYXNzPSJzZWciIGlkPSJzbC1tb2RlIj48YnV0"
    "dG9uIGRhdGEtdj0id2l0aGluIiBjbGFzcz0iYWN0aXZlIj5XaXRoaW4gc2xlZXZlPC9idXR0b24+PGJ1dHRvbiBkYXRhLXY9ImFt"
    "YyI+JSBvZiBBTUMgQVVNPC9idXR0b24+PC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9InNwYWNlciI+PC9kaXY+CiAgICAg"
    "IDxidXR0b24gY2xhc3M9ImJ0biBidG4tZGFyayIgaWQ9InNsLWV4cCI+JHtpY29Eb3duKCl9IEV4cG9ydCBDU1Y8L2J1dHRvbj4K"
    "ICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iY29udGVudCI+CiAgICAgIDxkaXYgY2xhc3M9InNoZWFkIj48aDIgaWQ9InNsLXRp"
    "dGxlIj5TbGVldmUgY29tcG9zaXRpb248L2gyPjxkaXYgY2xhc3M9ImN0eCIgaWQ9InNsLWN0eCI+PC9kaXY+PC9kaXY+CiAgICAg"
    "IDxkaXYgY2xhc3M9ImtwaXMiIGlkPSJzbC1rcGlzIj48L2Rpdj4KCiAgICAgIDxkaXYgY2xhc3M9ImNhcmQiPgogICAgICAgIDxk"
    "aXYgY2xhc3M9ImNhcmQtaCI+CiAgICAgICAgICA8ZGl2PjxoMyBpZD0ic2wtYzEtdGl0bGUiPlF1YXJ0aWxlIGRpc3RyaWJ1dGlv"
    "biB3aXRoaW4gc2xlZXZlPC9oMz48ZGl2IGNsYXNzPSJzdWIiIGlkPSJzbC1jMS1zdWIiPjwvZGl2PjwvZGl2PgogICAgICAgICAg"
    "PGRpdiBjbGFzcz0ibGVnZW5kIj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9Imxkb3QiIHN0eWxlPSJiYWNrZ3JvdW5k"
    "OiR7Qy5xMX0iPjwvc3Bhbj5RMTwvc3Bhbj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9Imxkb3QiIHN0eWxlPSJiYWNr"
    "Z3JvdW5kOiR7Qy5xMn0iPjwvc3Bhbj5RMjwvc3Bhbj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9Imxkb3QiIHN0eWxl"
    "PSJiYWNrZ3JvdW5kOiR7Qy5xM30iPjwvc3Bhbj5RMzwvc3Bhbj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9Imxkb3Qi"
    "IHN0eWxlPSJiYWNrZ3JvdW5kOiR7Qy5xNH0iPjwvc3Bhbj5RNDwvc3Bhbj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9"
    "ImxsaW5lIiBzdHlsZT0iYmFja2dyb3VuZDoke0MuYWJzbH0iPjwvc3Bhbj5TbGVldmUgd2VpZ2h0IChSKTwvc3Bhbj4KICAgICAg"
    "ICAgIDwvZGl2PgogICAgICAgIDwvZGl2PgogICAgICAgIDxkaXYgY2xhc3M9ImNoYXJ0Ym94IHRhbGwiPjxjYW52YXMgaWQ9InNs"
    "LXN0YWNrIj48L2NhbnZhcz48L2Rpdj4KICAgICAgICA8ZGl2IGlkPSJzbC1jMS1tZXRob2QiPjwvZGl2PgogICAgICA8L2Rpdj4K"
    "CiAgICAgIDxkaXYgY2xhc3M9ImdyaWQtMiI+CiAgICAgICAgPGRpdiBjbGFzcz0iY2FyZCI+CiAgICAgICAgICA8ZGl2IGNsYXNz"
    "PSJjYXJkLWgiPjxkaXY+PGgzPlRvcC1oYWxmIHNoYXJlICZhbXA7IHF1YWxpdHkgdHJlbmQ8L2gzPjxkaXYgY2xhc3M9InN1YiI+"
    "UTErUTIgc2hhcmUgd2l0aGluIHNsZWV2ZSwgd2l0aCBBVU0td2VpZ2h0ZWQgcXVhcnRpbGUgc2NvcmU8L2Rpdj48L2Rpdj4KICAg"
    "ICAgICAgICAgPGRpdiBjbGFzcz0ibGVnZW5kIj48c3Bhbj48c3BhbiBjbGFzcz0ibGxpbmUiIHN0eWxlPSJiYWNrZ3JvdW5kOiR7"
    "Qy5xMX0iPjwvc3Bhbj5RMStRMiAlPC9zcGFuPjxzcGFuPjxzcGFuIGNsYXNzPSJsbGluZSIgc3R5bGU9ImJhY2tncm91bmQ6JHtD"
    "Lm5hdnl9Ij48L3NwYW4+U2NvcmUgKFIpPC9zcGFuPjwvZGl2PgogICAgICAgICAgPC9kaXY+CiAgICAgICAgICA8ZGl2IGNsYXNz"
    "PSJjaGFydGJveCI+PGNhbnZhcyBpZD0ic2wtdHJlbmQiPjwvY2FudmFzPjwvZGl2PgogICAgICAgICAgPGRpdiBpZD0ic2wtdHJl"
    "bmQtbWV0aG9kIj48L2Rpdj4KICAgICAgICA8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJjYXJkIj4KICAgICAgICAgIDxkaXYg"
    "Y2xhc3M9ImNhcmQtaCI+PGRpdj48aDM+UmFua2luZyB2cyBUb3AtMTU8L2gzPjxkaXYgY2xhc3M9InN1YiI+UmUtcmFuayBvbiBh"
    "bnkgbW9udGgtZW5kPC9kaXY+PC9kaXY+CiAgICAgICAgICAgIDxkaXYgY2xhc3M9ImZnIj48bGFiZWw+QXMtb2YgZGF0ZTwvbGFi"
    "ZWw+PHNlbGVjdCBpZD0ic2wtcmFua2RhdGUiPjwvc2VsZWN0PjwvZGl2PgogICAgICAgICAgPC9kaXY+CiAgICAgICAgICA8ZGl2"
    "IGNsYXNzPSJ0Ymwtd3JhcCIgaWQ9InNsLXJhbmsiPjwvZGl2PgogICAgICAgICAgPGRpdiBpZD0ic2wtcmFuay1tZXRob2QiPjwv"
    "ZGl2PgogICAgICAgIDwvZGl2PgogICAgICA8L2Rpdj4KCiAgICAgIDxkaXYgY2xhc3M9ImNhcmQiPgogICAgICAgIDxkaXYgY2xh"
    "c3M9ImNhcmQtaCI+PGRpdj48aDM+Q3Jvc3Mtc2VjdGlvbmFsIHNuYXBzaG90PC9oMz48ZGl2IGNsYXNzPSJzdWIiIGlkPSJzbC1i"
    "YXItc3ViIj48L2Rpdj48L2Rpdj48L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNzPSJjaGFydGJveCBzaG9ydCI+PGNhbnZhcyBpZD0i"
    "c2wtYmFyIj48L2NhbnZhcz48L2Rpdj4KICAgICAgICA8ZGl2IGlkPSJzbC1iYXItbWV0aG9kIj48L2Rpdj4KICAgICAgPC9kaXY+"
    "CiAgICA8L2Rpdj4KICA8L3NlY3Rpb24+CgogIDwhLS0g4paR4paRIEFNQyDilpHilpEgLS0+CiAgPHNlY3Rpb24gY2xhc3M9InZp"
    "ZXciIGlkPSJ2aWV3LWFtYyI+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXJzIj4KICAgICAgPGRpdiBjbGFzcz0iZmciPjxsYWJlbD5G"
    "dW5kIEhvdXNlPC9sYWJlbD48c2VsZWN0IGlkPSJhbS1maCI+PC9zZWxlY3Q+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImZnIj48"
    "bGFiZWw+UGVlciBVbml2ZXJzZTwvbGFiZWw+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VnIiBpZD0iYW0tdW5pdiI+PGJ1dHRvbiBk"
    "YXRhLXY9ImFsbCIgY2xhc3M9ImFjdGl2ZSI+QWxsIFBlZXJzIChNRkkpPC9idXR0b24+PGJ1dHRvbiBkYXRhLXY9InZyIj5FeGFj"
    "dCBQZWVycyAoVlIpPC9idXR0b24+PC9kaXY+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImZnIj48bGFiZWw+Um9sbGluZyBXaW5k"
    "b3c8L2xhYmVsPgogICAgICAgIDxkaXYgY2xhc3M9InNlZyIgaWQ9ImFtLXJ3Ij48YnV0dG9uIGRhdGEtdj0iMSBZZWFyIiBjbGFz"
    "cz0iYWN0aXZlIj4xWTwvYnV0dG9uPjxidXR0b24gZGF0YS12PSIzIFllYXIiPjNZPC9idXR0b24+PC9kaXY+PC9kaXY+CiAgICAg"
    "IDxkaXYgY2xhc3M9InNwYWNlciI+PC9kaXY+CiAgICAgIDxidXR0b24gY2xhc3M9ImJ0biBidG4tZGFyayIgaWQ9ImFtLWV4cCI+"
    "JHtpY29Eb3duKCl9IEV4cG9ydCBDU1Y8L2J1dHRvbj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iY29udGVudCI+CiAgICAg"
    "IDxkaXYgY2xhc3M9InNoZWFkIj48aDI+QU1DLWxldmVsIHF1YXJ0aWxlIGRpc3RyaWJ1dGlvbiA8ZW0+4oCUIGFjcm9zcyBhbGwg"
    "c2xlZXZlczwvZW0+PC9oMj48ZGl2IGNsYXNzPSJjdHgiIGlkPSJhbS1jdHgiPjwvZGl2PjwvZGl2PgogICAgICA8ZGl2IGNsYXNz"
    "PSJrcGlzIiBpZD0iYW0ta3BpcyI+PC9kaXY+CgogICAgICA8ZGl2IGNsYXNzPSJjYXJkIj4KICAgICAgICA8ZGl2IGNsYXNzPSJj"
    "YXJkLWgiPjxkaXY+PGgzPkhvdXNlLXdpZGUgcXVhcnRpbGUgbWl4IG92ZXIgdGltZTwvaDM+PGRpdiBjbGFzcz0ic3ViIj5BbGwg"
    "ZXF1aXR5ICZhbXA7IGh5YnJpZCBBVU0sIG5vcm1hbGlzZWQgdG8gMTAwJTwvZGl2PjwvZGl2PgogICAgICAgICAgPGRpdiBjbGFz"
    "cz0ibGVnZW5kIj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9Imxkb3QiIHN0eWxlPSJiYWNrZ3JvdW5kOiR7Qy5xMX0i"
    "Pjwvc3Bhbj5RMTwvc3Bhbj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9Imxkb3QiIHN0eWxlPSJiYWNrZ3JvdW5kOiR7"
    "Qy5xMn0iPjwvc3Bhbj5RMjwvc3Bhbj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9Imxkb3QiIHN0eWxlPSJiYWNrZ3Jv"
    "dW5kOiR7Qy5xM30iPjwvc3Bhbj5RMzwvc3Bhbj4KICAgICAgICAgICAgPHNwYW4+PHNwYW4gY2xhc3M9Imxkb3QiIHN0eWxlPSJi"
    "YWNrZ3JvdW5kOiR7Qy5xNH0iPjwvc3Bhbj5RNDwvc3Bhbj48L2Rpdj4KICAgICAgICA8L2Rpdj4KICAgICAgICA8ZGl2IGNsYXNz"
    "PSJjaGFydGJveCB0YWxsIj48Y2FudmFzIGlkPSJhbS1zdGFjayI+PC9jYW52YXM+PC9kaXY+CiAgICAgICAgPGRpdiBpZD0iYW0t"
    "c3RhY2stbWV0aG9kIj48L2Rpdj4KICAgICAgPC9kaXY+CgogICAgICA8ZGl2IGNsYXNzPSJjYXJkIj4KICAgICAgICA8ZGl2IGNs"
    "YXNzPSJjYXJkLWgiPjxkaXY+PGgzPkFkaXR5YSBCaXJsYSB2cyBUb3AtMTUg4oCUIFExK1EyIHNoYXJlPC9oMz48ZGl2IGNsYXNz"
    "PSJzdWIiPlNvbGlkIHJlZCA9IEFkaXR5YSBCaXJsYTsgZmFkZWQgPSBwZWVyczwvZGl2PjwvZGl2PjwvZGl2PgogICAgICAgIDxk"
    "aXYgY2xhc3M9ImNoYXJ0Ym94IHRhbGwiPjxjYW52YXMgaWQ9ImFtLXJhbmsiPjwvY2FudmFzPjwvZGl2PgogICAgICAgIDxkaXYg"
    "aWQ9ImFtLXJhbmstbWV0aG9kIj48L2Rpdj4KICAgICAgPC9kaXY+CgogICAgICA8ZGl2IGNsYXNzPSJjYXJkIj4KICAgICAgICA8"
    "ZGl2IGNsYXNzPSJjYXJkLWgiPjxkaXY+PGgzPkxlYWd1ZSB0YWJsZTwvaDM+PGRpdiBjbGFzcz0ic3ViIj5BbGwgQU1DcyByYW5r"
    "ZWQgb24gdGhlIGNob3NlbiBkYXRlPC9kaXY+PC9kaXY+CiAgICAgICAgICA8ZGl2IGNsYXNzPSJmZyI+PGxhYmVsPkFzLW9mIGRh"
    "dGU8L2xhYmVsPjxzZWxlY3QgaWQ9ImFtLXJhbmtkYXRlIj48L3NlbGVjdD48L2Rpdj4KICAgICAgICA8L2Rpdj4KICAgICAgICA8"
    "ZGl2IGNsYXNzPSJ0Ymwtd3JhcCIgaWQ9ImFtLXRhYmxlIj48L2Rpdj4KICAgICAgICA8ZGl2IGlkPSJhbS10YWJsZS1tZXRob2Qi"
    "PjwvZGl2PgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgogIDwvc2VjdGlvbj4KCiAgPCEtLSDilpHilpEgU0NIRU1FIOKWkeKWkSAt"
    "LT4KICA8c2VjdGlvbiBjbGFzcz0idmlldyIgaWQ9InZpZXctc2NoZW1lIj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlcnMiPgogICAg"
    "ICA8ZGl2IGNsYXNzPSJmZyI+PGxhYmVsPkNhdGVnb3J5PC9sYWJlbD48c2VsZWN0IGlkPSJzYy1jYXQiPjwvc2VsZWN0PjwvZGl2"
    "PgogICAgICA8ZGl2IGNsYXNzPSJmZyB3aWRlIj48bGFiZWw+U2NoZW1lPC9sYWJlbD48c2VsZWN0IGlkPSJzYy1uYW1lIj48L3Nl"
    "bGVjdD48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic3BhY2VyIj48L2Rpdj4KICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuIGJ0bi1k"
    "YXJrIiBpZD0ic2MtZXhwIj4ke2ljb0Rvd24oKX0gRXhwb3J0IENTVjwvYnV0dG9uPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNz"
    "PSJjb250ZW50Ij4KICAgICAgPGRpdiBjbGFzcz0ic2hlYWQiPjxoMiBpZD0ic2MtdGl0bGUiPlNlbGVjdCBhIHNjaGVtZTwvaDI+"
    "PGRpdiBjbGFzcz0iY3R4IiBpZD0ic2MtY3R4Ij48L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ia3BpcyIgaWQ9InNjLWtw"
    "aXMiPjwvZGl2PgoKICAgICAgPGRpdiBjbGFzcz0iY2FyZCI+CiAgICAgICAgPGRpdiBjbGFzcz0iY2FyZC1oIj48ZGl2PjxoMz5R"
    "dWFydGlsZSBzY29yZSBoaXN0b3J5PC9oMz48ZGl2IGNsYXNzPSJzdWIiPjEgKGJvdHRvbSBxdWFydGlsZSkg4oaSIDQgKHRvcCBx"
    "dWFydGlsZSk8L2Rpdj48L2Rpdj4KICAgICAgICAgIDxkaXYgY2xhc3M9ImxlZ2VuZCI+PHNwYW4+PHNwYW4gY2xhc3M9ImxsaW5l"
    "IiBzdHlsZT0iYmFja2dyb3VuZDoke0MuYWJzbH0iPjwvc3Bhbj5Db21wb3NpdGU8L3NwYW4+PHNwYW4+PHNwYW4gY2xhc3M9Imxs"
    "aW5lIiBzdHlsZT0iYmFja2dyb3VuZDoke0MubmF2eX0iPjwvc3Bhbj4xWTwvc3Bhbj48c3Bhbj48c3BhbiBjbGFzcz0ibGxpbmUi"
    "IHN0eWxlPSJiYWNrZ3JvdW5kOiR7Qy5xMn0iPjwvc3Bhbj4zWTwvc3Bhbj48L2Rpdj4KICAgICAgICA8L2Rpdj4KICAgICAgICA8"
    "ZGl2IGNsYXNzPSJjaGFydGJveCB0YWxsIj48Y2FudmFzIGlkPSJzYy1zY29yZSI+PC9jYW52YXM+PC9kaXY+CiAgICAgICAgPGRp"
    "diBpZD0ic2Mtc2NvcmUtbWV0aG9kIj48L2Rpdj4KICAgICAgPC9kaXY+CgogICAgICA8ZGl2IGNsYXNzPSJjYXJkIj4KICAgICAg"
    "ICA8ZGl2IGNsYXNzPSJjYXJkLWgiPjxkaXY+PGgzPlBlZXIgY29tcGFyaXNvbjwvaDM+PGRpdiBjbGFzcz0ic3ViIj5BbGwgc2No"
    "ZW1lcyBpbiB0aGUgZXhhY3QtcGVlciBjYXRlZ29yeSwgcmFua2VkIG9uIHRoZSBjaG9zZW4gZGF0ZTwvZGl2PjwvZGl2PgogICAg"
    "ICAgICAgPGRpdiBjbGFzcz0iZmciPjxsYWJlbD5Bcy1vZiBkYXRlPC9sYWJlbD48c2VsZWN0IGlkPSJzYy1wZWVyZGF0ZSI+PC9z"
    "ZWxlY3Q+PC9kaXY+CiAgICAgICAgPC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0idGJsLXdyYXAiIGlkPSJzYy1wZWVycyI+PC9k"
    "aXY+CiAgICAgICAgPGRpdiBpZD0ic2MtcGVlcnMtbWV0aG9kIj48L2Rpdj4KICAgICAgPC9kaXY+CiAgICA8L2Rpdj4KICA8L3Nl"
    "Y3Rpb24+CgogIDwhLS0g4paR4paRIE1BVFJJWCDilpHilpEgLS0+CiAgPHNlY3Rpb24gY2xhc3M9InZpZXciIGlkPSJ2aWV3LW1h"
    "dHJpeCI+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXJzIj4KICAgICAgPGRpdiBjbGFzcz0iZmciPjxsYWJlbD5Sb2xsaW5nIFdpbmRv"
    "dzwvbGFiZWw+CiAgICAgICAgPGRpdiBjbGFzcz0ic2VnIiBpZD0ibXgtcnciPjxidXR0b24gZGF0YS12PSIxIFllYXIiIGNsYXNz"
    "PSJhY3RpdmUiPjFZPC9idXR0b24+PGJ1dHRvbiBkYXRhLXY9IjMgWWVhciI+M1k8L2J1dHRvbj48L2Rpdj48L2Rpdj4KICAgICAg"
    "PGRpdiBjbGFzcz0iZmciPjxsYWJlbD5NZXRyaWM8L2xhYmVsPgogICAgICAgIDxkaXYgY2xhc3M9InNlZyIgaWQ9Im14LW1ldHJp"
    "YyI+PGJ1dHRvbiBkYXRhLXY9InRvcGhhbGYiIGNsYXNzPSJhY3RpdmUiPlExK1EyICU8L2J1dHRvbj48YnV0dG9uIGRhdGEtdj0i"
    "c2NvcmUiPlF1YWxpdHkgc2NvcmU8L2J1dHRvbj48YnV0dG9uIGRhdGEtdj0icTEiPlExICU8L2J1dHRvbj48L2Rpdj48L2Rpdj4K"
    "ICAgICAgPGRpdiBjbGFzcz0iZmciPjxsYWJlbD5Bcy1vZiBkYXRlPC9sYWJlbD48c2VsZWN0IGlkPSJteC1kYXRlIj48L3NlbGVj"
    "dD48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0ic3BhY2VyIj48L2Rpdj4KICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuIGJ0bi1kYXJr"
    "IiBpZD0ibXgtZXhwIj4ke2ljb0Rvd24oKX0gRXhwb3J0IENTVjwvYnV0dG9uPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJj"
    "b250ZW50Ij4KICAgICAgPGRpdiBjbGFzcz0ic2hlYWQiPjxoMj5RdWFydGlsZSBtYXRyaXggPGVtPuKAlCBob3VzZSDDlyBzbGVl"
    "dmU8L2VtPjwvaDI+PGRpdiBjbGFzcz0iY3R4IiBpZD0ibXgtY3R4Ij48L2Rpdj48L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iY2Fy"
    "ZCI+CiAgICAgICAgPGRpdiBjbGFzcz0iY2FyZC1oIj48ZGl2PjxoMyBpZD0ibXgtdGl0bGUiPlRvcC1oYWxmIHNoYXJlIGJ5IGZ1"
    "bmQgaG91c2UgYW5kIHNsZWV2ZTwvaDM+PGRpdiBjbGFzcz0ic3ViIiBpZD0ibXgtc3ViIj48L2Rpdj48L2Rpdj48L2Rpdj4KICAg"
    "ICAgICA8ZGl2IGNsYXNzPSJ0Ymwtd3JhcCIgaWQ9Im14LWdyaWQiPjwvZGl2PgogICAgICAgIDxkaXYgaWQ9Im14LW1ldGhvZCI+"
    "PC9kaXY+CiAgICAgIDwvZGl2PgogICAgPC9kaXY+CiAgPC9zZWN0aW9uPgoKICA8IS0tIOKWkeKWkSBQRUVSIOKWkeKWkSAtLT4K"
    "ICA8c2VjdGlvbiBjbGFzcz0idmlldyIgaWQ9InZpZXctcGVlciI+CiAgICA8ZGl2IGNsYXNzPSJmaWx0ZXJzIj4KICAgICAgPGRp"
    "diBjbGFzcz0iZmciIHN0eWxlPSJmbGV4OjE7bWF4LXdpZHRoOjUyMHB4Ij48bGFiZWw+U2VhcmNoIHNjaGVtZSwgQU1DIG9yIGNh"
    "dGVnb3J5PC9sYWJlbD48aW5wdXQgaWQ9InBtLXNlYXJjaCIgcGxhY2Vob2xkZXI9IlR5cGUgdG8gZmlsdGVy4oCmIj48L2Rpdj4K"
    "ICAgICAgPGRpdiBjbGFzcz0ic3BhY2VyIj48L2Rpdj4KICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuIGJ0bi1kYXJrIiBpZD0icG0t"
    "ZXhwIj4ke2ljb0Rvd24oKX0gRXhwb3J0IENTVjwvYnV0dG9uPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJjb250ZW50Ij4K"
    "ICAgICAgPGRpdiBjbGFzcz0ic2hlYWQiPjxoMj5QZWVyICZhbXA7IGJlbmNobWFyayBtYXBwaW5nIDxlbT7igJQgVmFsdWUgUmVz"
    "ZWFyY2ggZXhhY3QtcGVlciBzZXQ8L2VtPjwvaDI+PGRpdiBjbGFzcz0iY3R4IiBpZD0icG0tY291bnQiPjwvZGl2PjwvZGl2Pgog"
    "ICAgICA8ZGl2IGNsYXNzPSJjYXJkIiBzdHlsZT0icGFkZGluZy10b3A6MjBweCI+CiAgICAgICAgPGRpdiBjbGFzcz0ic2Nyb2xs"
    "LXRhbGwiPjx0YWJsZSBjbGFzcz0icnQiIGlkPSJwbS10YWJsZSI+PHRoZWFkPjx0cj4KICAgICAgICAgIDx0aCBzdHlsZT0id2lk"
    "dGg6NDIlIj5TY2hlbWU8L3RoPjx0aCBzdHlsZT0id2lkdGg6MTIlIj5BTUZJPC90aD48dGggc3R5bGU9IndpZHRoOjI0JSI+Q2F0"
    "ZWdvcnk8L3RoPjx0aCBzdHlsZT0id2lkdGg6MjIlIj5OYW1lIGluIE1GSTwvdGg+CiAgICAgICAgPC90cj48L3RoZWFkPjx0Ym9k"
    "eT48L3Rib2R5PjwvdGFibGU+PC9kaXY+CiAgICAgICAgPGRpdiBpZD0icG0tbWV0aG9kIj48L2Rpdj4KICAgICAgPC9kaXY+CiAg"
    "ICA8L2Rpdj4KICA8L3NlY3Rpb24+YDsKfQoKZnVuY3Rpb24gaWNvRG93bigpeyByZXR1cm4gJzxzdmcgdmlld0JveD0iMCAwIDI0"
    "IDI0IiBmaWxsPSJub25lIiBzdHJva2U9ImN1cnJlbnRDb2xvciIgc3Ryb2tlLXdpZHRoPSIyIj48cGF0aCBkPSJNMTIgM3YxMm0w"
    "IDBsLTQtNG00IDRsNC00TTQgMjFoMTYiLz48L3N2Zz4nOyB9CgovL+KUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKU"
    "gOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgCBTTEVFVkUgVklFVyDilIDilIDilIDilIDi"
    "lIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDi"
    "lIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVuZGVyU2xlZXZlKCl7CiAgY29uc3QgRCA9IFMuZGF0YTsKICAvLyBwb3B1bGF0ZSBm"
    "aWx0ZXJzIG9uY2UKICBjb25zdCBmaFNlbCA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzbC1maCcpOwogIGlmKCFmaFNlbC5v"
    "cHRpb25zLmxlbmd0aCl7CiAgICBjb25zdCBmaHMgPSB1bmlxKEQuc2xlZXZlLCdmaCcpLnNvcnQoZmhTb3J0KTsKICAgIGNvbnN0"
    "IGJycyA9IHVuaXEoRC5zbGVldmUsJ2JyJykuc29ydCgpOwogICAgZmhTZWwuaW5uZXJIVE1MID0gZmhzLm1hcChmPT5gPG9wdGlv"
    "biAke2Y9PT1TLnNsZWV2ZS5maD8nc2VsZWN0ZWQnOicnfT4ke2Z9PC9vcHRpb24+YCkuam9pbignJyk7CiAgICBkb2N1bWVudC5n"
    "ZXRFbGVtZW50QnlJZCgnc2wtYnInKS5pbm5lckhUTUwgPSBicnMubWFwKGI9PmA8b3B0aW9uICR7Yj09PVMuc2xlZXZlLmJyPydz"
    "ZWxlY3RlZCc6Jyd9PiR7Yn08L29wdGlvbj5gKS5qb2luKCcnKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzbC1yYW5r"
    "ZGF0ZScpLmlubmVySFRNTCA9IGRhdGVPcHRpb25zKEQuYXVtX2RhdGVzLCBTLnNsZWV2ZS5yYW5rRGF0ZSk7CiAgICAvLyB3aXJl"
    "CiAgICBmaFNlbC5vbmNoYW5nZSA9IGU9PnsgUy5zbGVldmUuZmg9ZS50YXJnZXQudmFsdWU7IHJlbmRlclNsZWV2ZSgpOyB9Owog"
    "ICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NsLWJyJykub25jaGFuZ2UgPSBlPT57IFMuc2xlZXZlLmJyPWUudGFyZ2V0LnZh"
    "bHVlOyByZW5kZXJTbGVldmUoKTsgfTsKICAgIHNlZ1dpcmUoJ3NsLXJ3Jywgdj0+eyBTLnNsZWV2ZS5ydz12OyByZW5kZXJTbGVl"
    "dmUoKTsgfSk7CiAgICBzZWdXaXJlKCdzbC1tb2RlJywgdj0+eyBTLnNsZWV2ZS5tb2RlPXY7IHJlbmRlclNsZWV2ZSgpOyB9KTsK"
    "ICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzbC1yYW5rZGF0ZScpLm9uY2hhbmdlID0gZT0+eyBTLnNsZWV2ZS5yYW5rRGF0"
    "ZT0rZS50YXJnZXQudmFsdWU7IHJlbmRlclNsZWV2ZVJhbmsoKTsgcmVuZGVyU2xlZXZlQmFyKCk7IH07CiAgICBkb2N1bWVudC5n"
    "ZXRFbGVtZW50QnlJZCgnc2wtZXhwJykub25jbGljayA9IGV4cG9ydFNsZWV2ZTsKICB9CgogIGNvbnN0IHtmaCxicixydyxtb2Rl"
    "fSA9IFMuc2xlZXZlOwogIGNvbnN0IGJ5USA9IHNsZWV2ZVNsaWNlKGZoLGJyLHJ3KTsKICBjb25zdCBkYXRlcyA9IEQuYXVtX2Rh"
    "dGVzOwogIGNvbnN0IGxpID0gZGF0ZXMubGVuZ3RoLTE7CiAgY29uc3Qgd2VpZ2h0ID0gYnlRWyclIG9mIEFNQyBBVU0nXSB8fCBb"
    "XTsKCiAgLy8gbm9ybWFsaXNlZCBzZXJpZXMgKHdpdGhpbiBzbGVldmUpIE9SIHJhdyAoJSBvZiBBTUMpCiAgY29uc3Qgbm9ybSA9"
    "IHE9PnsKICAgIGNvbnN0IHJhdyA9IGJ5UVtxXXx8W107CiAgICBpZihtb2RlPT09J2FtYycpIHJldHVybiByYXcubWFwKHY9PiB2"
    "PT1udWxsP251bGwgOiB2KjEwMCk7CiAgICByZXR1cm4gcmF3Lm1hcCgodixpKT0+eyBjb25zdCB3PXdlaWdodFtpXTsgcmV0dXJu"
    "ICh2PT1udWxsfHwhdyk/bnVsbCA6ICh2L3cpKjEwMDsgfSk7CiAgfTsKCiAgLy8g4pSA4pSAIEtQSXMKICBjb25zdCBxMT1ieVFb"
    "JyUgaW4gUTEnXT8uW2xpXSwgcTI9YnlRWyclIGluIFEyJ10/LltsaV0sIHEzPWJ5UVsnJSBpbiBRMyddPy5bbGldLCBxND1ieVFb"
    "JyUgaW4gUTQnXT8uW2xpXTsKICBjb25zdCB3PXdlaWdodFtsaV07CiAgY29uc3QgdG9wVyA9IChxMXx8MCkrKHEyfHwwKTsKICBj"
    "b25zdCB0b3BTaGFyZSA9IHc/IHRvcFcvdyA6IG51bGw7CiAgY29uc3Qgc2NvcmUgPSBxdWFsaXR5U2NvcmUoYnlRLCBsaSk7CiAg"
    "ZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NsLXRpdGxlJykuaW5uZXJIVE1MID0gYCR7Zmh9IMK3ICR7YnJ9YDsKICBkb2N1bWVu"
    "dC5nZXRFbGVtZW50QnlJZCgnc2wtY3R4JykuaW5uZXJIVE1MID0gYCR7cnd9IHJvbGxpbmcgwrcgcXVhcnRpbGVzIHdpdGhpbiBN"
    "RkkgYWxsLXBlZXIgY2F0ZWdvcnk8YnI+U25hcHNob3QgJHtmbXQuZGF0ZShkYXRlc1tsaV0pfWA7CiAgZG9jdW1lbnQuZ2V0RWxl"
    "bWVudEJ5SWQoJ3NsLWtwaXMnKS5pbm5lckhUTUwgPSBgCiAgICA8ZGl2IGNsYXNzPSJrcGkgZ29vZCI+PGRpdiBjbGFzcz0iay1s"
    "YmwiPlExK1EyIHdpdGhpbiBzbGVldmU8L2Rpdj48ZGl2IGNsYXNzPSJrLXZhbCBnb29kIj4ke2ZtdC5wY3QodG9wU2hhcmUpfTwv"
    "ZGl2PjxkaXYgY2xhc3M9Imstc3ViIj50b3AtaGFsZiBvZiBzbGVldmUgQVVNPC9kaXY+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJr"
    "cGkgYWNjZW50Ij48ZGl2IGNsYXNzPSJrLWxibCI+U2xlZXZlIHdlaWdodCBpbiBBTUM8L2Rpdj48ZGl2IGNsYXNzPSJrLXZhbCBh"
    "Y2NlbnQiPiR7Zm10LnBjdCh3KX08L2Rpdj48ZGl2IGNsYXNzPSJrLXN1YiI+b2YgdG90YWwgJHtmaH0gQVVNPC9kaXY+PC9kaXY+"
    "CiAgICA8ZGl2IGNsYXNzPSJrcGkiPjxkaXYgY2xhc3M9ImstbGJsIj5BVU0td3RkIHF1YXJ0aWxlIHNjb3JlPC9kaXY+PGRpdiBj"
    "bGFzcz0iay12YWwiPiR7Zm10Lm51bShzY29yZSl9PC9kaXY+PGRpdiBjbGFzcz0iay1zdWIiPjEgKHdvcnN0KSDihpIgNCAoYmVz"
    "dCk8L2Rpdj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImtwaSB3YXJuIj48ZGl2IGNsYXNzPSJrLWxibCI+UTQgd2l0aGluIHNsZWV2"
    "ZTwvZGl2PjxkaXYgY2xhc3M9ImstdmFsIHdhcm4iPiR7Zm10LnBjdCh3PyhxNHx8MCkvdzpudWxsKX08L2Rpdj48ZGl2IGNsYXNz"
    "PSJrLXN1YiI+Ym90dG9tLXF1YXJ0aWxlIEFVTTwvZGl2PjwvZGl2PmA7CgogIC8vIOKUgOKUgCBtYWluIHN0YWNrZWQgY2hhcnQg"
    "KyB3ZWlnaHQgb24gc2Vjb25kYXJ5IGF4aXMKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2wtYzEtdGl0bGUnKS50ZXh0Q29u"
    "dGVudCA9IG1vZGU9PT0nd2l0aGluJwogICAgPyAnUXVhcnRpbGUgZGlzdHJpYnV0aW9uIHdpdGhpbiBzbGVldmUnIDogJ1F1YXJ0"
    "aWxlIGRpc3RyaWJ1dGlvbiBhcyAlIG9mIEFNQyBBVU0nOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzbC1jMS1zdWInKS50"
    "ZXh0Q29udGVudCA9IG1vZGU9PT0nd2l0aGluJwogICAgPyAnRWFjaCBiYW5kID0gdGhhdCBxdWFydGlsZeKAmXMgc2hhcmUgb2Yg"
    "dGhlIHNsZWV2ZSAoc3VtcyB0byAxMDAlKTsgcmVkIGxpbmUgPSBzbGVldmUgd2VpZ2h0IGluIHRoZSBob3VzZScKICAgIDogJ0Vh"
    "Y2ggYmFuZCA9IHRoYXQgcXVhcnRpbGXigJlzIHNoYXJlIG9mIHRvdGFsIEFNQyBBVU0gKHN1bXMgdG8gdGhlIHNsZWV2ZSB3ZWln"
    "aHQpJzsKCiAgY29uc3QgY3R4ID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NsLXN0YWNrJykuZ2V0Q29udGV4dCgnMmQnKTsK"
    "ICBTLmNoYXJ0cy5zbFN0YWNrPy5kZXN0cm95KCk7CiAgY29uc3QgYXJlYURzID0gUU9SREVSLm1hcChxPT4oewogICAgbGFiZWw6"
    "cS5yZXBsYWNlKCclIGluICcsJycpLCBkYXRhOiBub3JtKHEpLm1hcCgoeSxpKT0+KHt4OmRhdGVzW2ldLHl9KSksCiAgICBiYWNr"
    "Z3JvdW5kQ29sb3I6KGMpPT5ncmFkaWVudChjLmNoYXJ0LmN0eCxjLmNoYXJ0LmNoYXJ0QXJlYSxRQ09MW3FdLDAuODUsMC41NSks"
    "CiAgICBib3JkZXJDb2xvcjpRQ09MW3FdLCBib3JkZXJXaWR0aDowLjYsIGZpbGw6dHJ1ZSwgdGVuc2lvbjowLjMyLCBwb2ludFJh"
    "ZGl1czowLCB5QXhpc0lEOid5Jywgc3RhY2s6J3EnCiAgfSkpOwogIGNvbnN0IHdlaWdodERzID0gewogICAgbGFiZWw6J1NsZWV2"
    "ZSB3ZWlnaHQnLCBkYXRhOiB3ZWlnaHQubWFwKCh2LGkpPT4oe3g6ZGF0ZXNbaV0seTp2PT1udWxsP251bGw6dioxMDB9KSksCiAg"
    "ICBib3JkZXJDb2xvcjpDLmFic2wsIGJvcmRlcldpZHRoOjIuNCwgYm9yZGVyRGFzaDpbNSwzXSwgZmlsbDpmYWxzZSwgdGVuc2lv"
    "bjowLjMyLCBwb2ludFJhZGl1czowLCB5QXhpc0lEOid5MScKICB9OwogIFMuY2hhcnRzLnNsU3RhY2sgPSBuZXcgQ2hhcnQoY3R4"
    "LHsgdHlwZTonbGluZScsIGRhdGE6e2RhdGFzZXRzOlsuLi5hcmVhRHMsIHdlaWdodERzXX0sCiAgICBvcHRpb25zOiBzdGFja09w"
    "dHMoeyBsZWZ0TWF4OiBtb2RlPT09J3dpdGhpbic/MTAwOnVuZGVmaW5lZCwgcmlnaHRMYWJlbDonU2xlZXZlIHdlaWdodCAlJywg"
    "cmlnaHRBdXRvOnRydWUgfSkgfSk7CgogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzbC1jMS1tZXRob2QnKS5pbm5lckhUTUwg"
    "PSBtZXRob2ROb3RlKAogICAgbW9kZT09PSd3aXRoaW4nCiAgICA/IGBGb3IgdGhlIHNlbGVjdGVkIGhvdXNlIGFuZCBzbGVldmUs"
    "IGVhY2ggc2NoZW1l4oCZcyBOQVYgaXMgcmFua2VkIGFnYWluc3QgaXRzIDxiPk1GSSBhbGwtcGVlciBjYXRlZ29yeTwvYj4gb3Zl"
    "ciB0aGUgJHtydy50b0xvd2VyQ2FzZSgpfSB3aW5kb3cgYW5kIGJ1Y2tldGVkIGludG8gcXVhcnRpbGVzIChRMSA9IHRvcCBwZXJm"
    "b3JtZXJzKS4gRWFjaCBzY2hlbWXigJlzIDxiPnNoYXJlIG9mIHRoZSBob3VzZeKAmXMgdG90YWwgQVVNPC9iPiBpcyBzdW1tZWQg"
    "YnkgcXVhcnRpbGUsIHRoZW4gPGI+ZGl2aWRlZCBieSB0aGUgc2xlZXZl4oCZcyB0b3RhbCB3ZWlnaHQ8L2I+IHNvIHRoZSBmb3Vy"
    "IGJhbmRzIHN1bSB0byAxMDAlIOKAlCBzaG93aW5nIGhvdyB0aGUgc2xlZXZl4oCZcyBhc3NldHMgYXJlIGRpc3RyaWJ1dGVkIGFj"
    "cm9zcyBwZXJmb3JtYW5jZSBxdWFydGlsZXMgaXJyZXNwZWN0aXZlIG9mIHNsZWV2ZSBzaXplLiBUaGUgZGFzaGVkIHJlZCBsaW5l"
    "IChyaWdodCBheGlzKSBpcyB0aGUgc2xlZXZl4oCZcyByYXcgd2VpZ2h0IGluIHRoZSBob3VzZSwgc28geW91IGNhbiByZWFkIGNv"
    "bmNlbnRyYXRpb24gYW5kIHNpemUgdG9nZXRoZXIuYAogICAgOiBgQmFuZHMgc2hvdyBlYWNoIHF1YXJ0aWxl4oCZcyBzaGFyZSBv"
    "ZiA8Yj50b3RhbCAke2ZofSBBVU08L2I+OyB0aGV5IHN1bSB0byB0aGUgc2xlZXZl4oCZcyBvdmVyYWxsIHdlaWdodCBpbiB0aGUg"
    "aG91c2UgKHRoZSBkYXNoZWQgbGluZSkuIFVzZSDigJxXaXRoaW4gc2xlZXZl4oCdIHRvIG5vcm1hbGlzZSB0aGUgYmFuZHMgdG8g"
    "MTAwJSBpbnN0ZWFkLmAKICApOwoKICByZW5kZXJTbGVldmVUcmVuZCgpOyByZW5kZXJTbGVldmVSYW5rKCk7IHJlbmRlclNsZWV2"
    "ZUJhcigpOwp9CgpmdW5jdGlvbiByZW5kZXJTbGVldmVUcmVuZCgpewogIGNvbnN0IEQ9Uy5kYXRhOyBjb25zdCB7ZmgsYnIscnd9"
    "PVMuc2xlZXZlOyBjb25zdCBieVE9c2xlZXZlU2xpY2UoZmgsYnIscncpOwogIGNvbnN0IGRhdGVzPUQuYXVtX2RhdGVzOyBjb25z"
    "dCB3ZWlnaHQ9YnlRWyclIG9mIEFNQyBBVU0nXXx8W107CiAgY29uc3QgdG9wSGFsZiA9IGRhdGVzLm1hcCgoZCxpKT0+eyBjb25z"
    "dCB3PXdlaWdodFtpXTsgY29uc3QgdD0oYnlRWyclIGluIFExJ10/LltpXXx8MCkrKGJ5UVsnJSBpbiBRMiddPy5baV18fDApOyBy"
    "ZXR1cm4ge3g6ZCwgeTp3PyAodC93KSoxMDAgOiBudWxsfTsgfSk7CiAgY29uc3Qgc2NvcmVTZXJpZXMgPSBkYXRlcy5tYXAoKGQs"
    "aSk9Pih7eDpkLCB5OnF1YWxpdHlTY29yZShieVEsaSl9KSk7CgogIGNvbnN0IGN0eD1kb2N1bWVudC5nZXRFbGVtZW50QnlJZCgn"
    "c2wtdHJlbmQnKS5nZXRDb250ZXh0KCcyZCcpOwogIFMuY2hhcnRzLnNsVHJlbmQ/LmRlc3Ryb3koKTsKICBTLmNoYXJ0cy5zbFRy"
    "ZW5kID0gbmV3IENoYXJ0KGN0eCx7IHR5cGU6J2xpbmUnLCBkYXRhOntkYXRhc2V0czpbCiAgICB7IGxhYmVsOidRMStRMiAlJywg"
    "ZGF0YTp0b3BIYWxmLCBib3JkZXJDb2xvcjpDLnExLCBib3JkZXJXaWR0aDoyLjIsIHRlbnNpb246MC4zMiwgcG9pbnRSYWRpdXM6"
    "MCwgeUF4aXNJRDoneScsCiAgICAgIGJhY2tncm91bmRDb2xvcjooYyk9PmdyYWRpZW50KGMuY2hhcnQuY3R4LGMuY2hhcnQuY2hh"
    "cnRBcmVhLEMucTEsMC4xOCwwLjAxKSwgZmlsbDp0cnVlIH0sCiAgICB7IGxhYmVsOidRdWFsaXR5IHNjb3JlJywgZGF0YTpzY29y"
    "ZVNlcmllcywgYm9yZGVyQ29sb3I6Qy5uYXZ5LCBib3JkZXJXaWR0aDoyLCB0ZW5zaW9uOjAuMzIsIHBvaW50UmFkaXVzOjAsIHlB"
    "eGlzSUQ6J3kxJywgYm9yZGVyRGFzaDpbNCwzXSwgZmlsbDpmYWxzZSB9CiAgXX0sIG9wdGlvbnM6IGR1YWxPcHRzKHsgbGVmdExh"
    "YmVsOidRMStRMiAlJywgbGVmdFN1ZmZpeDonJScsIGxlZnRBdXRvOnRydWUsIHJpZ2h0TGFiZWw6J1Njb3JlIDHigJM0Jywgcmln"
    "aHRNaW46MSwgcmlnaHRNYXg6NCB9KSB9KTsKCiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NsLXRyZW5kLW1ldGhvZCcpLmlu"
    "bmVySFRNTCA9IG1ldGhvZE5vdGUoCiAgICBgPGI+UTErUTIgJTwvYj4gaXMgdGhlIHRvcC1oYWxmIHNoYXJlIHdpdGhpbiB0aGUg"
    "c2xlZXZlIChsZWZ0LCBhdXRvLXNjYWxlZCB0byB0aGUgZGF0YSkuIDxiPlF1YWxpdHkgc2NvcmU8L2I+IChyaWdodCwgZml4ZWQg"
    "MeKAkzQpIGlzIHRoZSBBVU0td2VpZ2h0ZWQgbWVhbiBxdWFydGlsZSwgc2NvcmluZyBRMT00LCBRMj0zLCBRMz0yLCBRND0xIOKA"
    "lCBhIHNpbmdsZSByZWFkIG9uIHdoZXRoZXIgdGhlIGhvdXNl4oCZcyBtb25leSBpbiB0aGlzIHNsZWV2ZSBzaXRzIHdpdGggd2lu"
    "bmVycyBvciBsYWdnYXJkcy5gKTsKfQoKZnVuY3Rpb24gcmVuZGVyU2xlZXZlUmFuaygpewogIGNvbnN0IEQ9Uy5kYXRhOyBjb25z"
    "dCB7YnIscncsbW9kZSxyYW5rRGF0ZX09Uy5zbGVldmU7IGNvbnN0IGk9cmFua0RhdGU7CiAgY29uc3QgZmhzID0gdW5pcShELnNs"
    "ZWV2ZS5maWx0ZXIocj0+ci5icj09PWJyICYmIHIucnc9PT1ydyksJ2ZoJyk7CiAgY29uc3Qgcm93cyA9IGZocy5tYXAoZmg9PnsK"
    "ICAgIGNvbnN0IGJ5UT1zbGVldmVTbGljZShmaCxicixydyk7IGNvbnN0IHc9YnlRWyclIG9mIEFNQyBBVU0nXT8uW2ldfHwwOwog"
    "ICAgY29uc3QgdG9wPShieVFbJyUgaW4gUTEnXT8uW2ldfHwwKSsoYnlRWyclIGluIFEyJ10/LltpXXx8MCk7CiAgICByZXR1cm4g"
    "eyBmaCwgdG9wLCB0b3BTaGFyZTp3P3RvcC93OjAsIHdlaWdodDp3LCBxNDpieVFbJyUgaW4gUTQnXT8uW2ldfHwwLCBzY29yZTpx"
    "dWFsaXR5U2NvcmUoYnlRLGkpIH07CiAgfSkuZmlsdGVyKHI9PnIud2VpZ2h0PjApLnNvcnQoKGEsYik9PmIudG9wU2hhcmUtYS50"
    "b3BTaGFyZSk7CiAgcm93cy5mb3JFYWNoKChyLGlkeCk9PnIucmFuaz1pZHgrMSk7CiAgY29uc3Qgbj1yb3dzLmxlbmd0aDsKICBs"
    "ZXQgaHRtbD0nPHRhYmxlIGNsYXNzPSJydCI+PHRoZWFkPjx0cj48dGg+IzwvdGg+PHRoPkZ1bmQgSG91c2U8L3RoPjx0aCBjbGFz"
    "cz0ibnVtIj5RMStRMjwvdGg+PHRoIGNsYXNzPSJudW0iPlNjb3JlPC90aD48dGggY2xhc3M9Im51bSI+V2VpZ2h0PC90aD48L3Ry"
    "PjwvdGhlYWQ+PHRib2R5Pic7CiAgcm93cy5mb3JFYWNoKHI9PnsKICAgIGNvbnN0IGNscz1yLmZoPT09Uy5zbGVldmUuZmg/J2hs"
    "JzonJzsKICAgIGNvbnN0IHBjPXIucmFuaz09PTE/J3IxJzpyLnJhbms8PU1hdGguY2VpbChuLzQpPyd1cCc6ci5yYW5rPm4tTWF0"
    "aC5jZWlsKG4vNCk/J2RuJzonJzsKICAgIGh0bWwrPWA8dHIgY2xhc3M9IiR7Y2xzfSI+PHRkPjxzcGFuIGNsYXNzPSJwaWxsICR7"
    "cGN9Ij4ke3IucmFua308L3NwYW4+PC90ZD48dGQgY2xhc3M9ImxibCI+JHtyLmZofTwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10"
    "LnBjdChyLnRvcFNoYXJlKX08L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdC5udW0oci5zY29yZSl9PC90ZD48dGQgY2xhc3M9Im51"
    "bSI+JHtmbXQucGN0KHIud2VpZ2h0KX08L3RkPjwvdHI+YDsKICB9KTsKICBodG1sKz0nPC90Ym9keT48L3RhYmxlPic7CiAgZG9j"
    "dW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NsLXJhbmsnKS5pbm5lckhUTUw9aHRtbDsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgn"
    "c2wtcmFuay1tZXRob2QnKS5pbm5lckhUTUwgPSBtZXRob2ROb3RlKAogICAgYEhvdXNlcyByYW5rZWQgYnkgPGI+UTErUTIgc2hh"
    "cmUgd2l0aGluIHRoZSAke2JyfSBzbGVldmU8L2I+IG9uIDxiPiR7Zm10LmRhdGUoRC5hdW1fZGF0ZXNbaV0pfTwvYj4gKCR7cnd9"
    "IHdpbmRvdykuIOKAnFNjb3Jl4oCdIGlzIHRoZSBBVU0td2VpZ2h0ZWQgcXVhcnRpbGUgKDHigJM0KTsg4oCcV2VpZ2h04oCdIGlz"
    "IHRoZSBzbGVldmXigJlzIHNpemUgaW4gZWFjaCBob3VzZS4gQ2hhbmdlIHRoZSBhcy1vZiBkYXRlIHRvIHNlZSBob3cgdGhlIGxl"
    "YWd1ZSBoYXMgc2hpZnRlZCBoaXN0b3JpY2FsbHkuYCk7Cn0KCmZ1bmN0aW9uIHJlbmRlclNsZWV2ZUJhcigpewogIGNvbnN0IEQ9"
    "Uy5kYXRhOyBjb25zdCB7YnIscncscmFua0RhdGV9PVMuc2xlZXZlOyBjb25zdCBpPXJhbmtEYXRlOwogIGNvbnN0IGZocyA9IHVu"
    "aXEoRC5zbGVldmUuZmlsdGVyKHI9PnIuYnI9PT1iciAmJiByLnJ3PT09cncpLCdmaCcpOwogIGNvbnN0IHJvd3MgPSBmaHMubWFw"
    "KGZoPT57IGNvbnN0IGJ5UT1zbGVldmVTbGljZShmaCxicixydyk7IGNvbnN0IHc9YnlRWyclIG9mIEFNQyBBVU0nXT8uW2ldfHww"
    "OwogICAgY29uc3QgdG9wPShieVFbJyUgaW4gUTEnXT8uW2ldfHwwKSsoYnlRWyclIGluIFEyJ10/LltpXXx8MCk7IHJldHVybiB7"
    "ZmgsIHk6dz8odG9wL3cpKjEwMDowLCB3ZWlnaHQ6d307IH0pCiAgICAuZmlsdGVyKHI9PnIud2VpZ2h0PjApLnNvcnQoKGEsYik9"
    "PmIueS1hLnkpOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzbC1iYXItc3ViJykudGV4dENvbnRlbnQgPSBgUTErUTIgc2hh"
    "cmUgd2l0aGluICR7YnJ9IGFjcm9zcyBob3VzZXMgwrcgJHtmbXQuZGF0ZShELmF1bV9kYXRlc1tpXSl9YDsKICBjb25zdCBjdHg9"
    "ZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NsLWJhcicpLmdldENvbnRleHQoJzJkJyk7CiAgUy5jaGFydHMuc2xCYXI/LmRlc3Ry"
    "b3koKTsKICBTLmNoYXJ0cy5zbEJhciA9IG5ldyBDaGFydChjdHgseyB0eXBlOidiYXInLCBkYXRhOnsgbGFiZWxzOnJvd3MubWFw"
    "KHI9PnIuZmgpLCBkYXRhc2V0czpbewogICAgZGF0YTpyb3dzLm1hcChyPT5yLnkpLCBiYWNrZ3JvdW5kQ29sb3I6cm93cy5tYXAo"
    "cj0+ci5maD09PVMuc2xlZXZlLmZoP0MuYWJzbDpDLnBlZXIpLAogICAgYm9yZGVyUmFkaXVzOjQsIGJvcmRlclNraXBwZWQ6ZmFs"
    "c2UgfV19LAogICAgb3B0aW9uczp7IG1haW50YWluQXNwZWN0UmF0aW86ZmFsc2UsIHJlc3BvbnNpdmU6dHJ1ZSwgcGx1Z2luczp7"
    "bGVnZW5kOntkaXNwbGF5OmZhbHNlfSwKICAgICAgdG9vbHRpcDp7Y2FsbGJhY2tzOntsYWJlbDpjPT4nUTErUTI6ICcrYy5wYXJz"
    "ZWQueS50b0ZpeGVkKDEpKyclJ319fSwKICAgICAgc2NhbGVzOnsgeDp7Z3JpZDp7ZGlzcGxheTpmYWxzZX0sdGlja3M6e2NvbG9y"
    "OkMuaW5rMyxmb250OntzaXplOjEwfSxtYXhSb3RhdGlvbjo1NSxtaW5Sb3RhdGlvbjo1NX19LAogICAgICAgIHk6e2dyYWNlOic4"
    "JScsdGlja3M6e2NhbGxiYWNrOnY9PnYrJyUnLGNvbG9yOkMuaW5rM30sZ3JpZDp7Y29sb3I6Qy5saW5lMn19IH0gfSB9KTsKICBk"
    "b2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2wtYmFyLW1ldGhvZCcpLmlubmVySFRNTCA9IG1ldGhvZE5vdGUoCiAgICBgU2FtZSBt"
    "ZXRyaWMgYXMgdGhlIGxlYWd1ZSB0YWJsZSwgZHJhd24gYXMgYSBjcm9zcy1zZWN0aW9uIHNvIHlvdSBjYW4gZXllYmFsbCB3aGVy"
    "ZSA8Yj4ke1Muc2xlZXZlLmZofTwvYj4gKHJlZCkgc2l0cyBpbiB0aGUgZGlzdHJpYnV0aW9uIG9uIHRoZSBzZWxlY3RlZCBkYXRl"
    "LiBBdXRvLXNjYWxlcyB0byB0aGUgc3ByZWFkIG9mIGhvdXNlcyBwcmVzZW50IGluIHRoZSBzbGVldmUuYCk7Cn0KCmZ1bmN0aW9u"
    "IGV4cG9ydFNsZWV2ZSgpewogIGNvbnN0IEQ9Uy5kYXRhOyBjb25zdCB7ZmgsYnIscncsbW9kZX09Uy5zbGVldmU7IGNvbnN0IGJ5"
    "UT1zbGVldmVTbGljZShmaCxicixydyk7IGNvbnN0IHc9YnlRWyclIG9mIEFNQyBBVU0nXXx8W107CiAgbGV0IGNzdj0nRGF0ZSxR"
    "MSxRMixRMyxRNCxTbGVldmUgd2VpZ2h0LEJhc2lzXG4nOwogIEQuYXVtX2RhdGVzLmZvckVhY2goKGQsaSk9PnsKICAgIGNvbnN0"
    "IGY9cT0+eyBjb25zdCByYXc9YnlRW3FdPy5baV07IGlmKHJhdz09bnVsbCkgcmV0dXJuICcnOyByZXR1cm4gbW9kZT09PSdhbWMn"
    "P3Jhdzood1tpXT9yYXcvd1tpXTonJyk7IH07CiAgICBjc3YrPVtkLGYoJyUgaW4gUTEnKSxmKCclIGluIFEyJyksZignJSBpbiBR"
    "MycpLGYoJyUgaW4gUTQnKSx3W2ldPz8nJyxtb2RlXS5qb2luKCcsJykrJ1xuJzsKICB9KTsKICBkb3dubG9hZENTVihjc3YsIGBz"
    "bGVldmVfJHtmaH1fJHticn1fJHtydy5yZXBsYWNlKCcgJywnJyl9XyR7bW9kZX0uY3N2YCk7Cn0KCmZ1bmN0aW9uIHNlZ1dpcmUo"
    "aWQsIGNiKXsKICBkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKGAjJHtpZH0gYnV0dG9uYCkuZm9yRWFjaChiPT57CiAgICBiLm9u"
    "Y2xpY2s9KCk9PnsgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbChgIyR7aWR9IGJ1dHRvbmApLmZvckVhY2goeD0+eC5jbGFzc0xp"
    "c3QucmVtb3ZlKCdhY3RpdmUnKSk7IGIuY2xhc3NMaXN0LmFkZCgnYWN0aXZlJyk7IGNiKGIuZGF0YXNldC52KTsgfTsKICB9KTsK"
    "fQoKLy/ilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDi"
    "lIDilIDilIDilIDilIAgQ0hBUlQgT1BUSU9OIEZBQ1RPUklFUyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDi"
    "lIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gYmFzZVRvb2x0aXAoKXsKICByZXR1cm4geyBtb2RlOidp"
    "bmRleCcsIGludGVyc2VjdDpmYWxzZSwgYmFja2dyb3VuZENvbG9yOicjMTYxMzBmJywgdGl0bGVDb2xvcjonI2YzZWFkOCcsIGJv"
    "ZHlDb2xvcjonI2U3ZGNjNicsCiAgICBib3JkZXJDb2xvcjoncmdiYSgyNDMsMjM0LDIxNiwwLjE1KScsIGJvcmRlcldpZHRoOjEs"
    "IHBhZGRpbmc6MTEsIGNvcm5lclJhZGl1czo2LCB0aXRsZUZvbnQ6e2ZhbWlseToiJ0pldEJyYWlucyBNb25vJyxtb25vc3BhY2Ui"
    "LHNpemU6MTF9LAogICAgYm9keUZvbnQ6e2ZhbWlseToiJ0ludGVyJyxzYW5zLXNlcmlmIixzaXplOjEyfSwgZGlzcGxheUNvbG9y"
    "czp0cnVlLCBib3hXaWR0aDo5LCBib3hIZWlnaHQ6OSwgdXNlUG9pbnRTdHlsZTp0cnVlIH07Cn0KZnVuY3Rpb24gdGltZVgoKXsg"
    "cmV0dXJuIHsgdHlwZTondGltZScsIHRpbWU6e3VuaXQ6J3llYXInfSwgZ3JpZDp7ZGlzcGxheTpmYWxzZX0sIHRpY2tzOntjb2xv"
    "cjpDLmluazMsIGZvbnQ6e3NpemU6MTB9fSwgYm9yZGVyOntjb2xvcjpDLmxpbmV9IH07IH0KCi8vIHN0YWNrZWQgYXJlYSAobGVm"
    "dCkgKyBvcHRpb25hbCBzZWNvbmRhcnkgbGluZSAocmlnaHQpCmZ1bmN0aW9uIHN0YWNrT3B0cyh7bGVmdE1heCwgcmlnaHRMYWJl"
    "bCwgcmlnaHRBdXRvfSl7CiAgcmV0dXJuIHsgbWFpbnRhaW5Bc3BlY3RSYXRpbzpmYWxzZSwgcmVzcG9uc2l2ZTp0cnVlLCBpbnRl"
    "cmFjdGlvbjp7bW9kZTonaW5kZXgnLGludGVyc2VjdDpmYWxzZX0sCiAgICBhbmltYXRpb246e2R1cmF0aW9uOjYwMCwgZWFzaW5n"
    "OidlYXNlT3V0Q3ViaWMnfSwKICAgIHBsdWdpbnM6eyBsZWdlbmQ6e2Rpc3BsYXk6ZmFsc2V9LCB0b29sdGlwOnsuLi5iYXNlVG9v"
    "bHRpcCgpLAogICAgICBjYWxsYmFja3M6eyBsYWJlbDpjPT4gYCR7Yy5kYXRhc2V0LmxhYmVsfTogJHtjLnBhcnNlZC55PT1udWxs"
    "PyfigJQnOmMucGFyc2VkLnkudG9GaXhlZCgxKSsoYy5kYXRhc2V0LnlBeGlzSUQ9PT0neTEnPyclJzonJScpfWAgfSB9IH0sCiAg"
    "ICBzY2FsZXM6ewogICAgICB4OiB0aW1lWCgpLAogICAgICB5OiB7IHN0YWNrZWQ6dHJ1ZSwgbWluOjAsIG1heDpsZWZ0TWF4LCBn"
    "cmFjZTogbGVmdE1heD91bmRlZmluZWQ6JzUlJywKICAgICAgICAgICB0aWNrczp7Y2FsbGJhY2s6dj0+disnJScsIGNvbG9yOkMu"
    "aW5rMywgZm9udDp7c2l6ZToxMH19LCBncmlkOntjb2xvcjpDLmxpbmUyfSwgYm9yZGVyOntkaXNwbGF5OmZhbHNlfSB9LAogICAg"
    "ICB5MTp7IHBvc2l0aW9uOidyaWdodCcsIGRpc3BsYXk6ISFyaWdodExhYmVsLCBncmFjZTonMTIlJywgbWluOjAsCiAgICAgICAg"
    "ICAgdGlja3M6e2NhbGxiYWNrOnY9PnYrJyUnLCBjb2xvcjpDLmFic2wsIGZvbnQ6e3NpemU6MTB9fSwgZ3JpZDp7ZHJhd09uQ2hh"
    "cnRBcmVhOmZhbHNlfSwgYm9yZGVyOntkaXNwbGF5OmZhbHNlfSwKICAgICAgICAgICB0aXRsZTp7ZGlzcGxheTohIXJpZ2h0TGFi"
    "ZWwsIHRleHQ6cmlnaHRMYWJlbCwgY29sb3I6Qy5hYnNsLCBmb250OntzaXplOjEwfX0gfQogICAgfSB9Owp9CgovLyBnZW5lcmlj"
    "IGR1YWwtYXhpcyBsaW5lIGNoYXJ0CmZ1bmN0aW9uIGR1YWxPcHRzKHtsZWZ0TGFiZWwsbGVmdFN1ZmZpeCxsZWZ0QXV0byxsZWZ0"
    "TWluLGxlZnRNYXgscmlnaHRMYWJlbCxyaWdodE1pbixyaWdodE1heH0pewogIHJldHVybiB7IG1haW50YWluQXNwZWN0UmF0aW86"
    "ZmFsc2UsIHJlc3BvbnNpdmU6dHJ1ZSwgaW50ZXJhY3Rpb246e21vZGU6J2luZGV4JyxpbnRlcnNlY3Q6ZmFsc2V9LAogICAgYW5p"
    "bWF0aW9uOntkdXJhdGlvbjo2MDAsIGVhc2luZzonZWFzZU91dEN1YmljJ30sCiAgICBwbHVnaW5zOnsgbGVnZW5kOntkaXNwbGF5"
    "OmZhbHNlfSwgdG9vbHRpcDp7Li4uYmFzZVRvb2x0aXAoKSwKICAgICAgY2FsbGJhY2tzOnsgbGFiZWw6Yz0+eyBjb25zdCB2PWMu"
    "cGFyc2VkLnk7IGlmKHY9PW51bGwpIHJldHVybiBgJHtjLmRhdGFzZXQubGFiZWx9OiDigJRgOwogICAgICAgIHJldHVybiBjLmRh"
    "dGFzZXQueUF4aXNJRD09PSd5MScgPyBgJHtjLmRhdGFzZXQubGFiZWx9OiAke3YudG9GaXhlZCgyKX1gIDogYCR7Yy5kYXRhc2V0"
    "LmxhYmVsfTogJHt2LnRvRml4ZWQoMSl9JHtsZWZ0U3VmZml4fHwnJ31gOyB9IH0gfSB9LAogICAgc2NhbGVzOnsKICAgICAgeDog"
    "dGltZVgoKSwKICAgICAgeTogeyBwb3NpdGlvbjonbGVmdCcsIG1pbjpsZWZ0TWluLCBtYXg6bGVmdE1heCwgZ3JhY2U6KGxlZnRN"
    "aW49PW51bGwmJmxlZnRNYXg9PW51bGwpPyc4JSc6dW5kZWZpbmVkLAogICAgICAgICAgIHRpY2tzOntjYWxsYmFjazp2PT52Kyhs"
    "ZWZ0U3VmZml4fHwnJyksIGNvbG9yOkMuaW5rMywgZm9udDp7c2l6ZToxMH19LCBncmlkOntjb2xvcjpDLmxpbmUyfSwgYm9yZGVy"
    "OntkaXNwbGF5OmZhbHNlfSwKICAgICAgICAgICB0aXRsZTp7ZGlzcGxheTohIWxlZnRMYWJlbCwgdGV4dDpsZWZ0TGFiZWwsIGNv"
    "bG9yOkMuaW5rMywgZm9udDp7c2l6ZToxMH19IH0sCiAgICAgIHkxOnsgcG9zaXRpb246J3JpZ2h0JywgbWluOnJpZ2h0TWluLCBt"
    "YXg6cmlnaHRNYXgsIGdyaWQ6e2RyYXdPbkNoYXJ0QXJlYTpmYWxzZX0sCiAgICAgICAgICAgdGlja3M6e2NvbG9yOkMubmF2eSwg"
    "Zm9udDp7c2l6ZToxMH19LCBib3JkZXI6e2Rpc3BsYXk6ZmFsc2V9LAogICAgICAgICAgIHRpdGxlOntkaXNwbGF5OiEhcmlnaHRM"
    "YWJlbCwgdGV4dDpyaWdodExhYmVsLCBjb2xvcjpDLm5hdnksIGZvbnQ6e3NpemU6MTB9fSB9CiAgICB9IH07Cn0KCi8vIHNpbXBs"
    "ZSBtdWx0aS1saW5lIChhdXRvLXNjYWxlZCkKZnVuY3Rpb24gbGluZU9wdHMoe3N1ZmZpeCwgbWluLCBtYXh9KXsKICByZXR1cm4g"
    "eyBtYWludGFpbkFzcGVjdFJhdGlvOmZhbHNlLCByZXNwb25zaXZlOnRydWUsIGludGVyYWN0aW9uOnttb2RlOiduZWFyZXN0Jyxp"
    "bnRlcnNlY3Q6ZmFsc2V9LAogICAgYW5pbWF0aW9uOntkdXJhdGlvbjo2MDAsIGVhc2luZzonZWFzZU91dEN1YmljJ30sCiAgICBw"
    "bHVnaW5zOnsgbGVnZW5kOntkaXNwbGF5OmZhbHNlfSwgdG9vbHRpcDp7Li4uYmFzZVRvb2x0aXAoKSwKICAgICAgY2FsbGJhY2tz"
    "OnsgbGFiZWw6Yz0+IGAke2MuZGF0YXNldC5sYWJlbH06ICR7Yy5wYXJzZWQueT09bnVsbD8n4oCUJzpjLnBhcnNlZC55LnRvRml4"
    "ZWQoMSkrKHN1ZmZpeHx8JycpfWAgfSB9IH0sCiAgICBzY2FsZXM6eyB4OnRpbWVYKCksIHk6eyBtaW4sIG1heCwgZ3JhY2U6KG1p"
    "bj09bnVsbCYmbWF4PT1udWxsKT8nNiUnOnVuZGVmaW5lZCwKICAgICAgdGlja3M6e2NhbGxiYWNrOnY9PnYrKHN1ZmZpeHx8Jycp"
    "LCBjb2xvcjpDLmluazMsIGZvbnQ6e3NpemU6MTB9fSwgZ3JpZDp7Y29sb3I6Qy5saW5lMn0sIGJvcmRlcjp7ZGlzcGxheTpmYWxz"
    "ZX0gfSB9IH07Cn0KCi8v4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSAIEFNQyBWSUVXIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKU"
    "gOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5j"
    "dGlvbiByZW5kZXJBTUMoKXsKICBjb25zdCBEPVMuZGF0YTsKICBjb25zdCBmaFNlbD1kb2N1bWVudC5nZXRFbGVtZW50QnlJZCgn"
    "YW0tZmgnKTsKICBpZighZmhTZWwub3B0aW9ucy5sZW5ndGgpewogICAgY29uc3QgZmhzPXVuaXEoRC5hbWNfYWxsLCdmaCcpLnNv"
    "cnQoZmhTb3J0KTsKICAgIGZoU2VsLmlubmVySFRNTD1maHMubWFwKGY9PmA8b3B0aW9uICR7Zj09PVMuYW1jLmZoPydzZWxlY3Rl"
    "ZCc6Jyd9PiR7Zn08L29wdGlvbj5gKS5qb2luKCcnKTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhbS1yYW5rZGF0ZScp"
    "LmlubmVySFRNTD1kYXRlT3B0aW9ucyhELmF1bV9kYXRlcyxTLmFtYy5yYW5rRGF0ZSk7CiAgICBmaFNlbC5vbmNoYW5nZT1lPT57"
    "IFMuYW1jLmZoPWUudGFyZ2V0LnZhbHVlOyByZW5kZXJBTUMoKTsgfTsKICAgIHNlZ1dpcmUoJ2FtLXVuaXYnLCB2PT57IFMuYW1j"
    "LnVuaXY9djsgcmVuZGVyQU1DKCk7IH0pOwogICAgc2VnV2lyZSgnYW0tcncnLCB2PT57IFMuYW1jLnJ3PXY7IHJlbmRlckFNQygp"
    "OyB9KTsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhbS1yYW5rZGF0ZScpLm9uY2hhbmdlPWU9PnsgUy5hbWMucmFua0Rh"
    "dGU9K2UudGFyZ2V0LnZhbHVlOyByZW5kZXJBTUNUYWJsZSgpOyB9OwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ2FtLWV4"
    "cCcpLm9uY2xpY2s9ZXhwb3J0QU1DOwogIH0KICBjb25zdCB7ZmgscncsdW5pdn09Uy5hbWM7IGNvbnN0IGJ5UT1hbWNTbGljZShm"
    "aCxydyx1bml2KTsgY29uc3QgZGF0ZXM9RC5hdW1fZGF0ZXM7IGNvbnN0IGxpPWRhdGVzLmxlbmd0aC0xOwogIGNvbnN0IHExPWJ5"
    "UVsnJSBpbiBRMSddPy5bbGldLHEyPWJ5UVsnJSBpbiBRMiddPy5bbGldLHEzPWJ5UVsnJSBpbiBRMyddPy5bbGldLHE0PWJ5UVsn"
    "JSBpbiBRNCddPy5bbGldOwogIGNvbnN0IHNjb3JlPXF1YWxpdHlTY29yZShieVEsbGkpOwogIGRvY3VtZW50LmdldEVsZW1lbnRC"
    "eUlkKCdhbS1jdHgnKS5pbm5lckhUTUw9YCR7dW5pdj09PSdhbGwnPydBbGwgUGVlcnMgKE1GSSknOidFeGFjdCBQZWVycyAoVlIp"
    "J30gwrcgJHtyd308YnI+U25hcHNob3QgJHtmbXQuZGF0ZShkYXRlc1tsaV0pfWA7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQo"
    "J2FtLWtwaXMnKS5pbm5lckhUTUw9YAogICAgPGRpdiBjbGFzcz0ia3BpIGdvb2QiPjxkaXYgY2xhc3M9ImstbGJsIj5RMSBzaGFy"
    "ZTwvZGl2PjxkaXYgY2xhc3M9ImstdmFsIGdvb2QiPiR7Zm10LnBjdChxMSl9PC9kaXY+PGRpdiBjbGFzcz0iay1zdWIiPnRvcCBx"
    "dWFydGlsZTwvZGl2PjwvZGl2PgogICAgPGRpdiBjbGFzcz0ia3BpIGFjY2VudCI+PGRpdiBjbGFzcz0iay1sYmwiPlExK1EyIHNo"
    "YXJlPC9kaXY+PGRpdiBjbGFzcz0iay12YWwgYWNjZW50Ij4ke2ZtdC5wY3QoKHExfHwwKSsocTJ8fDApKX08L2Rpdj48ZGl2IGNs"
    "YXNzPSJrLXN1YiI+dG9wIGhhbGYgb2YgaG91c2UgQVVNPC9kaXY+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGkiPjxkaXYgY2xh"
    "c3M9ImstbGJsIj5RdWFsaXR5IHNjb3JlPC9kaXY+PGRpdiBjbGFzcz0iay12YWwiPiR7Zm10Lm51bShzY29yZSl9PC9kaXY+PGRp"
    "diBjbGFzcz0iay1zdWIiPjEg4oaSIDQsIEFVTS13ZWlnaHRlZDwvZGl2PjwvZGl2PgogICAgPGRpdiBjbGFzcz0ia3BpIHdhcm4i"
    "PjxkaXYgY2xhc3M9ImstbGJsIj5RMytRNCBzaGFyZTwvZGl2PjxkaXYgY2xhc3M9ImstdmFsIHdhcm4iPiR7Zm10LnBjdCgocTN8"
    "fDApKyhxNHx8MCkpfTwvZGl2PjxkaXYgY2xhc3M9Imstc3ViIj5ib3R0b20gaGFsZjwvZGl2PjwvZGl2PmA7CgogIC8vIHN0YWNr"
    "ZWQgKGFscmVhZHkgfjEwMCUpCiAgY29uc3QgY3R4PWRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhbS1zdGFjaycpLmdldENvbnRl"
    "eHQoJzJkJyk7CiAgUy5jaGFydHMuYW1TdGFjaz8uZGVzdHJveSgpOwogIGNvbnN0IGRzPVFPUkRFUi5tYXAocT0+KHsgbGFiZWw6"
    "cS5yZXBsYWNlKCclIGluICcsJycpLCBkYXRhOihieVFbcV18fFtdKS5tYXAoKHYsaSk9Pih7eDpkYXRlc1tpXSx5OnY9PW51bGw/"
    "bnVsbDp2KjEwMH0pKSwKICAgIGJhY2tncm91bmRDb2xvcjooYyk9PmdyYWRpZW50KGMuY2hhcnQuY3R4LGMuY2hhcnQuY2hhcnRB"
    "cmVhLFFDT0xbcV0sMC44NSwwLjU1KSwgYm9yZGVyQ29sb3I6UUNPTFtxXSwgYm9yZGVyV2lkdGg6MC42LAogICAgZmlsbDp0cnVl"
    "LCB0ZW5zaW9uOjAuMzIsIHBvaW50UmFkaXVzOjAsIHN0YWNrOidxJyB9KSk7CiAgUy5jaGFydHMuYW1TdGFjaz1uZXcgQ2hhcnQo"
    "Y3R4LHt0eXBlOidsaW5lJyxkYXRhOntkYXRhc2V0czpkc30sb3B0aW9uczpzdGFja09wdHMoe2xlZnRNYXg6MTAwfSl9KTsKICBk"
    "b2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYW0tc3RhY2stbWV0aG9kJykuaW5uZXJIVE1MPW1ldGhvZE5vdGUoCiAgICBgRXZlcnkg"
    "ZXF1aXR5ICZhbXA7IGh5YnJpZCBzY2hlbWUgaW4gPGI+JHtmaH08L2I+IGlzIHF1YXJ0aWxlLXJhbmtlZCB3aXRoaW4gaXRzICR7"
    "dW5pdj09PSdhbGwnPydNRkkgYWxsLXBlZXInOidWYWx1ZSBSZXNlYXJjaCBleGFjdC1wZWVyJ30gY2F0ZWdvcnkgKCR7cnd9IHdp"
    "bmRvdyksIHRoZW4gYWdncmVnYXRlZCBieSA8Yj5BVU0gd2VpZ2h0PC9iPiBhY3Jvc3MgYWxsIHNsZWV2ZXMuIEJhbmRzIHN1bSB0"
    "byAxMDAlIG9mIHRoZSBob3VzZeKAmXMgY292ZXJlZCBBVU0uYCk7CgogIC8vIEFCU0wgdnMgcGVlcnMKICBjb25zdCBjdHgyPWRv"
    "Y3VtZW50LmdldEVsZW1lbnRCeUlkKCdhbS1yYW5rJykuZ2V0Q29udGV4dCgnMmQnKTsKICBTLmNoYXJ0cy5hbVJhbms/LmRlc3Ry"
    "b3koKTsKICBjb25zdCBVPXVuaXY9PT0nYWxsJz9ELmFtY19hbGw6RC5hbWNfdnI7CiAgY29uc3QgZmhzPXVuaXEoVSwnZmgnKTsK"
    "ICBjb25zdCBkc2V0cz1maHMubWFwKGY9PnsKICAgIGNvbnN0IHExcj1VLmZpbmQocj0+ci5maD09PWYmJnIucnc9PT1ydyYmci5x"
    "dD09PSclIGluIFExJyk7CiAgICBjb25zdCBxMnI9VS5maW5kKHI9PnIuZmg9PT1mJiZyLnJ3PT09cncmJnIucXQ9PT0nJSBpbiBR"
    "MicpOwogICAgaWYoIXExcnx8IXEycikgcmV0dXJuIG51bGw7CiAgICBjb25zdCBpc0E9Zj09PWZoOwogICAgcmV0dXJuIHsgbGFi"
    "ZWw6ZiwgZGF0YTpxMXIudi5tYXAoKHYsaSk9Pih7eDpkYXRlc1tpXSx5Oigodnx8MCkrKHEyci52W2ldfHwwKSkqMTAwfSkpLAog"
    "ICAgICBib3JkZXJDb2xvcjppc0E/Qy5hYnNsOkMucGVlciwgYm9yZGVyV2lkdGg6aXNBPzIuODoxLCBwb2ludFJhZGl1czowLCB0"
    "ZW5zaW9uOjAuMzIsIG9yZGVyOmlzQT8wOjEsCiAgICAgIGJhY2tncm91bmRDb2xvcjondHJhbnNwYXJlbnQnIH07CiAgfSkuZmls"
    "dGVyKEJvb2xlYW4pOwogIFMuY2hhcnRzLmFtUmFuaz1uZXcgQ2hhcnQoY3R4Mix7dHlwZTonbGluZScsZGF0YTp7ZGF0YXNldHM6"
    "ZHNldHN9LG9wdGlvbnM6bGluZU9wdHMoe3N1ZmZpeDonJSd9KX0pOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhbS1yYW5r"
    "LW1ldGhvZCcpLmlubmVySFRNTD1tZXRob2ROb3RlKAogICAgYEhvdXNlLXdpZGUgPGI+UTErUTIgc2hhcmU8L2I+IG92ZXIgdGlt"
    "ZSBmb3IgYWxsIFRvcC0xNSBBTUNzIChhdXRvLXNjYWxlZCkuIDxiPkFkaXR5YSBCaXJsYTwvYj4gaXMgdGhlIGJvbGQgcmVkIGxp"
    "bmU7IHBlZXJzIGFyZSBtdXRlZC4gSGlnaGVyID0gbW9yZSBvZiB0aGUgaG91c2XigJlzIEFVTSBpbiB0b3AtaGFsZiBwZXJmb3Jt"
    "ZXJzLmApOwoKICByZW5kZXJBTUNUYWJsZSgpOwp9CgpmdW5jdGlvbiByZW5kZXJBTUNUYWJsZSgpewogIGNvbnN0IEQ9Uy5kYXRh"
    "OyBjb25zdCB7cncsdW5pdixyYW5rRGF0ZX09Uy5hbWM7IGNvbnN0IGk9cmFua0RhdGU7CiAgY29uc3QgVT11bml2PT09J2FsbCc/"
    "RC5hbWNfYWxsOkQuYW1jX3ZyOyBjb25zdCBmaHM9dW5pcShVLCdmaCcpOwogIGNvbnN0IHJvd3M9ZmhzLm1hcChmaD0+eyBjb25z"
    "dCBieVE9YW1jU2xpY2UoZmgscncsdW5pdik7CiAgICByZXR1cm4geyBmaCwgcTE6YnlRWyclIGluIFExJ10/LltpXXx8MCwgdG9w"
    "OihieVFbJyUgaW4gUTEnXT8uW2ldfHwwKSsoYnlRWyclIGluIFEyJ10/LltpXXx8MCksCiAgICAgIHE0OmJ5UVsnJSBpbiBRNCdd"
    "Py5baV18fDAsIHNjb3JlOnF1YWxpdHlTY29yZShieVEsaSkgfTsgfSkKICAgIC5zb3J0KChhLGIpPT5iLnRvcC1hLnRvcCk7IHJv"
    "d3MuZm9yRWFjaCgocixpZHgpPT5yLnJhbms9aWR4KzEpOwogIGNvbnN0IG49cm93cy5sZW5ndGg7CiAgbGV0IGh0bWw9Jzx0YWJs"
    "ZSBjbGFzcz0icnQiPjx0aGVhZD48dHI+PHRoPiM8L3RoPjx0aD5GdW5kIEhvdXNlPC90aD48dGggY2xhc3M9Im51bSI+UTE8L3Ro"
    "Pjx0aCBjbGFzcz0ibnVtIj5RMStRMjwvdGg+PHRoIGNsYXNzPSJudW0iPlE0PC90aD48dGggY2xhc3M9Im51bSI+UXVhbGl0eSBz"
    "Y29yZTwvdGg+PC90cj48L3RoZWFkPjx0Ym9keT4nOwogIHJvd3MuZm9yRWFjaChyPT57IGNvbnN0IGNscz1yLmZoPT09Uy5hbWMu"
    "Zmg/J2hsJzonJzsgY29uc3QgcGM9ci5yYW5rPT09MT8ncjEnOnIucmFuazw9TWF0aC5jZWlsKG4vNCk/J3VwJzpyLnJhbms+bi1N"
    "YXRoLmNlaWwobi80KT8nZG4nOicnOwogICAgaHRtbCs9YDx0ciBjbGFzcz0iJHtjbHN9Ij48dGQ+PHNwYW4gY2xhc3M9InBpbGwg"
    "JHtwY30iPiR7ci5yYW5rfTwvc3Bhbj48L3RkPjx0ZCBjbGFzcz0ibGJsIj4ke3IuZmh9PC90ZD48dGQgY2xhc3M9Im51bSI+JHtm"
    "bXQucGN0KHIucTEpfTwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10LnBjdChyLnRvcCl9PC90ZD48dGQgY2xhc3M9Im51bSI+JHtm"
    "bXQucGN0KHIucTQpfTwvdGQ+PHRkIGNsYXNzPSJudW0iPiR7Zm10Lm51bShyLnNjb3JlKX08L3RkPjwvdHI+YDsgfSk7CiAgaHRt"
    "bCs9JzwvdGJvZHk+PC90YWJsZT4nOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhbS10YWJsZScpLmlubmVySFRNTD1odG1s"
    "OwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdhbS10YWJsZS1tZXRob2QnKS5pbm5lckhUTUw9bWV0aG9kTm90ZSgKICAgIGBB"
    "bGwgQU1DcyByYW5rZWQgYnkgPGI+UTErUTIgc2hhcmU8L2I+IG9uIDxiPiR7Zm10LmRhdGUoRC5hdW1fZGF0ZXNbaV0pfTwvYj4g"
    "KCR7dW5pdj09PSdhbGwnPydNRkkgYWxsLXBlZXInOidWUiBleGFjdC1wZWVyJ30sICR7cnd9KS4gUXVhbGl0eSBzY29yZSBpcyB0"
    "aGUgQVVNLXdlaWdodGVkIHF1YXJ0aWxlICgx4oCTNCkuIENoYW5nZSB0aGUgYXMtb2YgZGF0ZSB0byByZXBsYXkgYW55IG1vbnRo"
    "LWVuZC5gKTsKfQoKZnVuY3Rpb24gZXhwb3J0QU1DKCl7CiAgY29uc3QgRD1TLmRhdGE7IGNvbnN0IHtmaCxydyx1bml2fT1TLmFt"
    "YzsgY29uc3QgYnlRPWFtY1NsaWNlKGZoLHJ3LHVuaXYpOwogIGxldCBjc3Y9J0RhdGUsUTEsUTIsUTMsUTQsUXVhbGl0eSBzY29y"
    "ZVxuJzsKICBELmF1bV9kYXRlcy5mb3JFYWNoKChkLGkpPT57IGNzdis9W2QsYnlRWyclIGluIFExJ10/LltpXT8/JycsYnlRWycl"
    "IGluIFEyJ10/LltpXT8/JycsYnlRWyclIGluIFEzJ10/LltpXT8/JycsYnlRWyclIGluIFE0J10/LltpXT8/JycscXVhbGl0eVNj"
    "b3JlKGJ5USxpKT8/JyddLmpvaW4oJywnKSsnXG4nOyB9KTsKICBkb3dubG9hZENTVihjc3YsYGFtY18ke2ZofV8ke3VuaXZ9XyR7"
    "cncucmVwbGFjZSgnICcsJycpfS5jc3ZgKTsKfQoKLy/ilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDi"
    "lIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAgU0NIRU1FIFZJRVcg4pSA4pSA4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSACmZ1bmN0aW9uIHJlbmRlclNjaGVtZSgpewogIGNvbnN0IEQ9Uy5kYXRhOwogIGNvbnN0IGNhdFNlbD1kb2N1bWVudC5n"
    "ZXRFbGVtZW50QnlJZCgnc2MtY2F0Jyk7CiAgaWYoIWNhdFNlbC5vcHRpb25zLmxlbmd0aCl7CiAgICBjb25zdCBjYXRzPVsuLi5u"
    "ZXcgU2V0KEQuY29tcG9zaXRlLm1hcChzPT5zLmNhdCkpXS5maWx0ZXIoQm9vbGVhbikuc29ydCgpOwogICAgY2F0U2VsLmlubmVy"
    "SFRNTD1jYXRzLm1hcChjPT5gPG9wdGlvbj4ke2N9PC9vcHRpb24+YCkuam9pbignJyk7CiAgICAvLyBkZWZhdWx0IHRvIGZpcnN0"
    "IGNhdGVnb3J5IGNvbnRhaW5pbmcgYW4gQUJTTCBzY2hlbWUKICAgIGxldCBkZWZDYXQ9Y2F0c1swXSwgZGVmTmFtZT1udWxsOwog"
    "ICAgZm9yKGNvbnN0IGMgb2YgY2F0cyl7IGNvbnN0IGE9RC5jb21wb3NpdGUuZmluZChzPT5zLmNhdD09PWMgJiYgcy5zY2guc3Rh"
    "cnRzV2l0aCgnQWRpdHlhJykpOyBpZihhKXtkZWZDYXQ9YztkZWZOYW1lPWEuc2NoO2JyZWFrO30gfQogICAgUy5zY2hlbWUuY2F0"
    "PWRlZkNhdDsgUy5zY2hlbWUubmFtZT1kZWZOYW1lOyBjYXRTZWwudmFsdWU9ZGVmQ2F0OwogICAgZG9jdW1lbnQuZ2V0RWxlbWVu"
    "dEJ5SWQoJ3NjLXBlZXJkYXRlJykuaW5uZXJIVE1MPWRhdGVPcHRpb25zKEQuc2NvcmVfZGF0ZXMsUy5zY2hlbWUucGVlckRhdGUp"
    "OwogICAgY2F0U2VsLm9uY2hhbmdlPWU9PnsgUy5zY2hlbWUuY2F0PWUudGFyZ2V0LnZhbHVlOyBTLnNjaGVtZS5uYW1lPW51bGw7"
    "IGZpbGxTY2hlbWVOYW1lcygpOyByZW5kZXJTY2hlbWVCb2R5KCk7IH07CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Mt"
    "bmFtZScpLm9uY2hhbmdlPWU9PnsgUy5zY2hlbWUubmFtZT1lLnRhcmdldC52YWx1ZTsgcmVuZGVyU2NoZW1lQm9keSgpOyB9Owog"
    "ICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NjLXBlZXJkYXRlJykub25jaGFuZ2U9ZT0+eyBTLnNjaGVtZS5wZWVyRGF0ZT0r"
    "ZS50YXJnZXQudmFsdWU7IHJlbmRlclNjaGVtZVBlZXJzKCk7IH07CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2MtZXhw"
    "Jykub25jbGljaz1leHBvcnRTY2hlbWU7CiAgICBmaWxsU2NoZW1lTmFtZXMoKTsKICB9CiAgcmVuZGVyU2NoZW1lQm9keSgpOwp9"
    "CmZ1bmN0aW9uIGZpbGxTY2hlbWVOYW1lcygpewogIGNvbnN0IEQ9Uy5kYXRhOyBjb25zdCBjYXQ9Uy5zY2hlbWUuY2F0OwogIGNv"
    "bnN0IG5hbWVzPUQuY29tcG9zaXRlLmZpbHRlcihzPT5zLmNhdD09PWNhdCkubWFwKHM9PnMuc2NoKQogICAgLnNvcnQoKGEsYik9"
    "PmEuc3RhcnRzV2l0aCgnQWRpdHlhJyk/LTE6Yi5zdGFydHNXaXRoKCdBZGl0eWEnKT8xOmEubG9jYWxlQ29tcGFyZShiKSk7CiAg"
    "ZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3NjLW5hbWUnKS5pbm5lckhUTUw9bmFtZXMubWFwKG49PmA8b3B0aW9uICR7bj09PVMu"
    "c2NoZW1lLm5hbWU/J3NlbGVjdGVkJzonJ30+JHtufTwvb3B0aW9uPmApLmpvaW4oJycpOwogIGlmKCFTLnNjaGVtZS5uYW1lKSBT"
    "LnNjaGVtZS5uYW1lPW5hbWVzWzBdOwp9CmZ1bmN0aW9uIGxhc3ROb25OdWxsKGEpeyBpZighYSkgcmV0dXJuIG51bGw7IGZvcihs"
    "ZXQgaT1hLmxlbmd0aC0xO2k+PTA7aS0tKSBpZihhW2ldIT1udWxsKSByZXR1cm4gYVtpXTsgcmV0dXJuIG51bGw7IH0KZnVuY3Rp"
    "b24gcmVuZGVyU2NoZW1lQm9keSgpewogIGNvbnN0IEQ9Uy5kYXRhOyBjb25zdCB7Y2F0LG5hbWV9PVMuc2NoZW1lOyBpZighbmFt"
    "ZSkgcmV0dXJuOwogIGNvbnN0IGNvbXA9RC5jb21wb3NpdGUuZmluZChzPT5zLmNhdD09PWNhdCYmcy5zY2g9PT1uYW1lKTsKICBj"
    "b25zdCB5MT1ELnNjb3JlXzF5LmZpbmQocz0+cy5jYXQ9PT1jYXQmJnMuc2NoPT09bmFtZSk7CiAgY29uc3QgeTM9RC5zY29yZV8z"
    "eS5maW5kKHM9PnMuY2F0PT09Y2F0JiZzLnNjaD09PW5hbWUpOwogIGNvbnN0IGRhdGVzPUQuc2NvcmVfZGF0ZXM7CiAgZG9jdW1l"
    "bnQuZ2V0RWxlbWVudEJ5SWQoJ3NjLXRpdGxlJykudGV4dENvbnRlbnQ9bmFtZTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgn"
    "c2MtY3R4JykuaW5uZXJIVE1MPWA8c3BhbiBjbGFzcz0idGFnIj4ke2NhdH08L3NwYW4+YDsKICBjb25zdCBjTD1sYXN0Tm9uTnVs"
    "bChjb21wPy52KSwgeTFMPWxhc3ROb25OdWxsKHkxPy52KSwgeTNMPWxhc3ROb25OdWxsKHkzPy52KTsKICBjb25zdCBhdmcxMj1h"
    "PT57IGlmKCFhKSByZXR1cm4gbnVsbDsgY29uc3Qgdj1hLnNsaWNlKC0xMikuZmlsdGVyKHg9PnghPW51bGwpOyByZXR1cm4gdi5s"
    "ZW5ndGg/di5yZWR1Y2UoKHMseCk9PnMreCwwKS92Lmxlbmd0aDpudWxsOyB9OwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdz"
    "Yy1rcGlzJykuaW5uZXJIVE1MPWAKICAgIDxkaXYgY2xhc3M9ImtwaSBhY2NlbnQiPjxkaXYgY2xhc3M9ImstbGJsIj5Db21wb3Np"
    "dGUgKGxhdGVzdCk8L2Rpdj48ZGl2IGNsYXNzPSJrLXZhbCBhY2NlbnQiPiR7Zm10Lm51bShjTCl9PC9kaXY+PGRpdiBjbGFzcz0i"
    "ay1zdWIiPi8gNC4wPC9kaXY+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGkiPjxkaXYgY2xhc3M9ImstbGJsIj4xWSBzY29yZTwv"
    "ZGl2PjxkaXYgY2xhc3M9ImstdmFsIj4ke2ZtdC5udW0oeTFMKX08L2Rpdj48ZGl2IGNsYXNzPSJrLXN1YiI+bGF0ZXN0PC9kaXY+"
    "PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGkiPjxkaXYgY2xhc3M9ImstbGJsIj4zWSBzY29yZTwvZGl2PjxkaXYgY2xhc3M9Imst"
    "dmFsIj4ke2ZtdC5udW0oeTNMKX08L2Rpdj48ZGl2IGNsYXNzPSJrLXN1YiI+bGF0ZXN0PC9kaXY+PC9kaXY+CiAgICA8ZGl2IGNs"
    "YXNzPSJrcGkgZ29vZCI+PGRpdiBjbGFzcz0iay1sYmwiPjEyTSBhdmcgY29tcG9zaXRlPC9kaXY+PGRpdiBjbGFzcz0iay12YWwg"
    "Z29vZCI+JHtmbXQubnVtKGF2ZzEyKGNvbXA/LnYpKX08L2Rpdj48ZGl2IGNsYXNzPSJrLXN1YiI+dHJhaWxpbmcgeWVhcjwvZGl2"
    "PjwvZGl2PmA7CgogIGNvbnN0IGN0eD1kb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnc2Mtc2NvcmUnKS5nZXRDb250ZXh0KCcyZCcp"
    "OwogIFMuY2hhcnRzLnNjU2NvcmU/LmRlc3Ryb3koKTsKICBjb25zdCBtaz0ocyxjb2wsdyxkYXNoKT0+KHsgbGFiZWw6JycsIGRh"
    "dGE6KHM/LnZ8fFtdKS5tYXAoKHYsaSk9Pih7eDpkYXRlc1tpXSx5OnZ9KSksCiAgICBib3JkZXJDb2xvcjpjb2wsIGJvcmRlcldp"
    "ZHRoOncsIHRlbnNpb246MC4zLCBwb2ludFJhZGl1czowLCBzcGFuR2FwczpmYWxzZSwgYm9yZGVyRGFzaDpkYXNofHxbXSwKICAg"
    "IGJhY2tncm91bmRDb2xvcjpjb2w9PT1DLmFic2w/KGMpPT5ncmFkaWVudChjLmNoYXJ0LmN0eCxjLmNoYXJ0LmNoYXJ0QXJlYSxD"
    "LmFic2wsMC4xNCwwLjAxKTondHJhbnNwYXJlbnQnLCBmaWxsOmNvbD09PUMuYWJzbCB9KTsKICBjb25zdCBkMT1tayhjb21wLEMu"
    "YWJzbCwyLjgpOyBkMS5sYWJlbD0nQ29tcG9zaXRlJzsKICBjb25zdCBkMj1tayh5MSxDLm5hdnksMS43LFs0LDNdKTsgZDIubGFi"
    "ZWw9JzFZJzsKICBjb25zdCBkMz1tayh5MyxDLnEyLDEuNyxbNCwzXSk7IGQzLmxhYmVsPSczWSc7CiAgUy5jaGFydHMuc2NTY29y"
    "ZT1uZXcgQ2hhcnQoY3R4LHt0eXBlOidsaW5lJyxkYXRhOntkYXRhc2V0czpbZDEsZDIsZDNdfSxvcHRpb25zOmxpbmVPcHRzKHtt"
    "aW46MC44LG1heDo0LjJ9KX0pOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdzYy1zY29yZS1tZXRob2QnKS5pbm5lckhUTUw9"
    "bWV0aG9kTm90ZSgKICAgIGBEYWlseSBxdWFydGlsZSByYW5rIG9mIHRoZSBzY2hlbWUgd2l0aGluIGl0cyA8Yj5WYWx1ZSBSZXNl"
    "YXJjaCBleGFjdC1wZWVyPC9iPiBjYXRlZ29yeSwgbWFwcGVkIHRvIGEgc2NvcmUgKFExPTQg4oCmIFE0PTEpIGFuZCBzYW1wbGVk"
    "IG1vbnRoLWVuZC4gPGI+Q29tcG9zaXRlPC9iPiA9IDAuNMOXMVkgKyAwLjbDlzNZLiBHYXBzIG1lYW4gdGhlIHNjaGVtZSBoYWQg"
    "aW5zdWZmaWNpZW50IGhpc3RvcnkgZm9yIHRoYXQgd2luZG93LmApOwogIHJlbmRlclNjaGVtZVBlZXJzKCk7Cn0KZnVuY3Rpb24g"
    "cmVuZGVyU2NoZW1lUGVlcnMoKXsKICBjb25zdCBEPVMuZGF0YTsgY29uc3Qge2NhdCxuYW1lLHBlZXJEYXRlfT1TLnNjaGVtZTsg"
    "Y29uc3QgaT1wZWVyRGF0ZTsKICBjb25zdCBwZWVycz1ELmNvbXBvc2l0ZS5maWx0ZXIocz0+cy5jYXQ9PT1jYXQpLm1hcChzPT4o"
    "e3NjaDpzLnNjaCwgc2NvcmU6cy52W2ldfSkpLmZpbHRlcihwPT5wLnNjb3JlIT1udWxsKS5zb3J0KChhLGIpPT5iLnNjb3JlLWEu"
    "c2NvcmUpOwogIHBlZXJzLmZvckVhY2goKHAsaWR4KT0+cC5yYW5rPWlkeCsxKTsKICBjb25zdCBhdmc9cGVlcnMubGVuZ3RoP3Bl"
    "ZXJzLnJlZHVjZSgocyxwKT0+cytwLnNjb3JlLDApL3BlZXJzLmxlbmd0aDowOyBjb25zdCBuPXBlZXJzLmxlbmd0aDsKICBsZXQg"
    "aHRtbD0nPHRhYmxlIGNsYXNzPSJydCI+PHRoZWFkPjx0cj48dGg+IzwvdGg+PHRoPlNjaGVtZTwvdGg+PHRoIGNsYXNzPSJudW0i"
    "PkNvbXBvc2l0ZTwvdGg+PHRoIGNsYXNzPSJudW0iPnZzIHBlZXIgYXZnPC90aD48L3RyPjwvdGhlYWQ+PHRib2R5Pic7CiAgcGVl"
    "cnMuZm9yRWFjaChwPT57IGNvbnN0IGNscz1wLnNjaD09PW5hbWU/J2hsJzonJzsgY29uc3QgcGM9cC5yYW5rPT09MT8ncjEnOnAu"
    "cmFuazw9TWF0aC5jZWlsKG4vNCk/J3VwJzpwLnJhbms+bi1NYXRoLmNlaWwobi80KT8nZG4nOicnOwogICAgY29uc3QgZD1wLnNj"
    "b3JlLWF2ZzsgY29uc3QgZGM9ZD49MD8nZGVsdGEtcG9zJzonZGVsdGEtbmVnJzsKICAgIGh0bWwrPWA8dHIgY2xhc3M9IiR7Y2xz"
    "fSI+PHRkPjxzcGFuIGNsYXNzPSJwaWxsICR7cGN9Ij4ke3AucmFua308L3NwYW4+PC90ZD48dGQgY2xhc3M9ImxibCI+JHtwLnNj"
    "aH08L3RkPjx0ZCBjbGFzcz0ibnVtIj4ke2ZtdC5udW0ocC5zY29yZSl9PC90ZD48dGQgY2xhc3M9Im51bSAke2RjfSI+JHsoZD49"
    "MD8nKyc6JycpK2QudG9GaXhlZCgyKX08L3RkPjwvdHI+YDsgfSk7CiAgaHRtbCs9JzwvdGJvZHk+PC90YWJsZT4nOwogIGRvY3Vt"
    "ZW50LmdldEVsZW1lbnRCeUlkKCdzYy1wZWVycycpLmlubmVySFRNTD1odG1sOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdz"
    "Yy1wZWVycy1tZXRob2QnKS5pbm5lckhUTUw9bWV0aG9kTm90ZSgKICAgIGBFdmVyeSBzY2hlbWUgaW4gdGhlIDxiPiR7Y2F0fTwv"
    "Yj4gZXhhY3QtcGVlciBzZXQgcmFua2VkIGJ5IGNvbXBvc2l0ZSBzY29yZSBvbiA8Yj4ke2ZtdC5kYXRlKEQuc2NvcmVfZGF0ZXNb"
    "aV0pfTwvYj4uIOKAnHZzIHBlZXIgYXZn4oCdIGlzIHRoZSBnYXAgdG8gdGhlIGNhdGVnb3J5IG1lYW4gdGhhdCBkYXkuIE1vdmUg"
    "dGhlIGFzLW9mIGRhdGUgdG8gc2VlIHJhbmsgbWlncmF0aW9uLmApOwp9CmZ1bmN0aW9uIGV4cG9ydFNjaGVtZSgpewogIGNvbnN0"
    "IEQ9Uy5kYXRhOyBjb25zdCB7Y2F0LG5hbWV9PVMuc2NoZW1lOwogIGNvbnN0IGNvbXA9RC5jb21wb3NpdGUuZmluZChzPT5zLmNh"
    "dD09PWNhdCYmcy5zY2g9PT1uYW1lKTsKICBjb25zdCB5MT1ELnNjb3JlXzF5LmZpbmQocz0+cy5jYXQ9PT1jYXQmJnMuc2NoPT09"
    "bmFtZSk7CiAgY29uc3QgeTM9RC5zY29yZV8zeS5maW5kKHM9PnMuY2F0PT09Y2F0JiZzLnNjaD09PW5hbWUpOwogIGxldCBjc3Y9"
    "J0RhdGUsQ29tcG9zaXRlLDFZLDNZXG4nOwogIEQuc2NvcmVfZGF0ZXMuZm9yRWFjaCgoZCxpKT0+eyBjc3YrPVtkLGNvbXA/LnZb"
    "aV0/PycnLHkxPy52W2ldPz8nJyx5Mz8udltpXT8/JyddLmpvaW4oJywnKSsnXG4nOyB9KTsKICBkb3dubG9hZENTVihjc3YsYHNj"
    "aGVtZV8ke25hbWUucmVwbGFjZSgvW15cd10rL2csJ18nKS5zbGljZSgwLDUwKX0uY3N2YCk7Cn0KCi8v4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSAIE1BVFJJ"
    "WCBWSUVXIOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKU"
    "gOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgApmdW5jdGlvbiByZW5kZXJNYXRyaXgoKXsKICBjb25zdCBEPVMuZGF0"
    "YTsKICBpZighZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ214LWRhdGUnKS5vcHRpb25zLmxlbmd0aCl7CiAgICBkb2N1bWVudC5n"
    "ZXRFbGVtZW50QnlJZCgnbXgtZGF0ZScpLmlubmVySFRNTD1kYXRlT3B0aW9ucyhELmF1bV9kYXRlcyxTLm1hdHJpeC5kYXRlKTsK"
    "ICAgIHNlZ1dpcmUoJ214LXJ3Jywgdj0+eyBTLm1hdHJpeC5ydz12OyByZW5kZXJNYXRyaXhHcmlkKCk7IH0pOwogICAgc2VnV2ly"
    "ZSgnbXgtbWV0cmljJywgdj0+eyBTLm1hdHJpeC5tZXRyaWM9djsgcmVuZGVyTWF0cml4R3JpZCgpOyB9KTsKICAgIGRvY3VtZW50"
    "LmdldEVsZW1lbnRCeUlkKCdteC1kYXRlJykub25jaGFuZ2U9ZT0+eyBTLm1hdHJpeC5kYXRlPStlLnRhcmdldC52YWx1ZTsgcmVu"
    "ZGVyTWF0cml4R3JpZCgpOyB9OwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ214LWV4cCcpLm9uY2xpY2s9ZXhwb3J0TWF0"
    "cml4OwogIH0KICByZW5kZXJNYXRyaXhHcmlkKCk7Cn0KZnVuY3Rpb24gbWF0cml4VmFsdWUoZmgsYnIscncsaSxtZXRyaWMpewog"
    "IGNvbnN0IGJ5UT1zbGVldmVTbGljZShmaCxicixydyk7IGNvbnN0IHc9YnlRWyclIG9mIEFNQyBBVU0nXT8uW2ldfHwwOwogIGlm"
    "KCF3KSByZXR1cm4gbnVsbDsKICBpZihtZXRyaWM9PT0nc2NvcmUnKSByZXR1cm4gcXVhbGl0eVNjb3JlKGJ5USxpKTsKICBpZiht"
    "ZXRyaWM9PT0ncTEnKSByZXR1cm4gKGJ5UVsnJSBpbiBRMSddPy5baV18fDApL3c7CiAgcmV0dXJuICgoYnlRWyclIGluIFExJ10/"
    "LltpXXx8MCkrKGJ5UVsnJSBpbiBRMiddPy5baV18fDApKS93OyAvLyB0b3BoYWxmCn0KZnVuY3Rpb24gaGVhdENvbG9yKHQpeyAv"
    "LyB0IGluIDAuLjEg4oaSIHJlZOKGkmFtYmVy4oaSZ3JlZW4KICBjb25zdCBzdG9wcz1bWzE5Miw4Miw2Ml0sWzIxNCwxNjIsNjNd"
    "LFszMSw5Miw2MV1dOwogIGNvbnN0IHNlZz10PDAuNT8wOjE7IGNvbnN0IGx0PXQ8MC41P3QvMC41Oih0LTAuNSkvMC41OwogIGNv"
    "bnN0IGE9c3RvcHNbc2VnXSwgYj1zdG9wc1tzZWcrMV07CiAgY29uc3Qgcj1NYXRoLnJvdW5kKGFbMF0rKGJbMF0tYVswXSkqbHQp"
    "LCBnPU1hdGgucm91bmQoYVsxXSsoYlsxXS1hWzFdKSpsdCksIGJsPU1hdGgucm91bmQoYVsyXSsoYlsyXS1hWzJdKSpsdCk7CiAg"
    "cmV0dXJuIGByZ2JhKCR7cn0sJHtnfSwke2JsfSwwLjkyKWA7Cn0KZnVuY3Rpb24gcmVuZGVyTWF0cml4R3JpZCgpewogIGNvbnN0"
    "IEQ9Uy5kYXRhOyBjb25zdCB7cncsbWV0cmljLGRhdGU6aX09Uy5tYXRyaXg7CiAgY29uc3QgYnJzPXVuaXEoRC5zbGVldmUsJ2Jy"
    "Jykuc29ydCgpOwogIGNvbnN0IGZocz11bmlxKEQuc2xlZXZlLCdmaCcpLnNvcnQoZmhTb3J0KTsKICBjb25zdCBpc1Njb3JlPW1l"
    "dHJpYz09PSdzY29yZSc7CiAgY29uc3QgdGl0bGVzPXt0b3BoYWxmOidUb3AtaGFsZiBzaGFyZSAoUTErUTIpIGJ5IGZ1bmQgaG91"
    "c2UgYW5kIHNsZWV2ZScsc2NvcmU6J0FVTS13ZWlnaHRlZCBxdWFydGlsZSBzY29yZSBieSBmdW5kIGhvdXNlIGFuZCBzbGVldmUn"
    "LHExOidRMSBzaGFyZSBieSBmdW5kIGhvdXNlIGFuZCBzbGVldmUnfTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbXgtdGl0"
    "bGUnKS50ZXh0Q29udGVudD10aXRsZXNbbWV0cmljXTsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnbXgtc3ViJykudGV4dENv"
    "bnRlbnQ9YCR7cnd9IMK3ICR7Zm10LmRhdGUoRC5hdW1fZGF0ZXNbaV0pfWA7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ214"
    "LWN0eCcpLmlubmVySFRNTD1gJHtyd30gcm9sbGluZzxicj4ke2ZtdC5kYXRlKEQuYXVtX2RhdGVzW2ldKX1gOwoKICAvLyBnYXRo"
    "ZXIgdmFsdWVzIGZvciBjb2xvciBzY2FsaW5nCiAgY29uc3QgY2VsbHM9e307IGxldCBtbj1JbmZpbml0eSxteD0tSW5maW5pdHk7"
    "CiAgZmhzLmZvckVhY2goZmg9PmJycy5mb3JFYWNoKGJyPT57IGNvbnN0IHY9bWF0cml4VmFsdWUoZmgsYnIscncsaSxtZXRyaWMp"
    "OyBjZWxsc1tmaCsnfCcrYnJdPXY7IGlmKHYhPW51bGwpe21uPU1hdGgubWluKG1uLHYpO214PU1hdGgubWF4KG14LHYpO30gfSkp"
    "OwogIGNvbnN0IG5vcm09dj0+ICh2PT1udWxsKT9udWxsIDogaXNTY29yZSA/ICh2LTEpLzMgOiAobXg+bW4/KHYtbW4pLyhteC1t"
    "bik6MC41KTsKICBjb25zdCBkaXNwPXY9PiB2PT1udWxsPyfCtycgOiBpc1Njb3JlID8gdi50b0ZpeGVkKDIpIDogKHYqMTAwKS50"
    "b0ZpeGVkKDApKyclJzsKCiAgbGV0IGh0bWw9Jzx0YWJsZSBjbGFzcz0icnQiPjx0aGVhZD48dHI+PHRoPkZ1bmQgSG91c2U8L3Ro"
    "Pic7CiAgYnJzLmZvckVhY2goYj0+aHRtbCs9YDx0aCBjbGFzcz0ibnVtIj4ke2J9PC90aD5gKTsgaHRtbCs9JzwvdHI+PC90aGVh"
    "ZD48dGJvZHk+JzsKICBmaHMuZm9yRWFjaChmaD0+ewogICAgY29uc3QgY2xzPWZoPT09J0FkaXR5YSc/J2hsJzonJzsKICAgIGh0"
    "bWwrPWA8dHIgY2xhc3M9IiR7Y2xzfSI+PHRkIGNsYXNzPSJsYmwiPiR7Zmh9PC90ZD5gOwogICAgYnJzLmZvckVhY2goYnI9Pnsg"
    "Y29uc3Qgdj1jZWxsc1tmaCsnfCcrYnJdOyBjb25zdCB0PW5vcm0odik7CiAgICAgIGNvbnN0IGJnPXY9PW51bGw/J3RyYW5zcGFy"
    "ZW50JzpoZWF0Q29sb3IodCk7IGNvbnN0IGNvbD12PT1udWxsP0MuaW5rNDoodD4wLjMyJiZ0PDAuNz8nIzNiMmYxNSc6JyNmZmYn"
    "KTsKICAgICAgaHRtbCs9YDx0ZCBjbGFzcz0ibnVtIiBzdHlsZT0iYmFja2dyb3VuZDoke2JnfTtjb2xvcjoke2NvbH07Zm9udC13"
    "ZWlnaHQ6NjAwIj4ke2Rpc3Aodil9PC90ZD5gOyB9KTsKICAgIGh0bWwrPSc8L3RyPic7CiAgfSk7CiAgaHRtbCs9JzwvdGJvZHk+"
    "PC90YWJsZT4nOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdteC1ncmlkJykuaW5uZXJIVE1MPWh0bWw7CiAgZG9jdW1lbnQu"
    "Z2V0RWxlbWVudEJ5SWQoJ214LW1ldGhvZCcpLmlubmVySFRNTD1tZXRob2ROb3RlKAogICAgaXNTY29yZQogICAgPyBgRWFjaCBj"
    "ZWxsIGlzIHRoZSA8Yj5BVU0td2VpZ2h0ZWQgcXVhcnRpbGUgc2NvcmU8L2I+IChRMT00IOKApiBRND0xKSBmb3IgdGhhdCBob3Vz"
    "ZeKAmXMgc2xlZXZlIG9uIDxiPiR7Zm10LmRhdGUoRC5hdW1fZGF0ZXNbaV0pfTwvYj4uIEdyZWVuZXIgPSBtb3JlIEFVTSB3aXRo"
    "IHRvcCBwZXJmb3JtZXJzOyByZWQgPSBsYWdnYXJkcy4gRW1wdHkgY2VsbHMgbWVhbiB0aGUgaG91c2UgaGFzIG5vIEFVTSBpbiB0"
    "aGF0IHNsZWV2ZS5gCiAgICA6IGBFYWNoIGNlbGwgaXMgdGhlIDxiPiR7bWV0cmljPT09J3ExJz8nUTEnOidRMStRMid9IHNoYXJl"
    "IHdpdGhpbiB0aGUgc2xlZXZlPC9iPiBmb3IgdGhhdCBob3VzZSBvbiA8Yj4ke2ZtdC5kYXRlKEQuYXVtX2RhdGVzW2ldKX08L2I+"
    "ICgke3J3fSkuIENvbG91ciBzY2FsZXMgYWNyb3NzIHRoZSB2aXNpYmxlIHJhbmdlIOKAlCBncmVlbiA9IHN0cm9uZywgcmVkID0g"
    "d2Vhay4gQWRpdHlhIEJpcmxhIHJvdyBoaWdobGlnaHRlZC5gKTsKfQpmdW5jdGlvbiBleHBvcnRNYXRyaXgoKXsKICBjb25zdCBE"
    "PVMuZGF0YTsgY29uc3Qge3J3LG1ldHJpYyxkYXRlOml9PVMubWF0cml4OwogIGNvbnN0IGJycz11bmlxKEQuc2xlZXZlLCdicicp"
    "LnNvcnQoKTsgY29uc3QgZmhzPXVuaXEoRC5zbGVldmUsJ2ZoJykuc29ydChmaFNvcnQpOwogIGxldCBjc3Y9J0Z1bmQgSG91c2Us"
    "JyticnMuam9pbignLCcpKydcbic7CiAgZmhzLmZvckVhY2goZmg9PnsgY3N2Kz1maCsnLCcrYnJzLm1hcChicj0+eyBjb25zdCB2"
    "PW1hdHJpeFZhbHVlKGZoLGJyLHJ3LGksbWV0cmljKTsgcmV0dXJuIHY9PW51bGw/Jyc6djsgfSkuam9pbignLCcpKydcbic7IH0p"
    "OwogIGRvd25sb2FkQ1NWKGNzdixgbWF0cml4XyR7bWV0cmljfV8ke3J3LnJlcGxhY2UoJyAnLCcnKX1fJHtELmF1bV9kYXRlc1tp"
    "XX0uY3N2YCk7Cn0KCi8v4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSA"
    "4pSA4pSA4pSA4pSA4pSA4pSA4pSA4pSAIFBFRVIgTUFQUElORyDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDi"
    "lIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIDilIAKZnVuY3Rpb24gcmVu"
    "ZGVyUGVlcigpewogIGNvbnN0IEQ9Uy5kYXRhOwogIGNvbnN0IGRyYXc9Zj0+ewogICAgY29uc3Qgcm93cz1ELnBlZXJfbWFwLmZp"
    "bHRlcihyPT57IGlmKCFmKSByZXR1cm4gdHJ1ZTsgY29uc3Qgcz1mLnRvTG93ZXJDYXNlKCk7CiAgICAgIHJldHVybiAoci5zY2hl"
    "bWV8fCcnKS50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHMpfHwoci5jYXR8fCcnKS50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKHMpfHwo"
    "ci5tZmlfbmFtZXx8JycpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocyk7IH0pOwogICAgZG9jdW1lbnQucXVlcnlTZWxlY3Rvcign"
    "I3BtLXRhYmxlIHRib2R5JykuaW5uZXJIVE1MPXJvd3MubWFwKHI9PnsKICAgICAgY29uc3QgY2xzPShyLnNjaGVtZXx8JycpLnN0"
    "YXJ0c1dpdGgoJ0FkaXR5YScpPydobCc6Jyc7CiAgICAgIHJldHVybiBgPHRyIGNsYXNzPSIke2Nsc30iPjx0ZCBjbGFzcz0ibGJs"
    "Ij4ke3Iuc2NoZW1lfHwnJ308L3RkPjx0ZD4ke3IuYW1maXx8Jyd9PC90ZD48dGQ+JHtyLmNhdHx8Jyd9PC90ZD48dGQ+JHtyLm1m"
    "aV9uYW1lfHwnJ308L3RkPjwvdHI+YDsKICAgIH0pLmpvaW4oJycpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3BtLWNv"
    "dW50JykudGV4dENvbnRlbnQ9YCR7cm93cy5sZW5ndGh9IG9mICR7RC5wZWVyX21hcC5sZW5ndGh9IHNjaGVtZXNgOwogIH07CiAg"
    "ZHJhdygnJyk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3BtLXNlYXJjaCcpLm9uaW5wdXQ9ZT0+ZHJhdyhlLnRhcmdldC52"
    "YWx1ZSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3BtLWV4cCcpLm9uY2xpY2s9KCk9PnsKICAgIGxldCBjc3Y9J1NjaGVt"
    "ZSxBTUZJLENhdGVnb3J5LE5hbWUgaW4gTUZJXG4nOwogICAgRC5wZWVyX21hcC5mb3JFYWNoKHI9PnsgY3N2Kz1bci5zY2hlbWUs"
    "ci5hbWZpLHIuY2F0LHIubWZpX25hbWVdLm1hcCh2PT5gIiR7KHZ8fCcnKS50b1N0cmluZygpLnJlcGxhY2UoLyIvZywnIiInKX0i"
    "YCkuam9pbignLCcpKydcbic7IH0pOwogICAgZG93bmxvYWRDU1YoY3N2LCdwZWVyX21hcHBpbmcuY3N2Jyk7CiAgfTsKICBkb2N1"
    "bWVudC5nZXRFbGVtZW50QnlJZCgncG0tbWV0aG9kJykuaW5uZXJIVE1MPW1ldGhvZE5vdGUoCiAgICBgVGhlIDxiPmV4YWN0LXBl"
    "ZXIgdW5pdmVyc2U8L2I+IGZyb20gVmFsdWUgUmVzZWFyY2ggdXNlZCBmb3Igc2NoZW1lIHNjb3Jpbmcg4oCUIGVhY2ggc2NoZW1l"
    "IG1hcHBlZCB0byBpdHMgVlIgY2F0ZWdvcnkgYW5kIG1hdGNoZWQgdG8gdGhlIE1GSSBOQVYgbmFtZS4gQWRpdHlhIEJpcmxhIHNj"
    "aGVtZXMgaGlnaGxpZ2h0ZWQuYCk7Cn0KCi8vIOKUgOKUgCBraWNrIGV2ZXJ5dGhpbmcgb2ZmIChhZnRlciBhbGwgZGVjbGFyYXRp"
    "b25zIGFyZSBpbml0aWFsaXNlZCkg4pSA4pSACmxvYWREYXRhKCk7Cjwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4K"
)


if __name__ == "__main__":
    main()
