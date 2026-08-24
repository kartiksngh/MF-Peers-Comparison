# %% [markdown]
# # Peer Performance Monitor — canonical source (loaders + data cleaning)
# Single `# %%`-cell source -> exported to the audit notebook and the runnable .py.
# This file is built up section by section per BUILD_SPEC.md. Runs offline, no Claude.

# %% imports + config
import os
import re
import sys
import json
import numpy as np
import pandas as pd
from pathlib import Path
from joblib import Parallel, delayed

try:  # make console output UTF-8 safe on Windows (cp1252 default chokes on ×, ->, etc.)
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

pd.set_option("future.no_silent_downcasting", True)  # quiet replace()-downcast deprecation

ENGINE = "calamine"  # fast Excel reader; falls back to openpyxl if missing (set below)

DATA       = Path("Data")
SCHEME_DIR = DATA / "Scheme NAV and AUM"
MAP_DIR    = DATA / "Mapping"
BENCH_DIR  = DATA / "Benchmark NAV"

# ── Data-cleaning knobs (BUILD_SPEC §2) ──────────────────────────────
REPEAT_FRAC  = 0.90   # date is "non-trading" if >= this frac of POPULATED series repeat prior day's value exactly
MAX_FILL_GAP = 5      # forward-fill an internal gap ONLY if it spans < this many trading days AND a real value resumes


# %% readers
def read_mfi_nav(path) -> pd.DataFrame:
    """Read an MFI-format NAV sheet (scheme OR benchmark) -> dates × names, '--'/blank -> NaN.

    Robust to the differing metadata-row counts: scheme files have 3 metadata rows
    (Scheme/Index Code, AMFI Code, Fund Name) after the 'Date|names' header; the
    benchmark-from-MFI file has only 1 (Scheme/Index Code). We don't slice by count —
    we coerce the index to dates and keep only real-date rows, so metadata AND footer
    rows (which become NaT) drop automatically.
    """
    df = pd.read_excel(path, skiprows=5, header=0, index_col=0, engine=ENGINE)
    df.index = pd.to_datetime(df.index, errors="coerce")
    df = df.loc[df.index.notna()]
    df.index.name = "Date"
    df = df.apply(pd.to_numeric, errors="coerce")
    return df.replace(0, np.nan)


def read_aum(path) -> pd.DataFrame:
    """AUM xlsx -> long [Fund Name, Scheme Name, Date, Value]."""
    df = pd.read_excel(path, skiprows=5, header=0, engine=ENGINE)
    df = df[["Fund Name", "Scheme Name", "Date", "Value"]].copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Value"] = pd.to_numeric(df["Value"], errors="coerce")
    return df.dropna(subset=["Date"])


def read_simple_nav(path, sheet=0) -> pd.DataFrame:
    """Read the stale NASDAQ/S&P fallback (Dates + ticker columns) -> dates × tickers."""
    df = pd.read_excel(path, sheet_name=sheet, engine=ENGINE)
    date_col = next(c for c in df.columns if str(c).lower().startswith("date"))
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=[date_col]).set_index(date_col)
    df.index.name = "Date"
    return df.apply(pd.to_numeric, errors="coerce")


def _glob(folder, *needles):
    return sorted(folder / f for f in os.listdir(folder)
                  if f.endswith(".xlsx") and not f.startswith("~")
                  and all(n.lower() in f.lower() for n in needles))


def _coalesce_dup_cols(df: pd.DataFrame) -> pd.DataFrame:
    """When the same scheme/series appears in >1 file, concat creates duplicate columns.
    Coalesce them: first non-null value per (date, name). Keeps a single column per name."""
    if not df.columns.duplicated().any():
        return df
    n_dup = int(df.columns.duplicated().sum())
    df = df.T.groupby(level=0).first().T
    df.attrs["coalesced_dupes"] = n_dup
    return df


def load_scheme_nav(folder=SCHEME_DIR) -> pd.DataFrame:
    """Load + combine all MFI scheme NAV files; coalesce duplicate schemes; sort by date."""
    paths = _glob(folder, "nav")
    nav = pd.concat([read_mfi_nav(p) for p in paths], axis=1).sort_index()
    nav = _coalesce_dup_cols(nav)
    return nav


def _norm_code(c):
    """Normalize an AMFI code (handle float-read '120377.0', whitespace)."""
    s = str(c).strip()
    return s[:-2] if s.endswith(".0") else s


def scheme_code_map(folder=SCHEME_DIR) -> dict:
    """Map AMFI Code -> NAV column name, read from the 'AMFI Code' metadata row of each
    MFI NAV file. The robust join key across vendors (codes have no spelling variants)."""
    code2name = {}
    for p in _glob(folder, "nav"):
        raw = pd.read_excel(p, skiprows=5, header=None, engine=ENGINE)
        names = raw.iloc[0, 1:].tolist()
        rows = raw.index[raw[0].astype(str).str.strip() == "AMFI Code"]
        if len(rows):
            codes = raw.iloc[rows[0], 1:].tolist()
            for c, n in zip(codes, names):
                if pd.notna(c) and pd.notna(n):
                    code2name.setdefault(_norm_code(c), n)
    return code2name


# %% Value Research peer map + benchmark NAV wiring  (BUILD_SPEC §4a, §1)
VR_MAP_FILE = "Mapping VR to MFI names.xlsx"
VR_PEER_SHEET = "Ret. Compr.(Equity) - Dir"
AVOID_EXACT_CATS = ["Domestic + International", "Multi Index FoF",
                    "Other Competitor Thematic/Contra Funds"]
# Categories with duplicate peer-sets, dropped from the VR %AUM roll-up (nb 1686-1689)
REPEATED_PEER_CATS = ["Bal Bhavishya", "Retirement Fund 40"]

# ── The BROAD-MARKET benchmark (Kyser 2026-08-12) ────────────────────────────────────
# Every category already ships ITS OWN benchmark, so a scheme could be read against its
# category index but never against the market. This adds NIFTY 500 TRI as an extra series
# on the benchmark side, carried as a PSEUDO-CATEGORY so it rides the existing
# `returns["bench"]` block with no new plumbing, no new axis and no packer change.
# It costs nothing on the wire: 7 real categories already use this exact column, so the
# series is in `bcols` either way — only the extra key mapping is new.
# The key is deliberately un-category-like so it can never collide with a real VR category.
MARKET_KEY   = "__MKT__"
MARKET_BENCH = "NIFTY 500 Total Return Index"      # must match the benchmark NAV column exactly


def load_vr_mapping(map_path=None, sheet=VR_PEER_SHEET):
    """Return (category_map_exact, exact_bench, dropped_no_mfi).
    - category_map_exact: indexed by 'Scheme Name From MFI', cols incl Scheme, AMFI Code,
      Category, fund house. Benchmark rows removed.
    - exact_bench: Category -> benchmark DISPLAY name (from AMFI Code=='bench' rows).
    """
    map_path = map_path or (MAP_DIR / VR_MAP_FILE)
    cme = pd.read_excel(map_path, sheet_name=sheet, engine=ENGINE)
    cme.index = cme.pop("Scheme")
    is_bench = cme["AMFI Code"].astype(str).str.strip() == "bench"
    dropped = cme.index[cme["Scheme Name From MFI"].isna() & ~is_bench].tolist()
    cme = cme.drop(index=dropped)
    cme["fund house"] = house_series(cme.index).values

    is_bench = cme["AMFI Code"].astype(str).str.strip() == "bench"
    bench_rows = cme[is_bench]
    exact_bench = bench_rows.reset_index()[["Category", "Scheme"]].copy()
    exact_bench.columns = ["Category", "Benchmark"]
    # Strip whitespace — bench-row names occasionally carry a trailing space that
    # otherwise fails the join to the (stripped) benchmark-NAV column names.
    exact_bench["Benchmark"] = exact_bench["Benchmark"].astype(str).str.strip()
    exact_bench["Category"] = exact_bench["Category"].astype(str).str.strip()
    exact_bench = exact_bench.set_index("Category")["Benchmark"]

    cme = cme[~is_bench].reset_index().set_index("Scheme Name From MFI")
    return cme, exact_bench, dropped


def _norm_name(s):
    """Normalize a scheme name for matching: collapse whitespace, strip, lowercase."""
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def align_vr_to_nav(cme: pd.DataFrame, nav_columns, code2name: dict | None = None):
    """Remap cme's index (VR 'Scheme Name From MFI') to ACTUAL NAV column names, robustly:
       1) by AMFI Code (primary — no spelling variants), 2) by whitespace/case-normalized
       name (fallback). Recovers peers dropped by cross-vendor name-format mismatches.
       Returns (cme_aligned, unmatched_names)."""
    code2name = code2name or {}
    norm_to_actual = {}
    for col in nav_columns:
        norm_to_actual.setdefault(_norm_name(col), col)
    nav_set = set(nav_columns)
    codes = (cme["AMFI Code"].map(_norm_code) if "AMFI Code" in cme.columns
             else pd.Series([None] * len(cme), index=cme.index))
    new_idx, unmatched, recovered = [], [], 0
    for name, code in zip(cme.index, codes):
        actual = code2name.get(code) or norm_to_actual.get(_norm_name(name))
        if actual is not None:
            if name not in nav_set and actual in nav_set:
                recovered += 1
            new_idx.append(actual)
        else:
            new_idx.append(name)
            if pd.notna(name):
                unmatched.append(name)
    cme = cme.copy()
    cme.index = new_idx
    cme.attrs["recovered_by_match"] = recovered
    return cme, unmatched


def load_bench_nav(map_path=None, bench_dir=BENCH_DIR):
    """Assemble benchmark NAV (dates × DISPLAY names) from the two new files, using
    'Benchmarks and sources' [Benchmarks, MFI Identifier, File] to route each benchmark
    to its NAV column (MFI-format file, or the stale NASDAQ/S&P 'Missing' file).
    Returns (bench_nav, resolution_report)."""
    map_path = map_path or (MAP_DIR / VR_MAP_FILE)
    bs = pd.read_excel(map_path, sheet_name="Benchmarks and sources", engine=ENGINE)
    mfi = read_mfi_nav(_glob(bench_dir, "from mfi")[0])
    miss = read_simple_nav(_glob(bench_dir, "missing")[0])

    def _find(src, ident):
        ident = str(ident).strip()
        for c in src.columns:
            if str(c).strip() == ident:
                return c
        return None

    cols, report = {}, []
    for _, r in bs.iterrows():
        disp = str(r["Benchmarks"]).strip()
        ident, fil = r["MFI Identifier"], str(r["File"])
        if pd.isna(ident):
            report.append((disp, "NO SOURCE (composite, no NAV)")); continue
        src, srcname = (miss, "missing") if "missing" in fil.lower() else (mfi, "mfi")
        match = _find(src, ident)
        if match is None:
            report.append((disp, f"identifier '{str(ident).strip()}' not found in {srcname} file")); continue
        s = src[match]
        last = s.dropna().index.max()
        report.append((disp, f"OK [{srcname}] last={last.date() if pd.notna(last) else 'EMPTY'}"))
        cols[disp] = s
    bench_nav = pd.DataFrame(cols).sort_index()
    return bench_nav, report


# %% category map (MFI / all-peer)  (BUILD_SPEC §4a)
BREAKDOWN_RULES = {
    "Bottoms Up": ["Large & Mid Cap Fund", "Small cap Fund", "Mid Cap Fund", "Multi Cap Fund"],
    "Top Down": ["Large Cap Fund", "Flexi Cap Fund", "Focused Fund", "ELSS"],
    "Asset Allocation": ["Aggressive Hybrid Fund", "Balanced Hybrid Fund",
                         "Dynamic Asset Allocation or Balanced Advantage",
                         "Conservative Hybrid Fund", "Multi Asset Allocation"],
    "Arbitrage+": ["Equity Savings", "Arbitrage Fund"],
}
MULTI_ASSET_DROP = ["WhiteOak", "Edelweiss", "Mahindra"]  # debt-like taxation -> exclude
TOP15 = ["Axis", "Franklin", "Kotak", "HDFC", "Aditya", "DSP", "UTI", "Invesco",
         "Canara", "SBI", "Mirae", "Nippon", "Tata", "HSBC", "ICICI"]

# ── FUND-HOUSE IDENTITY (2026-08-13) ───────────────────────────────────────────────────────
# A scheme's "fund house" is its name's FIRST WORD. That is a deliberate, cheap convention and
# it is right for almost every AMC, but the raw first word splits three real houses in two, so
# any view that enumerates ALL houses (rather than the 15 hard-coded below) shows the same firm
# twice with two different denominators:
#     BARODA    (3 upper-cased FoFs)      vs Baroda   (28 schemes)   -> Baroda BNP Paribas
#     quant     (2 lower-cased schemes)   vs Quant    (23 schemes)   -> Quant
#     Templeton (Templeton India Value)   vs Franklin (23 schemes)   -> Franklin Templeton
# Folding them is a CORRECTION, not a cosmetic rename: percent_aum_share() divides each scheme
# by its house's total AUM, so while "Templeton India Value Fund" sat in its own one-fund house
# it was 100% of that house and contributed nothing to Franklin's quartile shares. Every place
# that derives a house now goes through house_of()/house_series() so the fold is applied once
# and cannot drift between the map, the AUM shares and the payload.
HOUSE_CANON = {"BARODA": "Baroda", "quant": "Quant", "Templeton": "Franklin"}


def house_of(name) -> str:
    """Canonical fund house for one scheme name. Splits on ANY whitespace (so a stray leading
    or doubled space cannot yield an empty house) and folds the known aliases."""
    parts = str(name).split()
    h = parts[0] if parts else ""
    return HOUSE_CANON.get(h, h)


def house_series(index) -> pd.Series:
    """house_of() over a pandas Index/Series, returned as a Series aligned to that index."""
    s = pd.Series(list(index), index=index).astype(str).str.split().str[0].fillna("")
    return s.replace(HOUSE_CANON)


def build_category_map(raw_map: pd.DataFrame, code2name: dict | None = None):
    """Clean MFI Map -> (cmap[Scheme Sub Nature, Category, fund house, Breakdown], ordered cats).
    Keeps categories with >=4 fund houses. Records drops in cmap.attrs['exclusions'].

    If code2name (AMFI Code -> NAV column name) is given, the map is JOINED BY AMFI CODE and
    indexed by the ACTUAL NAV name — robust to MFI-vs-Map name-format mismatches (the
    name join loses ~30 schemes; the code join keeps all). Falls back to Scheme Name."""
    excl = []
    cm = raw_map.copy()
    if code2name and "AMFI Code" in cm.columns:
        cm["__nav"] = cm["AMFI Code"].map(_norm_code).map(code2name)
        cm = cm[cm["__nav"].notna()].copy()
        cm.index = cm.pop("__nav")
    else:
        cm.index = cm["Scheme Name"]
    cm = cm[["Scheme Sub Nature", "Category"]].copy()
    cm["Category"] = cm["Category"].fillna("Unknown").astype(str)
    cm.loc[cm["Category"].isin(["Contra Fund", "Value Fund"]), "Category"] = "Value/Contra"

    categories = sorted(set(cm["Category"]))
    thematic = set(cm.loc[cm["Scheme Sub Nature"] == "Thematic", "Category"])
    sectoral = set(cm.loc[cm["Scheme Sub Nature"] == "Sectoral", "Category"])
    thematic -= sectoral
    categories = ([c for c in categories if c not in thematic | sectoral]
                  + sorted(sectoral) + sorted(thematic))

    cm["fund house"] = house_series(cm.index).values
    assigned = set()
    for bd, subs in BREAKDOWN_RULES.items():
        cm.loc[cm["Scheme Sub Nature"].isin(subs), "Breakdown"] = bd
        assigned |= set(subs)
    cm.loc[~cm["Scheme Sub Nature"].isin(assigned), "Breakdown"] = "Thematic"

    drop = cm.loc[(cm["Scheme Sub Nature"] == "Multi Asset Allocation")
                  & (cm["fund house"].isin(MULTI_ASSET_DROP))].index
    for s in drop:
        excl.append({"item": s, "type": "scheme", "reason": "Multi-Asset with debt-like taxation (MULTI_ASSET_DROP)"})
    cm = cm.drop(drop)

    counts = cm.reset_index().groupby("Category")["fund house"].count()
    valid = set(counts[counts >= 4].index)
    for c in categories:
        if c not in valid and c in set(cm["Category"]):
            excl.append({"item": c, "type": "category(all-peer)", "reason": f"<4 fund houses ({int(counts.get(c,0))})"})
    categories = [c for c in categories if c in valid]
    cm.attrs["exclusions"] = excl
    return cm, categories


# %% AUM share (daily, ffilled from monthly)  (BUILD_SPEC §4g)
def load_aum(folder=SCHEME_DIR) -> pd.DataFrame:
    paths = _glob(folder, "aum")
    aum_long = pd.concat([read_aum(p) for p in paths]).drop_duplicates()
    aum = aum_long.pivot_table(values="Value", index="Scheme Name", columns="Date")
    aum = aum.drop([i for i in aum.index if "Adjusted" in i or "Segregated" in i])
    return aum


def percent_aum_share(aum: pd.DataFrame, nav_index: pd.Index):
    """Daily ffilled scheme AUM (dates×schemes) and each scheme's share of its fund house's
    AUM, reindexed to the (daily) NAV calendar. Returns (aum_daily, pct_aum) both dates×schemes."""
    aum1 = aum.copy().T.sort_index().replace("--", np.nan).ffill().T
    aum1.insert(0, "fund house", house_series(aum1.index).values)
    fund_aum = aum1.groupby("fund house").sum(numeric_only=True)
    aum_daily = aum1.iloc[:, 1:].T  # dates × schemes (monthly)

    pct = pd.DataFrame()
    for fund in fund_aum.index:
        sub = aum1.loc[aum1["fund house"] == fund].iloc[:, 1:]
        share = (sub.div(fund_aum.loc[fund])).T  # dates × schemes
        daily = pd.DataFrame(index=nav_index)
        daily = daily.merge(share, left_index=True, right_index=True, how="outer").ffill()
        daily = daily.loc[nav_index]
        pct = pct.merge(daily, left_index=True, right_index=True, how="outer")
    # reindex monthly aum onto daily nav calendar too (ffilled)
    aum_daily_nav = aum_daily.reindex(aum_daily.index.union(nav_index)).sort_index().ffill().loc[nav_index]
    return aum_daily_nav, pct.loc[nav_index]


# %% calendar alignment + returns  (BUILD_SPEC §3, §4b)
def align_benchmarks(bench_nav: pd.DataFrame, master_index: pd.Index):
    """Put benchmarks on the scheme trading calendar and CARRY the last value forward (so the
    bonus is computable at every scheme date using the latest available benchmark level — fine
    for periodic benchmarks like the monthly CRISIL Hybrid indices). Genuine staleness is
    handled downstream by capping a category only when its benchmark's `last_real` is far
    behind (see exact_peer_scoring/STALE_BENCH_DAYS). Returns (bench_on_cal, last_real dict)."""
    out, last_real = {}, {}
    for c in bench_nav.columns:
        s = bench_nav[c]
        last_real[c] = s.dropna().index.max()
        out[c] = (s.reindex(s.index.union(master_index)).sort_index().ffill().reindex(master_index))
    return pd.DataFrame(out), last_real


def rolling_returns(nav: pd.DataFrame, win_1y=250, win_3y=750):
    """Shift-based rolling returns (ALL-PEER path, nb cells 25/26): 1Y cumulative, 3Y CAGR."""
    r1 = nav / nav.shift(win_1y) - 1
    r3 = (nav / nav.shift(win_3y)) ** (1 / 3) - 1
    return r1, r3


def _calendar_lookback_positions(index: pd.DatetimeIndex, years: int):
    """For each date t, position of the last trading day <= the SAME calendar date `years`
    before t (nb cell 49: data.loc[:t_minus].index[-1]). -1 where target precedes data."""
    def shift_y(ts):
        try:
            return ts.replace(year=ts.year - years)
        except ValueError:  # Feb 29 -> day-1 (matches nb 1377-1378)
            return ts.replace(year=ts.year - years, day=ts.day - 1)
    targets = pd.DatetimeIndex([shift_y(t) for t in index])
    return index.get_indexer(targets, method="ffill")


def calendar_returns(nav: pd.DataFrame):
    """EXACT-CALENDAR point-to-point returns (EXACT-PEER scoring, nb cell 49 'Final code'):
    1Y cumulative, 3Y CAGR, lookback = last trading day on/before the same calendar date
    1y/3y earlier. Quartile ranks & alpha sign match a cumulative basis; values match the
    notebook's annualized 3Y."""
    idx = nav.index
    pos1 = _calendar_lookback_positions(idx, 1)
    pos3 = _calendar_lookback_positions(idx, 3)
    vals = nav.values.astype(float)
    r1 = np.full(vals.shape, np.nan)
    r3 = np.full(vals.shape, np.nan)
    ok1, ok3 = pos1 >= 0, pos3 >= 0
    r1[ok1] = vals[ok1] / vals[pos1[ok1]] - 1.0
    r3[ok3] = (vals[ok3] / vals[pos3[ok3]]) ** (1.0 / 3.0) - 1.0
    return (pd.DataFrame(r1, index=idx, columns=nav.columns),
            pd.DataFrame(r3, index=idx, columns=nav.columns))


# ── Extra rolling windows (1/3/6/9 months + 2Y/5Y) — RAW quartiles, no composite/reward ──
# KV (2026-06-24): in ADDITION to the 1Y/3Y windows, evaluate every scheme on 1M/3M/6M/9M
# rolling windows so the whole deck (sleeve/AMC/league/scheme) can be viewed on any window.
# KV (2026-07-14): also 2Y and 5Y. Returns here are CUMULATIVE point-to-point
# (NAV_t/NAV_{t_m}-1), NOT annualized — a sub-year CAGR would be misleading, and for the
# multi-year windows the DISPLAYED object is the quartile RANK, which is identical either
# way (annualizing by ^(12/m) is monotonic). Conceptually 2Y/5Y are annualized (CAGR) like
# 3Y; only ranks are stored/shown, so the cumulative computation is exact for them.
MONTH_WINS = [1, 3, 6, 9, 24, 60]                           # the extra raw windows (months)
RES_WINS   = [1, 3, 6, 9, 12, 24, 36, 60]                   # all windows (incl. 1Y/3Y) for residency
WIN_LABEL  = {1: "1 Month", 3: "3 Month", 6: "6 Month", 9: "9 Month",
              12: "1 Year", 24: "2 Year", 36: "3 Year", 60: "5 Year"}

# ── THE SIP INVESTOR BOOKS (Kyser 2026-08-13) ──────────────────────────────────────────────
# The deck used to offer two SIP *instalment-day* conventions — 1st of the month and month-end
# — which answered a question nobody asks. It now offers two INVESTOR BOOKS instead: someone who
# has been running a monthly SIP for three years, and someone who has been running one for five.
# The rolling-window control stops meaning "how long was the SIP" and starts meaning "over what
# period are we measuring it", which is what an investor actually wants to know: *my* 3-year
# book, judged on the last six months.
#
# Both books use ONE instalment day — the trading day nearest the 15th (see sip_month_grids).
#
# A window can never be longer than the book it measures, so the 5-Year window exists only on
# the 5-year book; on the 3-year book it ships as all-null and the panel shows "—". That is
# arithmetic, not a display choice. Every other window stays available on both books, including
# 1M/3M/6M — Kyser was shown that those sit within ~1.5 quartile-agreement points of lumpsum and
# chose to keep them anyway (2026-08-13).
SIP_BOOKS = {"sip3y": 36, "sip5y": 60}                      # payload key -> book length, months
SIP_BOOK_LABEL = {"sip3y": "3-year SIP book", "sip5y": "5-year SIP book"}


def calendar_returns_m(nav: pd.DataFrame, months: int) -> pd.DataFrame:
    """Cumulative point-to-point return over a trailing `months`-month CALENDAR window:
    for each date t, t_m = last trading day on/before the SAME calendar day `months` months
    before t (DateOffset clamps the day to the target month's length, e.g. Mar-31 - 1M ->
    Feb-28). `R = NAV_t / NAV_{t_m} - 1` (cumulative). Mirrors calendar_returns()'s exact-
    calendar lookback, but month-based and never annualized (sub-year)."""
    idx = nav.index
    targets = pd.DatetimeIndex(idx) - pd.DateOffset(months=months)
    pos = idx.get_indexer(pd.DatetimeIndex(targets), method="ffill")
    vals = nav.values.astype(float)
    r = np.full(vals.shape, np.nan)
    ok = pos >= 0
    r[ok] = vals[ok] / vals[pos[ok]] - 1.0
    return pd.DataFrame(r, index=idx, columns=nav.columns)


def all_peer_quartiles_m(nav, cmap, categories, months, n_jobs=-1):
    """Daily all-peer (MFI) quartile labels for a `months`-month window (raw, like 1Y/3Y)."""
    r = calendar_returns_m(nav, months)
    members = {c: list(cmap.index[cmap["Category"] == c]) for c in categories}
    return _quartiles_for_returns(r, members, n_jobs=n_jobs)


def vr_quartiles_m(nav, cme, months, avoid=AVOID_EXACT_CATS, min_peers=1,
                   only_cats=None, n_jobs=-1):
    """Daily VR exact-peer quartile labels for a `months`-month window, as a MultiIndex
    (cat, scheme) frame (mirrors exact_peer_scoring's qy_1y/qy_3y shape, raw quartiles only)."""
    r = calendar_returns_m(nav, months)
    cats = sorted(c for c in cme["Category"].dropna().unique() if c not in avoid)
    if only_cats is not None:
        cats = [c for c in cats if c in only_cats]
    jobs = []
    for c in cats:
        fs = list(dict.fromkeys(s for s in cme.index[cme["Category"] == c] if s in nav.columns))
        if len(fs) < min_peers:
            continue
        jobs.append((c, r[fs]))
    labs = Parallel(n_jobs=n_jobs)(delayed(_quartile_block)(sub) for c, sub in jobs)
    return _assemble_mi({c: lab for (c, _), lab in zip(jobs, labs)})


def _vr_daily_by_scheme(qy_mi: pd.DataFrame) -> pd.DataFrame:
    """Collapse a MultiIndex(cat, scheme) daily-quartile frame to dates x scheme-NAME, keeping
    each scheme's HOME membership (KV ruling 2026-07-16: scheme-level reporting always comes
    from the home peerset): the non-widened category when one exists; else the widened set
    itself (native funds like Bal Bhavishya Yojna / Retirement 40s, whose only membership is
    their widened home). Borrowed memberships exist only inside their table's analysis."""
    home, rep = {}, {}
    for (c, sch) in qy_mi.columns:
        (rep if c in REPEATED_PEER_CATS else home).setdefault(sch, (c, sch))
    pick = {**rep, **home}                      # a home (non-widened) membership wins
    cols = list(pick.values())
    sub = qy_mi.loc[:, cols].copy()
    sub.columns = [sch for (_, sch) in cols]
    return sub


def window_residency(qy_daily: pd.DataFrame, asof_dates, months_back: int) -> dict:
    """Quartile RESIDENCY: for each scheme, the number of trading days spent in Q1/Q2/Q3/Q4
    over the trailing `months_back`-month CALENDAR window ending at each as-of date.
       window = (asof - months_back months, asof]  (same calendar-month lookback as the returns).
    `qy_daily`: daily dates x schemes labels (q1..q4 / NaN). Returns
       {scheme: {"f": firstAsofIdx, "v": [[q1,q2,q3,q4] | None, ...]}}  aligned to asof_dates,
    with leading/trailing all-empty as-of dates trimmed. n_days_traded = q1+q2+q3+q4."""
    didx = pd.DatetimeIndex(qy_daily.index)
    asof = pd.DatetimeIndex(asof_dates)
    qnum = qy_daily.replace({"q1": 1, "q2": 2, "q3": 3, "q4": 4}).apply(pd.to_numeric, errors="coerce")

    def _padcum(mask):                      # cumulative day-count, with a leading 0 row so
        c = mask.cumsum().values            # padded[p] = count over the first p trading days,
        return np.vstack([np.zeros((1, c.shape[1])), c])   # i.e. dates <= didx[p-1]
    cums = {q: _padcum(qnum == q) for q in (1, 2, 3, 4)}
    # padded[searchsorted(D,'right')] = count over all trading days on/before D
    end_pos = didx.searchsorted(asof, side="right")
    start_pos = didx.searchsorted(pd.DatetimeIndex(asof - pd.DateOffset(months=months_back)), side="right")

    out = {}
    for j, sch in enumerate(qy_daily.columns):
        rows = []
        for k in range(len(asof)):
            cnt = [int(cums[q][end_pos[k], j] - cums[q][start_pos[k], j]) for q in (1, 2, 3, 4)]
            rows.append(cnt if sum(cnt) > 0 else None)
        f = next((i for i, r in enumerate(rows) if r is not None), None)
        if f is None:
            continue
        last = len(rows) - next(i for i, r in enumerate(reversed(rows)) if r is not None)
        out[sch] = {"f": f, "v": rows[f:last]}
    return out


def quartiles_roundup(sorted_desc_index):
    """Given an index sorted best->worst, return [q1,q2,q3,q4] lists; remainder to TOP buckets
    (nb 1286-1296)."""
    n = len(sorted_desc_index)
    base, extra = divmod(n, 4)
    sizes = [base + (1 if i < extra else 0) for i in range(4)]
    out, start = [], 0
    for sz in sizes:
        out.append(list(sorted_desc_index[start:start + sz]))
        start += sz
    return out


def _sort_desc_stable(s: pd.Series):
    """Index of `s` sorted by value DESCENDING with a STABLE, position-aware tie-break (ties
    keep their original column order). This is deterministic — unlike pandas' default unstable
    quicksort — so the dashboard's client-side re-bucket (sort by score desc, ties by original
    index) reproduces the engine's quartiles EXACTLY, not just to within tie-noise."""
    order = np.argsort(-s.values, kind="stable")
    return s.index[order]


# %% ALL-PEER (MFI categories) quartiles + Sleeve/AMC %AUM tables  (BUILD_SPEC §4c, §4h)
def _quartile_block(sub: pd.DataFrame, with_ranks: bool = False):
    """Per date: rank a category's schemes by return desc -> q1..q4 (round-up).

    with_ranks=False (the default — every pre-existing call site) returns just the label
    frame, exactly as before. with_ranks=True (STANDING_SIP_DESIGN.md, 2026-08-11)
    returns a tuple (labels, ranks): `ranks` is a FLOAT frame holding each rated scheme's
    1-based position in the SAME `_sort_desc_stable` descending order that assigns the
    quartiles (NaN = unrated that date). Because label and rank come from ONE sort, a
    rank can never disagree with its quartile bucket — that coherence is what the
    standing sanity gate asserts."""
    dates = sub.index[sub.notna().any(axis=1)]
    lab = pd.DataFrame(index=dates, columns=sub.columns, dtype=object)
    rnk = pd.DataFrame(index=dates, columns=sub.columns, dtype=float) if with_ranks else None
    for d in dates:
        a = sub.loc[d].dropna()
        if len(a):
            srt = _sort_desc_stable(a)
            for l, names in zip(["q1", "q2", "q3", "q4"], quartiles_roundup(srt)):
                lab.loc[d, names] = l
            if with_ranks:
                rnk.loc[d, srt] = np.arange(1, len(srt) + 1, dtype=float)
    return (lab, rnk) if with_ranks else lab


def _quartiles_for_returns(rets: pd.DataFrame, members: dict, n_jobs=-1):
    """Per-category quartile labels (parallel). Returns dates×schemes of labels."""
    subs = [rets[[s for s in fs if s in rets.columns]]
            for fs in members.values() if any(s in rets.columns for s in fs)]
    blocks = Parallel(n_jobs=n_jobs)(delayed(_quartile_block)(s) for s in subs)
    out = pd.concat(blocks, axis=1) if blocks else pd.DataFrame(index=rets.index)
    return out.reindex(rets.index)


def all_peer_quartiles(nav, cmap, categories, basis="calendar"):
    """All-peer (MFI) quartile labels for 1Y and 3Y. Members = cmap Category groups.
    basis='calendar' (exact calendar dates, vendor-matching — KV's standard) or 'shift'
    (250/750 trading-day, the notebook's original all-peer basis)."""
    r1, r3 = calendar_returns(nav) if basis == "calendar" else rolling_returns(nav)
    members = {c: list(cmap.index[cmap["Category"] == c]) for c in categories}
    return _quartiles_for_returns(r1, members), _quartiles_for_returns(r3, members)


def sleeve_amc_tables(qy_1y, qy_3y, pct_aum, cmap, top15=TOP15, truncate="2013-12-31"):
    """Build Sleevewise + AMCwise %AUM-in-Q1..Q4 (mirror old .py build_quartile_aum_tables).
    qy_*: dates×schemes labels (cols=schemes). pct_aum: dates×schemes share. Returns
    (sleeve_df, amc_df) with dates in COLUMNS.

    `top15=None` keeps EVERY fund house (the scope the dashboard now ships, 2026-08-13, so its
    Top-15 / All-AMCs toggle has real numbers to switch to). Passing a list narrows it again —
    the Excel %AUM workbook still does exactly that, so no existing deliverable changes shape.
    Each house's shares are self-contained: pct_aum is already a share of that house's OWN AUM,
    so adding houses cannot move an existing house's numbers."""
    pct_T = pct_aum.T  # schemes × dates

    def slice_q(qy, label):
        outs = []
        for q in ["q1", "q2", "q3", "q4"]:
            # cast to float so date columns stay numeric through groupby(numeric_only=True)
            df = ((qy.T == q).astype(float) * pct_T.astype(float)).copy()  # schemes × dates
            df["Quartile"] = f"% in Q{q[-1]}"
            df["Rolling Window"] = label
            outs.append(df)
        return pd.concat(outs)

    pq1, pq3 = slice_q(qy_1y, "1 Year"), slice_q(qy_3y, "3 Year")

    def nonempty(df):
        dc = [c for c in df.columns if isinstance(c, pd.Timestamp)]
        return df.loc[df[dc].notna().any(axis=1)]

    pq1, pq3 = nonempty(pq1), nonempty(pq3)
    cmeta = cmap[["Breakdown", "Category", "Scheme Sub Nature", "fund house"]]
    m1 = pq1.merge(cmeta, left_index=True, right_index=True)
    m3 = pq3.merge(cmeta, left_index=True, right_index=True)
    m1["Scheme"] = m1.index; m3["Scheme"] = m3.index
    gk = ["Rolling Window", "Breakdown", "Category", "Scheme Sub Nature", "fund house", "Scheme", "Quartile"]
    merged = pd.concat([m3.groupby(gk).sum(numeric_only=True), m1.groupby(gk).sum(numeric_only=True)])
    tcut = pd.to_datetime(truncate)
    merged = merged[[c for c in merged.columns if isinstance(c, pd.Timestamp) and c >= tcut]]

    df = merged.reset_index()
    visual = df.groupby(["Rolling Window", "fund house", "Breakdown", "Quartile"]).sum(numeric_only=True)
    dcols = [c for c in visual.columns if isinstance(c, pd.Timestamp)]
    visual = visual[dcols]

    # "% of AMC AUM" rows
    aum_rows = cmeta.merge(pct_T[dcols], left_index=True, right_index=True, how="right")
    if top15 is not None:
        aum_rows = aum_rows.loc[aum_rows["fund house"].isin(top15)]
    parts = []
    for rw in ("1 Year", "3 Year"):
        a = aum_rows.copy(); a["Quartile"] = "% of AMC AUM"; a["Rolling Window"] = rw
        parts.append(a)
    aum_tot = pd.concat(parts).groupby(["Rolling Window", "fund house", "Breakdown", "Quartile"]).sum(numeric_only=True)[dcols]

    sleeve = pd.concat([visual, aum_tot]).sort_index(level=["Rolling Window", "fund house", "Breakdown", "Quartile"])
    if top15 is not None:
        sleeve = sleeve.loc[sleeve.index.get_level_values("fund house").isin(top15)]
    amc = sleeve.reset_index().groupby(["fund house", "Rolling Window", "Quartile"]).sum(numeric_only=True)
    return sleeve, amc


# %% EXACT-PEER (Value Research) scoring  (BUILD_SPEC §4e/4f) — the corrected method
def _assemble_mi(dct):
    """dict{cat: df[dates×schemes]} -> single frame with MultiIndex (cat, scheme) columns."""
    parts = []
    for c, df in dct.items():
        df = df.copy()
        df.columns = pd.MultiIndex.from_product([[c], df.columns])
        parts.append(df)
    return pd.concat(parts, axis=1) if parts else pd.DataFrame()


def _exact_cat_worker(c, f1, f3, b3):
    """Per-category: 1Y & 3Y quartile labels + 3Y alpha. Returns (c, l1, l3, a3)."""
    a3 = f3.sub(b3, axis=0)
    dates = f3.index[f3.notna().any(axis=1)]
    l1 = pd.DataFrame(index=dates, columns=f1.columns, dtype=object)
    l3 = pd.DataFrame(index=dates, columns=f3.columns, dtype=object)
    for d in dates:
        a = f1.loc[d].dropna()
        if len(a):
            for lab, names in zip(["q1", "q2", "q3", "q4"],
                                  quartiles_roundup(_sort_desc_stable(a))):
                l1.loc[d, names] = lab
        a = f3.loc[d].dropna()
        if len(a):
            for lab, names in zip(["q1", "q2", "q3", "q4"],
                                  quartiles_roundup(_sort_desc_stable(a))):
                l3.loc[d, names] = lab
    return c, l1, l3, a3.loc[dates]


def _rebucket_worker(c, sub):
    """Per-category: re-sort composite into fresh q1..q4 + within-cat rank percentile."""
    dates = sub.index[sub.notna().any(axis=1)]
    lab = pd.DataFrame(index=dates, columns=sub.columns, dtype=object)
    per = pd.DataFrame(index=dates, columns=sub.columns, dtype=float)
    for d in dates:
        a = sub.loc[d].dropna()
        if len(a):
            srt_idx = _sort_desc_stable(a)
            for l, names in zip(["q1", "q2", "q3", "q4"], quartiles_roundup(srt_idx)):
                lab.loc[d, names] = l
            per.loc[d, a.index] = (a.rank() / len(a)).values
    return c, lab, per


STALE_BENCH_DAYS = 45   # cap a category's composite only if its benchmark is >this many days
                        # behind scheme-latest (NASDAQ/S&P, broken composites); monthly indices ok


def exact_peer_scoring(nav, bench_cal, cme, exact_bench, w_1y=0.8, use_bonus=True,
                       avoid=AVOID_EXACT_CATS, only_cats=None, min_peers=1, n_jobs=-1,
                       bench_last=None, stale_days=STALE_BENCH_DAYS):
    """Mirror notebook cells 49-66 exactly (parallel). Returns a dict of MultiIndex(cat,scheme)
       frames: qy_1y, qy_3y, s1 (5/4/3/2), s3_raw (4/3/2/1), beat (1/0/None), alpha_3y,
       composite, qy_score (composite quartiles), score_per, plus `members`."""
    # EXACT-PEER uses exact-calendar returns (nb cell 49), for schemes AND benchmarks
    r1, r3 = calendar_returns(nav)
    _, rb3 = calendar_returns(bench_cal)   # benchmark 3Y CAGR on the scheme calendar

    cats = sorted(c for c in cme["Category"].dropna().unique() if c not in avoid)
    if only_cats is not None:
        cats = [c for c in cats if c in only_cats]

    members, jobs = {}, []
    for c in cats:
        fs = list(dict.fromkeys(s for s in cme.index[cme["Category"] == c] if s in nav.columns))
        if len(fs) < min_peers:   # 'fewer than N peers excluded from quartile computation'
            continue
        members[c] = fs
        bname = exact_bench.get(c)
        b3 = rb3[bname] if (bname in rb3.columns) else pd.Series(np.nan, index=nav.index)
        jobs.append((c, r1[fs], r3[fs], b3))

    out = Parallel(n_jobs=n_jobs)(delayed(_exact_cat_worker)(*j) for j in jobs)
    qy1 = {c: l1 for c, l1, l3, a3 in out}
    qy3 = {c: l3 for c, l1, l3, a3 in out}
    alp3 = {c: a3 for c, l1, l3, a3 in out}

    QY1, QY3, A3 = _assemble_mi(qy1), _assemble_mi(qy3), _assemble_mi(alp3)
    # numeric scores (notebook: 1Y -> 5/4/3/2, 3Y -> 4/3/2/1, ffill, then +1 if beat benchmark)
    s1 = QY1.replace({"q1": 5.0, "q2": 4.0, "q3": 3.0, "q4": 2.0}).apply(pd.to_numeric, errors="coerce").ffill()
    s3_raw = QY3.replace({"q1": 4.0, "q2": 3.0, "q3": 2.0, "q4": 1.0}).apply(pd.to_numeric, errors="coerce").ffill()
    bonus = (A3 > 0).astype(float)            # NaN alpha -> False -> 0 (matches notebook)
    s3 = s3_raw + (1.0 if use_bonus else 0.0) * bonus
    composite = w_1y * s1 + (1 - w_1y) * s3
    # beat flag for the dashboard: 1 beat / 0 not-beat / None bench-unavailable
    beat = A3.where(A3.isna(), other=(A3 > 0).astype(float))  # 1.0/0.0 where avail, NaN where not

    # Cap each category's bonus-dependent composite at its benchmark's last real date, so the
    # +1 bonus is always valid where shown (NASDAQ/S&P stop at 2026-03-11, Multi-Asset 11-28,
    # etc.). Categories whose benchmark has NO data (e.g. Make-in-India) are NOT capped — their
    # bonus is simply always 0 and the peer-only composite runs to scheme-current.
    if use_bonus and bench_last:
        scheme_latest = nav.index.max()
        thresh = scheme_latest - pd.Timedelta(days=stale_days)
        comp_cats = set(composite.columns.get_level_values(0))
        for c in members:
            if c not in comp_cats:
                continue
            lr = bench_last.get(exact_bench.get(c))
            if lr is not None and pd.notna(lr) and lr < thresh:
                composite.loc[composite.index > lr, c] = np.nan

    # re-bucket composite into fresh quartiles per category/date (nb 1626-1678) — parallel
    cats_in = [c for c in members if c in composite.columns.get_level_values(0)]
    reb = Parallel(n_jobs=n_jobs)(delayed(_rebucket_worker)(c, composite[c]) for c in cats_in)
    qsc = {c: lab for c, lab, per in reb}
    sper = {c: per for c, lab, per in reb}

    return dict(qy_1y=QY1, qy_3y=QY3, s1=s1, s3_raw=s3_raw, beat=beat, alpha_3y=A3,
                composite=composite, qy_score=_assemble_mi(qsc), score_per=_assemble_mi(sper),
                members=members)


def benchmark_exclusions(cme, exact_bench, bench_cal, bench_last, scheme_latest,
                         stale_days=STALE_BENCH_DAYS, avoid=AVOID_EXACT_CATS):
    """Per-category benchmark disclosures for the Data Quality & Exclusions tab.

    TWO different things can be wrong with a scored category's benchmark, and the tab has to
    say WHICH, because they damage different numbers:

      1. STALE — the benchmark has NAV, but it stopped more than `stale_days` (45) behind
         scheme-latest: NASDAQ/S&P frozen at 2026-03-11, the broken Multi-Asset composite at
         2025-11-28, the monthly CRISIL indices when a month-end has not landed yet. The +1
         bonus is still VALID up to that last real date, so `exact_peer_scoring` CAPS the
         composite after it (sets NaN) and the category simply ends early. The reason string
         therefore carries the cap date.

      2. EMPTY — the vendor ships no usable NAV for that benchmark AT ALL. Either the mapping
         sheet routes the category to a benchmark that never resolved (no source, or the MFI
         identifier is not in the file, so `load_bench_nav` never creates the column), or the
         column exists but every value is NaN. There is no date to cap at, so nothing is
         capped; instead the benchmark return, the 3Y excess/alpha and the +1 beats-benchmark
         bonus are unavailable for EVERY date, and the category is scored peer-only for its
         whole history. That is a permanent 1-point handicap versus every category that can
         earn the bonus, so it must be disclosed as loudly as the stale case — it is not
         visible anywhere else in the output (`beat` is simply null and `bonus` simply 0).

    DETECTED GENERICALLY, never from a hard-coded list of category names: case 2 is exactly
    "the mapped benchmark resolves to no column in the aligned frame, or to a column whose
    `last_real` is NaT". So the day a third benchmark goes empty, it discloses itself. (As of
    the August 2026 data this flags 'Make in India' and 'Equity Savings'.)

    Categories in `avoid` (AVOID_EXACT_CATS) are skipped: `exact_peer_scoring` never scores
    them, so they have no composite for a benchmark to help or handicap — saying their
    benchmark is missing would be true but misleading. Nothing else changes for them here.
    """
    out = []
    have = set(bench_cal.columns) if bench_cal is not None else set()
    skip = set(avoid)
    thresh = scheme_latest - pd.Timedelta(days=stale_days)
    for c in sorted(set(cme["Category"].dropna())):
        if c in skip:
            continue
        bn = exact_bench.get(c)
        # `load_bench_nav` keys every column by a stripped STRING display name, so anything
        # that is not a non-blank string (None, NaN from a blank mapping cell) means the
        # category has no benchmark mapped at all.
        named = isinstance(bn, str) and bn.strip() != ""
        lr = bench_last.get(bn) if named else None
        n = int((cme["Category"] == c).sum())          # VR memberships riding on this benchmark
        if (not named) or (bn not in have) or lr is None or pd.isna(lr):
            what = (f"benchmark '{str(bn).strip()}' resolves to NO usable NAV at all "
                    f"(empty vendor series)" if named
                    else "no benchmark is mapped for this category")
            out.append({"item": c, "type": "category (no benchmark NAV)",
                        "reason": f"{what}; "
                                  f"benchmark return, 3Y excess and the +1 beats-benchmark bonus "
                                  f"unavailable on every date - {n} VR peers scored peer-only "
                                  f"(nothing capped; composite runs to scheme-current)"})
        elif lr < thresh:
            out.append({"item": c, "type": "category (VR composite cap)",
                        "reason": f"benchmark stale (last {lr.date()}); composite capped >{stale_days}d"})
    return out


# %% VR-peer %AUM in Q1..Q4 (composite-quartile based)  (BUILD_SPEC §4h; nb 1689-1712)
def vr_amc_table(qy_score, pct_aum, cmap, top15=TOP15, repeated=REPEATED_PEER_CATS):
    """%AUM in Q1..Q4 by fund house, from the COMPOSITE-score quartiles (qy_score, a
    MultiIndex(cat,scheme) frame). Drops BORROWED memberships of 'repeated peer-set'
    (widened) categories — a scheme whose home is another category is counted there,
    once; a NATIVE fund whose ONLY membership is the widened set (e.g. Bal Bhavishya
    Yojna, Retirement 40s Plans) reports from that set (KV ruling 2026-07-03/16: each
    scheme counts exactly once, nobody counts zero times). dates in COLUMNS.
    Also returns aum-weighted score-percentile by fund house."""
    cats = qy_score.columns.get_level_values(0)
    schs = qy_score.columns.get_level_values(1)
    nonrep_schemes = set(schs[~cats.isin(repeated)])
    borrowed = cats.isin(repeated) & schs.isin(nonrep_schemes)
    qs = qy_score.loc[:, ~borrowed]
    schemes = list(qs.columns.get_level_values(1))
    pa = pct_aum.loc[qs.index, schemes].copy()
    pa.columns = qs.columns
    house = cmap["fund house"]

    def _by_house(mat):  # mat: dates×(cat,scheme) -> fund house × dates
        t = mat.T.reset_index()
        t = t.rename(columns={t.columns[0]: "cat", t.columns[1]: "scheme"}).drop(columns=["cat"])
        t["fund house"] = t["scheme"].map(house)
        return t.drop(columns=["scheme"]).groupby("fund house").sum(numeric_only=True)

    rows = {}
    for q in ["q1", "q2", "q3", "q4"]:
        rows[f"% in Q{q[-1]}"] = _by_house((qs == q).astype(float) * pa)
    exposure = _by_house(qs.notna().astype(float) * pa)          # total VR-peer AUM exposure
    rows["% of AMC AUM (VR peers)"] = exposure

    out = pd.concat({k: v for k, v in rows.items()}, names=["Quartile", "fund house"])
    out = out.loc[out.index.get_level_values("fund house").isin(top15)]
    # reorder to fund house × Quartile
    out = out.reorder_levels(["fund house", "Quartile"]).sort_index()
    return out


# %% month-end sampling + Excel writers  (BUILD_SPEC §4h)
def month_end_dates(index):
    """One Timestamp per (year, month) — the latest present — from a DatetimeIndex."""
    by = {}
    for v in index:
        if isinstance(v, pd.Timestamp):
            by[(v.year, v.month)] = max(v, by.get((v.year, v.month), v))
    return pd.DatetimeIndex(sorted(by.values()))


def month_end_asof(df: pd.DataFrame) -> pd.DataFrame:
    """Index-dated frame -> one row per (year,month): the LAST non-NaN value per column
    on/before that month-end (as-of). Handles benchmark-staleness caps and reporting lags:
    a column that ends mid-month shows its last value at that month-end; a column with no
    data in a month stays NaN that month."""
    idx = pd.DatetimeIndex(df.index)
    out_rows, out_idx = [], []
    for _, grp in df.groupby([idx.year, idx.month], sort=True):
        out_rows.append(grp.ffill().iloc[-1])
        out_idx.append(pd.DatetimeIndex(grp.index)[-1])
    res = pd.DataFrame(out_rows)
    res.index = pd.DatetimeIndex(out_idx)
    return res


def _sample(df, resolution, axis):
    """Sample a frame to month-end along `axis` (0=index dates, 1=column dates)."""
    if resolution == "daily":
        return df
    if axis == 1:
        me = month_end_dates([c for c in df.columns if isinstance(c, pd.Timestamp)])
        keep = [c for c in df.columns if not isinstance(c, pd.Timestamp) or c in set(me)]
        return df[keep]
    me = set(month_end_dates(df.index))
    return df.loc[[i for i in df.index if i in me]]


def write_aum_workbook(path, sleeve, amc_all, vr, resolution="monthly"):
    """File 1: %AUM in Q1..Q4 — Sleevewise + AMCwise(All) + AMCwise(VR peers). Dates in columns."""
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        _sample(sleeve, resolution, axis=1).reset_index().to_excel(w, "Sleevewise %AUM Q1-Q4 - All", index=False)
        _sample(amc_all, resolution, axis=1).reset_index().to_excel(w, "AMCwise %AUM Q1-Q4 - All", index=False)
        _sample(vr, resolution, axis=1).reset_index().to_excel(w, "AMCwise %AUM Q1-Q4 - VR peer", index=False)
    return path


def write_scoring_workbook(path, s1, s3_raw, beat, composite, qy_score, peer_map,
                           resolution="monthly", w_1y=0.8):
    """File 2: exact-peer score COMPONENTS (so any weight/reward combo is reconstructable) +
    the default composite + mapping. dates in index, MultiIndex cat/scheme columns.
       composite_any = w*s1 + (1-w)*(s3_raw + reward*beat==1).  Default shown: w=0.8, reward on."""
    s3_reward = s3_raw + (beat == 1).astype(float)
    def smp(df):
        return _sample(df, resolution, axis=0)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        smp(s1).to_excel(w, "1Y score (q1=5..q4=2)")
        smp(s3_raw).to_excel(w, "3Y score no reward (q1=4..1)")
        smp(s3_reward).to_excel(w, "3Y score +bench reward")
        smp(beat).to_excel(w, "Beats benchmark 3Y (1=yes)")
        _wp = int(round(w_1y * 100))
        smp(composite).to_excel(w, f"Composite {_wp}-{100 - _wp} +reward")
        smp(qy_score).to_excel(w, "Composite quartile (default)")
        peer_map.reset_index().to_excel(w, "Peer & Bench Mapping", index=False)
    return path


def write_dashboard_json(path, res, pct, cmap, amc_all, sleeve, exclusions, peer_map,
                         exact_bench, bench_last, aum_daily=None, w_1y=0.8,
                         qy1=None, qy3=None, qy_all_w=None, qy_vr_w=None,
                         standing=None, sip=None, returns=None):
    """Emit dashboard_data.json. Per VR scheme × month-end it ships the score COMPONENTS and
    AUM so the dashboard recomputes ANY basis live (1Y / 3Y±reward / Composite with weight):
        per scheme: q1y,q3y (return quartiles 1..4), y(=1Y score 5..2), t(=3Y score 4..1),
                    b(beats-benchmark 1/0/null), a(% of its AMC AUM), cr(absolute AUM, Rs cr),
                    sl(sleeve/Breakdown).
    Plus AMC and AMC×sleeve total AUM (Rs cr) so the deck shows %-of-AMC, %-of-sleeve, Rs cr.
    Also ships all-peer quartile %AUM, exclusions, peer/benchmark mapping.

    ADDITIVE (STANDING_SIP_DESIGN.md, 2026-08-11): `standing` (build_standing's dict) and
    `sip` (build_sip's dict) are attached as new top-level keys when given, with the meta
    flags has_standing / sip_conventions; when None (old caller, --no-standing/--no-sip)
    the keys are simply not emitted and the output stays byte-identical to before.
    ADDITIVE (STANDING_SIP_DESIGN_ADDENDUM.md section C, 2026-08-11): `returns`
    (build_returns' dict) is attached the same way, with meta flag has_returns.
    SCOPE WIDENING (STANDING_SIP_DESIGN_ADDENDUM3.md section I, 2026-08-11, ratification R1):
    `residency` now covers EVERY scheme/membership rather than the Top-15 houses — the peer
    analytics grid reads a whole category cohort, which contains non-Top-15 peers. Every
    previously-shipped residency record is unchanged; only new keys appear. Which schemes the
    AGGREGATE tabs surface is decided elsewhere (sleeve_amc_tables / vr_amc_table) and is
    deliberately still Top-15."""
    comp = res["composite"]
    months = month_end_dates(comp.index)
    mstr = [d.date().isoformat() for d in months]
    QMAP = {"q1": 1, "q2": 2, "q3": 3, "q4": 4}
    s1m = month_end_asof(res["s1"]).reindex(months)
    s3m = month_end_asof(res["s3_raw"]).reindex(months)
    btm = month_end_asof(res["beat"]).reindex(months)
    q1ym = month_end_asof(res["qy_1y"]).reindex(months)
    q3ym = month_end_asof(res["qy_3y"]).reindex(months)
    qvm = {m: month_end_asof(qy_vr_w[m]).reindex(months) for m in MONTH_WINS} if qy_vr_w else {}
    pctm = month_end_asof(pct).reindex(months)
    crm = (month_end_asof(aum_daily).reindex(months) if aum_daily is not None
           else pd.DataFrame(index=months))
    sleeve_of = cmap["Breakdown"].to_dict()

    def _q(v):
        return None if (isinstance(v, float) and pd.isna(v)) else QMAP.get(v)

    seen, vr = set(), []
    for (c, sch) in comp.columns:
        if (c, sch) in seen:
            continue
        seen.add((c, sch))
        col1, col3, colb = s1m[(c, sch)].values, s3m[(c, sch)].values, btm[(c, sch)].values
        cq1, cq3 = q1ym[(c, sch)].values, q3ym[(c, sch)].values
        cola = (pctm[sch].reindex(months).values if sch in pctm.columns else np.full(len(months), np.nan))
        colcr = (crm[sch].reindex(months).values if sch in crm.columns else np.full(len(months), np.nan))
        _vrec = {
            "c": c, "s": sch, "h": house_of(sch), "sl": sleeve_of.get(sch, "Other"),
            "q1y": [_q(v) for v in cq1], "q3y": [_q(v) for v in cq3],
            "y": [None if pd.isna(v) else int(round(v)) for v in col1],
            "t": [None if pd.isna(v) else int(round(v)) for v in col3],
            "b": [None if pd.isna(v) else int(round(v)) for v in colb],
            "a": [None if pd.isna(v) else round(float(v), 4) for v in cola],
            "cr": [None if pd.isna(v) else round(float(v), 2) for v in colcr],
        }
        for m in MONTH_WINS:        # 1M/3M/6M/9M VR exact-peer quartile (raw, 1..4)
            cm = (qvm[m][(c, sch)].values if (m in qvm and (c, sch) in qvm[m].columns)
                  else np.full(len(months), np.nan))
            _vrec[f"q{m}m"] = [_q(v) for v in cm]
        vr.append(_vrec)

    # AMC-total and AMC x sleeve-total AUM (Rs cr, month-end) for EVERY fund house (widened
    # 2026-08-13 alongside the sleeve/amc tables — the All-AMCs view needs absolute AUM for the
    # houses it adds, and a house with no schemes here is simply skipped)
    amc_total, amc_sleeve_total = {}, {}
    if aum_daily is not None:
        hcol = pd.Series({s: house_of(s) for s in crm.columns})
        scol = pd.Series({s: sleeve_of.get(s, "Other") for s in crm.columns})
        for h in sorted(set(hcol.values)):
            hs = [s for s in crm.columns if hcol[s] == h]
            if not hs:
                continue
            amc_total[h] = [None if pd.isna(v) else round(float(v), 2) for v in crm[hs].sum(axis=1).values]
            for sl in sorted(set(scol[s] for s in hs)):
                cols = [s for s in hs if scol[s] == sl]
                amc_sleeve_total.setdefault(h, {})[sl] = [None if pd.isna(v) else round(float(v), 2)
                                                          for v in crm[cols].sum(axis=1).values]

    # ── ALL-PEER (MFI) month-end tables. aum_dates is the canonical all-peer calendar:
    # the SLEEVE table's month-end columns (fund house x Breakdown x Quartile x date). The
    # AMC roll-up reuses the same dates so the two All-Peers views are perfectly aligned.
    sleeve_me = _sample(sleeve, "monthly", axis=1)
    aum_dcols = [c for c in sleeve_me.columns if isinstance(c, pd.Timestamp)]
    aum_dates = [d.date().isoformat() for d in aum_dcols]

    def _series(row):
        return [None if pd.isna(x) else round(float(x), 4) for x in row[aum_dcols].values]

    # sleeve index = (Rolling Window, fund house, Breakdown, Quartile)
    sleeve_recs = [{"rw": ix[0], "fh": ix[1], "br": ix[2], "qt": ix[3], "v": _series(row)}
                   for ix, row in sleeve_me.iterrows()]

    amc_me = _sample(amc_all, "monthly", axis=1).reindex(columns=sleeve_me.columns)
    # amc index = (fund house, Rolling Window, Quartile)
    amc_recs = [{"fh": ix[0], "rw": ix[1], "qt": ix[2], "v": _series(row)}
                for ix, row in amc_me.iterrows()]

    # ── PER-SCHEME ALL-PEER rows for the FULL all-peer universe (every scheme that has a
    # category), each with its raw 1Y/3Y all-peer quartile (1..4), % of AMC AUM and absolute
    # Rs cr, on the same aum_dates. The full membership is needed so the dashboard can re-bucket
    # an ALL-PEER COMPOSITE (weighted 1Y/3Y blend) within each complete MFI category client-side
    # — aggregation/drill-down still only surface the Top-15 houses. (The 1Y/3Y Sleeve/AMC/Matrix
    # views read the precomputed sleeve/amc tables above; these rows power composite + drill-down.)
    allpeer = []
    if qy1 is not None and qy3 is not None:
        cat_of = cmap["Category"].to_dict()
        aq1 = month_end_asof(qy1).reindex(aum_dcols)
        aq3 = month_end_asof(qy3).reindex(aum_dcols)
        aqm = {m: month_end_asof(qy_all_w[m]).reindex(aum_dcols) for m in MONTH_WINS} if qy_all_w else {}
        pct_a = month_end_asof(pct).reindex(aum_dcols)
        cr_a = (month_end_asof(aum_daily).reindex(aum_dcols) if aum_daily is not None
                else pd.DataFrame(index=aum_dcols))
        for s in sorted(set(aq1.columns) | set(aq3.columns)):
            hh = house_of(s)
            if s not in cat_of:
                continue
            a1 = aq1[s].values if s in aq1.columns else np.full(len(aum_dcols), np.nan)
            a3 = aq3[s].values if s in aq3.columns else np.full(len(aum_dcols), np.nan)
            pa = pct_a[s].values if s in pct_a.columns else np.full(len(aum_dcols), np.nan)
            ca = cr_a[s].values if s in cr_a.columns else np.full(len(aum_dcols), np.nan)
            _arec = {
                "s": s, "h": hh, "sl": sleeve_of.get(s, "Other"), "c": cat_of.get(s),
                "aq1y": [_q(v) for v in a1], "aq3y": [_q(v) for v in a3],
                "a": [None if pd.isna(v) else round(float(v), 4) for v in pa],
                "cr": [None if pd.isna(v) else round(float(v), 2) for v in ca],
            }
            for m in MONTH_WINS:        # 1M/3M/6M/9M MFI all-peer quartile (raw, 1..4)
                am = (aqm[m][s].values if (m in aqm and s in aqm[m].columns)
                      else np.full(len(aum_dcols), np.nan))
                _arec[f"aq{m}m"] = [_q(v) for v in am]
            allpeer.append(_arec)

    # ── QUARTILE RESIDENCY: for EVERY scheme, the count of trading days spent in Q1/Q2/Q3/Q4
    # over the trailing window ENDING AT each as-of month-end (so it follows the Scheme-Detail
    # as-of date selector). Both universes; all 8 windows. Aligned to aum_dates (all-peer) /
    # months (VR).
    #
    # ★ SCOPE WIDENED 2026-08-11 (STANDING_SIP_DESIGN_ADDENDUM3.md section I.1, orchestrator
    # ratification R1): this block used to be filtered to Top-15 houses because the scheme
    # PICKER only lists those. The peer-analytics grid broke that assumption — it shows the
    # SELECTED scheme's category cohort, and a cohort contains non-Top-15 peers (e.g. the
    # Technology category's Edelweiss / Motilal / Quant funds), each of which needs its own
    # four residency percentages. So residency now covers every scheme (all-peer) and every
    # home membership (VR); a scheme that was never rated in a window is still simply absent.
    # The AGGREGATE tabs (sleeve / amc / matrix / league) are unaffected — their Top-15
    # narrowing lives in sleeve_amc_tables / vr_amc_table, not here.
    residency = {"all": {}, "vr": {}}
    if qy_all_w:
        for m in RES_WINS:
            for sch, rec in window_residency(qy_all_w[m], list(aum_dcols), m).items():
                residency["all"].setdefault(sch, {})[WIN_LABEL[m]] = rec
    if qy_vr_w:
        for m in RES_WINS:
            daily = _vr_daily_by_scheme(qy_vr_w[m])
            for sch, rec in window_residency(daily, list(months), m).items():
                residency["vr"].setdefault(sch, {})[WIN_LABEL[m]] = rec

    # Capped categories must be the VR (exact-peer) categories whose benchmark is stale — the
    # composite is NaN'd after the benchmark's last real date. Iterate the composite's OWN VR
    # categories (NOT cmap["Category"], which are the all-peer names and would miss VR-only
    # categories like Pure International Plan / Nasdaq 100 FOFs, leaving the dashboard to score
    # schemes the engine had already capped).
    scheme_latest = comp.index.max()
    capped = {}
    for c in set(comp.columns.get_level_values(0)):
        lr = bench_last.get(exact_bench.get(c))
        if lr is not None and pd.notna(lr) and lr < scheme_latest - pd.Timedelta(days=STALE_BENCH_DAYS):
            capped[c] = lr.date().isoformat()

    n_cats_vr = len({r["c"] for r in vr})
    # Which houses the dashboard MAY show. `top15` stays the default scope (and the deck's
    # ABSL-vs-Top-15 framing); `houses_all` is every house that actually has a precomputed
    # 1Y/3Y roll-up, so the client can offer the All-AMCs toggle only when the payload can
    # honour it — an older payload simply hides the control instead of drawing empty houses.
    houses_all = sorted({r["fh"] for r in amc_recs})
    house_tables_scope = "all" if len(houses_all) > len(TOP15) else "top15"
    payload = {
        "meta": {"latest": mstr[-1] if mstr else None,
                 "latest_aum_date": aum_dates[-1] if aum_dates else None,
                 "w_default": int(w_1y * 100),
                 "n_vr_schemes": len(vr), "n_vr_categories": n_cats_vr,
                 "n_allpeer_schemes": len(allpeer), "top15": TOP15,
                 "houses_all": houses_all, "house_tables_scope": house_tables_scope,
                 "repeated_cats": REPEATED_PEER_CATS},
        "months": mstr, "amc_dates": aum_dates, "aum_dates": aum_dates,
        "vr": vr, "amc_all": amc_recs, "sleeve": sleeve_recs, "allpeer": allpeer,
        "residency": residency, "residency_windows": [WIN_LABEL[m] for m in RES_WINS],
        "capped_cats": capped,
        "amc_total_cr": amc_total, "amc_sleeve_total_cr": amc_sleeve_total,
        "exclusions": exclusions,
        "peer_map": [{k: (None if (isinstance(v, float) and pd.isna(v)) else v)
                      for k, v in r.items()}
                     for r in peer_map.reset_index().to_dict(orient="records")],
    }
    # ── ADDITIVE blocks (STANDING_SIP_DESIGN.md 2026-08-11 + its ADDENDUM section C): the
    # standing rank records, the SIP mode cube and the return profile, prebuilt by
    # build_standing()/build_sip()/build_returns() in run(). Appended AFTER the payload
    # literal so every pre-existing key serializes byte-identically; when all three are None
    # nothing at all changes.
    if standing is not None:
        payload["meta"]["has_standing"] = 1
        payload["standing"] = standing
    if sip is not None:
        payload["meta"]["sip_conventions"] = list(sip.keys())
        payload["sip"] = sip
    if returns is not None:
        payload["meta"]["has_returns"] = 1
        payload["returns"] = returns
    with open(path, "w", encoding="utf-8") as f:
        # allow_nan=False => emit STRICT/valid JSON (NaN/Infinity are not valid JSON and
        # break browsers' JSON.parse); raises if any stray NaN slipped through, so we catch it.
        json.dump(payload, f, separators=(",", ":"), default=str, allow_nan=False)
    return path, path.stat().st_size / 1e6


# %% CATEGORY STANDING + SIP MODE  (STANDING_SIP_DESIGN.md, 2026-08-11 — ADDITIVE)
# Two engine features, contracted in STANDING_SIP_DESIGN.md (the binding design doc):
#   FEATURE A — "standing": per-scheme 1-based category RANKS per window per month-end,
#     so the deck can say "rank 8 of 24; 70.8% of rated funds / 80.3% of rated AUM at or
#     below". Ranks come from the SAME per-date descending stable sort that assigns the
#     quartiles (_quartile_block(with_ranks=True)), so a rank can never disagree with its
#     quartile digit.
#   FEATURE B — "sip": every quartile/score/residency/standing view recomputed on SIP
#     returns (Rs 100 invested on the first OR last trading day of each calendar month in
#     the window) instead of the lumpsum point-to-point return. The ranking metric is the
#     SIP value ratio VRt = (sum_i NAV_t/NAV_{d_i}) / m — with common installment dates and
#     amounts across schemes this orders IDENTICALLY to XIRR (both are monotone in the
#     final value), so quartiles/ranks from VRt are exact; the actual XIRR is solved only
#     for display, at the latest as-of date.
# Everything in this cell is a standalone additive path: the validated lumpsum builders
# above are untouched, and when run() skips these steps (--no-standing / --no-sip) the
# dashboard JSON is byte-identical to the pre-feature output.

def _all_member_sets(cmap, cats, nav_columns):
    """All-peer member lists per category — the SAME sets the all-peer quartile builders
    use (cmap 'Category' groups, see all_peer_quartiles/_m), filtered to schemes actually
    present in the NAV panel so a mapped-but-absent scheme can't create an empty column."""
    navset = set(nav_columns)
    return {c: [s for s in cmap.index[cmap["Category"] == c] if s in navset] for c in cats}


def _vr_member_sets(cme, nav_columns, avoid=AVOID_EXACT_CATS, min_peers=1):
    """VR exact-peer member lists per category — mirrors exact_peer_scoring/vr_quartiles_m
    exactly: drop the AVOID categories, de-dup with dict.fromkeys preserving order (a
    scheme can appear on two map rows of one category), keep only schemes with NAV, and
    honour min_peers. Using the same recipe guarantees standing/SIP rank the same people
    the quartile builders rank."""
    navset = set(nav_columns)
    out = {}
    for c in sorted(c for c in cme["Category"].dropna().unique() if c not in avoid):
        fs = list(dict.fromkeys(s for s in cme.index[cme["Category"] == c] if s in navset))
        if len(fs) >= min_peers:
            out[c] = fs
    return out


def _win_field(m, allpeer=False):
    """JSON field name for a rolling window, matching the existing lumpsum names exactly:
    12 -> 'q1y', 36 -> 'q3y', anything else -> 'q{m}m'; all-peer fields get an 'a' prefix
    ('aq1y', 'aq3m', ...)."""
    base = {12: "q1y", 36: "q3y"}.get(m, f"q{m}m")
    return ("a" + base) if allpeer else base


def _qdigit(v):
    """'q1'..'q4' / NaN -> 1..4 / None — the deck's digit convention. Module-level twin of
    write_dashboard_json's inner `_q` (same semantics) so the standing/SIP builders can
    reuse it without touching that function."""
    return None if (isinstance(v, float) and pd.isna(v)) else {"q1": 1, "q2": 2, "q3": 3, "q4": 4}.get(v)


def _trim_int_series(vals):
    """A rank column sampled onto an axis -> the residency-style trimmed record
    {"f": firstNonNullAxisIdx, "v": [int|None, ...]} where v[i - f] is the value at axis
    index i. Leading AND trailing all-null stretches are trimmed (interior nulls kept);
    returns None when the scheme was never ranked, so the caller can omit it entirely
    ('absent, like residency')."""
    rows = [None if pd.isna(v) else int(round(float(v))) for v in vals]
    f = next((i for i, r in enumerate(rows) if r is not None), None)
    if f is None:
        return None
    last = len(rows) - next(i for i, r in enumerate(reversed(rows)) if r is not None)
    return {"f": f, "v": rows[f:last]}


def _rank_blocks_for_window(ret, members_all, members_vr, n_jobs=-1):
    """One rolling window's returns frame -> combined quartile-label + rank frames for BOTH
    universes, in ONE joblib Parallel batch (the existing parallel pattern — this is where
    the extra per-category daily sorts run, so they use every core).

    Returns (lab_all, rnk_all, lab_vr, rnk_vr):
      lab_all/rnk_all — dates x scheme-name, all-peer categories concatenated then
        reindexed to the full returns calendar (mirrors _quartiles_for_returns, so
        month-end as-of sampling later behaves identically to the shipped digits);
      lab_vr/rnk_vr — MultiIndex (cat, scheme) frames (mirrors vr_quartiles_m /
        exact_peer_scoring assembly via _assemble_mi)."""
    jobs_all = [(c, ret[[s for s in fs if s in ret.columns]])
                for c, fs in members_all.items() if any(s in ret.columns for s in fs)]
    jobs_vr = [(c, ret[fs]) for c, fs in members_vr.items()]
    outs = Parallel(n_jobs=n_jobs)(
        delayed(_quartile_block)(sub, True) for _, sub in (jobs_all + jobs_vr))
    na = len(jobs_all)
    labs_a = [o[0] for o in outs[:na]]
    rnks_a = [o[1] for o in outs[:na]]
    lab_all = (pd.concat(labs_a, axis=1) if labs_a else pd.DataFrame(index=ret.index)).reindex(ret.index)
    rnk_all = (pd.concat(rnks_a, axis=1) if rnks_a else pd.DataFrame(index=ret.index)).reindex(ret.index)
    lab_vr = _assemble_mi({c: outs[na + i][0] for i, (c, _) in enumerate(jobs_vr)})
    rnk_vr = _assemble_mi({c: outs[na + i][1] for i, (c, _) in enumerate(jobs_vr)})
    return lab_all, rnk_all, lab_vr, rnk_vr


def _standing_fill(dst, rnk_frame, axis, win_label, vr_keys=False):
    """Month-end-sample a daily rank frame onto `axis` (the same as-of sampling the
    quartile digits get: month_end_asof then reindex) and store trimmed {f, v} records
    into `dst`. Keys: scheme name (all-peer) or '<cat>|<scheme>' per VR MEMBERSHIP when
    vr_keys=True (a borrowed membership has its own rank in that category)."""
    if rnk_frame is None or rnk_frame.shape[0] == 0 or rnk_frame.shape[1] == 0:
        return
    samp = month_end_asof(rnk_frame).reindex(list(axis))
    for col in samp.columns:
        key = f"{col[0]}|{col[1]}" if vr_keys else str(col)
        rec = _trim_int_series(samp[col].values)
        if rec is not None:
            dst.setdefault(key, {})[win_label] = rec


def build_standing(nav, cmap, cats, cme, months_axis, aum_axis,
                   avoid=AVOID_EXACT_CATS, min_peers=1, n_jobs=-1):
    """FEATURE A (STANDING_SIP_DESIGN.md): the `standing` block for the dashboard JSON.

    For BOTH universes and ALL 8 RES_WINS windows, recompute the window returns (1Y/3Y via
    calendar_returns — identical to the quartile builders' inputs — the rest via
    calendar_returns_m), run the combined label+rank quartile block per category (labels
    are DISCARDED here; coherence with the shipped quartile digits is asserted by the
    sanity gate instead), month-end-sample the daily rank frames onto the same axes the
    quartile digit series use (all-peer -> the sleeve table's month-end date columns;
    VR -> the composite's month-ends), and trim to {f, v} records.

    Output: {"all": {scheme: {winLabel: {f, v}}}, "vr": {"cat|scheme": {...}}} — ALL
    schemes/memberships (not Top-15-only: the client needs every member's rank to sum
    beaten AUM and to sort the peer table by true rank)."""
    members_all = _all_member_sets(cmap, cats, nav.columns)
    members_vr = _vr_member_sets(cme, nav.columns, avoid, min_peers)
    r1, r3 = calendar_returns(nav)          # exact-calendar 1Y cum / 3Y CAGR, like the quartiles
    standing = {"all": {}, "vr": {}}
    for m in RES_WINS:
        ret = r1 if m == 12 else r3 if m == 36 else calendar_returns_m(nav, m)
        _, rnk_all, _, rnk_vr = _rank_blocks_for_window(ret, members_all, members_vr, n_jobs)
        _standing_fill(standing["all"], rnk_all, aum_axis, WIN_LABEL[m], vr_keys=False)
        _standing_fill(standing["vr"], rnk_vr, months_axis, WIN_LABEL[m], vr_keys=True)
    return standing


def sip_month_grids(nav_index, nav=None, min_cov=0.5):
    """The two SIP installment calendars, taken from the CLEANED trading calendar itself:
    'first' = the first trading day of each calendar month present in `nav_index`,
    'last'  = the last trading day of each month. Returns {"first": DatetimeIndex,
    "last": DatetimeIndex}, both sorted.

    ★ The FINAL month needs care under the month-end convention (deviation from the
    contract's literal grid, reported to the orchestrator): its "last trading day" is just
    the latest day observed so far — when the panel ends MID-month that is not a month-end
    at all but an installment that hasn't happened yet. Keeping it would place an (m+1)-th
    phantom grid date inside every trailing window evaluated AT the as-of date, and the
    strict ==m installment count would then NaN the whole latest cross-section (empirically:
    sip.last.sipret came out EMPTY on any mid-month as-of, which the SIP sanity guard
    hard-fails). So the final month's entry is kept ONLY when it genuinely is its month's
    last business day (then the month-end installment has really happened — e.g. a panel
    ending Jul-31). Historical months are complete by construction, so their entries are
    always true month-ends and are always kept; the 'first' grid needs no such care (the
    current month's first-day installment has always already happened)."""
    idx = pd.DatetimeIndex(nav_index).sort_values()
    first, last = {}, {}
    for v in idx:                    # idx is sorted, so per month the first hit is the
        k = (v.year, v.month)        # month's first trading day and the latest hit is
        if k not in first:           # its last trading day
            first[k] = v
        last[k] = v
    last_idx = pd.DatetimeIndex(sorted(last.values()))
    if len(last_idx):
        t_end = last_idx[-1]
        # BMonthEnd().rollforward(t) == t exactly when t IS its month's last business day.
        # (A rare month-end holiday makes a true month-end look incomplete — that errs on
        # the safe side: the as-of then shows the previous m completed installments.)
        if pd.offsets.BMonthEnd().rollforward(t_end) != t_end:
            last_idx = last_idx[:-1]

    # ── 'mid' — THE SINGLE INSTALMENT CONVENTION (Kyser 2026-08-13) ────────────────────
    # Replaces the first/last pair. Two defects died with that pair:
    #   (a) the BOUNDARY ARTIFACT — a fund whose history starts on the 1st of a month
    #       qualifies on the 'first' grid and misses the 'last' grid by ONE DAY, which
    #       moved a whole category's quartile cut (TRUSTMF Arbitrage, 2026-08-12);
    #   (b) the DEGENERATE 1-MONTH WINDOW — under 'last' the single instalment is bought
    #       on the as-of date itself, so NAV_t/NAV_t = 1 and EVERY fund's gain is exactly
    #       zero, leaving the quartiles to be decided by floating-point noise.
    # Mid-month is immune to both: it is never adjacent to a month boundary, so a fund's
    # inception cannot straddle it the way it straddles the 1st/last, and a mid-month
    # instalment always has ~2 weeks of exposure by the time the month is measured.
    # "Nearest trading day to the 15th" — ties (equidistant either side) break EARLIER,
    # so the instalment has strictly more exposure rather than less.
    # ★ THE INSTALMENT DAY MUST BE A DAY THE UNIVERSE ACTUALLY TRADES (fix, 2026-08-18).
    # `nav_index` is not a trading calendar: a market HOLIDAY survives in it whenever even ONE
    # scheme reports that day (debt/liquid funds price 365 days a year). 2026-01-15 was such a
    # holiday — 1 of the 951 deck schemes had a NAV, the other 950 did not — and because it is
    # the 15th itself, "nearest the 15th" locked January's instalment onto it. The rating rule
    # then demands a non-null NAV at EVERY instalment, so ONE holiday voided the whole opening
    # book: of 951 schemes, 640 had 29 of 30 valid instalments and exactly 1 had all 30. Every
    # mixed window (win < book) of the 3y/5y books silently collapsed to a single rated scheme
    # while the win == book diagonal — which never touches the opening slice — kept working, so
    # the payload still looked populated and no gate caught it.
    #
    # So candidate days are filtered to those reporting at least `min_cov` of the BUSIEST day in
    # their OWN month. Month-local on purpose: an absolute bar ("half of every scheme that ever
    # reports") is not scale-free — in 2006 the panel had a few dozen live schemes, so an absolute
    # bar failed 160 of 245 months and silently truncated SIP history to 85 instalments. A holiday
    # collapses to ~10% of its month's normal level in any era, so the local ratio separates it
    # cleanly while leaving thin early years alone. Pass nav=None for the old index-only behaviour.
    cnt = nav.notna().sum(axis=1) if nav is not None else None
    mid = {}
    for k, days in _group_by_month(idx).items():
        target = pd.Timestamp(year=k[0], month=k[1], day=15)
        pool = list(days)
        if cnt is not None:
            local = [int(cnt.get(d, 0)) for d in days]
            mx = max(local) if local else 0
            if mx:
                keep = [d for d, c in zip(days, local) if c >= min_cov * mx]
                pool = keep or pool     # a month that is ALL holiday keeps its days rather than
                                        # vanishing: dropping a month shifts every later window
        mid[k] = min(pool, key=lambda d: (abs((d - target).days), d))
    mid_idx = pd.DatetimeIndex(sorted(mid.values()))
    # Same phantom-instalment care the 'last' grid needs, for the same reason: when the panel
    # ends BEFORE the 15th of its final month, the "nearest day to the 15th" is just the last
    # day observed so far — an instalment that has not happened yet. Keeping it would place an
    # (m+1)-th grid date inside every trailing window and NaN the whole latest cross-section.
    if len(mid_idx):
        t_end = idx[-1]
        if t_end.day < 15 and mid_idx[-1].month == t_end.month and mid_idx[-1].year == t_end.year:
            mid_idx = mid_idx[:-1]
    return {"first": pd.DatetimeIndex(sorted(first.values())),
            "last": last_idx,
            "mid": mid_idx}


def _group_by_month(idx):
    """{(year, month): [timestamps]} for a sorted DatetimeIndex."""
    out = {}
    for v in idx:
        out.setdefault((v.year, v.month), []).append(v)
    return out


def sip_value_ratio(nav, grid, months):
    """SIP cumulative gain per date x scheme: VRt - 1, where
        VRt = (sum_i NAV_t / NAV_{d_i}) / m
    over the m installment dates d_i = grid dates in the calendar window
    (t - m months, t]  (pd.DateOffset month arithmetic, the SAME lookback convention as
    calendar_returns_m). Intuition: each installment of Rs 100 buys 100/NAV_{d_i} units;
    at t the pot is worth 100 * sum(NAV_t/NAV_{d_i}) against 100*m paid in, so VRt is the
    value-to-cost ratio and VRt - 1 the SIP gain.

    A scheme is RATED at t only if NAV_t is non-null AND its NAV is non-null at ALL m
    installment dates AND the window contains exactly m grid dates — no partial SIPs, no
    fabricated fills beyond the panel's existing careful-ffill. Everything else is NaN.

    Vectorized: reciprocal panel at the grid dates, nan-aware cumulative sums plus a
    parallel valid-count cumsum; per t the window sum is S[j_hi]-S[j_lo] with the count
    C[j_hi]-C[j_lo] required to equal m."""
    idx = pd.DatetimeIndex(nav.index)
    grid = pd.DatetimeIndex(grid)
    vals = nav.values.astype(float)
    gvals = nav.reindex(grid).values.astype(float)     # NAV on installment dates (grid ⊆ calendar)
    with np.errstate(divide="ignore", invalid="ignore"):
        recip = 1.0 / gvals
    valid = np.isfinite(recip)
    r0 = np.where(valid, recip, 0.0)
    S = np.vstack([np.zeros((1, r0.shape[1])), np.cumsum(r0, axis=0)])      # leading 0 row so
    C = np.vstack([np.zeros((1, r0.shape[1])), np.cumsum(valid, axis=0)])   # S[j]=sum grid[0..j-1]
    j_hi = grid.searchsorted(idx, side="right")                             # grid dates <= t
    # POSITIONAL window, matching sip_book_value_ratio (fix 2026-08-18 — see the post-mortem in
    # that function). Kept identical here so the ENDPOINT IDENTITY win==book stays exact on the
    # rated MASK as well as on the values; _verify_sip_book.py asserts both.
    j_lo = np.clip(j_hi - months, 0, None)
    full = (j_hi - months) >= 0                        # the last m installments must all exist
    sums = S[j_hi] - S[j_lo]
    cnts = C[j_hi] - C[j_lo]
    vr = vals * sums / float(months) - 1.0
    vr[(~np.isfinite(vals)) | (cnts != months) | (~full[:, None])] = np.nan
    return pd.DataFrame(vr, index=idx, columns=nav.columns)


def sip_book_value_ratio(nav, grid, book_months, win_months):
    """THE INVESTOR-BOOK MEASURE (Kyser 2026-08-13).

    An investor who has been running a SIP for `book_months`, measured over the LAST
    `win_months`. Returns (gain, w_open) — two date x scheme frames:

        S_open  = sum over grid dates in (t-book, t-win]   of 1/NAV_d      units held at window open
        dS      = sum over grid dates in (t-win,  t]       of 1/NAV_d      units bought inside it
        V_open  = NAV_{t-win} * S_open                                     opening value
        n_new   = number of instalments inside the window                  new money (Rs 1 each)
        gain    = NAV_t*(S_open+dS) / (V_open + n_new) - 1                 value-to-cost, the SAME
                                                                           convention as VRt-1
        w_open  = V_open / (V_open + n_new)                                opening share of cost

    `w_open` is shipped because the client CANNOT recover a rate from the gain alone once
    there is an opening lump: it must solve
        w*(1+r)^(T/365) + ((1-w)/n_new) * sum_j (1+r)^(a_j/365) = 1 + gain
    where a_j are the instalment ages already shipped in `sipdays`. Without w the browser
    would have to guess the split between old and new money, and would silently be wrong.

    ENDPOINT (worth knowing, and asserted in the tests): at win_months == book_months there is
    no opening book at all — S_open = 0, V_open = 0, w_open = 0 — and the formula collapses to
    exactly `sip_value_ratio`, the fresh SIP the deck shipped before this change. So the book
    measure GENERALISES the old one rather than replacing it; the old behaviour is the
    win == book diagonal of the new grid.

    RATING RULE, deliberately strict and identical in spirit to sip_value_ratio: a scheme is
    rated at t only if NAV_t is present, the opening valuation NAV is present, at least
    book_months instalments exist at or before t, and NAV is non-null at EVERY one of them.
    No partial books, no fabricated fills.

    ★ The book and window are counted POSITIONALLY — the LAST book_months / win_months
    instalments at or before t — not as "exactly N grid dates inside the trailing N calendar
    months". The calendar form silently NaN'd the whole latest cross-section whenever a holiday
    roll put the boundary instalment on the excluded edge of the half-open interval while the
    current month's instalment had not happened yet; it cost five days of publishing in Aug-2026.
    The full post-mortem is in the implementation comment below."""
    if win_months > book_months:
        raise ValueError(f"window ({win_months}m) exceeds the book ({book_months}m) — undefined")
    idx = pd.DatetimeIndex(nav.index)
    grid = pd.DatetimeIndex(grid)
    vals = nav.values.astype(float)
    gvals = nav.reindex(grid).values.astype(float)
    with np.errstate(divide="ignore", invalid="ignore"):
        recip = 1.0 / gvals
    valid = np.isfinite(recip)
    r0 = np.where(valid, recip, 0.0)
    S = np.vstack([np.zeros((1, r0.shape[1])), np.cumsum(r0, axis=0)])
    C = np.vstack([np.zeros((1, r0.shape[1])), np.cumsum(valid, axis=0)])

    j_hi = grid.searchsorted(idx, side="right")                                   # grid dates <= t
    # ★ POSITIONAL BOOK (fix, 2026-08-18). The book is "the LAST `book_months` instalments at or
    # before t", not "exactly `book_months` grid dates inside the trailing `book_months` CALENDAR
    # months". The calendar form had a one-day cliff that NaN'd the entire latest cross-section:
    #
    #   as-of 2026-08-14 counted 35 instalments for the 36m book, not 36, because
    #     (a) the grid has NO Aug-2026 instalment — the store ended Fri 14 Aug (15 Aug 2026 was a
    #         Saturday) and the phantom-instalment guard in sip_month_grids correctly drops a
    #         "nearest the 15th" date when the panel ends before the 15th; and
    #     (b) the boundary instalment was 2023-08-14 — 15 Aug 2023 was Independence Day so
    #         "nearest the 15th" rolled BACKWARD — and the interval (t-36m, t] is half-open, so
    #         (2023-08-14, ...] excludes it.
    #   One lost at the bottom, none gained at the top => 35 => `full` False everywhere =>
    #   sip3y.sipret empty => the sanity gate hard-failed 7 consecutive runs and the deck sat at
    #   2026-08-13 for five days. The 60m book survived only by luck: ITS boundary instalment is
    #   2021-08-16 (15 Aug 2021 was a Sunday, so it rolled FORWARD) and so stayed inside the
    #   interval. Two holidays landing on opposite weekdays was the whole difference.
    #   Simulated forward, the calendar rule refuses again at as-of 2026-10-15 (37 instalments).
    #
    # Counting positionally is what "I have been running a SIP for `book_months` months" actually
    # means, and it cannot be moved by a holiday roll or by the phantom guard. The strictness the
    # docstring promises is untouched: cnt_new/cnt_open below still demand a non-null NAV at EVERY
    # instalment, so there are still no partial books and no fabricated fills.
    j_win = j_hi - win_months                      # first instalment INSIDE the window
    j_bok = j_hi - book_months                     # first instalment inside the book
    full = j_bok >= 0                              # enough instalments exist to fill the book
    j_win = np.clip(j_win, 0, None)
    j_bok = np.clip(j_bok, 0, None)

    units_new = S[j_hi] - S[j_win]                 # dS
    units_open = S[j_win] - S[j_bok]               # S_open
    cnt_new = C[j_hi] - C[j_win]
    cnt_open = C[j_win] - C[j_bok]
    n_new = float(win_months)

    if book_months > win_months:
        # opening valuation = the last trading day STRICTLY BEFORE the window's first instalment.
        # That is the positional analogue of "the day the window opens": units carried in from the
        # book are valued the day before the window starts buying, so the value being split into
        # V_open and n_new is exactly the money that was already invested. (Previously this was the
        # calendar date t-win_months, which no longer matches a positionally-defined window.)
        win_open_day = grid.values[np.clip(j_win, 0, len(grid) - 1)]
        pos_open = idx.searchsorted(pd.DatetimeIndex(win_open_day), side="left") - 1
        ok_open = pos_open >= 0
        pos_safe = np.where(ok_open, pos_open, 0)
        nav_open = vals[pos_safe]                   # date x scheme, gathered per row
        V_open = nav_open * units_open
        open_bad = (~np.isfinite(nav_open)) | (~ok_open[:, None])
    else:
        # win == book: there is NO opening book — units_open is 0 by construction, so the
        # opening valuation multiplies zero. Requiring it anyway would refuse to rate a fund
        # whose very first instalment opens the book (no NAV exists before it), which the plain
        # SIP rates happily — and that one difference is what stopped the endpoint from being
        # EXACT. Caught by the endpoint test on 2026-08-13: values matched to 0.00e+00 but the
        # rated/unrated masks disagreed. The endpoint is the whole justification for calling
        # this a generalisation of sip_value_ratio, so it has to hold on the mask too.
        V_open = np.zeros_like(vals)
        open_bad = np.zeros_like(vals, dtype=bool)

    cost = V_open + n_new
    with np.errstate(divide="ignore", invalid="ignore"):
        gain = vals * (units_open + units_new) / cost - 1.0
        w_open = V_open / cost

    bad = ((~np.isfinite(vals)) | open_bad
           | (cnt_new != win_months) | (cnt_open != (book_months - win_months))
           | (~full[:, None]) | (~np.isfinite(cost)) | (cost <= 0))
    gain[bad] = np.nan
    w_open[bad] = np.nan
    return (pd.DataFrame(gain, index=idx, columns=nav.columns),
            pd.DataFrame(w_open, index=idx, columns=nav.columns))


def sip_xirr(install_dates, t, amount, value_t, lo=-0.99, hi=10.0):
    """Hand-rolled XIRR for a level SIP (display only; NO scipy): solve for the annual
    rate r in
        sum_i amount * (1+r)^((t - d_i)/365) = value_t.
    f(r) is STRICTLY INCREASING in r (every exponent >= 0), so Newton from a CAGR-style
    initial guess is tried first, and a guaranteed bisection on the monotone bracket
    (lo, hi] is the fallback. Returns r as a decimal; **None if the true root lies OUTSIDE
    the bracket** (censored — see the note at the bisection: shipping the bound looked like a
    real rate and was not); 0.0 if no time has elapsed (all money in on day t, any rate fits
    — 0 by convention); None if value_t is non-positive/non-finite."""
    return xirr_flows(install_dates, [amount] * len(install_dates), t, value_t, lo=lo, hi=hi)


def xirr_flows(dates, amounts, t, value_t, lo=-0.99, hi=10.0):
    """The general form of the solver above: contributions of DIFFERENT sizes on different
    dates, all positive, valued at t. Solve for the annual rate r in
        sum_i amount_i * (1+r)^((t - d_i)/365) = value_t.

    Needed by the investor-book measure (2026-08-13): a 3-year SIP judged over the last six
    months is not a level SIP over those six months — it is an OPENING LUMP (what the book was
    already worth when the window opened) plus six ordinary instalments. Pricing that as six
    equal payments would report a wildly wrong rate, because almost all of the money was
    already invested before the window began.

    All amounts must be positive, which keeps f(r) strictly increasing in r (every exponent is
    >= 0) — so Newton-from-a-sensible-guess with a guaranteed bisection fallback cannot miss the
    root. Same censoring contract as sip_xirr: None when the root lies outside (lo, hi]."""
    if value_t is None or not np.isfinite(value_t) or value_t <= 0 or not len(dates):
        return None
    e = np.array([(t - d).days for d in dates], dtype=float) / 365.0
    A = np.asarray(amounts, dtype=float)
    if len(A) != len(e) or not np.isfinite(A).all() or (A <= 0).any():
        return None
    V = float(value_t)
    if e.max() <= 0:
        return 0.0
    Asum = float(A.sum())

    def f(r):
        return float((A * np.power(1.0 + r, e)).sum()) - V

    ebar = float((A * e).sum() / Asum)          # MONEY-weighted average age, not a plain mean:
    vratio = V / Asum                           # the opening lump dominates and must dominate
                                                # the starting guess too, or Newton wanders
    r = (vratio ** (1.0 / ebar) - 1.0) if (ebar > 0 and vratio > 0) else 0.0
    r = float(min(max(r, lo + 1e-6), hi))
    for _ in range(60):                                # Newton
        fr = f(r)
        if abs(fr) <= 1e-10 * max(1.0, V):
            return r
        # derivative of f: the amount belongs INSIDE the sum now that amounts differ per date
        # (it used to be a scalar factor outside it — leaving it there would make d1 an array
        # and `d1 <= 0` raise on the next line).
        d1 = float((A * e * np.power(1.0 + r, e - 1.0)).sum())
        if not np.isfinite(d1) or d1 <= 0:
            break
        rn = r - fr / d1
        if not np.isfinite(rn) or rn <= lo or rn > hi or abs(rn - r) < 1e-14:
            r = rn if np.isfinite(rn) and lo < rn <= hi else r
            break
        r = rn
    if abs(f(r)) <= 1e-9 * max(1.0, V):
        return r
    a, b = lo + 1e-9, hi                               # bisection on the monotone bracket
    # ★ 2026-08-12 — CENSORED SOLVES NOW RETURN None, they used to return the BOUND ITSELF.
    # `hi` is 10.0, so a root above the bracket was shipped as the number 10.0 and rendered as
    # "1000.00% p.a." — a printed rate that is not the scheme's rate, indistinguishable from a
    # genuine solve. 62 schemes shipped exactly that. The true values behind them run to
    # astronomic figures (a few days of one instalment annualized), which is precisely why they
    # cannot be shown as a rate. None is already part of this function's contract and the one
    # caller (build_sip -> sipret) already writes null for it, so the panel now prints "—" —
    # honestly missing rather than confidently wrong. The GAIN beside it is unaffected and still
    # carries the information; only the annualized rate is withheld.
    if f(a) >= 0:
        return None                                    # root at/below the lower bound: censored
    if f(b) <= 0:
        return None                                    # root above the upper bound: censored
    for _ in range(200):
        mid = 0.5 * (a + b)
        fm = f(mid)
        if fm == 0 or (b - a) < 1e-12:
            return mid
        if fm < 0:
            a = mid
        else:
            b = mid
    return 0.5 * (a + b)


def build_sip(nav, bench_cal, cmap, cats, cme, exact_bench, bench_last,
              months_axis, aum_axis, avoid=AVOID_EXACT_CATS, min_peers=1,
              stale_days=STALE_BENCH_DAYS, n_jobs=-1, top15=None, verbose=True):
    """FEATURE B (STANDING_SIP_DESIGN.md): the `sip` block for the dashboard JSON —
    the whole lumpsum analytics machinery re-run on SIP returns, for BOTH INVESTOR BOOKS.

    ★ REWRITTEN 2026-08-13 (Kyser). The two blocks used to be instalment-DAY conventions —
    'first' and 'last' trading day of the month — which answered a question nobody asks and
    carried two defects (a one-day inception boundary artifact that moved real quartile cuts,
    and a degenerate 1-month window under 'last' where every fund's gain was exactly zero).
    They are now `sip3y` and `sip5y`: an investor three or five years into a monthly SIP, both
    buying on the trading day nearest the 15th. The rolling window stops meaning "how long the
    SIP ran" and starts meaning "over what period we are measuring it". Per book:

      vr        — per VR MEMBERSHIP '<cat>|<scheme>': month-end SIP quartile digits for
                  all 8 windows (q1m..q60m with q1y/q3y named like lumpsum) + the score
                  analogues y (SIP-1Y quartile -> 5/4/3/2), t (SIP-3Y -> 4/3/2/1), and
                  b (1 if the scheme's SIP-3Y value ratio beats its category benchmark's
                  SIP-3Y on the SAME grid; NaN'd after a stale benchmark's cap date —
                  same trigger and threshold as the lumpsum composite cap). y/t are
                  ffilled before sampling, mirroring the lumpsum s1/s3_raw construction.
      allpeer   — per scheme: aq* month-end digit series on the all-peer axes.
      residency — window_residency on the SIP daily quartile labels (ALL schemes, both
                  universes, VR collapsed to home membership via _vr_daily_by_scheme —
                  identical to the lumpsum residency path). ★ SCOPE WIDENED 2026-08-11
                  (ADDENDUM3 section I.1, ratification R1): was Top-15 houses only,
                  because the scheme PICKER only lists those; the peer-analytics grid
                  needs the residency of every peer in the selected scheme's category,
                  and categories contain non-Top-15 houses. `top15=None` (the default) is
                  the all-schemes scope; passing a house list narrows it again, which is
                  the one-argument escape hatch if the payload ever has to shrink.
      standing  — SIP rank records, same structure as Feature A's standing block.
      sipret    — LATEST as-of only: {scheme: {winLabel: [gainPct, xirrPct]}}, gain =
                  (VRt-1)*100 and the hand-rolled XIRR, both rounded to 2dp (SIPRET_DP —
                  this is a DISPLAY figure, and it is also the reference
                  returns_xirr_roundtrip compares the client's re-solve against, so its
                  own +-0.005pp is subtracted there before the contract bar is applied).
                  Keyed by scheme name (universe-independent). Historical SIP analytics
                  stay ranks/quartiles/residency only — same philosophy as lumpsum, which
                  ships no returns at all.

    Series use the SAME null/int/digit conventions as the existing vr/allpeer fields; a
    scheme/membership with no rated dates anywhere is simply absent."""
    members_all = _all_member_sets(cmap, cats, nav.columns)
    members_vr = _vr_member_sets(cme, nav.columns, avoid, min_peers)
    grids = sip_month_grids(nav.index, nav=nav)
    months_list, aum_list = list(months_axis), list(aum_axis)
    t_latest = months_list[-1] if months_list else None
    top15set = None if top15 is None else set(top15)     # None = every scheme (the default)
    # A category's b is trustworthy only while its benchmark is fresh: same staleness
    # trigger as the lumpsum composite cap (benchmark's last real date > stale_days behind
    # scheme-latest -> NaN every b after that date).
    stale_thresh = nav.index.max() - pd.Timedelta(days=stale_days)

    out = {}
    grid = grids["mid"]                      # ONE instalment day for both books (nearest the 15th)
    for bkey, book in SIP_BOOKS.items():
        vr_fields, all_fields = {}, {}          # (cat,scheme)->{field: [...]}, scheme->{...}
        residency = {"all": {}, "vr": {}}
        standing = {"all": {}, "vr": {}}
        sipret = {}
        for m in RES_WINS:
            if m > book:
                # a measurement window cannot be longer than the book it measures. The field is
                # simply never filled, so the record ships an all-null array for it and the
                # panel prints "—" — the same shape every other absent window already has.
                continue
            ret, w_open = sip_book_value_ratio(nav, grid, book, m)
            lab_all, rnk_all, lab_vr, rnk_vr = _rank_blocks_for_window(
                ret, members_all, members_vr, n_jobs)
            fld_vr, fld_all = _win_field(m), _win_field(m, allpeer=True)

            # month-end digit series, sampled exactly like the lumpsum digits
            if lab_vr.shape[1]:
                dv = month_end_asof(lab_vr).reindex(months_list)
                for col in dv.columns:
                    vr_fields.setdefault(col, {})[fld_vr] = [_qdigit(v) for v in dv[col].values]
            if lab_all.shape[1]:
                da = month_end_asof(lab_all).reindex(aum_list)
                for col in da.columns:
                    all_fields.setdefault(col, {})[fld_all] = [_qdigit(v) for v in da[col].values]

            # SIP residency: same recipe as the lumpsum block in write_dashboard_json —
            # every scheme by default (top15set is None), narrowed only if a caller asks
            cols = [s for s in lab_all.columns
                    if top15set is None or house_of(s) in top15set]
            if cols:
                for sch, rec in window_residency(lab_all[cols], aum_list, m).items():
                    residency["all"].setdefault(sch, {})[WIN_LABEL[m]] = rec
            if lab_vr.shape[1]:
                daily = _vr_daily_by_scheme(lab_vr)
                vcols = [s for s in daily.columns
                         if top15set is None or house_of(s) in top15set]
                if vcols:
                    for sch, rec in window_residency(daily[vcols], months_list, m).items():
                        residency["vr"].setdefault(sch, {})[WIN_LABEL[m]] = rec

            # SIP standing: rank records, same structure as Feature A
            _standing_fill(standing["all"], rnk_all, aum_list, WIN_LABEL[m], vr_keys=False)
            _standing_fill(standing["vr"], rnk_vr, months_list, WIN_LABEL[m], vr_keys=True)

            # score analogues at the 1Y/3Y windows (mirror s1/s3_raw incl. the ffill)
            if m == 12 and lab_vr.shape[1]:
                y = (lab_vr.replace({"q1": 5.0, "q2": 4.0, "q3": 3.0, "q4": 2.0})
                     .apply(pd.to_numeric, errors="coerce").ffill())
                ym = month_end_asof(y).reindex(months_list)
                for col in ym.columns:
                    vr_fields.setdefault(col, {})["y"] = [
                        None if pd.isna(v) else int(round(v)) for v in ym[col].values]
            if m == 36:
                if lab_vr.shape[1]:
                    tsc = (lab_vr.replace({"q1": 4.0, "q2": 3.0, "q3": 2.0, "q4": 1.0})
                           .apply(pd.to_numeric, errors="coerce").ffill())
                    tm = month_end_asof(tsc).reindex(months_list)
                    for col in tm.columns:
                        vr_fields.setdefault(col, {})["t"] = [
                            None if pd.isna(v) else int(round(v)) for v in tm[col].values]
                # b: scheme SIP-3Y vs its category benchmark's SIP-3Y on the SAME grid
                # (benchmark SIP on bench_cal, the carried-forward benchmark panel);
                # construction mirrors lumpsum `beat` (NaN where either side is NaN),
                # then the stale-benchmark cap NaNs everything past the cap date.
                bparts = {}
                # the benchmark is measured on the SAME book, so the comparison is like-for-like:
                # a 5-year-book scheme is judged against a 5-year-book benchmark, not against a
                # fresh 3-year SIP in the index.
                bench_sip3 = sip_book_value_ratio(bench_cal, grid, book, 36)[0]
                for c, fs in members_vr.items():
                    bname = exact_bench.get(c)
                    b3s = (bench_sip3[bname] if bname in bench_sip3.columns
                           else pd.Series(np.nan, index=nav.index))
                    diff = ret[fs].sub(b3s, axis=0)
                    bframe = diff.where(diff.isna(), other=(diff > 0).astype(float))
                    lr = bench_last.get(bname) if bench_last else None
                    if lr is not None and pd.notna(lr) and lr < stale_thresh:
                        bframe = bframe.copy()
                        bframe.loc[bframe.index > lr] = np.nan
                    bparts[c] = bframe
                bmi = _assemble_mi(bparts)
                if bmi.shape[1]:
                    bm = month_end_asof(bmi).reindex(months_list)
                    for col in bm.columns:
                        vr_fields.setdefault(col, {})["b"] = [
                            None if pd.isna(v) else int(round(v)) for v in bm[col].values]

            # sipret: SIP return VALUES at the latest as-of only (gain % + XIRR %).
            # ★ The rate is now solved on the REAL cash flows of a book, not on m level
            # instalments: an opening lump on the window's first day (what the book was already
            # worth) plus one Rs 100 instalment per month inside the window. Pricing a 3-year
            # book's 6-month window as if it were a fresh 6-month SIP overstates the rate
            # enormously — on a worked example at a true 10.0% it reports 23.9%, because it
            # pretends money that had been invested for years arrived last month.
            if t_latest is not None and t_latest in ret.index:
                dts = list(grid[(grid > t_latest - pd.DateOffset(months=m)) & (grid <= t_latest)])
                prior = nav.index[nav.index <= t_latest - pd.DateOffset(months=m)]
                open_date = prior[-1] if len(prior) else None      # same valuation day the book uses
                wrow = w_open.loc[t_latest]
                for sch, v in ret.loc[t_latest].items():
                    if pd.isna(v):
                        continue
                    w = wrow.get(sch, 0.0)
                    w = float(w) if (w is not None and np.isfinite(w)) else 0.0
                    # normalise to Rs 100 per instalment: total cost = 100*m/(1-w), of which the
                    # opening lump is the share w. w == 0 (window == book) reproduces the old
                    # level-SIP flows exactly, so the previous behaviour is the endpoint here too.
                    if 0.0 < w < 1.0 and open_date is not None:
                        f_dates = [open_date] + dts
                        f_amts = [100.0 * m * w / (1.0 - w)] + [100.0] * len(dts)
                    else:
                        f_dates, f_amts = dts, [100.0] * len(dts)
                    cost = float(np.sum(f_amts))
                    xr = xirr_flows(f_dates, f_amts, t_latest, cost * (1.0 + float(v)))
                    sipret.setdefault(sch, {})[WIN_LABEL[m]] = [
                        round(float(v) * 100.0, 2),
                        None if xr is None else round(float(xr) * 100.0, 2)]
        if verbose:
            print(f"      SIP [{bkey} = {SIP_BOOK_LABEL[bkey]}]: {len(vr_fields)} VR memberships, "
                  f"{len(all_fields)} all-peer schemes, {len(sipret)} sipret schemes; "
                  f"residency {len(residency['all'])} all / {len(residency['vr'])} vr "
                  f"({'ALL schemes' if top15set is None else 'top15 only'}), "
                  f"standing {len(standing['all'])} all / {len(standing['vr'])} vr")

        # assemble the JSON-ready records: a membership/scheme with NO rated date anywhere
        # is simply absent (contract: 'like residency'), but an EMITTED record carries ALL
        # its fields — a window with no data (e.g. 5Y on a young panel) becomes an all-null
        # array, matching the existing lumpsum vr/allpeer rows the client indexes directly.
        vr_field_order = [_win_field(m) for m in RES_WINS] + ["y", "t", "b"]
        ap_field_order = [_win_field(m, allpeer=True) for m in RES_WINS]
        null_m, null_a = [None] * len(months_list), [None] * len(aum_list)
        vr_recs = {f"{c}|{s}": {f_: flds.get(f_, null_m) for f_ in vr_field_order}
                   for (c, s), flds in vr_fields.items()
                   if any(x is not None for arr in flds.values() for x in arr)}
        all_recs = {s: {f_: flds.get(f_, null_a) for f_ in ap_field_order}
                    for s, flds in all_fields.items()
                    if any(x is not None for arr in flds.values() for x in arr)}
        out[bkey] = {"vr": vr_recs, "allpeer": all_recs,
                     "residency": residency, "standing": standing, "sipret": sipret,
                     "book_months": book}
    return out


# %% RETURN PROFILE — lumpsum + benchmark + SIP returns
# (STANDING_SIP_DESIGN_ADDENDUM.md section C, 2026-08-11 — ADDITIVE)
#
# Until now the deck shipped only RANKS (quartile digits, standing ranks): Scheme Detail
# could say "you are 8th of 24" but never "you made 12.4%". This cell builds the `returns`
# block so the Return-profile panel can show, for the selected scheme / window / as-of /
# peerset, all six of its rows — lumpsum scheme return, the category benchmark's return,
# the excess, the SIP gain, the benchmark's SIP gain and the SIP excess — with no second
# engine pass and no client-side return maths beyond an annualization.
#
# Four conventions everything below depends on, stated once:
#   1. VALUES ARE CUMULATIVE, ALWAYS, IN PERCENT. LUMPSUM AND BENCHMARK ROUND TO 2dp; SIP
#      GAIN ROUNDS TO 4dp. One convention per measure, no special cases: each window ships
#      `NAV_t / NAV_{t-m} - 1` (point to point). The deck's
#      RANKING source for the 3Y window is a CAGR (`calendar_returns`), so the 3Y cumulative
#      here is rebuilt from the SAME lookback positions and then ASSERTED to reproduce that
#      CAGR exactly under `(1+cum)**(1/3)-1` (tolerance 1e-9) — if that assert ever fires,
#      the displayed value and the shipped quartile have drifted apart and the build stops.
#      The client annualizes for display when the window is >= 1Y (`(1+cum)**(12/m)-1`);
#      annualizing is monotone, so a displayed value can never contradict its quartile.
#   2. THE SAME LOOKBACK MACHINERY AS THE RANKS. 1Y/3Y use `_calendar_lookback_positions`
#      (the exact-calendar "same day N years earlier" rule); the other six windows use
#      `calendar_returns_m`'s `DateOffset(months=m)` rule. The number shown beside a
#      quartile is therefore the very number that quartile was computed from.
#   3. THE SAME AS-OF SAMPLING AS EVERY OTHER SERIES: `month_end_asof` (per column, the last
#      non-NaN value on/before the month-end) reindexed onto the axis, then the residency-
#      style {f, v} trim — so index arithmetic in the client is identical everywhere.
#   4. SCOPE = ALL SCHEMES / ALL MEMBERSHIPS (`top15=None`, the default), identical to
#      `residency` and `standing`. ★ WIDENED 2026-08-11 (the contract conflict flagged by
#      the first build is now RESOLVED): Addendum 2 scoped this to the Top-15 houses because
#      the scheme PICKER only lists those, but STANDING_SIP_DESIGN_ADDENDUM3.md section I.2
#      (orchestrator ratification R1) supersedes it — the peer-analytics grid shows the
#      SELECTED scheme's whole category cohort, and cohorts contain non-Top-15 peers (the
#      Technology category's Edelweiss / Motilal / Quant funds are the worked example), each
#      needing its own return row. Passing a house list to `top15` narrows it again; that is
#      the one-argument size escape hatch, and the size table below prints what the widening
#      actually costs. The AGGREGATE tabs (sleeve / amc / matrix / league) are NOT affected —
#      their Top-15 narrowing lives in sleeve_amc_tables / vr_amc_table, not here.
#   5. SIP GAIN PRECISION = 4dp (ratification R2). The client re-solves the XIRR from the
#      shipped gain, and an XIRR AMPLIFIES the gain's rounding by ~1/e, where e is the
#      money-weighted holding period in years: at 2dp a 1-Month SIP whose single instalment
#      is four days old (e ~ 0.011y) could only pin the annual rate to ~+-0.5pp. Two extra
#      decimals divide that by 100 and bring EVERY window inside the contract's 0.01pp
#      round-trip bar. Lumpsum and benchmark returns stay at 2dp — they are DISPLAYED as
#      shipped, never re-solved, so extra digits would be bytes with no information.


# Decimal places, named once so the builder and the round-trip check can never drift apart:
# LUMP_DP is what a human READS (the panel prints it as shipped), SIPGAIN_DP is what a
# machine SOLVES from (the client's XIRR), SIPRET_DP is the precision of `sip[conv].sipret`'s
# stored XIRR — the reference the round-trip compares against, so its own +-0.5 * 10^-DP is
# an irreducible part of every measured difference.
LUMP_DP, SIPGAIN_DP, SIPRET_DP = 2, 4, 2
# The opening weight is a share in 0..1 that only splits the cash flows for the rate solve;
# 4dp is 0.01% of the book, far finer than the gain's own rounding contributes to that solve.
SIPOPEN_DP = 4


def calendar_returns_cum(nav: pd.DataFrame, months: int) -> pd.DataFrame:
    """CUMULATIVE point-to-point return over a trailing `months`-month window, for ANY of
    the 8 windows, using EXACTLY the lookback each window's ranked returns already use:

      * 12 / 36 months -> `_calendar_lookback_positions(idx, 1|3)`, the exact-calendar rule
        of `calendar_returns` (same calendar date one/three years earlier, Feb-29 -> day-1);
        the difference from `calendar_returns` is only that the 3Y result is left CUMULATIVE
        instead of being cube-rooted into a CAGR.
      * every other window -> `calendar_returns_m`, which is already cumulative, so we simply
        call it (no second implementation of the DateOffset lookback to drift out of sync).

    Why this matters: the ranking (and therefore the quartile digit and the standing rank)
    comes from those positions. Recomputing the same ratio from the same positions is what
    makes the displayed return and the displayed quartile provably the same measurement."""
    if months not in (12, 36):
        return calendar_returns_m(nav, months)
    idx = pd.DatetimeIndex(nav.index)
    pos = _calendar_lookback_positions(idx, months // 12)
    vals = nav.values.astype(float)
    r = np.full(vals.shape, np.nan)
    ok = pos >= 0
    r[ok] = vals[ok] / vals[pos[ok]] - 1.0
    return pd.DataFrame(r, index=nav.index, columns=nav.columns)


def _trim_float_series(vals, nd=2):
    """A RETURN column sampled onto an axis -> the residency/standing-style trimmed record
    {"f": firstNonNullAxisIdx, "v": [float|None, ...]}, where `v[i - f]` is the value at
    axis index i. Float twin of `_trim_int_series`: leading AND trailing all-null stretches
    are trimmed, interior nulls are kept, and an all-null column returns None so the caller
    can omit the record entirely (absent = never rated, exactly like residency).

    `+ 0.0` normalizes the negative zero that `round(-0.001, 2)` produces, so the JSON
    carries `0.0` rather than a puzzling `-0.0`."""
    rows = []
    for v in vals:
        rows.append(None if (v is None or pd.isna(v)) else round(float(v), nd) + 0.0)
    f = next((i for i, r in enumerate(rows) if r is not None), None)
    if f is None:
        return None
    last = len(rows) - next(i for i, r in enumerate(reversed(rows)) if r is not None)
    return {"f": f, "v": rows[f:last]}


def _fill_returns(dst, samp, keys, win_label, nd=2):
    """Store one window's already-sampled value frame into `dst[recordKey][win_label]` as a
    trimmed {f, v} record. `keys` is a list of (recordKey, column) pairs — that indirection
    is what lets the same helper serve all three keyings the contract asks for: scheme name
    (all-peer), '<cat>|<scheme>' per VR MEMBERSHIP (a borrowed membership gets its own
    record), and VR category -> its benchmark column.

    `nd` = decimal places. 2 for the DISPLAYED lumpsum/benchmark returns; 4 for the SIP gain,
    which the client re-solves an XIRR from (ratification R2) — see convention 5 above."""
    if samp is None or samp.shape[1] == 0:
        return
    have = set(samp.columns)
    for key, col in keys:
        if col not in have:
            continue
        rec = _trim_float_series(samp[col].values, nd)
        if rec is not None:
            dst.setdefault(key, {})[win_label] = rec


def build_returns(nav, bench_cal, cmap, cats, cme, exact_bench, aum_axis, months_axis,
                  bench_last=None, avoid=AVOID_EXACT_CATS, min_peers=1,
                  stale_days=STALE_BENCH_DAYS, top15=None, scope="full", verbose=True):
    """The `returns` block (STANDING_SIP_DESIGN_ADDENDUM.md section C):

        returns = {
          "lump":    {"all": {scheme:        {winLabel: {f, v}}},      # axis: aum_dates
                      "vr":  {"cat|scheme":  {winLabel: {f, v}}}},     # axis: months
          "bench":   {VRcategory:            {winLabel: {f, v}}},      # axis: months
          "sipgain": {book: {"all": {...}, "vr": {...}, "bench": {...},
                             "wopen": {"all": {...}, "vr": {...}}}},
          "sipdays": {book: {winLabel: [[dayOffsets at as-of 0], ...]}},   # axis: months
          "sipopen": {book: {winLabel: [ageOfOpeningValuationInDays, ...]}},  # axis: months
        }

    `book` is a key of SIP_BOOKS — "sip3y" / "sip5y" (2026-08-13; it used to be an
    instalment-day convention, "first" / "last").

    `lump`/`bench` values are cumulative percent returns rounded to LUMP_DP (2) decimals;
    `sipgain` values are the SIP gain `(VRt - 1) * 100` on the identical installment grids
    `build_sip` ranks on — so the panel's SIP row and its SIP quartile are the same
    measurement — rounded to SIPGAIN_DP (4) decimals because the client re-solves an XIRR
    from them and the solve amplifies the rounding (ratification R2, convention 5 above).
    SCOPE = every scheme and every membership by default (`top15=None`, ratification R1);
    pass a house list to narrow it. `sipdays[book][win][asofIdx]` is
    the list of `(asof - installment_date).days` offsets for that window at that as-of —
    shipping the OFFSETS instead of a second float series lets the client solve the XIRR
    for a few hundred kilobytes instead of a few megabytes. The list is empty when the window
    does not hold exactly m grid dates, which is precisely when the book measure refuses to
    rate anyone.

    ★ THE OPENING LUMP (2026-08-13). Under a book the window does not start from zero: the
    investor already held something when it opened. Two extra pieces make that solvable in the
    browser — `wopen`, the share of the position's cost that was already invested (per scheme
    x window x as-of, same {f,v} shape as the gain), and `sipopen`, the AGE IN DAYS of the
    opening valuation (one number per window x as-of, because the window opens on a date, not
    per fund). The client then solves
        w*(1+r)^(open/365) + ((1-w)/m) * sum_i (1+r)^(off_i/365) = 1 + gain/100
    which is the level-SIP equation it already solved whenever w == 0 — and w IS 0 exactly when
    the window equals the book, i.e. when the book is a fresh SIP. Without `wopen` the browser
    would have to guess how much of the money was old, and would be silently, largely wrong: on
    a worked 6-month window of a 3-year book, pricing it as a fresh SIP turns a true 10.0% into
    23.9%.

    AXIS NOTE (verified at build time and printed): the all-peer as-of dates are a subset of
    `months`, so `bench` ships ONCE on the `months` axis and the client maps by date string.
    If that ever stops holding, a second copy `bench_aum` on the all-peer axis is emitted and
    a loud warning printed — the contract's fallback.

    BENCHMARK STALENESS (engine-side, deliberately conservative — mirrors the SIP `b` flag):
    a benchmark more than `stale_days` behind scheme-latest has its returns NaN'd after its
    last real observation, so the panel can never show a return computed off a frozen index
    level. The client still shows "—" plus the cap date from `capped_cats`; this just makes
    that outcome impossible to get wrong.

    `scope="lumpsum"` ships only `lump` + `bench` (the contract's size escape hatch)."""
    houses = None if top15 is None else set(top15)
    m_list, a_list = list(months_axis), list(aum_axis)
    if not m_list:
        return {"lump": {"all": {}, "vr": {}}, "bench": {}}

    # ── member sets: the SAME recipes the quartile/standing builders use, optionally narrowed
    # to a house list (`top15`; None = every scheme, the shipped scope). Returns are a
    # per-SCHEME quantity (they do not depend on the peerset), so the two universes differ
    # only in axis and keying — see the size table at the end for what that costs.
    members_all = _all_member_sets(cmap, cats, nav.columns)
    members_vr = _vr_member_sets(cme, nav.columns, avoid, min_peers)
    if houses is not None:
        members_all = {c: [s for s in fs if house_of(s) in houses]
                       for c, fs in members_all.items()}
        members_vr = {c: [s for s in fs if house_of(s) in houses]
                      for c, fs in members_vr.items()}
    all_cols = sorted({s for fs in members_all.values() for s in fs})
    vr_pairs = [(c, s) for c, fs in sorted(members_vr.items()) for s in fs]
    keys_all = [(s, s) for s in all_cols]
    keys_vr = [(f"{c}|{s}", s) for (c, s) in vr_pairs]
    need = sorted(set(all_cols) | {s for _, s in vr_pairs})
    navx = nav[need] if need else nav.iloc[:, :0]

    # ── benchmark panel: one column per DISTINCT benchmark actually present (several
    # categories share a benchmark), plus the category -> benchmark keying for the output.
    bname_of = {c: exact_bench.get(c) for c in members_vr}
    # Broad-market series (Kyser 2026-08-12): one extra pseudo-category so NIFTY 500 TRI is
    # always available on the benchmark side regardless of which category is on screen. Added
    # here — AFTER the real categories — so it can only ever ADD a key, never shadow one.
    if bench_cal is not None and MARKET_BENCH in getattr(bench_cal, "columns", []):
        bname_of[MARKET_KEY] = MARKET_BENCH
    elif verbose:
        print(f"      WARN: market benchmark {MARKET_BENCH!r} absent from the benchmark panel "
              f"— the broad-market series will not ship")
    bcols = sorted({b for b in bname_of.values() if b is not None and b in bench_cal.columns})
    bpanel = bench_cal[bcols] if bcols else None
    keys_bench = [(c, b) for c, b in sorted(bname_of.items()) if b in set(bcols)]
    stale_thresh = nav.index.max() - pd.Timedelta(days=stale_days)

    def _cap_bench(frame):
        """NaN a stale benchmark's values after its last real observation — same trigger and
        threshold as the lumpsum composite cap and the SIP `b` flag."""
        if not bench_last or frame is None:
            return frame
        out_f = frame.copy()
        for b in out_f.columns:
            lr = bench_last.get(b)
            if lr is not None and pd.notna(lr) and lr < stale_thresh:
                out_f.loc[out_f.index > lr, b] = np.nan
        return out_f

    # ── axis relationship. The all-peer as-of dates being a subset of the VR month-ends is
    # what lets one sampling pass serve both axes (the all-peer sample is then just a row
    # selection of the months sample, so the two can never disagree) and lets `bench` ship
    # once. Verified here, every build, rather than assumed.
    mset = set(m_list)
    missing = [d for d in a_list if d not in mset]
    aum_subset = not missing
    if verbose:
        print(f"      axis check: aum_dates ({len(a_list)}) subset of months ({len(m_list)}): "
              f"{aum_subset}" + ("" if aum_subset else f"  ★ {len(missing)} missing, "
                                 f"e.g. {[str(d.date()) for d in missing[:3]]} -> shipping bench twice"))

    def _samp(frame):
        """month-end as-of sample onto (months axis, all-peer axis)."""
        sm = month_end_asof(frame).reindex(m_list)
        sa = sm.reindex(a_list) if aum_subset else month_end_asof(frame).reindex(a_list)
        return sm, sa

    # ── LUMPSUM + BENCHMARK ────────────────────────────────────────────────────────────
    lump = {"all": {}, "vr": {}}
    bench, bench_aum = {}, {}
    _r1x, _r3x = calendar_returns(navx)      # the SHIPPED ranking sources, for the assert
    for m in RES_WINS:
        wl = WIN_LABEL[m]
        ret = calendar_returns_cum(navx, m)
        if m in (12, 36):
            # ★ Round-trip assert: annualizing the cumulative must reproduce the frame the
            # quartiles/ranks were built from, value for value AND null for null. 1Y is an
            # identity (both are the same ratio); 3Y is the real test of the CAGR conversion.
            src = _r1x if m == 12 else _r3x
            chk = ret if m == 12 else (1.0 + ret) ** (1.0 / 3.0) - 1.0
            cv, sv = chk.values.astype(float), src.values.astype(float)
            both = np.isfinite(cv) & np.isfinite(sv)
            dmax = float(np.abs(cv[both] - sv[both]).max()) if both.any() else 0.0
            nmis = int((np.isfinite(cv) != np.isfinite(sv)).sum())
            if dmax > 1e-9 or nmis:
                raise ValueError(
                    f"returns: {wl} cumulative does not round-trip to the ranked returns "
                    f"(max abs diff {dmax:.3e} over {int(both.sum())} cells, "
                    f"{nmis} null-pattern mismatches) — displayed value and quartile would disagree")
            if verbose:
                print(f"      {wl} cumulative <-> ranked-return round-trip: max abs diff "
                      f"{dmax:.2e} over {int(both.sum())} cells (tolerance 1e-9)")
        sm, sa = _samp(ret.mul(100.0))
        _fill_returns(lump["all"], sa, keys_all, wl)
        _fill_returns(lump["vr"], sm, keys_vr, wl)
        if bpanel is not None:
            bsm, bsa = _samp(_cap_bench(calendar_returns_cum(bpanel, m)).mul(100.0))
            _fill_returns(bench, bsm, keys_bench, wl)
            if not aum_subset:
                _fill_returns(bench_aum, bsa, keys_bench, wl)

    out = {"lump": lump, "bench": bench}
    if not aum_subset:
        out["bench_aum"] = bench_aum

    # ── SIP GAIN + INSTALLMENT DAY OFFSETS ─────────────────────────────────────────────
    n_carry = 0
    if scope == "full":
        grids = sip_month_grids(nav.index, nav=nav)
        sipgain, sipdays, sipopen = {}, {}, {}
        grid = pd.DatetimeIndex(grids["mid"])       # ONE instalment day now, for both books
        hi_pos = grid.searchsorted(pd.DatetimeIndex(m_list), side="right")
        for bkey, book in SIP_BOOKS.items():
            g_all, g_vr, g_bench, days, w_all, w_vr = {}, {}, {}, {}, {}, {}
            for m in RES_WINS:
                if m > book:
                    continue          # a window cannot exceed its book — see SIP_BOOKS
                wl = WIN_LABEL[m]
                # ★ 4dp, not 2 (ratification R2): the client SOLVES an XIRR from this number,
                # and the solve amplifies the rounding by ~1/e (e = money-weighted holding
                # period in years). See convention 5 at the top of this cell; the round-trip
                # check below measures the resulting band per window and prints it.
                gainf, wof = sip_book_value_ratio(navx, grid, book, m)
                gain = gainf.mul(100.0)
                sm, sa = _samp(gain)
                _fill_returns(g_all, sa, keys_all, wl, nd=SIPGAIN_DP)
                _fill_returns(g_vr, sm, keys_vr, wl, nd=SIPGAIN_DP)
                # ★ OPENING WEIGHT (2026-08-13). The browser re-solves the annualized rate from
                # the shipped gain + instalment ages. With a book that is no longer enough: most
                # of the money was already invested when the window opened, and the split between
                # that opening lump and the new instalments cannot be recovered from the gain. So
                # ship it. 4dp; 0 exactly when window == book, where the book IS a fresh SIP and
                # the client's existing level-SIP solve is already correct.
                wsm, wsa = _samp(wof)
                _fill_returns(w_all, wsa, keys_all, wl, nd=SIPOPEN_DP)
                _fill_returns(w_vr, wsm, keys_vr, wl, nd=SIPOPEN_DP)
                if bpanel is not None:
                    bsm, _ = _samp(_cap_bench(sip_book_value_ratio(bpanel, grid, book, m)[0].mul(100.0)))
                    _fill_returns(g_bench, bsm, keys_bench, wl, nd=SIPGAIN_DP)
                # How many schemes at the LATEST as-of carry a gain from an earlier day (no
                # NAV on the as-of itself)? Their offsets below are measured to the as-of, so
                # the client's XIRR is ~1 trading day out for them — small on long windows,
                # visible on 1M. Counted, not hidden.
                if len(m_list) and m_list[-1] in gain.index:
                    live = gain.loc[m_list[-1]]
                    n_carry += int((sm.iloc[-1].notna() & live.reindex(sm.columns).isna()).sum())
                # installment day offsets, on the months axis. The window bounds are the SAME
                # searchsorted/DateOffset bounds sip_value_ratio uses, so the offsets always
                # describe exactly the installments that produced the gain above; a window
                # that does not hold exactly m grid dates rates nobody, and ships [].
                lo_pos = grid.searchsorted(
                    pd.DatetimeIndex(m_list) - pd.DateOffset(months=m), side="right")
                # ...and the age of the OPENING VALUATION itself, so the client can date the
                # lump. It is the last trading day at or before t-m: exactly the day
                # sip_book_value_ratio values the opening book on, so the client's cash-flow
                # dates and the engine's are the same dates, not merely similar ones.
                nav_idx = pd.DatetimeIndex(navx.index)
                op_pos = nav_idx.searchsorted(
                    pd.DatetimeIndex(m_list) - pd.DateOffset(months=m), side="right") - 1
                rows, orows = [], []
                for k, t in enumerate(m_list):
                    if hi_pos[k] - lo_pos[k] != m:
                        rows.append([]); orows.append(None)
                    else:
                        rows.append([int((t - d).days) for d in grid[lo_pos[k]:hi_pos[k]]])
                        orows.append(int((t - nav_idx[op_pos[k]]).days) if op_pos[k] >= 0 else None)
                days[wl] = rows
                sipopen.setdefault(bkey, {})[wl] = orows
            sipgain[bkey] = {"all": g_all, "vr": g_vr, "bench": g_bench,
                             "wopen": {"all": w_all, "vr": w_vr}}
            sipdays[bkey] = days
        out["sipgain"] = sipgain
        out["sipdays"] = sipdays
        out["sipopen"] = sipopen        # age in days of the window's opening valuation

    if verbose:
        _compact = lambda o: len(json.dumps(o, separators=(",", ":"), default=str))
        rows = [("lump.all", lump["all"]), ("lump.vr", lump["vr"]), ("bench", bench)]
        if not aum_subset:
            rows.append(("bench_aum", bench_aum))
        if scope == "full":
            rows += [(f"sipgain.{k}", out["sipgain"][k]) for k in SIP_BOOKS]
            rows += [("sipdays", out["sipdays"]), ("sipopen", out["sipopen"])]
        print(f"      returns: {len(lump['all'])} all-peer schemes, {len(lump['vr'])} VR "
              f"memberships, {len(bench)} benchmarks, scope={scope}, "
              f"houses={'top15 only' if houses else 'ALL schemes'}, "
              f"dp lump/bench={LUMP_DP} sipgain={SIPGAIN_DP}; "
              f"SIP as-of carry {n_carry} scheme-windows")
        print("      returns size (compact JSON bytes):")
        for name, obj in rows:
            n = _compact(obj)
            print(f"        {name:<16s} {n:>12,d}   {n / 1e6:7.2f} MB")
        tot = _compact(out)
        print(f"        {'TOTAL (block)':<16s} {tot:>12,d}   {tot / 1e6:7.2f} MB")
    return out


def _xirr_from_offsets(offsets, gain_pct, w_open=0.0, off_open=None, lo=-0.99, hi=10.0):
    """CLIENT-STYLE XIRR from the SHIPPED data only: given the day offsets and the gain %,
    solve `sum_i (1+r)^(off_i/365) = m * (1 + gain%/100)` (both sides divided by the Rs 100
    installment) by plain bisection on the monotone bracket (lo, hi]. Every exponent is >= 0
    so the left side rises strictly with r and bisection cannot miss the root.

    Deliberately NOT `sip_xirr`: this is the INDEPENDENT re-solve the round-trip check
    compares against, and sharing the engine's solver would prove nothing."""
    e = np.asarray(offsets, dtype=float) / 365.0
    m = len(e)
    if m == 0 or gain_pct is None:
        return None
    # ── the investor-book generalisation (2026-08-13) ──────────────────────────────────────
    # Normalise so the whole position cost 1 rupee. `w_open` of that rupee was already invested
    # when the window opened (age `off_open` days); the rest arrived as m equal instalments at
    # the ages in `offsets`. Solve
    #     w*(1+r)^(off_open/365) + ((1-w)/m) * sum_i (1+r)^(e_i)  =  1 + gain/100
    # w = 0 collapses to the level-SIP equation this function has always solved, so the deck's
    # pre-book behaviour is the w == 0 case rather than a separate code path. This is exactly
    # the equation the BROWSER implements — writing it here is what lets the round-trip check
    # prove the browser can rebuild the engine's rate from the shipped numbers alone.
    w = 0.0 if w_open is None else float(w_open)
    if not np.isfinite(w) or w < 0.0 or w >= 1.0:
        w = 0.0
    if w > 0.0 and (off_open is None or not np.isfinite(off_open) or off_open < 0):
        return None                     # an opening lump with no date cannot be priced
    eo = 0.0 if off_open is None else float(off_open) / 365.0
    V = 1.0 + float(gain_pct) / 100.0
    if V <= 0 or max(e.max(), eo if w > 0 else 0.0) <= 0:
        return None
    coef = (1.0 - w) / m

    def f(r):
        tot = coef * float(np.power(1.0 + r, e).sum())
        if w > 0.0:
            tot += w * float(np.power(1.0 + r, eo))
        return tot - V

    a, b = lo + 1e-9, hi
    if f(a) >= 0:
        return a
    if f(b) <= 0:
        return b
    for _ in range(80):
        mid = 0.5 * (a + b)
        if f(mid) < 0:
            a = mid
        else:
            b = mid
        if b - a < 1e-13:
            break
    return 0.5 * (a + b)


def returns_xirr_roundtrip(returns, sip, months_axis, aum_axis, tol=0.01, verbose=True):
    """PROOF that the client can rebuild the SIP XIRR from the shipped block alone
    (STANDING_SIP_DESIGN_ADDENDUM.md section C: "verify the client's solve reproduces the
    engine's `sipret` XIRR at the latest as-of to 0.01 for >=20 schemes and PRINT the check").

    Method — exactly the client's path, nothing else: take `sipgain[conv]["all"][scheme][win]`
    at the LAST all-peer as-of, take `sipdays[conv][win]` at the position of that same DATE on
    the `months` axis (the date-string mapping the client does), re-solve the XIRR by
    bisection, and compare against `sip[conv]["sipret"][scheme][win][1]` — the engine's
    Newton/bisection answer computed from the raw NAVs.

    ★ THE 0.01pp BAR NOW APPLIES TO ALL EIGHT WINDOWS (ratification R2, 2026-08-11). The
    first build could only judge the long windows, because the gain shipped at 2dp and an
    XIRR AMPLIFIES that rounding by ~1/e, where `e` is the money-weighted average holding
    period in years: a 5Y SIP has e ~ 2.5y (band ~0.003pp, comfortably inside 0.01) but a
    1-Month SIP at a MID-MONTH as-of can have its single installment four days old (as-of
    2026-08-07, first-of-month grid date 2026-08-03), e ~ 0.011y, which opened a ~0.5pp band
    — no implementation could beat it, so the bar was applied only where it was reachable.
    Shipping the gain at 4dp (SIPGAIN_DP) divides that by 100 and the bar becomes reachable
    everywhere, which is exactly why R2 raised the precision.

    TWO MEASUREMENTS, because they answer different questions:
      1. INSIDE-THE-BAND (the defect detector, exact and impossible to false-alarm): the
         engine's XIRR must lie inside the band the shipped data can physically express —
         re-solve at gain +- half a 4dp cell and widen by half a cell of `sipret`'s own
         stored precision. Anything outside that is a real disagreement, not rounding.
      2. THE CONTRACT'S 0.01pp BAR, on EVERY window. It is measured on the difference NET OF
         THE REFERENCE'S OWN STORAGE (`max(0, |diff| - 0.5*10^-SIPRET_DP)`), because the
         reference `sipret` xirr is itself stored to SIPRET_DP=2 decimals: demanding 0.01pp
         agreement against a number that is only KNOWN to +-0.005pp would be testing the
         storage format, not the client's solve. The raw max is printed beside it, so the
         subtraction hides nothing.

    ★ RESIDUAL, stated because it is arithmetic and not a bug: the amplification factor for a
    single four-day installment is ~91*(1+g/100)^90 points of annual rate per point of gain,
    so a fund that moved ~1% in those four days still carries ~0.011pp of irreducible XIRR
    uncertainty at 4dp (a >2.7% four-day move is clamped by the solver's own 10.0 upper bound
    on BOTH sides, so it agrees exactly again). If the 1-Month row of the per-window table
    ever breaches the bar on real data, the fix is one constant — SIPGAIN_DP = 6, which costs
    two bytes per sipgain value and nothing else. This function only REPORTS; it never stops
    a build, so a breach is information, not an outage.

    Consequence for the client (already ratified as R3): on windows < 1Y show the TOTAL GAIN
    as the headline with the XIRR as subtext — annualizing a four-day holding period into a
    'p.a.' number is arithmetic theatre whatever its precision."""
    log = print if verbose else (lambda *a, **k: None)
    if not returns or not sip or "sipgain" not in returns or "sipdays" not in returns:
        log("      XIRR round-trip: SKIPPED (needs the sip block and --returns-scope full)")
        return None
    m_list, a_list = list(months_axis), list(aum_axis)
    if not m_list or not a_list or a_list[-1] not in set(m_list):
        log("      XIRR round-trip: SKIPPED (the latest all-peer as-of is not on the months axis)")
        return None
    i_all = len(a_list) - 1
    i_m = m_list.index(a_list[-1])
    # half-cells of the two stored precisions — the only two sources of unavoidable spread
    gain_half = 0.5 * 10.0 ** (-SIPGAIN_DP)     # the gain the client re-solves from (4dp)
    open_half = 0.5 * 10.0 ** (-SIPOPEN_DP)     # the opening weight, the solve's other input (4dp)
    ref_half = 0.5 * 10.0 ** (-SIPRET_DP)       # sipret's stored xirr, the reference (2dp)

    diffs, nets, bands, schemes, worst, outside = {}, {}, {}, set(), (0.0, None), []
    for bkey in SIP_BOOKS:
        gains = returns["sipgain"].get(bkey, {}).get("all", {})
        wopens = returns["sipgain"].get(bkey, {}).get("wopen", {}).get("all", {})
        offs_by_win = returns["sipdays"].get(bkey, {})
        open_by_win = (returns.get("sipopen") or {}).get(bkey, {})
        sipret = sip.get(bkey, {}).get("sipret", {})
        for sch, bywin in gains.items():
            ref = sipret.get(sch)
            if not ref:
                continue
            for wl, rec in bywin.items():
                j = i_all - rec["f"]
                g = rec["v"][j] if 0 <= j < len(rec["v"]) else None
                rows_d = offs_by_win.get(wl)
                offs = rows_d[i_m] if (rows_d and 0 <= i_m < len(rows_d)) else None
                eng = (ref.get(wl) or [None, None])[1]
                # the opening lump, read the same way the client reads it: the scheme's own
                # opening WEIGHT out of the sipgain block, and the window's opening AGE (one
                # number for every scheme, since the window opens on a date, not per fund)
                wrec = wopens.get(sch, {}).get(wl)
                jw = (i_all - wrec["f"]) if wrec else None
                w = (wrec["v"][jw] if (wrec and 0 <= jw < len(wrec["v"])) else None) or 0.0
                orow = open_by_win.get(wl)
                off_open = orow[i_m] if (orow and 0 <= i_m < len(orow)) else None
                if g is None or not offs or eng is None:
                    continue
                r = _xirr_from_offsets(offs, g, w, off_open)
                if r is None:
                    continue
                # The band the SHIPPED precision can express. ★ 2026-08-13: the solve now
                # depends on TWO rounded numbers, not one — the gain (4dp) and the opening
                # weight (4dp) — so the band is the corner of both cells, not just the gain's.
                # Measured before changing it: with the gain alone, 81 of 8,902 scheme-windows
                # fell outside a band that simply did not include a source of spread that is
                # really there; allowing the opening weight's own half-cell takes that to 0 and
                # absorbs nothing larger. This is completed accounting, not a loosened bar — the
                # contract check below is untouched and already passed at 0.0016 of its 0.01 pp.
                edges = [_xirr_from_offsets(offs, g + dg, max(0.0, w + dw), off_open)
                         for dg in (-gain_half, gain_half)
                         for dw in ((-open_half, open_half) if w else (0.0,))]
                band = ref_half + 100.0 * max(abs((e if e is not None else r) - r) for e in edges)
                diff = abs(r * 100.0 - float(eng))
                # the part of the disagreement the REFERENCE's 2dp storage cannot explain —
                # this is what the contract's 0.01pp bar is measured on (see the docstring)
                net = max(0.0, diff - ref_half)
                diffs.setdefault(wl, []).append(diff)
                nets.setdefault(wl, []).append(net)
                bands.setdefault(wl, []).append(band)
                schemes.add(sch)
                if diff > band + 1e-9:
                    outside.append((bkey, sch, wl, round(diff, 4), round(band, 4)))
                if diff > worst[0]:
                    worst = (diff, (bkey, sch, wl, round(r * 100.0, 4), eng))
    if not diffs:
        log("      XIRR round-trip: NO comparable scheme-windows found")
        return None
    order = [WIN_LABEL[m] for m in RES_WINS if WIN_LABEL[m] in diffs]
    n_cmp = sum(len(v) for v in diffs.values())
    judged = list(order)                    # ★ R2: the bar now applies to EVERY window
    bar_max = max((max(nets[w]) for w in judged), default=0.0)
    breached = [w for w in judged if max(nets[w]) > tol]
    log(f"      XIRR round-trip: {len(schemes)} schemes, max abs diff {worst[0]:.4f} pp "
        f"(sipgain at {SIPGAIN_DP}dp, reference sipret at {SIPRET_DP}dp)")
    log(f"        {n_cmp} scheme-windows re-solved from sipgain + sipdays alone; "
        f"{n_cmp - len(outside)}/{n_cmp} inside the shipped precision band -> "
        f"{'PASS' if not outside else 'FAIL'}" + (f"; e.g. {outside[:3]}" if outside else ""))
    log("        per window (max diff / max net-of-reference / max band, pp): "
        + " | ".join(f"{w} {max(diffs[w]):.4f}/{max(nets[w]):.4f}/{max(bands[w]):.4f}"
                     for w in order))
    log(f"        contract bar {tol} pp on ALL {len(judged)} windows -> max {bar_max:.4f} -> "
        f"{'PASS' if not breached else 'FAIL on ' + ', '.join(breached)}")
    return {"n_schemes": len(schemes), "n_compared": n_cmp, "max_abs_diff": worst[0],
            "n_outside_band": len(outside), "judged_windows": judged, "max_on_judged": bar_max,
            "breached_windows": breached,
            "per_window": {w: max(diffs[w]) for w in order},
            "per_window_net": {w: max(nets[w]) for w in order},
            "per_window_band": {w: max(bands[w]) for w in order}, "worst": worst[1]}


# %% data cleaning  (BUILD_SPEC §2)
def find_nontrading_days(nav: pd.DataFrame, frac: float = REPEAT_FRAC):
    """Dates where >= `frac` of POPULATED series equal the prior day's value exactly.
    Returns (drop_index, frac_same_series) for auditing. NaN==NaN is False, so gaps
    don't count as 'repeats'."""
    nav = nav.sort_index()
    eq_prev = (nav == nav.shift(1))
    pop = nav.notna().sum(axis=1)
    frac_same = eq_prev.sum(axis=1) / pop.replace(0, np.nan)
    drop_mask = (frac_same >= frac).fillna(False)
    if len(drop_mask):
        drop_mask.iloc[0] = False  # first row has no prior day
    return nav.index[drop_mask], frac_same


def drop_nontrading_days(nav: pd.DataFrame, frac: float = REPEAT_FRAC):
    drop_idx, _ = find_nontrading_days(nav, frac)
    return nav.drop(index=drop_idx), list(drop_idx)


def _careful_ffill_series(s: pd.Series, max_gap: int):
    """Fill internal NaN runs shorter than max_gap; never fill before first or after
    last real value (not-yet-launched / discontinued -> leave NaN)."""
    valid = s.notna()
    if not valid.any():
        return s, 0
    first = valid.idxmax()
    last = valid[::-1].idxmax()
    s = s.copy()
    seg = s.loc[first:last]
    na = seg.isna()
    runs = (na != na.shift()).cumsum()
    ff = seg.ffill()
    filled = 0
    for _, idx in seg.groupby(runs).groups.items():
        blk = seg.loc[idx]
        if blk.isna().all() and len(blk) < max_gap:
            seg.loc[idx] = ff.loc[idx]
            filled += len(blk)
    s.loc[first:last] = seg
    return s, filled


def careful_ffill(nav: pd.DataFrame, max_gap: int = MAX_FILL_GAP):
    """Apply careful ffill per column. Returns (filled_nav, audit_dict)."""
    out = {}
    audit = {"gaps_filled_cells": 0, "series_with_fills": 0, "discontinued": []}
    global_last = nav.sort_index().index.max()
    for col in nav.columns:
        s2, filled = _careful_ffill_series(nav[col], max_gap)
        out[col] = s2
        if filled:
            audit["gaps_filled_cells"] += filled
            audit["series_with_fills"] += 1
        lastv = nav[col].dropna()
        if len(lastv) and lastv.index.max() < global_last:
            audit["discontinued"].append((col, lastv.index.max()))
    return pd.DataFrame(out), audit


def clean_nav(nav: pd.DataFrame, frac=REPEAT_FRAC, max_gap=MAX_FILL_GAP, label="", verbose=True):
    """Full clean: drop non-trading days, then careful ffill. Returns (clean_nav, report)."""
    nav = nav.sort_index()
    clean, dropped = drop_nontrading_days(nav, frac)
    filled_nav, audit = careful_ffill(clean, max_gap)
    rep = {"label": label, "n_dropped_nontrading": len(dropped),
           "dropped_sample": [d.date().isoformat() for d in dropped[-10:]],
           **audit}
    if verbose:
        print(f"  [{label}] dropped {len(dropped)} non-trading dates (frac>={frac}); "
              f"filled {audit['gaps_filled_cells']} gap-cells across {audit['series_with_fills']} series; "
              f"{len(audit['discontinued'])} discontinued series (last real < {global_last_str(nav)})")
    return filled_nav, rep


def global_last_str(nav):
    return nav.sort_index().index.max().date().isoformat()


# %% [markdown]
# # 9. FULL PIPELINE — one command produces both Excel files + the dashboard JSON
# `python peer_monitor.py` (or the generated generate_report.py). Runs offline.

# %%
def run(data_dir="Data", out_dir="out", w_1y=0.8, min_peers=1, repeat_frac=REPEAT_FRAC,
        max_fill_gap=MAX_FILL_GAP, verbose=True, with_standing=True, with_sip=True,
        with_returns=True, returns_scope="full"):
    """End-to-end monthly run. Returns a dict of all intermediate frames (for audit/notebook).
    with_standing / with_sip (STANDING_SIP_DESIGN.md, 2026-08-11) gate the two additive
    blocks; when False the corresponding JSON keys are absent and the rest is unchanged.
    with_returns / returns_scope (its ADDENDUM section C) gate the third: 'full' ships
    lump+bench+sipgain+sipdays, 'lumpsum' ships only lump+bench (the size escape hatch)."""
    global DATA, SCHEME_DIR, MAP_DIR, BENCH_DIR
    DATA = Path(data_dir); SCHEME_DIR = DATA / "Scheme NAV and AUM"
    MAP_DIR = DATA / "Mapping"; BENCH_DIR = DATA / "Benchmark NAV"
    out = Path(out_dir); out.mkdir(parents=True, exist_ok=True)
    log = print if verbose else (lambda *a, **k: None)

    log("[1/8] Load + clean scheme NAV (calendar basis); AMFI-code category join")
    nav = clean_nav(load_scheme_nav(), frac=repeat_frac, max_gap=max_fill_gap, verbose=verbose)[0]
    code2name = scheme_code_map()
    raw_map = pd.read_excel(MAP_DIR / "Map MFI Scheme to Category.xlsx", engine=ENGINE)
    cmap, cats = build_category_map(raw_map, code2name)
    # The VR peerset is ABSOLUTE (KV 2026-07-03): align it against the FULL NAV first, so
    # All-Peers exclusions (MULTI_ASSET_DROP, unmapped schemes) can never drop a VR peer
    # from the universe. All-Peers categorisation itself is untouched (cmap unchanged).
    cme, exact_bench, vr_dropped = load_vr_mapping()
    cme, vr_unmatched = align_vr_to_nav(cme, nav.columns, code2name)
    vr_names = set(cme.index)
    nav = nav[[s for s in nav.columns if s in cmap.index or s in vr_names]]
    cmap = cmap.loc[[s for s in cmap.index if s in nav.columns]]
    log(f"      universe: {nav.shape[1]} schemes, {len(cats)} all-peer categories, "
        f"to {nav.index.max().date()}")

    log("[2/8] Daily AUM share")
    aum = load_aum()
    aum_daily, pct = percent_aum_share(aum, nav.index)
    pct = pct[[c for c in pct.columns if c in nav.columns]]

    log("[3/8] Benchmarks (carry-forward align) + VR peer map (AMFI-code join)")
    bench_nav, bench_report = load_bench_nav()
    bench_nav = clean_nav(bench_nav, frac=repeat_frac, max_gap=max_fill_gap, verbose=False)[0]
    bench_cal, bench_last = align_benchmarks(bench_nav, nav.index)
    log(f"      VR peers recovered by code: {cme.attrs.get('recovered_by_match', 0)}; "
        f"dropped (no MFI NAV): {len(vr_dropped)}")

    log("[4/8] All-peer (MFI) quartiles -> Sleeve/AMC %AUM")
    qy1, qy3 = all_peer_quartiles(nav, cmap, cats)
    # top15=None -> EVERY fund house. The dashboard's Top-15 / All-AMCs toggle needs the wide
    # table; which houses a VIEW surfaces is now the client's choice, not the engine's.
    sleeve, amc = sleeve_amc_tables(qy1, qy3, pct, cmap, top15=None)

    log("[5/8] Exact-peer (VR) scoring: 5/4/3/2 & 4/3/2/1, +reward, composite, re-bucket")
    res = exact_peer_scoring(nav, bench_cal, cme, exact_bench, w_1y=w_1y,
                             min_peers=min_peers, bench_last=bench_last)

    log("[5b/8] Extra rolling windows 1M/3M/6M/9M/2Y/5Y (raw quartiles, both universes)")
    qy_all_w = {12: qy1, 36: qy3}                  # 1Y/3Y daily quartiles already computed
    qy_vr_w  = {12: res["qy_1y"], 36: res["qy_3y"]}
    for m in MONTH_WINS:
        qy_all_w[m] = all_peer_quartiles_m(nav, cmap, cats, m)
        qy_vr_w[m]  = vr_quartiles_m(nav, cme, m, min_peers=min_peers)

    # ── ADDITIVE steps (STANDING_SIP_DESIGN.md 2026-08-11). Both need the month axes the
    # dashboard JSON uses, derived EXACTLY the way write_dashboard_json derives them
    # internally (months = month-ends of the composite's index; aum axis = the sleeve
    # table's month-end date columns) so ranks/SIP series land on the same grid as the
    # quartile digit series. If you change one derivation, change the other.
    months_axis = month_end_dates(res["composite"].index)
    aum_axis = [c for c in _sample(sleeve, "monthly", axis=1).columns if isinstance(c, pd.Timestamp)]
    standing = None
    if with_standing:
        log("[5c/8] Category standing: per-date 1-based ranks (same sort as the quartiles), both universes, 8 windows")
        standing = build_standing(nav, cmap, cats, cme, months_axis, aum_axis, min_peers=min_peers)
    else:
        log("[5c/8] Category standing: SKIPPED (--no-standing)")
    sip = None
    if with_sip:
        log("[5d/8] SIP mode: Rs100/month first/last-day grids -> value-ratio quartiles, ranks, "
            "residency (ALL schemes), standing, XIRR")
        sip = build_sip(nav, bench_cal, cmap, cats, cme, exact_bench, bench_last,
                        months_axis, aum_axis, min_peers=min_peers, verbose=verbose)
    else:
        log("[5d/8] SIP mode: SKIPPED (--no-sip)")
    returns = None
    if with_returns:
        log("[5e/8] Return profile: cumulative lumpsum + category-benchmark returns (ALL schemes)"
            + (f" + SIP gain ({SIPGAIN_DP}dp) and installment day-offsets"
               if returns_scope == "full" else ""))
        returns = build_returns(nav, bench_cal, cmap, cats, cme, exact_bench,
                                aum_axis, months_axis, bench_last=bench_last,
                                min_peers=min_peers, scope=returns_scope, verbose=verbose)
        # The client solves the SIP XIRR from `sipgain` + `sipdays` alone; this proves it can,
        # by re-solving with an independent bisection and comparing to the engine's `sipret`.
        returns_xirr_roundtrip(returns, sip, months_axis, aum_axis, verbose=verbose)
    else:
        log("[5e/8] Return profile: SKIPPED (--no-returns)")

    log("[6/8] Month-end sampling + VR composite %AUM")
    qy_me = month_end_asof(res["qy_score"])
    vr = vr_amc_table(qy_me, pct, cmap)
    peer_map = cme[["Scheme", "AMFI Code", "Category"]].copy()
    peer_map["Benchmark"] = peer_map["Category"].map(exact_bench)

    log("[7/8] Exclusions log")
    exclusions = [{"item": s, "type": "VR scheme", "reason": "no matching MFI NAV (AMFI code absent)"}
                  for s in vr_dropped]
    sl = nav.index.max()
    # Both benchmark defects — STALE (composite capped after a date) and EMPTY (no NAV at all,
    # so the +1 bonus can never be earned) — come from one generic scan, so a benchmark that
    # goes empty next month discloses itself instead of staying invisible. See
    # `benchmark_exclusions` for why the two cases must read differently on the tab.
    exclusions += benchmark_exclusions(cme, exact_bench, bench_cal, bench_last, sl)
    exclusions += list(cmap.attrs.get("exclusions", []))

    log("[8/8] Write Excel + dashboard JSON")
    asof = sl.strftime("%B %d, %Y")
    f1 = out / f"percent AUM in Q1 to Q4 - {asof}.xlsx"
    f2 = out / f"Scheme Scoring on Exact Peer Set - Calendar 1Y 3Y - {asof}.xlsx"
    write_aum_workbook(f1, _sample(sleeve, "monthly", 1), _sample(amc, "monthly", 1), vr, resolution="daily")
    write_scoring_workbook(f2, month_end_asof(res["s1"]), month_end_asof(res["s3_raw"]),
                           month_end_asof(res["beat"]), month_end_asof(res["composite"]),
                           qy_me, peer_map, resolution="daily", w_1y=w_1y)
    jp, sz = write_dashboard_json(out / "dashboard_data.json", res, pct, cmap, amc, sleeve,
                                  exclusions, peer_map, exact_bench, bench_last,
                                  aum_daily=aum_daily, w_1y=w_1y, qy1=qy1, qy3=qy3,
                                  qy_all_w=qy_all_w, qy_vr_w=qy_vr_w,
                                  standing=standing, sip=sip, returns=returns)
    write_dashboards(out)  # dashboard.html (+ offline) from embedded template, if available
    # The OFFLINE deck size is a standing reporting requirement (the deck inlines the whole
    # JSON, so it is the payload's real-world cost — email-ability is KV's call).
    _deck = out / "dashboard.html"
    _dsz = (_deck.stat().st_size / 1e6) if _deck.exists() else 0.0
    log(f"\nDone. Outputs in {out.resolve()}:\n  {f1.name}\n  {f2.name}\n  "
        f"dashboard_data.json ({sz:.2f} MB) + dashboard.html ({_dsz:.2f} MB offline deck)")
    return dict(nav=nav, cmap=cmap, cats=cats, pct=pct, bench_cal=bench_cal, bench_last=bench_last,
                cme=cme, exact_bench=exact_bench, res=res, sleeve=sleeve, amc=amc, vr=vr,
                peer_map=peer_map, exclusions=exclusions, standing=standing, sip=sip,
                returns=returns)


def write_dashboards(out: Path, template=None):
    """Build the SELF-CONTAINED OFFLINE dashboard by inlining this run's dashboard_data.json
    into the template's `<script id="peer-data">` placeholder (Chart.js is already inlined in
    the template). Writes out/dashboard.html and out/dashboard_offline.html (identical, fully
    offline — double-clickable, no server, no internet). The template (with a __PEER_DATA__
    placeholder) is the auditable HTML at the project root."""
    cands = ([Path(template)] if template else []) + [Path("dashboard.html").resolve()]
    p = out.resolve()
    for _ in range(6):
        p = p.parent
        cands.append(p / "dashboard.html")
    tmpl_path = next((c for c in cands if c.exists()
                      and "__PEER_DATA__" in c.read_text(encoding="utf-8", errors="ignore")), None)
    if tmpl_path is None:
        print("  (no offline dashboard template with __PEER_DATA__ placeholder found; skipped)")
        return
    tmpl = tmpl_path.read_text(encoding="utf-8")
    data = (out / "dashboard_data.json").read_text(encoding="utf-8").replace("</", "<\\/")
    html = tmpl.replace("__PEER_DATA__", data, 1)
    (out / "dashboard.html").write_text(html, encoding="utf-8")
    (out / "dashboard_offline.html").write_text(html, encoding="utf-8")


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="Peer Performance Monitor — monthly run")
    ap.add_argument("--data", default="Data")
    ap.add_argument("--out", default="out")
    ap.add_argument("--w1y", type=float, default=0.8, help="default 1Y weight in composite (0..1)")
    ap.add_argument("--min-peers", type=int, default=1, help="drop VR categories with fewer peers")
    ap.add_argument("--no-standing", action="store_true",
                    help="skip the category-standing rank block (STANDING_SIP_DESIGN.md Feature A)")
    ap.add_argument("--no-sip", action="store_true",
                    help="skip the SIP investment-mode block (STANDING_SIP_DESIGN.md Feature B)")
    ap.add_argument("--no-returns", action="store_true",
                    help="skip the return-profile block (STANDING_SIP_DESIGN_ADDENDUM.md section C)")
    ap.add_argument("--returns-scope", choices=["full", "lumpsum"], default="full",
                    help="'full' = lump+bench+sipgain+sipdays; 'lumpsum' = lump+bench only "
                         "(the size escape hatch)")
    args = ap.parse_args()
    run(args.data, args.out, w_1y=args.w1y, min_peers=args.min_peers,
        with_standing=not args.no_standing, with_sip=not args.no_sip,
        with_returns=not args.no_returns, returns_scope=args.returns_scope)
