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
    cme["fund house"] = cme.index.to_series().str.split(" ").str[0]

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

    cm["fund house"] = cm.index.to_series().str.split(" ").str[0]
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
    aum1.insert(0, "fund house", aum1.index.to_series().str.split().str[0])
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


# ── Short rolling windows (1/3/6/9 months) — RAW quartiles, no composite/reward ──
# KV (2026-06-24): in ADDITION to the 1Y/3Y windows, evaluate every scheme on 1M/3M/6M/9M
# rolling windows so the whole deck (sleeve/AMC/league/scheme) can be viewed on any window.
# Returns are CUMULATIVE point-to-point (NAV_t/NAV_{t_m}-1), NOT annualized — a sub-year CAGR
# would be misleading; quartile RANKS are identical either way (annualizing is monotonic).
MONTH_WINS = [1, 3, 6, 9]                                   # the NEW sub-year windows
RES_WINS   = [1, 3, 6, 9, 12, 36]                           # all windows (incl. 1Y/3Y) for residency
WIN_LABEL  = {1: "1 Month", 3: "3 Month", 6: "6 Month", 9: "9 Month", 12: "1 Year", 36: "3 Year"}


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
    the FIRST category for any scheme that sits in >1 VR category (matches the dashboard's
    `vr.find(v=>v.s===name)` first-match convention)."""
    seen, cols, names = set(), [], []
    for (c, sch) in qy_mi.columns:
        if sch in seen:
            continue
        seen.add(sch); cols.append((c, sch)); names.append(sch)
    sub = qy_mi.loc[:, cols].copy()
    sub.columns = names
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
def _quartile_block(sub: pd.DataFrame) -> pd.DataFrame:
    """Per date: rank a category's schemes by return desc -> q1..q4 (round-up)."""
    dates = sub.index[sub.notna().any(axis=1)]
    lab = pd.DataFrame(index=dates, columns=sub.columns, dtype=object)
    for d in dates:
        a = sub.loc[d].dropna()
        if len(a):
            for l, names in zip(["q1", "q2", "q3", "q4"],
                                quartiles_roundup(_sort_desc_stable(a))):
                lab.loc[d, names] = l
    return lab


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
    (sleeve_df, amc_df) with dates in COLUMNS."""
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
    aum_rows = aum_rows.loc[aum_rows["fund house"].isin(top15)]
    parts = []
    for rw in ("1 Year", "3 Year"):
        a = aum_rows.copy(); a["Quartile"] = "% of AMC AUM"; a["Rolling Window"] = rw
        parts.append(a)
    aum_tot = pd.concat(parts).groupby(["Rolling Window", "fund house", "Breakdown", "Quartile"]).sum(numeric_only=True)[dcols]

    sleeve = pd.concat([visual, aum_tot]).sort_index(level=["Rolling Window", "fund house", "Breakdown", "Quartile"])
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


# %% VR-peer %AUM in Q1..Q4 (composite-quartile based)  (BUILD_SPEC §4h; nb 1689-1712)
def vr_amc_table(qy_score, pct_aum, cmap, top15=TOP15, repeated=REPEATED_PEER_CATS):
    """%AUM in Q1..Q4 by fund house, from the COMPOSITE-score quartiles (qy_score, a
    MultiIndex(cat,scheme) frame). Drops 'repeated peer-set' categories. dates in COLUMNS.
    Also returns aum-weighted score-percentile by fund house."""
    qs = qy_score.loc[:, ~qy_score.columns.get_level_values(0).isin(repeated)]
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
                         qy1=None, qy3=None, qy_all_w=None, qy_vr_w=None):
    """Emit dashboard_data.json. Per VR scheme × month-end it ships the score COMPONENTS and
    AUM so the dashboard recomputes ANY basis live (1Y / 3Y±reward / Composite with weight):
        per scheme: q1y,q3y (return quartiles 1..4), y(=1Y score 5..2), t(=3Y score 4..1),
                    b(beats-benchmark 1/0/null), a(% of its AMC AUM), cr(absolute AUM, Rs cr),
                    sl(sleeve/Breakdown).
    Plus AMC and AMC×sleeve total AUM (Rs cr) so the deck shows %-of-AMC, %-of-sleeve, Rs cr.
    Also ships all-peer quartile %AUM, exclusions, peer/benchmark mapping."""
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
            "c": c, "s": sch, "h": str(sch).split()[0], "sl": sleeve_of.get(sch, "Other"),
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

    # AMC-total and AMC x sleeve-total AUM (Rs cr, month-end) for the top-15 houses
    amc_total, amc_sleeve_total = {}, {}
    if aum_daily is not None:
        hcol = pd.Series({s: str(s).split()[0] for s in crm.columns})
        scol = pd.Series({s: sleeve_of.get(s, "Other") for s in crm.columns})
        for h in TOP15:
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
            hh = str(s).split()[0]
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

    # ── QUARTILE RESIDENCY: for every Top-15-house scheme, the count of trading days spent in
    # Q1/Q2/Q3/Q4 over the trailing window ENDING AT each as-of month-end (so it follows the
    # Scheme-Detail as-of date selector). Both universes; windows 1M/3M/6M/9M/1Y/3Y. Aligned to
    # aum_dates (all-peer) / months (VR). Only Top-15 schemes (the scheme picker only lists them).
    top15set = set(TOP15)
    residency = {"all": {}, "vr": {}}
    if qy_all_w:
        for m in RES_WINS:
            cols = [s for s in qy_all_w[m].columns if str(s).split()[0] in top15set]
            for sch, rec in window_residency(qy_all_w[m][cols], list(aum_dcols), m).items():
                residency["all"].setdefault(sch, {})[WIN_LABEL[m]] = rec
    if qy_vr_w:
        for m in RES_WINS:
            daily = _vr_daily_by_scheme(qy_vr_w[m])
            cols = [s for s in daily.columns if str(s).split()[0] in top15set]
            for sch, rec in window_residency(daily[cols], list(months), m).items():
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
    payload = {
        "meta": {"latest": mstr[-1] if mstr else None,
                 "latest_aum_date": aum_dates[-1] if aum_dates else None,
                 "w_default": int(w_1y * 100),
                 "n_vr_schemes": len(vr), "n_vr_categories": n_cats_vr,
                 "n_allpeer_schemes": len(allpeer), "top15": TOP15,
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
    with open(path, "w", encoding="utf-8") as f:
        # allow_nan=False => emit STRICT/valid JSON (NaN/Infinity are not valid JSON and
        # break browsers' JSON.parse); raises if any stray NaN slipped through, so we catch it.
        json.dump(payload, f, separators=(",", ":"), default=str, allow_nan=False)
    return path, path.stat().st_size / 1e6


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
        max_fill_gap=MAX_FILL_GAP, verbose=True):
    """End-to-end monthly run. Returns a dict of all intermediate frames (for audit/notebook)."""
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
    nav = nav[[s for s in nav.columns if s in cmap.index]]
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
    cme, exact_bench, vr_dropped = load_vr_mapping()
    cme, vr_unmatched = align_vr_to_nav(cme, nav.columns, code2name)
    log(f"      VR peers recovered by code: {cme.attrs.get('recovered_by_match', 0)}; "
        f"dropped (no MFI NAV): {len(vr_dropped)}")

    log("[4/8] All-peer (MFI) quartiles -> Sleeve/AMC %AUM")
    qy1, qy3 = all_peer_quartiles(nav, cmap, cats)
    sleeve, amc = sleeve_amc_tables(qy1, qy3, pct, cmap)

    log("[5/8] Exact-peer (VR) scoring: 5/4/3/2 & 4/3/2/1, +reward, composite, re-bucket")
    res = exact_peer_scoring(nav, bench_cal, cme, exact_bench, w_1y=w_1y,
                             min_peers=min_peers, bench_last=bench_last)

    log("[5b/8] Short rolling windows 1M/3M/6M/9M (raw quartiles, both universes)")
    qy_all_w = {12: qy1, 36: qy3}                  # 1Y/3Y daily quartiles already computed
    qy_vr_w  = {12: res["qy_1y"], 36: res["qy_3y"]}
    for m in MONTH_WINS:
        qy_all_w[m] = all_peer_quartiles_m(nav, cmap, cats, m)
        qy_vr_w[m]  = vr_quartiles_m(nav, cme, m, min_peers=min_peers)

    log("[6/8] Month-end sampling + VR composite %AUM")
    qy_me = month_end_asof(res["qy_score"])
    vr = vr_amc_table(qy_me, pct, cmap)
    peer_map = cme[["Scheme", "AMFI Code", "Category"]].copy()
    peer_map["Benchmark"] = peer_map["Category"].map(exact_bench)

    log("[7/8] Exclusions log")
    exclusions = [{"item": s, "type": "VR scheme", "reason": "no matching MFI NAV (AMFI code absent)"}
                  for s in vr_dropped]
    sl = nav.index.max()
    for c in sorted(set(cme["Category"])):
        lr = bench_last.get(exact_bench.get(c))
        if lr is not None and pd.notna(lr) and lr < sl - pd.Timedelta(days=STALE_BENCH_DAYS):
            exclusions.append({"item": c, "type": "category (VR composite cap)",
                               "reason": f"benchmark stale (last {lr.date()}); composite capped >{STALE_BENCH_DAYS}d"})
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
                                  qy_all_w=qy_all_w, qy_vr_w=qy_vr_w)
    write_dashboards(out)  # dashboard.html (+ offline) from embedded template, if available
    log(f"\nDone. Outputs in {out.resolve()}:\n  {f1.name}\n  {f2.name}\n  dashboard_data.json ({sz:.2f} MB) + dashboard.html")
    return dict(nav=nav, cmap=cmap, cats=cats, pct=pct, bench_cal=bench_cal, bench_last=bench_last,
                cme=cme, exact_bench=exact_bench, res=res, sleeve=sleeve, amc=amc, vr=vr,
                peer_map=peer_map, exclusions=exclusions)


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
    args = ap.parse_args()
    run(args.data, args.out, w_1y=args.w1y, min_peers=args.min_peers)
