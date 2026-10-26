# -*- coding: utf-8 -*-
"""Adhoc Excel dump: daily all-peer (MFI) quartile MATRICES (1Y & 3Y) + a t1->t2
"% of days in Q1/Q2/Q3/Q4" analysis with AMC / t1 / t2 dropdowns.

Reuses the ENGINE's own functions (peer_monitor steps [1]+[4]) so the matrices are
IDENTICAL to the quartile data feeding the published deck (BUILD_SPEC 4c; the deck's
residency feature 4i uses the same frames). Denominator convention matches the deck:
n = days the scheme was RATED (q1+q2+q3+q4), not calendar days.

Run from inside the refresh folder:   python _quartile_matrix_dump.py
Output: out/Quartile Matrix and Time-in-Quartile - <asof>.xlsx
Sheets: ReadMe | Map | Q 1Y matrix | Q 3Y matrix | Fund View | All Schemes | Lists
"""
import numpy as np
import pandas as pd
from pathlib import Path
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

QMAP = {"q1": 1, "q2": 2, "q3": 3, "q4": 4}
DEFAULT_T1, DEFAULT_T2 = "2025-04-01", "2026-03-31"   # KV's worked example
DEFAULT_AMC = "Aditya"


def build_frames():
    """Engine steps [1]+[4] exactly as peer_monitor.run() does them."""
    import peer_monitor as pm
    pm.DATA = Path("Data")
    pm.SCHEME_DIR = pm.DATA / "Scheme NAV and AUM"
    pm.MAP_DIR = pm.DATA / "Mapping"
    pm.BENCH_DIR = pm.DATA / "Benchmark NAV"

    nav = pm.clean_nav(pm.load_scheme_nav(), verbose=True)[0]
    code2name = pm.scheme_code_map()
    raw_map = pd.read_excel(pm.MAP_DIR / "Map MFI Scheme to Category.xlsx", engine=pm.ENGINE)
    cmap, cats = pm.build_category_map(raw_map, code2name)
    nav = nav[[s for s in nav.columns if s in cmap.index]]
    cmap = cmap.loc[[s for s in cmap.index if s in nav.columns]]
    print(f"universe: {nav.shape[1]} schemes, {len(cats)} categories, to {nav.index.max().date()}")

    qy1, qy3 = pm.all_peer_quartiles(nav, cmap, cats)
    name2code = {v: k for k, v in code2name.items()}
    return pm, nav, cmap, qy1, qy3, name2code


def verify(pm, qy1, qy3, cmap):
    """Independent spot-check: direct slice counts vs the engine's window_residency
    (the published deck's own residency function) for t1=2025-04-01..t2=2026-03-31,
    which equals the deck's trailing-12M window ending 2026-03-31: (t2-12m, t2]."""
    t1, t2 = pd.Timestamp(DEFAULT_T1), pd.Timestamp(DEFAULT_T2)
    rows = []
    aditya = cmap.index[(cmap["fund house"] == "Aditya")]
    picks = [s for s in aditya if "Flexi" in s][:1] + [s for s in aditya if "Frontline" in s][:1]
    for basis, qy in (("1Y", qy1), ("3Y", qy3)):
        res = pm.window_residency(qy[picks], [t2], 12)
        for s in picks:
            sl = qy[s].loc[t1:t2]
            direct = [int((sl == f"q{q}").sum()) for q in (1, 2, 3, 4)]
            eng = res.get(s, {}).get("v", [None])[0]
            ok = (eng == direct)
            n = sum(direct)
            pct = [d / n * 100 if n else float("nan") for d in direct]
            rows.append((basis, s, direct, eng, ok, n, [f"{p:.1f}%" for p in pct]))
            print(f"[verify {basis}] {s}\n   direct slice {direct}  engine residency {eng}  "
                  f"MATCH={ok}  n={n}  %={[f'{p:.1f}' for p in pct]}")
    assert all(r[4] for r in rows), "direct-slice vs engine window_residency MISMATCH"
    return rows


def write_workbook(nav, cmap, qy1, qy3, name2code):
    asof = nav.index.max().strftime("%B %d, %Y")
    out = Path("out") / f"Quartile Matrix and Time-in-Quartile - {asof}.xlsx"

    # ---- order everything by (fund house, category, scheme); numeric matrices ----
    m = cmap.copy()
    m["AMFI Code"] = [name2code.get(s, "") for s in m.index]
    # only schemes the engine actually quartiles (valid >=4-house categories)
    dropped = [s for s in m.index if s not in qy1.columns]
    if dropped:
        print(f"note: {len(dropped)} schemes in Map but in no valid all-peer category "
              f"(engine excludes categories with <4 fund houses) — left out of the dump")
    m = m.loc[[s for s in m.index if s in qy1.columns and s in qy3.columns]]
    m = m.sort_index()                                              # scheme name asc ...
    m = m.sort_values(["fund house", "Category"], kind="stable")    # ... within AMC, category
    schemes = list(m.index)
    n1 = qy1[schemes].replace(QMAP).apply(pd.to_numeric, errors="coerce")
    n3 = qy3[schemes].replace(QMAP).apply(pd.to_numeric, errors="coerce")
    keep = n1.notna().any(axis=1) | n3.notna().any(axis=1)
    n1, n3 = n1.loc[keep], n3.loc[keep]
    dates = n1.index
    nrows, ncols = len(dates), len(schemes)
    print(f"matrix: {nrows} dates x {ncols} schemes  ({dates[0].date()} .. {dates[-1].date()})")

    amcs = sorted(m["fund house"].unique())
    amc_counts = m.groupby("fund house").size()
    maxn = int(amc_counts.max())

    wb = xlsxwriter.Workbook(str(out), {"constant_memory": True,
                                        "default_date_format": "yyyy-mm-dd"})
    f_hdr = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white",
                           "border": 1})
    f_date = wb.add_format({"num_format": "yyyy-mm-dd"})
    f_date_in = wb.add_format({"num_format": "yyyy-mm-dd", "bg_color": "#FFF2CC", "border": 1})
    f_in = wb.add_format({"bg_color": "#FFF2CC", "border": 1})
    f_pct = wb.add_format({"num_format": "0.0%"})
    f_int = wb.add_format({"num_format": "0"})
    f_lbl = wb.add_format({"bold": True})
    f_wrap = wb.add_format({"text_wrap": True, "valign": "top"})
    f_note = wb.add_format({"italic": True, "font_color": "#7F7F7F"})

    # ---------- ReadMe ----------
    ws = wb.add_worksheet("ReadMe")
    ws.set_column("A:A", 130)
    readme = [
        ("Quartile Matrix and Time-in-Quartile — " + asof, f_lbl),
        ("", None),
        ("WHAT (definition)", f_lbl),
        ("For every scheme and every trading day: which quartile (1=Q1 best ... 4=Q4 worst) the scheme's "
         "return sat in WITHIN ITS MFI CATEGORY (the All-Peers universe), on a rolling 1-Year and a rolling "
         "3-Year basis. Then, for any window t1..t2 you pick: the % of rated days each fund spent in each quartile.", None),
        ("", None),
        ("METHOD (reproducible without the author)", f_lbl),
        ("Source: this refresh's Data/ folder (MFI NAV files + 'Map MFI Scheme to Category.xlsx'), processed by "
         "peer_monitor.py steps [1]+[4] — the SAME code that produced the published deck.", None),
        ("Returns: exact-calendar point-to-point. 1Y = NAV_t/NAV_{t-1y} - 1 (cumulative); "
         "3Y = (NAV_t/NAV_{t-3y})^(1/3) - 1 (annualised). t-Ny = last trading day on/before the same calendar date N years earlier.", None),
        ("Quartiles: per date, per MFI category, rank schemes by return descending; split into 4 buckets with the "
         "remainder going to the TOP buckets (round-up); ties broken by original column order (stable). "
         "1 = Q1 (best 25%), 4 = Q4 (worst 25%). Blank = the scheme had no return that day (younger than the window, "
         "discontinued, or NAV missing).", None),
        ("Row granularity of the matrices: one row = one trading day; one column = one scheme.", None),
        ("% of days in Qn between t1 and t2  =  (rated days in Qn within [t1,t2])  /  n,  where "
         "n = total RATED days in [t1,t2] (= days in Q1+Q2+Q3+Q4). Calendar days a scheme was unrated are "
         "EXCLUDED from the denominator — same convention as the published deck's residency panel.", None),
        ("t1/t2 are inclusive. If t1/t2 fall on non-trading days the formulas use the trading days inside [t1,t2].", None),
        ("", None),
        ("WHY", f_lbl),
        ("Answers: over a chosen period, how consistently did each fund of an AMC sit in the top/bottom quartile "
         "of its category — e.g. 'ABSL Flexi Cap was in Q1 62% of days between Apr-2025 and Mar-2026 on a 1Y basis'.", None),
        ("", None),
        ("SHEETS", f_lbl),
        ("Map — scheme -> AMFI code, category, fund house (AMC), sub-nature, breakdown; plus each scheme's matrix "
         "column number (ColIdx) used by the formulas.", None),
        ("Q 1Y matrix / Q 3Y matrix — the full m x n base data: trading dates x schemes, values 1..4 (blank = unrated). "
         "Both sheets have IDENTICAL row and column order.", None),
        ("Fund View — pick AMC, t1, t2 from dropdowns (yellow cells; t1/t2 can also be typed as any date). "
         "Shows every fund of that AMC: rated days n, days in each quartile, and % of days in each quartile, on 1Y and 3Y.", None),
        ("All Schemes — the same %s for EVERY scheme (uses the same t1/t2), with AMC + category columns and an "
         "autofilter, for any further aggregation/pivot.", None),
        ("Lists — the AMC dropdown source (with scheme counts).", None),
        ("", None),
        ("CAVEATS", f_lbl),
        ("Universe = All Peers (MFI categories), RAW return quartiles - NOT the VR exact-peer composite score. "
         "Categories with <4 fund houses are excluded by the engine. A fund younger than the window has fewer rated "
         "days (n shows this). Quartile ranks are within-category, so % in Q1 across different categories are "
         "comparable as percentiles, not as returns.", None),
        (f"Data as of {asof} (latest scheme NAV date). Generated by _quartile_matrix_dump.py.", f_note),
    ]
    for i, (txt, fmt) in enumerate(readme):
        ws.write(i, 0, txt, fmt or f_wrap)

    # ---------- Lists (dropdown sources) ----------
    ws = wb.add_worksheet("Lists")
    ws.write_row(0, 0, ["AMC (fund house)", "# schemes"], f_hdr)
    for i, a in enumerate(amcs):
        ws.write(i + 1, 0, a)
        ws.write(i + 1, 1, int(amc_counts[a]))
    ws.set_column(0, 0, 24)
    wb.define_name("AMCS", f"=Lists!$A$2:$A${len(amcs) + 1}")

    # ---------- Map ----------
    ws = wb.add_worksheet("Map")
    hdr = ["Scheme", "AMFI Code", "Category", "Fund house (AMC)", "Scheme Sub Nature",
           "Breakdown", "ColIdx", "PickRank (helper for Fund View)"]
    ws.write_row(0, 0, hdr, f_hdr)
    for i, s in enumerate(schemes):
        r = i + 1
        ws.write(r, 0, s)
        ws.write(r, 1, str(m.at[s, "AMFI Code"]))
        ws.write(r, 2, m.at[s, "Category"])
        ws.write(r, 3, m.at[s, "fund house"])
        ws.write(r, 4, m.at[s, "Scheme Sub Nature"])
        ws.write(r, 5, m.at[s, "Breakdown"])
        ws.write(r, 6, i + 1, f_int)                      # matrix column number (1-based)
        ws.write_formula(r, 7, f'=IF($D{r+1}=AMC_SEL,COUNTIF($D$2:$D{r+1},AMC_SEL),"")')
    ws.set_column(0, 0, 55); ws.set_column(1, 7, 16)
    ws.freeze_panes(1, 1)
    ws.autofilter(0, 0, len(schemes), 5)
    last = len(schemes) + 1
    wb.define_name("MAP_SCHEME", f"=Map!$A$2:$A${last}")
    wb.define_name("MAP_CAT",    f"=Map!$C$2:$C${last}")
    wb.define_name("MAP_COL",    f"=Map!$G$2:$G${last}")
    wb.define_name("MAP_RANK",   f"=Map!$H$2:$H${last}")

    # ---------- Quartile matrices ----------
    def write_matrix(name, mat):
        w = wb.add_worksheet(name)
        w.write(0, 0, "Date", f_hdr)
        for j, s in enumerate(schemes):
            w.write(0, j + 1, s, f_hdr)
        vals = mat.values
        for i in range(nrows):
            w.write_datetime(i + 1, 0, dates[i].to_pydatetime(), f_date)
            row = vals[i]
            w.write_row(i + 1, 1, [int(v) if v == v else None for v in row])
        w.freeze_panes(1, 1)
        w.set_column(0, 0, 12)
        return w

    write_matrix("Q 1Y matrix", n1)
    write_matrix("Q 3Y matrix", n3)
    endc = xl_col_to_name(ncols)                # last data column letter (B..)
    wb.define_name("DATES", f"='Q 1Y matrix'!$A$2:$A${nrows + 1}")
    # NOTE: names must NOT look like cell refs (M1/T1/R2 are cells -> Excel rejects them)
    wb.define_name("MAT_1Y", f"='Q 1Y matrix'!$B$2:${endc}${nrows + 1}")
    wb.define_name("MAT_3Y", f"='Q 3Y matrix'!$B$2:${endc}${nrows + 1}")

    # ---------- shared formula builders ----------
    def metric_formulas(row, colidx_ref):
        """Returns list of (col, formula, fmt) for n / dQ1..4 / %Q1..4 on both bases,
        starting at column E (index 4). colidx_ref = cell ref holding the matrix ColIdx."""
        out, c = [], 4
        for mname in ("MAT_1Y", "MAT_3Y"):
            rng = f"INDEX({mname},ROW_FROM,{colidx_ref}):INDEX({mname},ROW_TO,{colidx_ref})"
            ncell = f"{xl_col_to_name(c)}{row + 1}"
            out.append((c, f'=IF(OR(ROW_TO<ROW_FROM,{colidx_ref}=""),"",COUNT({rng}))', f_int)); c += 1
            for q in (1, 2, 3, 4):
                out.append((c, f'=IF(OR({ncell}="",{ncell}=0),"",COUNTIF({rng},{q}))', f_int)); c += 1
            for q in (1, 2, 3, 4):
                dcell = f"{xl_col_to_name(c - 4)}{row + 1}"
                out.append((c, f'=IF({ncell}="","",IF({ncell}=0,"",{dcell}/{ncell}))', f_pct)); c += 1
        return out

    METRIC_HDR = []
    for b in ("1Y", "3Y"):
        METRIC_HDR += [f"n rated days ({b})"] + [f"days Q{q} ({b})" for q in (1, 2, 3, 4)] \
                      + [f"% in Q{q} ({b})" for q in (1, 2, 3, 4)]

    # ---------- Fund View ----------
    ws = wb.add_worksheet("Fund View")
    ws.write(0, 0, "AMC", f_lbl);  ws.write(0, 1, DEFAULT_AMC, f_in)
    ws.write(1, 0, "t1 (from, inclusive)", f_lbl)
    ws.write_datetime(1, 1, pd.Timestamp(DEFAULT_T1).to_pydatetime(), f_date_in)
    ws.write(2, 0, "t2 (to, inclusive)", f_lbl)
    ws.write_datetime(2, 1, pd.Timestamp(DEFAULT_T2).to_pydatetime(), f_date_in)
    wb.define_name("AMC_SEL", "='Fund View'!$B$1")
    wb.define_name("T_FROM", "='Fund View'!$B$2")
    wb.define_name("T_TO", "='Fund View'!$B$3")
    ws.write(0, 3, "pick from dropdown", f_note)
    ws.write(1, 3, "dropdown of trading days — or TYPE any date", f_note)

    ws.write(3, 0, "matrix rows used", f_lbl)
    ws.write_formula(3, 1, "=IF(T_FROM<=INDEX(DATES,1),1,IFERROR(MATCH(T_FROM,DATES,0),MATCH(T_FROM,DATES,1)+1))", f_int)
    ws.write_formula(3, 2, f"=IF(T_TO>=INDEX(DATES,{nrows}),{nrows},IFERROR(MATCH(T_TO,DATES,0),MATCH(T_TO,DATES,1)))", f_int)
    wb.define_name("ROW_FROM", "='Fund View'!$B$4")
    wb.define_name("ROW_TO", "='Fund View'!$C$4")
    ws.write(4, 0, "trading days resolved", f_lbl)
    ws.write_formula(4, 1, '=IF(ROW_TO<ROW_FROM,"no trading day in range",INDEX(DATES,ROW_FROM))', f_date)
    ws.write_formula(4, 2, '=IF(ROW_TO<ROW_FROM,"",INDEX(DATES,ROW_TO))', f_date)

    ws.data_validation(0, 1, 0, 1, {"validate": "list", "source": "=AMCS"})
    for r in (1, 2):
        ws.data_validation(r, 1, r, 1, {"validate": "list", "source": "=DATES",
                                        "show_error": False})   # typing any date allowed

    hrow = 6
    ws.write_row(hrow, 0, ["#", "Scheme", "Category", "ColIdx"] + METRIC_HDR, f_hdr)
    for k in range(1, maxn + 1):
        r = hrow + k
        ws.write_number(r, 0, k, f_int)
        ws.write_formula(r, 1, f'=IFERROR(INDEX(MAP_SCHEME,MATCH($A{r+1},MAP_RANK,0)),"")')
        ws.write_formula(r, 2, f'=IF($B{r+1}="","",INDEX(MAP_CAT,MATCH($A{r+1},MAP_RANK,0)))')
        ws.write_formula(r, 3, f'=IF($B{r+1}="","",INDEX(MAP_COL,MATCH($A{r+1},MAP_RANK,0)))', f_int)
        for c, fml, fmt in metric_formulas(r, f"$D{r+1}"):
            ws.write_formula(r, c, fml, fmt)
    ws.set_column(1, 1, 55); ws.set_column(2, 2, 24); ws.set_column(3, 3, 7)
    ws.set_column(4, 4 + len(METRIC_HDR) - 1, 12)
    ws.freeze_panes(hrow + 1, 2)

    # ---------- All Schemes ----------
    ws = wb.add_worksheet("All Schemes")
    ws.write(0, 0, "Uses the SAME t1/t2 as Fund View (named cells T_FROM/T_TO on that sheet).", f_note)
    hrow = 1
    ws.write_row(hrow, 0, ["Scheme", "AMC", "Category", "ColIdx"] + METRIC_HDR, f_hdr)
    for i, s in enumerate(schemes):
        r = hrow + 1 + i
        ws.write(r, 0, s)
        ws.write(r, 1, m.at[s, "fund house"])
        ws.write(r, 2, m.at[s, "Category"])
        ws.write_number(r, 3, i + 1, f_int)
        for c, fml, fmt in metric_formulas(r, f"$D{r+1}"):
            ws.write_formula(r, c, fml, fmt)
    ws.set_column(0, 0, 55); ws.set_column(1, 2, 22); ws.set_column(3, 3, 7)
    ws.set_column(4, 4 + len(METRIC_HDR) - 1, 12)
    ws.freeze_panes(hrow + 1, 1)
    ws.autofilter(hrow, 0, hrow + len(schemes), 3 + len(METRIC_HDR))

    wb.close()
    print(f"written: {out}  ({out.stat().st_size/1e6:.1f} MB)")
    return out


def main():
    pm, nav, cmap, qy1, qy3, name2code = build_frames()
    verify(pm, qy1, qy3, cmap)
    write_workbook(nav, cmap, qy1, qy3, name2code)


if __name__ == "__main__":
    main()
