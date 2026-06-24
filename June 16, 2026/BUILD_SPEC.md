# Peer Performance Monitor — Build Spec & Audit Methodology
*Pinned 2026-06-17. Source of truth = the June-6 notebook (`All Peer & Exact Peer based Q1 to Q4 and Scroes - on June 6, 2026.ipynb`); the old `generate_report.py` is a DIVERGENT simplification and must NOT be trusted for the scoring logic.*

This document records, with notebook line citations (into `_nb_dump.txt`), the exact methodology the deliverable must reproduce, plus the agreed changes for the new data layout and data cleaning. It is the contract for the rebuild and doubles as the audit trail.

---
## 0. Deliverable architecture (decided 2026-06-17)
- **Single canonical source** in `# %%`-cell-delimited Python → generates BOTH:
  - `Peer Performance Monitor.ipynb` — logic visible in cells, for KV's audit; runs offline.
  - `generate_report.py` — same code, runnable as `python generate_report.py`; runs offline.
- No internet / no Claude needed to run either. (Dashboard's Chart.js CDN is the only online bit; offline dashboard handled separately.)
- Both regenerated every monthly run; kept in lock-step (one source) so they cannot drift again.

## 1. Input data layout (NEW — single `Data/` folder)
- `Data/Scheme NAV and AUM/` — MFI scheme NAV files (`HistoricalNav_*`) + AUM files (`Scheme wise AUM Report-*`). NAV = dates × schemes; header at row 5 ("Date | names"), rows 6-8 = Scheme/Index Code, AMFI Code, Fund Name, data row 9+, 3 footer rows. `--`/0 → NaN.
- `Data/Mapping/Map MFI Scheme to Category.xlsx` — sheet `Map`: cols [AMFI Code, Scheme Name, Scheme Nature, Scheme Sub Nature, Category, Is sector/theme, Scheme/Index Code, Fund Name]. (= old `Map.xlsx`.)
- `Data/Mapping/Mapping VR to MFI names.xlsx`:
  - sheet `Ret. Compr.(Equity) - Dir`: cols [Scheme, AMFI Code, Category, Scheme Name From MFI]. 536 rows; **36 rows with `AMFI Code == 'bench'`** (the per-category benchmark rows — convention preserved). 38 categories.
  - sheet `Benchmarks and sources`: cols [Benchmarks, MFI Identifier, File] — maps benchmark **display name** → **MFI Identifier** (= the column name in the MFI benchmark NAV file) → file. (Old file used cols [Identifier, Benchmarks].)
- `Data/Benchmark NAV/VR Benchmarks NAV from MFI - June 16 2026.xlsx` — benchmark NAVs in **MFI NAV format** (Output sheet; header row 5 = "Date | names", row 6 = Scheme/Index Code, **data row 7+** — only ONE metadata row, unlike scheme files' three → needs its own reader).
- `Data/Benchmark NAV/VR Benchmarks Missing MFI NAV June 16 2026.xlsx` — sheet `Sheet1`, cols [Dates, SPG1200T Index, XNDX Index] = S&P Global 1200 TRI & NASDAQ-100 TRI, stale Bloomberg fallback (MFI failed to supply).

## 2. Data cleaning (NEW — agreed 2026-06-17)
Applied to BOTH scheme NAV and benchmark NAV, with a full audit log (every dropped date, every filled gap):
1. **Drop non-trading days.** A date is non-trading if **≥ `REPEAT_FRAC` of populated series equal the prior day's value exactly**. `REPEAT_FRAC` is a single labeled constant; default **0.90** (KV's steer; 0.98 too strict). Report how many dates drop at 0.90/0.95/0.98 for tuning. (Calendar holidays unreliable → detect from data.)
2. **Careful forward-fill.** Fill a within-series gap **only if** it spans `< MAX_FILL_GAP` trading days (default **5**) **AND** a real value resumes after. **Never** fill past a series' last real value (discontinued scheme/index → leave NaN; do NOT fabricate). Blind `.ffill()` (old code, NAV line 86 / notebook line 106) is REMOVED.

## 3. Calendar alignment & staleness (NEW — agreed 2026-06-17)
- Old bug: `nav = nav.loc[nav.index.intersection(bench_nav.index)]` (py 679-681) truncated ALL analysis to the earliest common last-date. The INTENT was only to put schemes & benchmarks on a common trading-day calendar.
- **New rule:** align on common trading days **up to `min(last scheme update, last benchmark update)`; beyond that, use whatever is available.** Do NOT let a stale benchmark drag the global window down.
- **Per-benchmark staleness cap:** each category's benchmark-dependent outputs (alpha, the +1 bonus) run only to **that category's benchmark's last real date**. Last real dates (as of 2026-06-17 data): scheme NAV → 2026-06-16; most MFI benchmarks → 2026-06-15; CRISIL Hybrid ×3 → 2026-05-31 (monthly); "65% BSE200+CRISIL STBond" composite → 2025-11-28; Silver/BSE Select Group → 2026-05-29; BSE India Manufacturing TRI* → EMPTY; **NASDAQ-100 (XNDX) & S&P Global 1200 (SPG1200T) → 2026-03-11**.
- So: **peer-relative quartiles run to 2026-06-16; NASDAQ-100-FOF & S&P-Global-linked categories' benchmark metrics cap at 2026-03-11; everything else current.**

## 4. Canonical analysis (mirror the notebook EXACTLY)

### 4a-0. JOIN BY AMFI CODE, not name (CRITICAL, KV requirement 2026-06-17)
Cross-source scheme joins MUST key on **AMFI Code** (present in NAV's metadata row, the Map, and the VR map), NOT scheme name — names differ across vendors/formats (double spaces, "Dir" vs "Direct", en-dashes) and silently drop schemes. Measured coverage: NAV↔Map **876/906 by name → 906/906 by code**; VR↔NAV **479/497 by name → 497/497 by code**. NAV↔AUM matches 906/906 by name (same MFI vendor) so AUM stays name-joined. Implementation: `scheme_code_map()` (AMFI Code→NAV name from the NAV 'AMFI Code' row); `build_category_map(raw_map, code2name)` joins by code, indexes by NAV name; `align_vr_to_nav(cme, nav_cols, code2name)` remaps VR peers to NAV names by code (normalized-name fallback). Keep readable NAV names for display; use code only for the join.

### 4a. Universe & maps
- `fund house` = first token of scheme name (`split(" ")[0]`) — nb 173, 284.
- All-peer (MFI) categories: keep categories with ≥4 fund houses (nb 245); Value+Contra merged → "Value/Contra"; Breakdown rules + MULTI_ASSET_DROP per old code §2.
- Exact-peer (VR) categories from `category_map_exact['Category']`, **excluding** `["Domestic + International","Multi Index FoF","Other Competitor Thematic/Contra Funds"]` (nb 1235); sorted (nb 1238).
- `exact_bench`: Category → Benchmark name, from the `AMFI Code=='bench'` rows (nb 287-291). Benchmark name must match a column in `bench_nav` after name-mapping.

### 4b. Returns — EXACT CALENDAR (corrected 2026-06-17, KV requirement; canonical = nb Cell 49 "Final code… Point to Point Calendar Year")
- **USE EXACT-CALENDAR point-to-point returns, NOT 250/750-day shift.** KV gets vendor reports on a calendar basis — numbers must match. For each date `t`: `t_1y` = last trading day on/before the SAME calendar date 1 year earlier (Feb-29 → day-1); `t_3y` likewise. `1Y = NAV_t/NAV_{t_1y} − 1` (cumulative); `3Y = (NAV_t/NAV_{t_3y})**(1/3) − 1` (**CAGR/annualized**). Same for benchmarks, on the scheme calendar. (nb Cell 49 lines 1373-1396.)
- Cell 47 (`shift(250)/shift(750)`) is the EARLIER, superseded version — do not use.
- Applied to BOTH exact-peer scoring AND all-peer quartiles (`basis="calendar"`, default). Quartile ranks ≈ invariant to CAGR-vs-cumulative, but calendar-vs-shift IS material (e.g., ABSL Large Cap Q4→Q3).
- **VERIFIED 2026-06-17:** hand-reconciled ABSL Large Cap at t=2026-06-16 — manual `NAV_t/NAV_{2025-06-16}−1 = −0.024273` and `(NAV_t/NAV_{2023-06-16})**(1/3)−1 = 0.117145` BOTH exactly match `calendar_returns()`. Composite range 1.8–5.0 ✓. VR quartile bands sum == coverage ✓.

### 4c. Quartiles (per category, per date) — round-up bucketing
Sort schemes by return descending; bucket sizes = `n//4` with the remainder distributed to the TOP buckets (nb 1286-1296). Assign q1..q4. Done for 1Y and 3Y. (nb 1281-1324.)

### 4d. Benchmark alpha
`f_alpha_Ny = f_Ny_rets.sub(_Ny_rets_bench[f_bench], axis=0)` — scheme return minus its category benchmark's return, per category (nb 1256-1258, 1435-1437).

### 4e. SCHEME SCORE (exact-peer / VR) — the part the .py got wrong
- `df_1y = qy_df_1y` mapped **q1→5, q2→4, q3→3, q4→2** (nb 1529). *(1Y scored a point higher than 3Y.)*
- `df_3y = qy_df_3y` mapped **q1→4, q2→3, q3→2, q4→1** (nb 1531); keep `df_3y_raw` pre-bonus.
- **Benchmark bonus:** `df_3y = df_3y + 1*(alpha_df_3y > 0)` — +1 if scheme beats its benchmark on 3Y (nb 1540). *(For NASDAQ/S&P categories, alpha only valid to 2026-03-11; cap there.)*
- **Composite:** `score_df = 0.8*df_1y + 0.2*df_3y` (nb 1543).
- Drop a category if no fund has any score (<3y history) (nb 1547-1549).

### 4f. Composite re-bucket → VR quartiles
Re-sort `score_df` per category/date into fresh q1..q4 (round-up) → `qy_df_score`; `score_per` = within-category rank percentile of the composite score (nb 1626-1678).

### 4f-bis. INTERACTIVE composite-score knob (NEW — requested 2026-06-17)
The composite-score method becomes a **user control in the dashboard**; the VR-peer visuals recompute live. Mechanics:
- **JSON ships per scheme × month-end:** `s1` = 1Y quartile score (5/4/3/2), `s3raw` = 3Y quartile score (4/3/2/1), `beat` = 1 if 3Y alpha vs its benchmark > 0 (null when benchmark stale/unavailable at that date, e.g. NASDAQ/S&P after 2026-03-11), plus category, fund house, AUM share.
- **Controls:**
  - **1Y weight `w`** — integer %, 0–100; **3Y weight auto = 100−w** (single uniform knob). Free integer entry + presets 100 / 80 / 67 / 50 / 0.
  - **Benchmark bonus** toggle (default ON).
  - Preset shortcuts: "1Y only" (w=100), "3Y only" (w=0, bonus off), "3Y + benchmark reward" (w=0, bonus on), "Composite 80/20 + reward" (default).
- **Live recompute (JS), per category × date:** `s3 = s3raw + (bonus? beat:0)`; `composite = (w/100)*s1 + (1−w/100)*s3`; re-bucket schemes in category by composite desc → q1..q4 (round-up); aggregate %AUM per quartile by fund house, scheme scores, quality score.
- **Excel files are static** → use the DEFAULT method (w=80, bonus ON = notebook canonical) and label the method in the file. "All Peers" tab (raw quartiles, no score) is unaffected by the knob.

### 4f-ter. Dashboard requirements (accumulated)
1. Interactive composite-score knob (§4f-bis): 1Y weight integer % (3Y auto = 100−w) + benchmark-bonus toggle + presets; VR visuals recompute live.
2. **Data Quality & Exclusions tab** (requested 2026-06-17): list every drop/exclusion with reason — 3 Nasdaq-FOF schemes w/ no MFI NAV; 10 coalesced dup schemes; 43/1 non-trading dates dropped; lagging (1d) & discontinued series; stale benchmarks + per-category caps; Make-in-India empty benchmark; categories dropped (<4 houses / <3y / repeated peer-set). Also print at run-time + an Excel sheet.
3. **All chart axes ADAPTIVE** (requested 2026-06-17): no hardcoded caps anywhere. The score chart was capped at y=4, hiding scores up to **5.0** (1Y Q1 s1=5; 3Y Q1 s3_raw=4 +1 bonus = 5 → 0.8·5+0.2·5=5). Every axis must auto-scale to the data (incl. the score panels, quality-score, %AUM bands).

### 4g. AUM share
- `percent_aum_df` = each scheme's AUM ÷ its fund-house total AUM, daily (monthly AUM ffilled to daily) (nb 348-365). `percent_aum_df1` merges `fund house`.
- The old `.py` `percent_aum_daily()` is equivalent — reuse/verify.

### 4h. Outputs (mirror current shapes)
- **`percent AUM in Q1 to Q4 - <date>.xlsx`**: Sleevewise + AMCwise (**All Peers** = raw 1Y & 3Y quartiles) + AMCwise (**VR Peers** = composite `qy_df_score`, after dropping `reperated_peer_category` = Bal Bhavishya & Retirement Fund 40, "repeated peersets" — nb 1686-1689). Restricted to TOP15 houses.
- **`Scheme Scoring on Exact Peer Set - Calendar 1Y 3Y - <date>.xlsx`**: 1Y score, 3Y score, composite daily score, peer+bench mapping (exact-peer, composite-based).
- **Distance/alpha table**: Birla-vs-category (top/bottom/mean/Q1) AND Birla-vs-benchmark alpha + % outperforming (peer + benchmark), `% times +ve alpha` rolling 250 (nb 440-467, 647-648).
- **Dashboard** (`dashboard.html` + `dashboard_data.json` + `dashboard_offline.html`): reuse the embedded template; feed it the corrected composite/score/alpha + month-end sampling.

### 4i. Short rolling windows 1M/3M/6M/9M + quartile residency (NEW — KV 2026-06-24)
In ADDITION to the 1Y/3Y windows, every scheme is evaluated on **1-, 3-, 6- and 9-month** rolling
windows so the WHOLE deck (sleeve/AMC/league/matrix/scheme) can be viewed on any window via the
global **`Rolling window / basis`** selector (`1M·3M·6M·9M·1Y·3Y·3Y+rew·Composite`).
- **Returns** (`calendar_returns_m(nav, months)`): exact-calendar point-to-point, `NAV_t/NAV_{t_m}−1`,
  where `t_m` = last trading day on/before the same calendar day `months` months before `t`
  (`pd.DateOffset(months=...)`, day-clamped, e.g. Mar-31−1M→Feb-28). **CUMULATIVE, not annualized**
  for sub-year — a sub-year CAGR misleads, and quartile RANKS are identical either way (annualizing is
  monotonic). The 1Y/3Y windows are unchanged (1Y cumulative, 3Y CAGR).
- **No composite / no benchmark reward** for the sub-year windows — raw quartiles only; composite
  (0.8·1Y+0.2·3Y, +reward) applies ONLY when combining 1Y & 3Y.
- **Quartiles**: round-up bucketing (`quartiles_roundup`) within the **MFI all-peer** category
  (`all_peer_quartiles_m`) AND within the **VR exact-peer** category (`vr_quartiles_m`). Both universes.
- **Engine→JSON**: per-scheme month-end quartiles `aq1m/aq3m/aq6m/aq9m` (all-peer rows) and
  `q1m/q3m/q6m/q9m` (VR rows). Dashboard aggregates the all-peer short windows into the house bands
  **client-side** (a raw cube from the per-scheme quartiles — symmetric with the VR cube — so we DON'T
  ship 4 more precomputed sleeve/AMC tables; the validated 1Y/3Y precomputed-table path is untouched).
- **Quartile RESIDENCY** (`window_residency`): for every **Top-15-house** scheme, both universes, all
  6 windows, the count of trading days spent in Q1/Q2/Q3/Q4 over the **trailing window ending at each
  as-of month-end** (window length = lookback, e.g. 3M ≈ 62-65 trading days). Shipped as
  `residency[universe][scheme][windowLabel] = {f:firstAsofIdx, v:[[q1,q2,q3,q4]|null,…]}` aligned to
  `aum_dates` (all) / `months` (VR), leading/trailing-empty trimmed. **Definition** = days in each
  quartile over the trailing window; **denominator `n` = days the scheme was rated** (= q1+q2+q3+q4,
  NOT calendar days — a young scheme without the full window of history has fewer rated days);
  shown as days, fraction `d/n`, and `%` to 2 dp. Rendered in **Scheme Detail, beside the quartile chart**;
  **follows the as-of date selector**. Composite/3Y+reward show a "pick a single window" hint (a blend
  of 1Y & 3Y has no single window). Only Top-15 schemes (the scheme picker only lists those).
- **Size note**: residency + the per-scheme short-window quartiles roughly double the offline deck
  (~9.7 → ~18.7 MB; residency ≈ 6.3 MB). Knowingly accepted for the as-of-aware interactivity; GitHub
  Pages serves it gzipped (~3-4 MB on the wire). Trim levers if email size matters: drop 1Y/3Y residency,
  cap residency to recent N years, or revert residency to latest-date-only.

## 5. Open / parked
- Auto-fetch of MFI data (analogous to NSE-Indices fetch in the FFT project) — AFTER this run.
- Web fetch for NASDAQ-100 TRI & S&P Global 1200 TRI (replace stale Bloomberg) — parked, save for later.
- **`_app.js` is a STALE working copy** — the authoritative dashboard app code lives INLINE in
  `_dashboard_src.html` (the bake reads only that). Edit `_dashboard_src.html`; re-sync or delete
  `_app.js` in a future cleanup so it can't mislead.

## 7b. Benchmark staleness handling (finalized 2026-06-17)
`align_benchmarks` now CARRIES benchmarks forward (ffill) onto the scheme calendar, so the +1 bonus is computable at every date using the latest available benchmark level (correct for periodic indices like monthly CRISIL Hybrid). A category's composite is CAPPED (set NaN after its benchmark's last real date) only when that benchmark is **>`STALE_BENCH_DAYS` (45) behind scheme-latest** — i.e. genuinely stale. As of June data this caps exactly: Pure International & Nasdaq-100-FOFs (2026-03-11), Multi-Asset Allocation (2025-11-28, broken composite benchmark). Monthly CRISIL/Silver and 1-day-lagged benchmarks stay current. Month-end outputs use `month_end_asof` (last non-NaN ≤ month-end per column) so capped cats show their last valid date and 1-day reporting lags are absorbed. Excel files default to month-end resolution. Verified Aditya VR composite as-of 06-16: Q1 41.5/Q2 15.5/Q3 27.9/Q4 5.2, coverage 90.2%.

## 7. Build status / resume (updated 2026-06-17)
- **DONE & validated:** `peer_monitor.py` §loaders + §cleaning, tested on June data.
  - `load_scheme_nav()` reads both MFI files, coalesces **10 duplicate schemes** → 906 schemes × 5059 dates (2006-03-13 → 2026-06-16). `read_mfi_nav` is metadata-row-agnostic (works for scheme [3 meta rows] and benchmark [1] files).
  - Non-trading-day drop: **43** scheme dates (e.g. weekend rows 2022-08-20/21), **1** benchmark date; stable across 0.90/0.95/0.98 (≤1 date difference).
  - Careful ffill: filled internal <5d gaps; left tails unfilled. 7 scheme series lag 1 day (overseas FOFs, last real 2026-06-15 — correctly NOT fabricated for 06-16). Stale benchmarks confirmed: NASDAQ/SPG→2026-03-11; 65%BSE200 composite→2025-11-28; CRISIL Hybrid→2026-05-31; Silver/BSE Select→2026-05-29.
- **DONE & validated:** `peer_monitor.py` §VR-wiring. `load_vr_mapping()` → 497 peer schemes, 3 dropped (no MFI NAV), `exact_bench` (Category→benchmark display name, whitespace-stripped). `load_bench_nav()` assembles 26 benchmarks across the MFI file + stale NASDAQ/SPG file via `Benchmarks and sources` [Benchmarks, MFI Identifier, File]. **All 35 in-scope exact-peer categories resolve** to a benchmark NAV column.
  - **Per-category benchmark staleness caps** (benchmark metrics/+1 bonus cap at this date; peer quartiles unaffected, run to scheme-current 2026-06-16): Multi Asset Allocation → 2025-11-28; Nasdaq 100 FOFs & Pure International Plan → 2026-03-11; Conglomerate & Silver FoF → 2026-05-29; Aggressive Hybrid, Asset Allocator, Balanced Advantage, Retirement Fund 40 → 2026-05-31; **Make in India → benchmark EMPTY (no NAV in MFI; +1 bonus always 0, peer-only)**. All other categories' benchmarks current to 2026-06-15.
- **DONE & validated:** `peer_monitor.py` §category-map (`build_category_map`), §AUM-share (`load_aum`,`percent_aum_share`), §calendar-align (`align_benchmarks` — bridges ≤3d offsets, never carries past a benchmark's last real date), §returns (`rolling_returns` 250/750d cumulative), §quartiles (`quartiles_roundup`), and §exact-peer scoring (`exact_peer_scoring`): 5/4/3/2 & 4/3/2/1, +1 bonus, 0.8/0.2 composite, composite rebucket — validated on Large Cap (Kotak composite 4.8/Q1; ABSL Large Cap 2.0/Q4) and Nasdaq (alpha caps 2026-03-11). Exclusions threaded via `cmap.attrs['exclusions']` + the 3 dropped Nasdaq-FOF schemes.
  - **REFINEMENT (min-of-updates):** composite (bonus-dependent) output per category caps at its benchmark's last real date; peer-only quartiles run to 2026-06-16. Driven by per-(cat,date) `beat`-availability so the dashboard knob (bonus off) can extend to scheme-current.
- **Data Quality & Exclusions panel** (requested): dashboard tab + run-time print + Excel sheet listing every drop/exclusion w/ reason — 3 Nasdaq-FOF schemes (no MFI NAV: ABSL US Equity Passive FoF, Axis US Specific Equity Passive FoF, Motilal Oswal Nasdaq 100 FOF), 10 coalesced dup schemes, 43/1 non-trading dates, lagging/discontinued series, stale/empty benchmarks, dropped categories.
- **ENGINE VERIFIED 2026-06-17 (all PASS, calendar basis):** (A) calendar returns hand-reconciled on 3 schemes (1Y+3Y exact match); (B) all-peer quartile bucket sizes = round-up, bands sum ≤ AMC coverage; (C) composite range [1.800,5.000], `composite=0.8·s1+0.2·(s3_raw+bonus)` spot-check exact; (D) benchmark staleness caps exact (Nasdaq/SPG 03-11, Multi-Asset 11-28, CRISIL-Hybrid 05-31); (E) VR bands sum == coverage; (F) exclusions captured. Full pipeline ~110s. Caches: `_cache.pkl` (nav,cmap,cats,pct,qy1,qy3,sleeve,amc,bench_cal,bench_last), `_cache_res.pkl` (exact-peer res).
- **PENDING REFINEMENT (presentation, not logic):** VR/composite snapshot "latest" should use each category's benchmark-last date (so the +1 bonus is applied) — currently 06-16 shows bonus=0 (benchmark absent that day). Apply per-cat cap in Excel/dashboard snapshot.
- **RE-VERIFIED 2026-06-17 with AMFI-code joins (all PASS):** universe 873→**903 schemes** (NAV↔Map by code), **all 18 VR peers recovered (0 unmatched)**, Arbitrage 14→15. Checks A–F all PASS. Aditya all-peer 1Y mix (complete universe): Q1 42.4/Q2 25.5/Q3 12.4/Q4 19.5. VR composite: Q1 44.9/Q2 15.5/Q3 28.1/Q4 5.2. Caches regenerated.
- **Cross-checked vs KV's live 1Y report:** with the report's ≥8-peer filter applied, **Q1 count matched exactly (10 funds)** — logic confirmed. Residual %AUM gaps are peer-universe-definition differences (KV: don't force-match; run on full universes). `min_peers` param added to `exact_peer_scoring` (default 1; set 8 to mirror the report's "<8 peers excluded" note).
- **INTEGRATED & VALIDATED:** `peer_monitor.py` `run()` does the full pipeline end-to-end; `python peer_monitor.py --out X` exits 0 and writes both Excel + dashboard_data.json (903 schemes, 18 VR peers recovered). Excel writers expose score COMPONENTS (1Y, 3Y-no-reward, 3Y+reward, beats-benchmark flag, default composite+quartile) so any weight/reward combo is reconstructable. `write_dashboard_json` ships per-scheme components + exclusions + capped_cats + repeated_cats + peer_map (2.03 MB). **Dashboard JS recompute spec VERIFIED** (JS-mock reproduces Excel: Q1 41.5/Q2 15.5/Q4 5.2 exact, Q3 28.1 vs 27.9 tie-noise) — needs cap-skip + repeated-cat-skip + round-up bucketing.
- **Notebook:** `make_notebook.py` converts the `# %%` source -> `Peer Performance Monitor.ipynb` (15 cells, argparse guard dropped, explicit run() cell). Single source -> .py + .ipynb, no drift.
- **Knobs are the DASHBOARD's job** (interactive); Excel is static default (0.8/0.2+reward) + components. Dashboard build delegated to an agent (knob + exclusions tab + adaptive axes), with the verified recompute spec + test vector.
- **DELIVERABLE COMPLETE (2026-06-17):** `python peer_monitor.py --out out` (exit 0) produces the full package in `out/`: 2 Excel files, strictly-valid `dashboard_data.json` (NaN bug fixed — `allow_nan=False`), `dashboard.html`. Files: `peer_monitor.py` (45KB engine), `Peer Performance Monitor.ipynb` (15 cells, from `make_notebook.py`), `dashboard.html` (51KB template at root, copied to out/). Dashboard rendered headlessly — 5 tabs, live score knob (1Y-weight + reward toggle + presets), Data Quality & Exclusions tab, adaptive axes (scores fit to 5); Aditya VR strip 41.5/15.5/28.1/5.2 matches the engine.
- **Dashboard hosting:** fetches dashboard_data.json → run `python -m http.server` in `out/` (or host on intranet); not file://. Chart.js from CDN (one online load; self-host for full offline per README pattern).
- **OPEN / OPTIONAL:** (a) per-fund 1Y-quartile reconciliation vs KV's report still pending KV's per-fund data — differences are peer-universe composition (KV: don't force-match); (b) offline single-file dashboard (inline JSON + Chart.js) for double-click/email — not built; (c) final /workflows adversarial pass — not run (engine already hand-verified extensively); (d) cleanup dev scaffolding (_inspect*, _nb_dump*, _run_test, _verify*, _build_outputs, _cache*, out_test).

## 6. Verification (via /workflows, adversarial)
Independent agents check: (1) cleaning dropped only genuine holidays & ffill respected discontinuation; (2) calendar/staleness decoupling (peer to 06-16, NASDAQ cats to 03-11); (3) benchmark→category mapping correct across the two new files; (4) scoring constants (5/4/3/2, 4/3/2/1, +1 bonus, 0.8/0.2, composite rebucket) match the notebook; (5) outputs reconcile vs the notebook on sample categories.
