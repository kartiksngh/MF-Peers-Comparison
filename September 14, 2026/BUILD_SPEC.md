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
- **The VR peerset is ABSOLUTE and a SUBSET of All-Peers (KV 2026-07-03).** The NAV universe = Map schemes ∪ VR schemes; All-Peers exclusions (MULTI_ASSET_DROP, unmapped schemes) remove a scheme from the All-Peers categorisation ONLY — never from the NAV universe, so a VR peer is always scored (e.g. WhiteOak Capital Multi Asset Allocation: excluded from All-Peers, scored in VR). A scheme common to both universes must carry the same category concept in both (label synonyms allowed, e.g. Gennext↔Consumer); Map `Category=0` is not allowed — VR members take their VR category, others their SEBI (AMFI NAVAll) category.
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
- **`percent AUM in Q1 to Q4 - <date>.xlsx`**: Sleevewise + AMCwise (**All Peers** = raw 1Y & 3Y quartiles) + AMCwise (**VR Peers** = composite `qy_df_score`, after dropping the **BORROWED memberships** of the repeated/widened peer-sets = Bal Bhavishya & Retirement Fund 40 — a scheme whose home is another category counts there, once; a NATIVE fund whose only membership is the widened set (Bal Bhavishya Yojna, Retirement 40s Plans) reports FROM that set. Until 2026-07-16 the whole repeated category was dropped, which made the native funds count zero times — corrected per KV's each-scheme-counts-once ruling; the deck's client-side cube mirrors this exactly). Restricted to TOP15 houses.
- **`Scheme Scoring on Exact Peer Set - Calendar 1Y 3Y - <date>.xlsx`**: 1Y score, 3Y score, composite daily score, peer+bench mapping (exact-peer, composite-based).
- **Distance/alpha table**: Birla-vs-category (top/bottom/mean/Q1) AND Birla-vs-benchmark alpha + % outperforming (peer + benchmark), `% times +ve alpha` rolling 250 (nb 440-467, 647-648).
- **Dashboard** (`dashboard.html` + `dashboard_data.json` + `dashboard_offline.html`): reuse the embedded template; feed it the corrected composite/score/alpha + month-end sampling.

### 4i. Extra rolling windows 1M/3M/6M/9M/2Y/5Y + quartile residency (KV 2026-06-24; 2Y/5Y added KV 2026-07-14)
In ADDITION to the 1Y/3Y windows, every scheme is evaluated on **1-, 3-, 6-, 9-month, 2-year and
5-year** rolling windows so the WHOLE deck (sleeve/AMC/league/matrix/scheme) can be viewed on any
window via the global **`Rolling window / basis`** selector (`1M·3M·6M·9M·1Y·2Y·3Y·5Y·3Y+rew·Composite`).
- **Returns** (`calendar_returns_m(nav, months)`): exact-calendar point-to-point, `NAV_t/NAV_{t_m}−1`,
  where `t_m` = last trading day on/before the same calendar day `months` months before `t`
  (`pd.DateOffset(months=...)`, day-clamped, e.g. Mar-31−1M→Feb-28). **Computed CUMULATIVE, not
  annualized** — a sub-year CAGR misleads, and quartile RANKS are identical either way (annualizing
  by `^(12/m)` is monotonic). Conceptually the multi-year windows (2Y/5Y) are annualized (CAGR) like
  3Y; since only the RANK is stored and shown, the cumulative computation is exact for them. A scheme
  without the full window of history (e.g. <5y for 5Y) is NaN → unrated in that window, so the rated
  universe shrinks as the window lengthens. The 1Y/3Y windows are unchanged (1Y cumulative, 3Y CAGR).
- **No composite / no benchmark reward** for these extra windows — raw quartiles only; composite
  (0.8·1Y+0.2·3Y, +reward) applies ONLY when combining 1Y & 3Y.
- **Quartiles**: round-up bucketing (`quartiles_roundup`) within the **MFI all-peer** category
  (`all_peer_quartiles_m`) AND within the **VR exact-peer** category (`vr_quartiles_m`). Both universes.
- **Engine→JSON**: per-scheme month-end quartiles `aq1m/aq3m/aq6m/aq9m/aq24m/aq60m` (all-peer rows) and
  `q1m/q3m/q6m/q9m/q24m/q60m` (VR rows). Dashboard aggregates the all-peer extra windows into the house bands
  **client-side** (a raw cube from the per-scheme quartiles — symmetric with the VR cube — so we DON'T
  ship 4 more precomputed sleeve/AMC tables; the validated 1Y/3Y precomputed-table path is untouched).
- **Quartile RESIDENCY** (`window_residency`): for **EVERY scheme** (scope widened 2026-08-11 — see the
  §4j/§4k/§4l scope note; it was Top-15 houses only until the peer-analytics grid needed whole category
  cohorts), both universes, all
  8 windows (1M/3M/6M/9M/1Y/2Y/3Y/5Y). For a scheme with >1 VR membership the residency (and every
  scheme-level view: score history, KPIs, peer table) reports from its **HOME peerset** — the
  non-widened category when one exists, else its widened home (`_vr_daily_by_scheme`; deck
  `vrHomeRec`) — never from a borrowed membership (KV 2026-07-16). It counts the trading days spent in Q1/Q2/Q3/Q4 over the **trailing window ending at each
  as-of month-end** (window length = lookback, e.g. 3M ≈ 62-65 trading days). Shipped as
  `residency[universe][scheme][windowLabel] = {f:firstAsofIdx, v:[[q1,q2,q3,q4]|null,…]}` aligned to
  `aum_dates` (all) / `months` (VR), leading/trailing-empty trimmed. **Definition** = days in each
  quartile over the trailing window; **denominator `n` = days the scheme was rated** (= q1+q2+q3+q4,
  NOT calendar days — a young scheme without the full window of history has fewer rated days);
  shown as days, fraction `d/n`, and `%` to 2 dp. Rendered in **Scheme Detail, beside the quartile chart**;
  **follows the as-of date selector**. Composite/3Y+reward show a "pick a single window" hint (a blend
  of 1Y & 3Y has no single window). Only Top-15 schemes (the scheme picker only lists those).
- **Size note**: residency + the per-scheme short-window quartiles roughly double the offline deck
  (~9.7 → ~18.7 MB; residency ≈ 6.3 MB at the time). Knowingly accepted for the as-of-aware
  interactivity; GitHub Pages serves it gzipped (~3-4 MB on the wire). Trim levers if email size matters:
  drop 1Y/3Y residency, cap residency to recent N years, or revert residency to latest-date-only.
  **Updated 2026-08-11 (all-scheme scope, measured on that refresh's own data):** top-level `residency`
  8.23 → **12.07 MB** (499 → 924 all-peer scheme keys, 331 → 473 VR scheme keys), and each convention's
  `sip.<conv>.residency` 8.41 → **≈12.33 MB** on the same 1.4658× factor. Per window, the widened
  all-peer block runs 0.65 MB (5Y) to 1.07 MB (9M) and the VR block 0.42 to 0.67 MB. Full table in §4l.

### ★ 4j/4k/4l SCOPE RULE — the per-scheme analytics are ALL-SCHEME (ADDENDUM3 §I, ratified 2026-08-11)
One rule governs `residency`, `standing` and `returns`, stated once here because all three obey it:
**they cover EVERY scheme (all-peer universe) and EVERY membership (VR universe) — not just the Top-15
houses.** `standing` was always all-scheme; `residency` and `returns` were Top-15-filtered in their first
builds and were widened on 2026-08-11 (ratification R1, superseding Addendum 2 §C).
- **Why**: the filter's justification was "the Scheme-Detail picker only lists Top-15 houses". The PEER
  ANALYTICS GRID (Addendum 3 §G) broke it — the grid shows the selected scheme's whole CATEGORY COHORT
  side by side, and cohorts contain non-Top-15 houses (Technology's Edelweiss / Motilal / Quant funds are
  the worked example). Under the old filter those peers rendered as blank rows: data that silently isn't
  there, the failure mode this project forbids.
- **How**: one argument, `top15=None` (now the default) on `build_sip` / `build_returns`, and no filter at
  all in `write_dashboard_json`'s residency block. Passing a house list narrows it again — the size escape
  hatch, kept deliberately.
- **What did NOT change**: which schemes the AGGREGATE tabs surface. Sleeve / AMC / matrix / league stay
  Top-15, because that narrowing lives in `sleeve_amc_tables` / `vr_amc_table` / `meta.top15`, untouched.
- **Additive**: every previously-shipped record is byte-identical; only new keys appear (asserted in the
  smoke test at both universes).
- **Gate**: `_sanity_check.py` check 10 (`check_scope`) — coverage floors per block (≥80%: never-rated
  schemes are legitimately absent), "every standing cell has a residency and a return for the same window"
  (≥99%), and a NAMED non-Top-15 scheme probed for presence in both blocks. Verified with a positive
  control and two negative controls (the pre-R1 Top-15 shape, and a single non-Top-15 scheme's returns
  removed — both caught). **Run against a real-data-shaped 2026-08-11 deck: 0 hard failures** — every
  coverage line reads 100.0% (924/924 all-peer schemes, 473/473 VR schemes, 511/511 VR memberships) and
  the standing→residency implication reads 10,280/10,281 = 99.99%.
- **The one tolerated miss, and why `IMPLY_MIN` is 0.99 not 1.0** — it is PRE-EXISTING, not caused by the
  widening. `standing.vr['Conglomerate Fund|Aditya Birla Sun Life Conglomerate Fund - Dir - Growth']`
  carries 9 non-null **1-Year ranks**, but that scheme's shipped `vr` row has **zero** non-null `q1y`
  digits and the SHIPPED (pre-widening, Top-15) residency had no 1-Year record either. Cause: `standing`
  recomputes VR 1Y ranks from `calendar_returns` + `_vr_member_sets`, while the digits and residency come
  from `res["qy_1y"]` (`exact_peer_scoring`), which produces no labels for that category. **Consequence
  worth fixing separately: the Category-standing card can print "rank r of n" on a window whose quartile
  chip reads "—".** Logged here, out of scope for this wave.

### 4j. Category STANDING — per-scheme ranks (STANDING_SIP_DESIGN.md Feature A; engine built 2026-08-11)
The deck needs each scheme's exact 1-based RANK inside its category (not just the quartile digit) so
Scheme Detail can show "rank 8 of 24; 70.8% of rated funds / 80.3% of rated AUM at or below". The deck
ships no returns, so for the raw windows the ranks must come from the engine.
- **Definition**: rank = the scheme's 1-based position in `_sort_desc_stable(returns that date)` among
  the category's non-null members — the SAME per-date descending stable sort that assigns the quartiles
  (`_quartile_block` now takes `with_ranks=False`; `True` returns `(labels, ranks)` from one sort, so a
  rank can never disagree with its quartile digit; every pre-existing call site is untouched and the
  label output is bit-identical — asserted in the smoke test).
- **Method** (`build_standing(nav, cmap, cats, cme, months_axis, aum_axis)`, a standalone step — the
  validated quartile path is not touched): recompute the window returns (1Y/3Y via `calendar_returns`,
  the other 6 windows via `calendar_returns_m`) on the SAME member sets as the quartile builders
  (all-peer = cmap Category groups; VR = cme categories minus AVOID_EXACT_CATS, dict.fromkeys de-dup,
  min_peers); run combined label+rank blocks per category (one joblib Parallel batch per window,
  n_jobs=-1 — labels are discarded; coherence vs the shipped digits is asserted by sanity check 7);
  month-end-sample the daily rank frames with `month_end_asof` onto the SAME axes as the quartile digit
  series (all-peer → the sleeve table's month-end date columns = `aum_dates`; VR → the composite's
  month-ends = `months`); trim to residency-style `{f, v}` records (leading+trailing all-null trimmed,
  `v[i-f]` = rank at axis index i, interior nulls kept).
- **JSON**: top-level `standing = {"all": {scheme: {winLabel: {f,v}}}, "vr": {"cat|scheme": {...}}}` for
  ALL schemes/memberships (the scope rule above — the client sums beaten AUM over every member, and the
  peer grid ranks every cohort member); VR keyed per
  MEMBERSHIP so a borrowed membership has its own rank. `meta.has_standing = 1`. Run flag `--no-standing`
  omits the key entirely; every pre-existing key stays byte-identical either way (the new keys are
  appended to the payload after construction; when both blocks are None nothing changes).
- **Why**: quartiles compress 24 funds into 4 buckets; the standing card needs the exact position plus
  AUM-weighted outperformance, and the peer table needs a true-rank sort.

### 4k. SIP investment mode (STANDING_SIP_DESIGN.md Feature B; engine built 2026-08-11)
A third global deck control (Lumpsum · SIP·1st-of-month · SIP·month-end): every quartile / score /
residency / standing view re-readable from SIP-based analytics. Lumpsum outputs are bit-identical.
- **Return definition**: Rs 100 per month on the first ('first') or last ('last') trading day of each
  calendar month (`sip_month_grids`, grid from the cleaned `nav.index`). Window of m months ending at t:
  installment dates = grid dates in `(t − m·months, t]` (`pd.DateOffset`, the `calendar_returns_m`
  lookback). Rated at t iff NAV_t non-null AND NAV non-null at ALL m installments AND the window holds
  exactly m grid dates — no partial SIPs, no fills beyond the panel's careful-ffill. Ranking metric =
  **SIP value ratio** `VRt = (Σ_i NAV_t/NAV_{d_i}) / m`, shipped as `VRt − 1`; with common dates and
  amounts VRt orders identically to XIRR (both monotone in the final value), so quartiles/ranks are
  exact. Vectorized in `sip_value_ratio` (reciprocal panel at grid dates, nan-aware cumsums + a valid-
  count cumsum; sum = S[j_hi]−S[j_lo] with count required == m). Verified vs a hand loop to 1e-12.
- **★ Final-month guard ('last' grid — a deliberate deviation from the contract's literal grid, flagged
  for ratification)**: the final month's "last trading day" is just the latest observed day; mid-month
  that installment hasn't happened yet, and keeping it puts an (m+1)-th phantom date in every window at
  the as-of date → the ENTIRE latest cross-section went NaN and `sip.last.sipret` came out empty (the
  contract's own sanity guard would hard-fail every mid-month build). Fix in `sip_month_grids`: the final
  month's entry is kept only when it truly is its month's last business day (`BMonthEnd().rollforward`).
  Historical months are complete and always kept; 'first' needs no guard.
- **Analytics** (`build_sip`, per convention, mirrors the lumpsum machinery): SIP return frames × 8
  windows → combined `_quartile_block(with_ranks=True)` per category (same member sets / round-up /
  stable sort, one Parallel batch per window) → month-end sampling on the lumpsum axes. Ships per VR
  membership `q1m..q60m` (q1y/q3y named like lumpsum) + `y` (SIP-1Y quartile → 5/4/3/2), `t` (SIP-3Y →
  4/3/2/1) — both ffilled before sampling exactly like s1/s3_raw — and `b` (1 iff scheme SIP-3Y VRt >
  its category benchmark's SIP-3Y VRt, benchmark SIP on `bench_cal` same grid; NaN'd after a stale
  benchmark's cap date, same >45d trigger/threshold as the lumpsum composite cap). All-peer rows ship
  `aq*` digits. SIP residency = `window_residency` on the SIP daily labels (**ALL schemes** per the scope
  rule above — widened 2026-08-11 from Top-15; both universes, VR
  home-collapsed via `_vr_daily_by_scheme`). SIP standing = rank records like §4j. A rated record
  carries ALL its fields (a window with no data = all-null array, like the lumpsum rows); a never-rated
  membership is absent.
- **sipret** (LATEST as-of only): `{scheme: {winLabel: [gainPct, xirrPct]}}`, gain = (VRt−1)·100, xirr
  from the hand-rolled solver `sip_xirr` (solve Σ100·(1+r)^((t−d_i)/365) = V_t; Newton from a CAGR-style
  guess + guaranteed bisection on the monotone bracket, bounds (−0.99, 10], no scipy; verified to 1e-6
  against an independent bisection). Historical SIP values are NOT shipped (ranks/quartiles/residency
  carry the history — same philosophy as lumpsum, which ships no returns at all).
- **JSON**: top-level `sip = {"first": {vr, allpeer, residency, standing, sipret}, "last": {...}}`,
  `meta.sip_conventions = ["first","last"]`. Run flag `--no-sip` omits the key.
- **Gates** (`_sanity_check.py`, additive, skip with [info] when the block is absent so old folders
  pass): check 7 = standing coherence at the latest as-of (≥20 random cells: ranks a permutation of
  1..n, every rank-r digit == round-up bucket of r; cells whose ranks are NOT a permutation are skipped
  as "as-of-mixed" — month-end as-of carry can legitimately mix per-scheme underlying dates, e.g.
  1-day-lagged FOFs — and more cells are drawn; <20 clean cells only warns). Check 8 = SIP structural
  (both conventions, 5 sub-blocks, digit domains, 3 sampled schemes' sipret finite & sign-consistent
  with a 0.05pp rounding band). Verified with positive AND negative controls (an injected rank swap that
  preserves the permutation is caught).

### 4l. RETURN PROFILE — the `returns` block (STANDING_SIP_DESIGN_ADDENDUM.md §C; engine built 2026-08-11)
The deck shipped RANKS but never the RETURN itself, so Scheme Detail could say "8th of 24" but not
"you made 12.4%". `build_returns` adds the numbers behind the six rows of the Return-profile panel
(lumpsum scheme · lumpsum benchmark · excess · SIP scheme · SIP benchmark · SIP excess).
- **JSON**: top-level
  `returns = {"lump": {"all": {scheme: {winLabel: {f,v}}}, "vr": {"cat|scheme": {...}}},
  "bench": {VRcategory: {winLabel: {f,v}}}, "sipgain": {conv: {"all","vr","bench"}},
  "sipdays": {conv: {winLabel: [[dayOffsets per as-of], ...]}}}`, `meta.has_returns = 1`.
  Flags: `--no-returns` omits the key; `--returns-scope lumpsum` ships only `lump`+`bench` (the size
  escape hatch). Pre-existing keys stay byte-identical (the block is appended after the payload
  literal, exactly like `standing`/`sip`).
- **Values — ONE convention per measure, no special cases**: CUMULATIVE point-to-point
  `NAV_t/NAV_{t−m} − 1`, in PERCENT, for ALL 8 windows including 3Y (whose *ranking* source is a CAGR).
  The client annualizes for display when the window ≥ 1Y (`(1+cum)^(12/m) − 1`, cumulative as subtext);
  annualizing is monotone, so a displayed value can never contradict its quartile. `sipgain` values are
  the SIP gain `(VRt − 1)·100` on the identical grids `build_sip` ranks on.
- **★ PRECISION (`LUMP_DP`/`SIPGAIN_DP`/`SIPRET_DP` = 2/4/2, ratification R2, 2026-08-11)**: `lump` and
  `bench` round to **2dp** — they are DISPLAYED exactly as shipped, so extra digits would be bytes with no
  information. `sipgain` rounds to **4dp** (raised from 2 on 2026-08-11) because the client SOLVES an XIRR
  from it, and that solve amplifies the gain's rounding by ~1/e where e is the money-weighted holding
  period in years: a 1-Month SIP whose single instalment is four days old (e ≈ 0.011y) turned ±0.005pp of
  gain into ±0.5pp of annual rate. Two extra decimals divide that by 100. **Measured on the synthetic
  smoke panel with a deliberately mid-month as-of: max round-trip error 0.4743pp at 2dp (breaching four
  windows) → 0.0060pp at 4dp (breaching none).** Cost on the real Aug-11 panel: `sipgain` 6.16 → 15.41 MB
  across both conventions (the 4dp digits are the smaller half of the sipgain growth; the all-scheme scope
  is the larger).
- **Same lookback as the ranks** (`calendar_returns_cum`): 12/36 months reuse
  `_calendar_lookback_positions` (the exact-calendar rule of `calendar_returns`); the other six delegate
  to `calendar_returns_m`. **Round-trip assert (build-stopping, tolerance 1e-9)**: annualizing the 3Y
  cumulative must reproduce `calendar_returns`' 3Y CAGR frame value-for-value AND null-for-null; 1Y is
  asserted as an identity. If it ever fires, displayed value and shipped quartile have drifted apart.
- **Same as-of sampling and trim as residency/standing**: `month_end_asof` → reindex onto the axis →
  `{f, v}` with leading/trailing all-null trimmed, `v[i−f]` = value at axis index i.
- **Axes**: `lump.all` on `aum_dates`, `lump.vr` / `bench` / `sipdays` on `months`. **Verified every
  build and printed**: `aum_dates` ⊆ `months` (2026-08-11 data: 153 ⊆ 210, TRUE), so `bench` ships once
  and the client maps by DATE STRING; if that ever fails, a second copy `bench_aum` on the all-peer axis
  is emitted and a loud warning printed (contract fallback, exercised in the smoke test).
- **Scope**: ALL schemes / ALL memberships (`top15=None`, the default) — the §4j/4k/4l scope rule above.
  The Addendum-2-vs-Addendum-3 conflict the first build flagged is **RESOLVED in favour of Addendum 3 §I.2**
  (ratification R1, 2026-08-11): 499 → 924 all-peer schemes and 359 → 511 VR memberships on the Aug-11
  panel. Passing a house list narrows it again; no CLI flag was invented. Measured cost below.
- **Benchmark** (`bench`, keyed by VR category from `exact_bench`, computed once per DISTINCT benchmark
  on `bench_cal`): a benchmark >`STALE_BENCH_DAYS` behind scheme-latest has its returns NaN'd after its
  last real observation — the same trigger/threshold as the lumpsum composite cap and the SIP `b` flag.
  Deliberately more conservative than the contract text, which only asks the CLIENT to show "—" past the
  cap; this makes a return built on a frozen index level impossible to display at all.
- **`sipdays[conv][win][asofIdx]`** = the `(asof − installment_date).days` offsets for that window at
  that as-of, grid-ascending (so strictly decreasing), computed from the SAME searchsorted/`DateOffset`
  bounds `sip_value_ratio` uses; `[]` when the window does not hold exactly m grid dates (precisely when
  nobody is rated). Shipping OFFSETS instead of a second float series lets the client solve the XIRR
  exactly for ~0.3 MB instead of ~10 MB.
- **XIRR round-trip check** (`returns_xirr_roundtrip`, printed every build): re-solves the XIRR at the
  latest as-of from `sipgain` + `sipdays` ONLY, with an independent bisection (`_xirr_from_offsets`, not
  `sip_xirr`), and compares against `sip[conv].sipret`. It makes TWO statements:
  1. **inside-the-band** (the defect detector, cannot false-alarm): every re-solve must lie inside the band
     the shipped data can physically express — re-solved at gain ± half a `SIPGAIN_DP` cell and widened by
     half a `SIPRET_DP` cell of the reference.
  2. **the contract's 0.01pp bar, now on ALL EIGHT windows** (R2; the first build could only judge the long
     ones at 2dp). It is measured NET OF THE REFERENCE'S OWN STORAGE — `max(0, |diff| − 0.5·10⁻ᴰᴾ)` with
     DP = `SIPRET_DP` = 2 — because `sipret`'s xirr is itself stored to 2dp, so demanding 0.01pp against a
     number known only to ±0.005pp would test the storage format, not the client's solve. The raw max is
     printed beside the net one, so nothing is hidden.
  **Measured on the SYNTHETIC panel (mid-month as-of, 4-day-old 1M instalment): 2dp → max net 0.4743pp,
  four windows breached; 4dp → max net 0.0060pp, none breached.**
  ★ **Measured on the REAL 2026-08-11 panel (916 schemes, 12,801 scheme-windows re-solved), 4dp: seven
  of the eight windows come in at ≤0.0008pp — and the 1-Month window breaches at 0.0331pp.** That is
  arithmetic, not a defect. The amplification for a single four-day instalment is ~91·(1+g/100)⁹⁰ points
  of annual rate per point of gain, so the worst case is the fund with the largest four-day move: HDFC
  Defence Fund's 1-Month SIP annualizes to **984.89% p.a.**, and 0.0331pp of that is 0.003% in relative
  terms — invisible to any reader, but outside an ABSOLUTE 0.01pp bar. (A >2.7% four-day move saturates
  both solvers at their shared 10.0 upper bound and agrees exactly again, so the breach lives in a narrow
  band of very large — but not absurd — short-window moves.)
  **The remedy, if the orchestrator wants a clean sweep, is one constant: `SIPGAIN_DP = 6`** — measured
  cost and effect are in the size note below. The check only REPORTS; it never stops a build, so the
  breach is information, not an outage. **Client implication (ratified as R3): on windows < 1Y show the
  SIP GAIN as the headline with the XIRR as subtext — annualizing a four-day holding period into a
  '985% p.a.' headline is arithmetic theatre at any precision.**
- **★ SIZE (required deliverable) — MEASURED, not modelled.** Re-run on the REAL 2026-08-11 panel
  (`build_returns` and `window_residency` against that refresh's own data, on its own `months`/`aum_dates`
  axes; no publish, nothing written into its `out/`). Compact-JSON bytes:

  | sub-block | wave-1 shipped (Top-15, sipgain 2dp) | NOW (all schemes, sipgain 4dp) |
  |---|---:|---:|
  | `residency` (top level) | 8.23 MB | **12.07 MB** |
  | `sip.first.residency` | 8.41 MB | ≈12.33 MB *(projected at the measured 1.4658× factor)* |
  | `sip.last.residency` | 8.41 MB | ≈12.33 MB *(same)* |
  | `standing` (already all-scheme) | 2.64 MB | 2.64 MB (unchanged) |
  | `returns.lump.all` | 2.13 MB | **3.34 MB** |
  | `returns.lump.vr` | 1.77 MB | **2.37 MB** |
  | `returns.bench` | 0.29 MB | 0.29 MB (unchanged) |
  | `returns.sipgain.first` | 4.09 MB | **7.84 MB** |
  | `returns.sipgain.last` | 4.00 MB | **7.57 MB** |
  | `returns.sipdays` | 0.23 MB | 0.23 MB (unchanged) |
  | **`returns` TOTAL** | **12.52 MB** | **21.64 MB** |

  The `returns` growth splits cleanly: **+5.29 MB from the R1 scope widening** and **+3.83 MB from the R2
  4dp gain** (at Top-15 scope the 4dp alone would have cost +2.69 MB). Record counts: `lump.all`
  499 → 924 schemes, `lump.vr` 359 → 511 memberships; `residency` 499 → 924 all-peer and 331 → 473 VR
  scheme keys (i.e. exactly the whole universe — 924 is `len(allpeer)` and 473 is the distinct-scheme
  count behind the 511 VR memberships). **Additivity verified on the real panel: recomputing residency at
  the Top-15 scope reproduces the SHIPPED block byte-for-byte (8,231,223 bytes both), and every one of
  those records is byte-identical inside the widened block — the widening adds 425 all-peer and 142 VR
  keys and changes nothing that existed.**
  **Whole-payload effect: `dashboard_data.json` 63.70 → ≈97.0 MB; the OFFLINE single-file deck
  66.39 → ≈99.7 MB.** The hosted path is lazy + gzipped and unaffected in practice; the offline deck is
  the constraint and this is the number the Addendum-3 contract demands be reported rather than
  silently trimmed. Levers, NONE of them applied (a deck that carries different data from the hosted page
  is exactly what this project forbids, so this is Kyser's call): `--returns-scope lumpsum` ships
  `lump`+`bench` only (6.0 MB instead of 21.6); the residency/returns scope narrows again with one house
  list; or residency drops its longest windows. The engine prints the exact table every run.
- **`SIPGAIN_DP` — the precision/size/accuracy curve, measured end to end on the real panel** (all-scheme
  scope throughout; "max net" = the XIRR round-trip's contract statistic, worst over all 8 windows):

  | `SIPGAIN_DP` | `sipgain.first` | `sipgain.last` | `returns` TOTAL | max net (pp) | windows breaching 0.01pp |
  |---:|---:|---:|---:|---:|---|
  | 2 (wave 1) | 5.85 MB | 5.73 MB | 17.81 MB | 2.8335 | 1M, 3M, 6M, 9M, 1Y |
  | **4 (shipped)** | **7.84 MB** | **7.57 MB** | **21.64 MB** | **0.0331** | **1M only** |
  | 6 | 9.85 MB | 9.42 MB | 25.50 MB | 0.0001 | none |

  So 4dp buys an ~86× tightening for +3.83 MB, and 6dp would close the last window for a further
  +3.86 MB (offline deck ≈103.6 MB instead of ≈99.7 MB). 4dp is what R2 ratified and what ships.
- **★ `sipgain` values are NOT byte-identical to wave 1's — by design, and the difference is DOUBLE
  ROUNDING, not a moved measurement.** 9,430 of 2,011,052 real values (0.47%) disagree with the 2dp build
  after re-rounding to 2dp. Verified exhaustively: **every one of the 9,430 differs by exactly one 2dp cell
  (0.01) and in every single case the 4dp value sits exactly on the `.xx5` rounding tie** (e.g. 0.865 vs
  0.87, −2.975 vs −2.97, −0.495 vs −0.50). That is the textbook signature of `round(round(x,4),2)` vs
  `round(x,2)`, it occurs at the theoretically expected ~0.5% rate, and the 4dp value is the one closer to
  the true gain. The regression bar "pre-existing keys byte-identical" therefore does NOT apply to
  `sipgain` — R2 deliberately changes those values. Every other pre-existing key still holds it (residency
  was re-verified byte-for-byte on the real panel).
- **Gate**: `_sanity_check.py` check 9 (HARD, skips with `[info]` when absent) — inside one
  category/window/as-of, a strictly higher shipped return may never carry a worse quartile digit.
  Ties at 2dp are not compared (round-up bucketing legitimately splits a genuine tie); cells whose
  cross-section is not uniform are skipped as "as-of-mixed" using the same permutation gate as check 7
  (fallback when `standing` is absent: the digit histogram must equal the round-up bucket sizes).
  Verified with a positive control and a negative control (an injected return swap is caught).
- **Known residual**: `sipgain` is month-end AS-OF carried like every other series, but `sipdays` is
  measured to the as-of date, so a scheme with no NAV on the as-of day (1-day-lagged overseas FoFs) has
  offsets ~1 day out — negligible on long windows, visible on 1M. The engine COUNTS and prints those
  scheme-windows ("SIP as-of carry N") rather than hiding them.

## 5. Open / parked
- Auto-fetch of MFI data (analogous to NSE-Indices fetch in the FFT project) — AFTER this run.
- Web fetch for NASDAQ-100 TRI & S&P Global 1200 TRI (replace stale Bloomberg) — parked, save for later.
- **`_app.js` is a STALE working copy** — the authoritative dashboard app code lives INLINE in
  `_dashboard_src.html` (the bake reads only that). Edit `_dashboard_src.html`; re-sync or delete
  `_app.js` in a future cleanup so it can't mislead.

## 7b. Benchmark staleness handling (finalized 2026-06-17)
`align_benchmarks` now CARRIES benchmarks forward (ffill) onto the scheme calendar, so the +1 bonus is computable at every date using the latest available benchmark level (correct for periodic indices like monthly CRISIL Hybrid). A category's composite is CAPPED (set NaN after its benchmark's last real date) only when that benchmark is **>`STALE_BENCH_DAYS` (45) behind scheme-latest** — i.e. genuinely stale. As of June data this caps exactly: Pure International & Nasdaq-100-FOFs (2026-03-11), Multi-Asset Allocation (2025-11-28, broken composite benchmark). Monthly CRISIL/Silver and 1-day-lagged benchmarks stay current. Month-end outputs use `month_end_asof` (convention restated below). Excel files default to month-end resolution. Verified Aditya VR composite as-of 06-16: Q1 41.5/Q2 15.5/Q3 27.9/Q4 5.2, coverage 90.2%.

An EMPTY benchmark is a **different** defect from a stale one and is NOT capped — there is no date to cap at, so the +1 bonus is simply never earned and the category is scored peer-only for its whole history. Since 2026-08-11 `run()` step [7/8] discloses both through one generic scan (`benchmark_exclusions`), so §4f-ter requirement 2's "Make-in-India empty benchmark" line is actually emitted, and so a benchmark that empties next month discloses itself instead of staying invisible. On the August 2026 data the empty set is **Make in India** (11 VR peers) and **Equity Savings** (8) — both mapped to a benchmark column the vendor ships with no NAV in it at all; the tab labels them `category (no benchmark NAV)` against the stale case's `category (VR composite cap)`. Detection is on the resolved frame (no column, or a column whose `last_real` is NaT), never on a list of category names. Categories in `AVOID_EXACT_CATS` are skipped because they are never scored on an exact peer set at all.

**`month_end_asof` — corrected 2026-08-11. The CODE (`peer_monitor.month_end_asof`) is authoritative; this text was wrong and is being fixed after an independent re-implementation followed it literally and produced different numbers.** The sentence used to read *"last non-NaN ≤ month-end per column"*, which says the value may come from ANY earlier date. It may not. What the implementation actually does: group the frame's rows by the (year, month) **of their own date**, forward-fill inside that group only, take the group's **last row**, and label it with the **last date actually present in that month** — so a part-month is stamped with the last data date it has (August 2026 → `2026-08-07`), not with the calendar month-end. Three consequences that the old wording lost:
- **The carry is intra-month only.** The value shown at a month's as-of date must come from an observation **inside that same calendar month**; it is never carried across a month boundary.
- **A column with no observation in a month stays NaN for that month** — it does not repeat the previous month's value.
- **A month with no rows in the frame produces no output row at all** (the axis is the months the data has, not a synthetic calendar).

This is material, not cosmetic. Under the literal "last non-NaN ≤ month-end" reading a scheme's return would be carried straight across a NAV blackout: the 2026-08-11 verifier found ABSL Multi-Asset Omni FOF would gain **6 fabricated 1-Month as-of values across a 171-day gap**, bridging a 16.99 → 9.97 NAV reset as if it were a real return. What the implementation *does* absorb is the intended case only — a 1-day reporting lag, or a series that stops mid-month (capped categories, discontinued schemes) still showing its last real value at that month's as-of date, and NaN from the next month on.

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
