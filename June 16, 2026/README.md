# Peer Performance Monitor — Automation Package

Replaces the manual monthly workflow:
old: notebook → copy-paste to Excel → format → email
new: `python generate_report.py` → `out/` folder ready to publish + share

**One command does everything** — ingest MFI + VR data, run the quartile/scoring
analysis, and write the Excel files, the data JSON, and *both* dashboard files.
The dashboard template is embedded inside `generate_report.py`, so the script is
fully self-contained; no other files are required.

## What you get

Every run produces these in `out/`:

| File | Purpose |
|------|---------|
| `percent AUM in Q1 to Q4 - <date>.xlsx` | Internal sharing — same shape as today's file (Sleevewise + AMCwise — All Peers + AMCwise — VR Peers) |
| `Scheme Scoring on Exact Peer Set - Calendar 1Y 3Y - <date>.xlsx` | Internal sharing — 1Y score, 3Y score, composite daily score, peer mapping |
| `dashboard.html` + `dashboard_data.json` | Interactive web report. Drop both on any intranet path, share the URL |
| `dashboard_offline.html` | Single self-contained file — JSON inlined, double-click to open. Email this to management |

## Folder layout

```
project/
├── generate_report.py           ← run this every month (dashboard template embedded inside)
├── embed_data.py                ← optional helper (only to re-bundle a hand-edited template)
├── MFI Data/
│   ├── HistoricalNav_*.xlsx
│   ├── Scheme wise AUM Report-*.xlsx
│   └── Map.xlsx
├── Value Research Data/
│   └── Value Research Benchmarks NAV <date>.xlsx
└── out/                         ← created by the script
    ├── percent AUM in Q1 to Q4 - <date>.xlsx
    ├── Scheme Scoring on Exact Peer Set - <date>.xlsx
    ├── dashboard_data.json
    ├── dashboard.html           ← interactive report (host on intranet)
    └── dashboard_offline.html   ← single self-contained file (email / double-click)
```

The only file you maintain is `generate_report.py`. You do **not** need a
separate `dashboard.html` to run it — the template lives inside the script and
is written out fresh on every run.

## Monthly workflow

1. **Refresh raw data.** Drop the new MFI NAV/AUM/Map files into `MFI Data/`. Drop the new Value Research file into `Value Research Data/`. The script picks up the latest VR file automatically.
2. **Run:** `python generate_report.py` (~3–10 min depending on data size and cores; you'll see step-by-step progress).
3. **Publish / share.** Either host `dashboard.html` + `dashboard_data.json` on the intranet, or just email `dashboard_offline.html`. Email the two `.xlsx` files from `out/` as before.

## Hosting the dashboard — three modes

Pick whichever matches your environment:

**Mode A — Intranet HTTP (recommended).** Drop `dashboard.html` + `dashboard_data.json` onto any web folder on your intranet (SharePoint, IIS, Apache, even a Python server). Share the URL. Multiple users hit it; each browser caches `dashboard_data.json` separately.

**Mode B — Quick local test.** From inside `out/`:
```bash
python -m http.server 8000
```
Then open `http://localhost:8000/dashboard.html` in any browser.

**Mode C — Single-file (for email or shares without HTTP).** Just use `dashboard_offline.html` — the script already builds it (~1.5 MB, JSON inlined). A single file anyone can double-click to open; no server, no separate JSON. This is what you send to senior management.

> `embed_data.py` is only needed if you hand-edit `dashboard.html` and want to
> re-bundle it without re-running the full pipeline. Normal runs never need it.

## Chart.js — CDN vs offline

The dashboard pulls Chart.js from `cdn.jsdelivr.net` once on page load (~150 KB, cached after first visit). If your intranet blocks external CDNs, download Chart.js once and self-host:

```bash
# One-time download — keep next to dashboard.html
curl -o chart.umd.min.js https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js
curl -o chartjs-adapter.js https://cdn.jsdelivr.net/npm/chartjs-adapter-date-fns@3.0.0/dist/chartjs-adapter-date-fns.bundle.min.js
```

Then in `dashboard.html`, replace the two `<script src="https://cdn..."` lines near the top with `<script src="chart.umd.min.js"></script>` and `<script src="chartjs-adapter.js"></script>`. To make the change permanent across runs, edit a copy of `dashboard.html`, then either re-embed it into the script (base64) or pass it via `--dashboard-html` (see below).

## CLI options

```bash
python generate_report.py \
  --mfi "MFI Data" \
  --vr "Value Research Data" \
  --out "out" \
  --dashboard-html "my_custom_dashboard.html"   # optional
```

Defaults match the folder layout above; only override if your paths differ.
`--dashboard-html` is optional — omit it to use the dashboard template built
into the script. Supply it only if you've customised the template by hand.

## What the dashboard shows

Five tabs, all filterable, all exportable to CSV. Every visualization carries a
short **methodology note** explaining exactly how it's computed, and all axes
**auto-scale** as you change filters.

1. **Sleeve View** — quartile distribution *within* the chosen sleeve. The four bands are normalised by the sleeve's total weight, so they sum to 100% (e.g. if Bottoms Up is 11% of the house and half of that is Q1+Q2, you see 50% — not 5.5%). The sleeve's raw weight in the house is plotted on a **secondary axis**. Plus a top-half + quality-score trend, a **ranking table vs Top-15 with a date picker** (re-rank on any month-end), and a cross-sectional bar. Filters: Fund House × Breakdown × Rolling Window (1Y/3Y) × basis toggle (Within sleeve / % of AMC AUM).
2. **AMC Roll-up** — house-wide quartile mix (already sums to 100%), Aditya Birla vs Top-15 Q1+Q2 over time, and a **league table with a date picker**. Toggle All Peers (MFI) vs Exact Peers (VR).
3. **Scheme Detail** — pick a category, then a scheme. Composite + 1Y + 3Y score history, plus a **peer-comparison table with a date picker** (ranks every scheme in the exact-peer category on any month-end; defaults to the latest *populated* date).
4. **Quartile Matrix** — a house × sleeve heatmap (green = strong, red = weak) with a metric toggle (Q1+Q2 share / AUM-weighted quality score / Q1 share), rolling window, and date picker. One-glance read of where Aditya Birla leads or lags across every sleeve.
5. **Peer Mapping** — searchable table of the VR exact-peer mapping (scheme → category → MFI name).

**Quality score** (used in several places) is the AUM-weighted mean quartile on a 1–4 scale: Q1=4, Q2=3, Q3=2, Q4=1. A higher number means more of the house's money sits with top-performing schemes.

## Dependencies

```bash
pip install pandas numpy openpyxl joblib python-calamine
```

`python-calamine` is optional but makes Excel reads ~10× faster. If you don't have it, edit `generate_report.py`: change `ENGINE = "calamine"` to `ENGINE = "openpyxl"` at the top.

## Adapting the dashboard

The dashboard is a single HTML file using Chart.js (CDN), embedded inside
`generate_report.py` as a base64 blob. To restyle it:

1. Run the script once and grab `out/dashboard.html` (or decode the
   `DASHBOARD_TEMPLATE_B64` blob at the bottom of the script).
2. Edit it:
   - **Brand colors** are CSS variables at the top: `--absl` is the ABSL highlight red, `--q1/q2/q3/q4` are the quartile colors, `--paper` is the background. Change them once, everything updates.
   - **Fonts** load from Google Fonts (Fraunces, Inter, JetBrains Mono). Swap in your own.
   - **Tabs / filters / charts** are vanilla JS at the bottom of the file. No build step.
3. Use your edited file via `python generate_report.py --dashboard-html my_dashboard.html`, or re-embed it permanently (base64-encode and replace the blob in the script).

## Customizing the analysis

Most knobs live near the top of `generate_report.py`:

- `TOP15` — fund houses included in sleeve/AMC tables
- `BREAKDOWN_RULES` — how Scheme Sub Natures roll up into Breakdowns
- `MULTI_ASSET_DROP` — schemes excluded due to debt taxation
- `composite_score(w_1y, w_3y)` — change the 40/60 weights if needed
- Inside `_process_category` — the quartile-splitting logic and distance metrics

The structure mirrors the original notebook section-for-section, so you can map any specific calculation back to the source.

## Notes

- **Date granularity.** The Excel files keep daily resolution like today. The dashboard JSON samples to month-end snapshots — daily values barely change since AUM is monthly, and this keeps the JSON ~1.5 MB instead of ~50 MB. Edit `month_end_of()` in `write_dashboard_json()` if you want different granularity.
- **Missing data.** Schemes with no NAV in MFI but present in VR are reported and dropped at load time. New launches (in NAV but not yet in AUM or Map) are dropped from the universe.
- **Performance.** The quartile analysis is parallelized across categories using `joblib`. On a 16-core machine it takes ~2–3 minutes for the full universe.
