# MF Peer Quartile Monitor — Aditya Birla Sun Life

A self-contained, offline dashboard that ranks Indian mutual-fund schemes and fund houses
into performance quartiles (Q1 best … Q4 worst) on **1-year** and **3-year** returns, and
shows where **Aditya Birla Sun Life (ABSL)** stands versus the **Top-15** fund houses.

**▶ Live dashboard:** https://kartiksngh.github.io/MF-Peers-Comparison/ (always shows the latest refresh)

This refresh is as of **16 June 2026**. The deck is **interactive** — every snapshot view has a
time-slider with play / reverse / pause / step and an animated-**GIF export**, and the Composite
bar has per-quartile show/hide toggles. To update the live link with a new refresh, see
**`HOW_TO_PUBLISH.md`**.

Two peer universes are shown side by side:

- **All Peers (MFI)** — every scheme in a category; raw 1Y & 3Y return quartiles.
- **Exact Peers (VR)** — Value Research's hand-picked competitor set; a composite score
  (`0.8·1Y + 0.2·3Y`, with `+1` for beating the scheme's benchmark on 3Y), re-bucketed Q1–Q4.

## Contents

```
index.html              ← the LIVE dashboard (copy of the latest refresh's dashboard_offline.html)
publish_refresh.py      ← stages a new refresh + resets index.html to the latest (see HOW_TO_PUBLISH.md)
HOW_TO_PUBLISH.md       ← how to update the live link
June 16, 2026/          ← this refresh (archived) — self-contained
├── peer_monitor.py     ← the engine (one file; reads Data/, writes out/)
├── dashboard.html      ← baked offline template (__PEER_DATA__ + inlined Chart.js, html2canvas, gif.js, slider/GIF engine)
├── make_notebook.py    ← regenerates the audit notebook from the engine (no drift)
├── embed_data.py       ← helper to inject data into the template
├── BUILD_SPEC.md       ← the methodology / audit contract (read first)
├── Peer Performance Monitor.ipynb   ← audit view (same logic as the engine)
├── Data/               ← raw vendor inputs for this refresh (NAV, AUM, mapping, benchmarks)
└── out/                ← deliverables
    ├── dashboard_offline.html        ← self-contained offline deck (double-click to open)
    ├── dashboard.html                ← identical self-contained deck
    ├── dashboard_data.json           ← the underlying data (for audit / other tools)
    ├── percent AUM in Q1 to Q4 - June 16, 2026.xlsx
    └── Scheme Scoring on Exact Peer Set - Calendar 1Y 3Y - June 16, 2026.xlsx
```

## Rebuild

```bash
cd "June 16, 2026"
python peer_monitor.py --data Data --out out
```

Outputs land in `out/`: two Excel files plus `dashboard.html` / `dashboard_offline.html`
— both fully offline (open by double-click, no server, no internet).

## Note

The dashboard is **self-contained**: all chart data is inlined in the HTML. Methodology
and the exact scoring/staleness rules are in `June 16, 2026/BUILD_SPEC.md`.
