"""
embed_data.py  (optional helper)
================================
NOTE: generate_report.py already produces dashboard_offline.html automatically,
so you normally do NOT need this script. Use it only to (re)bundle a hand-edited
dashboard.html with an existing dashboard_data.json — e.g. after tweaking the
template's styling by hand without re-running the full pipeline.

Bundles dashboard_data.json into dashboard.html so the result is a single,
self-contained file that anyone can open by double-clicking (no HTTP server,
no separate JSON file).

Run:
    python embed_data.py
    python embed_data.py --html dashboard.html --json dashboard_data.json --out dashboard_offline.html
"""

import argparse
import json
import re
from pathlib import Path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--html", default="dashboard.html", help="Source dashboard HTML")
    ap.add_argument("--json", default="dashboard_data.json", help="Data JSON to inline")
    ap.add_argument("--out", default="dashboard_offline.html", help="Output single-file HTML")
    args = ap.parse_args()

    html_path = Path(args.html)
    json_path = Path(args.json)
    out_path = Path(args.out)

    if not html_path.exists():
        raise SystemExit(f"Source HTML not found: {html_path}")
    if not json_path.exists():
        raise SystemExit(f"Data JSON not found: {json_path}")

    html = html_path.read_text(encoding="utf-8")
    data = json.loads(json_path.read_text(encoding="utf-8"))

    # Re-serialize compactly (no spaces) — keeps file size minimal
    payload = json.dumps(data, separators=(",", ":"))

    # Find the placeholder script tag and inject
    pattern = r'(<script id="embedded-data" type="application/json">)([^<]*)(</script>)'
    if not re.search(pattern, html):
        raise SystemExit("Could not find <script id='embedded-data'> placeholder in HTML")

    new_html = re.sub(pattern, lambda m: m.group(1) + payload + m.group(3), html, count=1)
    out_path.write_text(new_html, encoding="utf-8")

    src_mb = html_path.stat().st_size / 1e6
    json_mb = json_path.stat().st_size / 1e6
    out_mb = out_path.stat().st_size / 1e6
    print(f"  Source HTML : {src_mb:.2f} MB")
    print(f"  Data JSON   : {json_mb:.2f} MB")
    print(f"  Output      : {out_mb:.2f} MB → {out_path.resolve()}")
    print(f"\n✓ Single-file dashboard ready. Double-click to open.")


if __name__ == "__main__":
    main()
