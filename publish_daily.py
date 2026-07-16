r"""DAILY publish — refresh ONLY the live page files, no dated-folder copy.

Writes, from a refresh folder's outputs, the three files the hosted page serves:
  index.html      = the dashboard template SHELL (data placeholder -> null)
  peer_data.json  = dashboard_data.json minus residency (page fetches it, gzipped by Pages)
  residency.json  = the residency block (page lazy-fetches on demand)

The dated ARCHIVE copy (whole folder) is publish_refresh.py's job (weekly). This script
exists so the DAILY cron can update the live deck without growing the repo by a folder
per day. Caller does the git add/commit/push (DAILY_REFRESH.ps1 amends a rolling
"daily deck" commit so history stays small).

Usage: python publish_daily.py "<path to refresh folder>"
"""
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent


def main():
    if len(sys.argv) < 2:
        sys.exit('usage: python publish_daily.py "<path to refresh folder>"')
    src = Path(sys.argv[1]).resolve()
    tpl = src / "dashboard.html"                      # template copy WITH __PEER_DATA__
    dj = src / "out" / "dashboard_data.json"
    if not tpl.exists() or "__PEER_DATA__" not in tpl.read_text(encoding="utf-8", errors="ignore")[:5_000_000]:
        sys.exit(f"no data-placeholder template at {tpl} — run the engine package copy first")
    if not dj.exists():
        sys.exit(f"no {dj} — run the engine first")

    sys.path.insert(0, str(REPO))
    from _page_speed import extract_fonts, pack_data, pack_residency

    data = json.loads(dj.read_text(encoding="utf-8"))
    residency = data.pop("residency", None)
    data["residency"] = None
    data = pack_data(data)                    # ~4x smaller month-series, page unpacks at boot
    (REPO / "peer_data.json").write_text(json.dumps(data, separators=(",", ":")), encoding="utf-8")
    (REPO / "residency.json").write_text(json.dumps(pack_residency(residency), separators=(",", ":")),
                                         encoding="utf-8")
    html = tpl.read_text(encoding="utf-8").replace("__PEER_DATA__", "null", 1)
    html = extract_fonts(html, REPO)          # ~2 MB of base64 fonts -> cached fonts/*.woff2
    (REPO / "index.html").write_text(html, encoding="utf-8")
    print(f"live page refreshed from {src.name}: index.html (fonts split) + peer_data.json + residency.json (packed)")


if __name__ == "__main__":
    main()
