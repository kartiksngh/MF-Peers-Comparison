r"""DAILY publish — refresh ONLY the live page files, no dated-folder copy.

Writes, from a refresh folder's outputs, the files the hosted page serves:
  index.html      = the dashboard template SHELL (data placeholder -> null)
  peer_data.json  = dashboard_data.json minus the lazy blocks (page fetches it, gzipped by Pages)
  residency.json  = the residency block (page lazy-fetches on demand)
  standing.json   = the category-standing rank block (lazy; written only when the engine
                    shipped a `standing` key — old engine outputs publish fine without it)
  sip_<book>.json = one file per SIP INVESTOR BOOK — sip_sip3y.json / sip_sip5y.json (lazy;
                    same tolerance — written only when the engine shipped a `sip` key). These
                    replaced sip_first.json / sip_last.json on 2026-08-13 when the two SIP
                    instalment-day conventions became two investor books; the retired files are
                    DELETED from the repo below so the live site cannot keep serving a block the
                    page no longer knows how to read.
  returns.json    = the return-profile block — lumpsum/benchmark/SIP-gain series + the SIP
                    day-offset grids the page solves XIRR from (lazy; same tolerance —
                    written only when the engine shipped a `returns` key)

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
    from _page_speed import (extract_fonts, pack_data, pack_residency, pack_returns,
                             pack_sip, pack_standing)

    data = json.loads(dj.read_text(encoding="utf-8"))
    residency = data.pop("residency", None)
    data["residency"] = None
    # Standing + SIP + returns (2026-08-11 features) get the same lazy-file treatment as
    # residency: popped out of peer_data.json, served as their own fetch-on-demand files.
    # Old engine outputs have NONE of those keys — then nothing new is written and the keys
    # stay absent (only when a block was actually popped does peer_data.json carry the null
    # marker, which is what the page gates its fetch on). Every marker is set BEFORE the
    # json.dumps below, so the shipped peer_data.json always matches the files written.
    standing = data.pop("standing", None)
    sip = data.pop("sip", None)
    returns = data.pop("returns", None)
    if standing is not None:
        data["standing"] = None               # page lazy-fetches standing.json
    if sip is not None:
        data["sip"] = None                    # page lazy-fetches sip_<book>.json
    if returns is not None:
        data["returns"] = None                # page lazy-fetches returns.json
    data = pack_data(data)                    # ~4x smaller month-series, page unpacks at boot
    (REPO / "peer_data.json").write_text(json.dumps(data, separators=(",", ":")), encoding="utf-8")
    (REPO / "residency.json").write_text(json.dumps(pack_residency(residency), separators=(",", ":")),
                                         encoding="utf-8")
    written = ["peer_data.json", "residency.json"]
    if standing is not None:
        (REPO / "standing.json").write_text(json.dumps(pack_standing(standing), separators=(",", ":")),
                                            encoding="utf-8")
        written.append("standing.json")
    if sip is not None:                       # one file per investor book, each its own lazy fetch
        for bkey in [k for k in sip if isinstance(sip.get(k), dict) and "sipret" in sip[k]]:
            (REPO / f"sip_{bkey}.json").write_text(
                json.dumps(pack_sip(sip.get(bkey)), separators=(",", ":")), encoding="utf-8")
            written.append(f"sip_{bkey}.json")
        # Retire the pre-2026-08-13 instalment-day files. Leaving them behind would leave ~32 MB
        # of a block on the live site that nothing fetches, and — worse — a stale deck could still
        # find and read them. Deleting is safe: they are regenerated from the engine, never edited.
        for old in ("sip_first.json", "sip_last.json"):
            f = REPO / old
            if f.exists():
                f.unlink()
                written.append(f"-{old}")
    if returns is not None:                   # return profile: one lazy file, same envelope
        (REPO / "returns.json").write_text(json.dumps(pack_returns(returns), separators=(",", ":")),
                                           encoding="utf-8")
        written.append("returns.json")
    html = tpl.read_text(encoding="utf-8").replace("__PEER_DATA__", "null", 1)
    html = extract_fonts(html, REPO)          # ~2 MB of base64 fonts -> cached fonts/*.woff2
    (REPO / "index.html").write_text(html, encoding="utf-8")
    print(f"live page refreshed from {src.name}: index.html (fonts split) + " + " + ".join(written) + " (packed)")


if __name__ == "__main__":
    main()
