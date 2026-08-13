"""
publish_refresh.py  —  stage a Peer Quartile Monitor refresh for GitHub Pages.

Run it from inside this repo (MF-Peers-Comparison) and point it at the WORKING project's
dated refresh folder (the one you just ran the engine in). It:
  1. copies that dated folder into this repo (kept as an archive),
  2. drops scratch/build artifacts from the copy,
  3. refreshes the fixed landing page  ->  root  index.html = that refresh's offline deck,
  4. runs `git add -A` so you can review, commit, and push.

The live link  https://kartiksngh.github.io/MF-Peers-Comparison/  serves root index.html,
so it ALWAYS shows the LATEST refresh — no URL change per month. Dated folders accumulate
as history; index.html is overwritten with the newest deck each time.

  python publish_refresh.py "C:/Users/Administrator/Documents/Projects/Peer NAV Quartiles comparison/2026/August 11, 2026"

Then publish (public MF data — just a deliberate publish step; Claude or KV can run the push):
  git -C "<this repo>" commit -m "August 11, 2026 deck"
  git -C "<this repo>" push origin main
"""
import shutil, sys, subprocess
from pathlib import Path

REPO = Path(__file__).resolve().parent
CRUFT_FILES = {"_template.html", "assemble.js", "verify_embedded.js"}
CRUFT_GLOBS = ["*- Copy.html", "* - Copy.html", "sample_*animation.gif", "dashboard_animated.html",
               "BirlaMFPR*.xls*",   # internal ABSL/VR competition workbook — never publish
               "MF Data*.xls*",     # vendor portfolio-holdings dump (NAVIndia) — never publish
               "_engine_run.log"]
CRUFT_DIRS  = {"_shots", "assets", "_verify", "__pycache__", ".ipynb_checkpoints"}


def main():
    if len(sys.argv) < 2:
        sys.exit('usage: python publish_refresh.py "<path to working refresh folder>"')
    src = Path(sys.argv[1]).resolve()
    deck = src / "out" / "dashboard_offline.html"
    if not deck.exists():
        sys.exit(f"no out/dashboard_offline.html under {src} — run the engine first")

    dest = REPO / src.name                       # e.g. "August 11, 2026"
    if dest.exists():
        shutil.rmtree(dest)
    shutil.copytree(src, dest)

    # strip scratch from the copied refresh folder
    for d in CRUFT_DIRS:
        for p in list(dest.rglob(d)):
            if p.is_dir():
                shutil.rmtree(p, ignore_errors=True)
    for pat in list(CRUFT_FILES) + CRUFT_GLOBS:
        for p in dest.rglob(pat):
            try: p.unlink()
            except OSError: pass

    # the fixed landing page = the newest deck -> link always shows the latest refresh.
    # HOSTED SPLIT (2026-07-16, page-speed): instead of the 20+ MB self-contained file,
    # index.html = the template SHELL (data placeholder -> null) and the data ships as
    # separate JSON files the page fetches (gzipped by GitHub Pages; residency — the
    # biggest block, Scheme-Detail-only — loads lazily; standing, the two sip convention
    # blocks and the return profile, added 2026-08-11, load the same lazy way and are
    # written only when the engine output actually carries them, so OLD refresh folders
    # still publish clean).
    # The emailed offline deck in the dated folder stays fully self-contained and untouched.
    import json
    tpl = dest / "dashboard.html"                       # template copy WITH __PEER_DATA__
    if tpl.exists() and "__PEER_DATA__" in tpl.read_text(encoding="utf-8", errors="ignore")[:5_000_000]:
        import sys as _sys
        _sys.path.insert(0, str(REPO))
        from _page_speed import (extract_fonts, pack_data, pack_residency, pack_returns,
                                 pack_sip, pack_standing)
        data = json.loads((dest / "out" / "dashboard_data.json").read_text(encoding="utf-8"))
        residency = data.pop("residency", None)
        data["residency"] = None                        # page lazy-fetches residency.json
        # Standing + SIP + returns: same pop-and-lazy-file treatment as residency (mirrors
        # publish_daily.py exactly). Absent on old outputs -> nothing written, keys stay
        # absent; the null markers go in ONLY when a block was actually popped, and always
        # before the json.dumps below.
        standing = data.pop("standing", None)
        sip = data.pop("sip", None)
        returns = data.pop("returns", None)
        if standing is not None:
            data["standing"] = None                     # page lazy-fetches standing.json
        if sip is not None:
            data["sip"] = None                          # page lazy-fetches sip_first/sip_last.json
        if returns is not None:
            data["returns"] = None                      # page lazy-fetches returns.json
        (REPO / "peer_data.json").write_text(json.dumps(pack_data(data), separators=(",", ":")),
                                             encoding="utf-8")
        (REPO / "residency.json").write_text(json.dumps(pack_residency(residency), separators=(",", ":")),
                                             encoding="utf-8")
        written = ["peer_data.json", "residency.json"]
        if standing is not None:
            (REPO / "standing.json").write_text(json.dumps(pack_standing(standing), separators=(",", ":")),
                                                encoding="utf-8")
            written.append("standing.json")
        if sip is not None:                             # one file per convention, each its own lazy fetch
            (REPO / "sip_first.json").write_text(json.dumps(pack_sip(sip.get("first")), separators=(",", ":")),
                                                 encoding="utf-8")
            (REPO / "sip_last.json").write_text(json.dumps(pack_sip(sip.get("last")), separators=(",", ":")),
                                                encoding="utf-8")
            written += ["sip_first.json", "sip_last.json"]
        if returns is not None:                         # return profile: one lazy file
            (REPO / "returns.json").write_text(json.dumps(pack_returns(returns), separators=(",", ":")),
                                               encoding="utf-8")
            written.append("returns.json")
        html = tpl.read_text(encoding="utf-8").replace("__PEER_DATA__", "null", 1)
        html = extract_fonts(html, REPO)
        (REPO / "index.html").write_text(html, encoding="utf-8")
        print("index.html = fast shell (fonts split); data packed into " + " + ".join(written))
    else:                                               # old refreshes: self-contained fallback
        shutil.copy2(dest / "out" / "dashboard_offline.html", REPO / "index.html")

    # ── the self-contained decks are no longer archivable (2026-08-11) ─────────────────
    # GitHub's hard limit is 100 MB/file and this script refuses anything over 95 MB. The
    # dated folder's two self-contained decks (dashboard.html and dashboard_offline.html —
    # identical files) used to be ~10 MB; once the category-standing, SIP and return-profile
    # blocks landed they inline ~97 MB of JSON and reach ~100 MB, which would abort the whole
    # weekly archive. They are also the ONE archived artefact that is fully regenerable: the
    # folder keeps Data/, the engine, the template and out/dashboard_data.json, so
    # `python peer_monitor.py --data Data --out out` reproduces both decks exactly.
    # So drop them from the ARCHIVE COPY only (this happens AFTER the fallback branch above
    # has already used dashboard_offline.html, and the live page is built from the template +
    # packed JSON, never from these files) and leave a note saying how to rebuild.
    # Any OTHER oversized file still aborts — that is what the guard is for.
    pruned = []
    for name in ("dashboard.html", "dashboard_offline.html"):
        p = dest / "out" / name
        if p.is_file() and p.stat().st_size > 95 * 1024 * 1024:
            pruned.append(f"{name} ({p.stat().st_size / 1e6:.0f} MB)")
            p.unlink()
    if pruned:
        (dest / "out" / "DECK_NOT_ARCHIVED.txt").write_text(
            "The self-contained offline deck(s) were NOT archived here:\n  "
            + "\n  ".join(pruned)
            + "\n\nThey exceed the 95 MB per-file publish limit since the standing/SIP/return\n"
              "blocks were added. Everything needed to rebuild them byte-for-byte is in this\n"
              "folder — from inside it run:\n\n    python peer_monitor.py --data Data --out out\n\n"
              "The live site is unaffected: it is served from index.html + the packed JSON\n"
              "files in the repo root, which are built from this folder's dashboard.html\n"
              "template and out/dashboard_data.json.\n", encoding="utf-8")
        print("archive: skipped oversized self-contained deck(s): " + ", ".join(pruned)
              + " (see out/DECK_NOT_ARCHIVED.txt — they are regenerable)")

    big = [p for p in REPO.rglob("*")
           if p.is_file() and ".git" not in p.parts and p.stat().st_size > 95 * 1024 * 1024]
    if big:
        sys.exit("ABORT — file(s) over 95 MB (GitHub limit):\n  " + "\n  ".join(map(str, big)))

    subprocess.run(["git", "-C", str(REPO), "add", "-A"], check=True)
    print(f"Staged '{src.name}' and refreshed index.html (= latest deck).")
    print(f'  review : git -C "{REPO}" status')
    print(f'  commit : git -C "{REPO}" commit -m "{src.name} deck"')
    print(f'  push   : git -C "{REPO}" push origin main')


if __name__ == "__main__":
    main()
