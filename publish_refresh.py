"""
publish_refresh.py  —  stage a Peer Quartile Monitor refresh for GitHub Pages.

Run it from inside this repo (MF-Peers-Comparison) and point it at the WORKING project's
dated refresh folder (the one you just ran the engine in). It:
  1. copies that dated folder into this repo (kept as an archive),
  2. drops scratch/build artifacts from the copy,
  3. refreshes the live page files through publish_daily.write_live_page() — the SAME writer
     the daily run uses, so the Monday page can never differ from a weekday page,
  4. keeps every archived file under GitHub's per-file limit (see "oversized" below),
  5. runs `git add -A` so you can review, commit, and push.

The live link  https://kartiksngh.github.io/MF-Peers-Comparison/  serves root index.html,
so it ALWAYS shows the LATEST refresh — no URL change per month. Dated folders accumulate
as history; index.html is overwritten with the newest deck each time.

  python publish_refresh.py "C:/Users/Administrator/Documents/Projects/Peer NAV Quartiles comparison/2026/August 11, 2026"

Then publish (public MF data — just a deliberate publish step; Claude or KV can run the push):
  git -C "<this repo>" commit -m "August 11, 2026 deck"
  git -C "<this repo>" push origin main
"""
import gzip, hashlib, shutil, sys, subprocess
from pathlib import Path

REPO = Path(__file__).resolve().parent
LIMIT = 95 * 1024 * 1024                 # GitHub refuses files over 100 MB; stay clear of it
CRUFT_FILES = {"_template.html", "assemble.js", "verify_embedded.js"}
CRUFT_GLOBS = ["*- Copy.html", "* - Copy.html", "sample_*animation.gif", "dashboard_animated.html",
               "BirlaMFPR*.xls*",   # internal ABSL/VR competition workbook — never publish
               "MF Data*.xls*",     # vendor portfolio-holdings dump (NAVIndia) — never publish
               "_engine_run.log"]
CRUFT_DIRS  = {"_shots", "assets", "_verify", "__pycache__", ".ipynb_checkpoints"}


def _sha256(stream):
    h = hashlib.sha256()
    for block in iter(lambda: stream.read(1 << 20), b""):
        h.update(block)
    return h.hexdigest()


def _gzip_verified(p):
    """Replace p with p.gz, but only after proving the .gz unpacks to the same bytes."""
    gz = p.with_name(p.name + ".gz")
    with open(p, "rb") as fi, gzip.GzipFile(gz, "wb", compresslevel=9, mtime=0) as fo:
        shutil.copyfileobj(fi, fo, 1 << 20)
    with open(p, "rb") as a, gzip.open(gz, "rb") as b:
        if _sha256(a) != _sha256(b):
            gz.unlink()
            sys.exit(f"ABORT — {gz.name} does not unpack to the same bytes as {p.name}; nothing archived")
    if gz.stat().st_size > LIMIT:
        sys.exit(f"ABORT — {gz.name} is still {gz.stat().st_size / 1e6:.0f} MB after compression")
    size_in, size_out = p.stat().st_size, gz.stat().st_size
    p.unlink()
    return f"{p.name} ({size_in / 1e6:.0f} MB -> {gz.name} {size_out / 1e6:.0f} MB)"


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
    # HOSTED SPLIT (2026-07-16, page-speed): index.html = the template SHELL and the data ships as
    # separate JSON files the page fetches. Written by publish_daily.write_live_page() — one writer
    # for both runs (2026-09-15; the copy that used to live here had drifted, see that docstring).
    tpl = src / "dashboard.html"                        # template copy WITH __PEER_DATA__
    if tpl.exists() and "__PEER_DATA__" in tpl.read_text(encoding="utf-8", errors="ignore")[:5_000_000]:
        sys.path.insert(0, str(REPO))
        from publish_daily import write_live_page
        write_live_page(src)
    else:                                               # old refreshes: self-contained fallback
        shutil.copy2(dest / "out" / "dashboard_offline.html", REPO / "index.html")

    # ── oversized archive files ────────────────────────────────────────────────────────
    # The self-contained decks (dashboard.html / dashboard_offline.html) outgrew the limit on
    # 2026-08-11 and are fully regenerable, so they are dropped from the ARCHIVE COPY (the live
    # page never reads them). On 2026-09-09 the engine output out/dashboard_data.json itself grew
    # from 85 to 141 MB, and the 2026-09-14 archive publish stopped on it: the live deck stayed at
    # 2026-09-09 data. That file is the archive's actual record of the day, so it is KEPT, gzipped
    # (141 MB -> ~34 MB), and only after the .gz is proven to unpack to identical bytes.
    # Any OTHER oversized file still aborts — that is what the guard is for.
    pruned = []
    for name in ("dashboard.html", "dashboard_offline.html"):
        p = dest / "out" / name
        if p.is_file() and p.stat().st_size > LIMIT:
            pruned.append(f"{name} ({p.stat().st_size / 1e6:.0f} MB)")
            p.unlink()
    squeezed = [_gzip_verified(p) for p in sorted((dest / "out").glob("*.json")) if p.stat().st_size > LIMIT]
    if pruned or squeezed:
        note = ["Some files in this archived folder are not stored as the engine wrote them, because",
                "GitHub refuses any file over 100 MB.", ""]
        if pruned:
            note += ["NOT ARCHIVED (the self-contained offline decks):", *("  " + s for s in pruned),
                     "Rebuild them byte-for-byte from inside this folder with:",
                     "    python peer_monitor.py --data Data --out out", ""]
        if squeezed:
            note += ["STORED GZIPPED (unpack with: python -c \"import gzip,shutil,sys; "
                     "shutil.copyfileobj(gzip.open(sys.argv[1]), open(sys.argv[1][:-3], 'wb'))\" <file.gz>):",
                     *("  " + s for s in squeezed), ""]
        note += ["The live site is unaffected: it is served from index.html + the packed JSON files in",
                 "the repo root, which are built from the working refresh folder."]
        (dest / "out" / "DECK_NOT_ARCHIVED.txt").write_text("\n".join(note) + "\n", encoding="utf-8")
        if pruned:
            print("archive: skipped oversized self-contained deck(s): " + ", ".join(pruned))
        if squeezed:
            print("archive: stored gzipped: " + ", ".join(squeezed))

    big = [p for p in REPO.rglob("*")
           if p.is_file() and ".git" not in p.parts and p.stat().st_size > LIMIT]
    if big:
        sys.exit("ABORT — file(s) over 95 MB (GitHub limit):\n  " + "\n  ".join(map(str, big)))

    subprocess.run(["git", "-C", str(REPO), "add", "-A"], check=True)
    print(f"Staged '{src.name}' and refreshed index.html (= latest deck).")
    print(f'  review : git -C "{REPO}" status')
    print(f'  commit : git -C "{REPO}" commit -m "{src.name} deck"')
    print(f'  push   : git -C "{REPO}" push origin main')


if __name__ == "__main__":
    main()
