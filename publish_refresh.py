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
CRUFT_GLOBS = ["*- Copy.html", "* - Copy.html", "sample_*animation.gif", "dashboard_animated.html"]
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

    # the fixed landing page = the newest deck  -> link always shows the latest refresh
    shutil.copy2(dest / "out" / "dashboard_offline.html", REPO / "index.html")

    # GitHub hard limit is 100 MB/file; the offline deck is ~10 MB
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
