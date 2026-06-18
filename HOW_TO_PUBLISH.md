# How to publish a refresh (operator note)

Plain-English guide for updating the live dashboard, so the pipeline runs hassle-free even if
no one who built it is around. **This folder is the *publish* repo** — a git repo wired to
GitHub. It is NOT where the numbers are produced; that happens in the sibling **working**
folder `…\Documents\Projects\Peer NAV Quartiles comparison\` (which runs the engine
`peer_monitor.py`). This repo just takes a finished refresh and puts it online.

---

## The live link always shows the LATEST refresh

The fixed URL **https://kartiksngh.github.io/MF-Peers-Comparison/** is served by GitHub Pages
from the repo-root **`index.html`**. Each refresh **overwrites `index.html`** with a copy of
the newest deck, so the URL never changes — it always renders whatever `index.html` currently
is = the latest. The dated folders (`June 16, 2026/`, then future months) just pile up as
**archives**; the link does not point into them.

## What's in here

```
index.html              ← the LIVE deck (a copy of the latest refresh's out/dashboard_offline.html)
publish_refresh.py      ← the tool that stages a new refresh (copies it in + resets index.html)
HOW_TO_PUBLISH.md       ← this note
README.md               ← public description of the project
June 16, 2026/          ← an archived refresh (engine, Data, out/ deliverables, notebook)
   └── out/             ← that month's dashboard + Excel + json
```

## Publish a new refresh (3 steps)

After the engine has been run in the working folder (so `<refresh>\out\dashboard_offline.html`
exists):

```bash
# 1. stage it (copies the dated folder in + resets index.html to the new deck + git add)
cd "C:\Users\Administrator\Documents\Projects\MF-Peers-Comparison"
python publish_refresh.py "C:\Users\Administrator\Documents\Projects\Peer NAV Quartiles comparison\2026\August 11, 2026"

# 2. review + commit
git status
git commit -m "August 11, 2026 deck"

# 3. push  —  A HUMAN RUNS THIS (see warning below)
git push origin main
```

GitHub Pages is already configured (Settings → Pages → branch `main`, folder `/ (root)`); it
redeploys automatically a minute or so after the push. Confirm the push landed with
`git ls-remote --heads origin`.

**Manual fallback** (if `publish_refresh.py` is unavailable): copy the new dated folder into
this repo, then `copy "<that folder>\out\dashboard_offline.html" index.html`, then commit +
push.

## Two safety rules (do not skip)

1. **A person runs `git push` — never automate it.** The deck is **public** and inlines all
   chart data, and this repo also contains the raw vendor `Data\`. Publishing is a deliberate
   human decision each time (proprietary data → public internet). This was chosen knowingly.
2. **Only ever push from THIS folder.** On this machine git is accidentally rooted at the home
   directory (`C:\Users\Administrator`), so a push from the working project folder would try to
   upload your entire profile (SSH keys, tokens, AppData). Before any git command, run
   `git rev-parse --show-toplevel` and confirm it prints this repo's path.

## Where the methodology lives

The numbers, cleaning rules, calendar-return definitions, and scoring constants are documented
in each refresh's `BUILD_SPEC.md` (e.g. `June 16, 2026/BUILD_SPEC.md`). The dashboard is
self-contained and interactive (time-slider + animated-GIF export on the snapshot views;
per-quartile toggles on the Composite bar).
