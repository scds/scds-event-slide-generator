# AGENTS.md

Personal automation scripts that scrape McMaster SCDS library events (LibCal) into a PowerPoint slide of upcoming workshops.

## Run

- `pull.py` is the pipeline: fetches the live LibCal feed (`https://libcal.mcmaster.ca/api_events.php?m=upc&cid=8132_7565&...`, URL in the script), parses events, buckets them into the four program-area boxes, and writes a dated copy of the deck (`SCDS Slide - YYYY-MM-DD.pptx`, auto-versioned `-v2`, `-v3`... if it already exists). It never modifies `SCDS Slide template.pptx` itself.
- The feed is table-based: one `<table class="s-lc-ea-tb">` per event with rows `s-lc-ea-ttit` (title/link), `s-lc-ea-from` (date like "11:00am Thursday, September 24, 2026" — parsed with `%I:%M%p %A, %B %d, %Y`), and `s-lc-ea-tcat` (comma-separated category names). Changes there break parsing.
- Category → box mapping is the editable `BOXES` priority list in `pull.py`: an event tagged with several categories lands in the first matching box; unmatched events fall through to "Do More with Digital Scholarship". Current LibCal category names seen: "DR", "Research Data Management", "DASH", "DMDS".
- Season in the slide title ("Fall/Winter/Spring/Summer YEAR") is derived from the earliest event date (Jan–Feb Winter, Mar–May Spring, Jun–Aug Summer, Sep–Dec Fall).
- The deck is a single slide named `SCDS Slide template.pptx`. Older two-slide decks are handled (extra first slide is dropped); don't assume TextBox 2 or the overflow layout still exists — the script clears TextBox 2 if present.

## Feed URL parameters (user-verified)

- `m=` mode: upc / today / week / fortnight / month. `cid=` calendar IDs, underscore-separated. `d=` **category IDs** filter (33846 RDM, 36203 DR). `audience=` numeric audience codes. `c=` campus (unused). `l=` limit, max 50. `tar=0` leave alone. Full table in README.md.

## Web tool (docs/, GitHub Pages)

- Static site: left panel builds/fetches the LibCal feed client-side (feed sends `Access-Control-Allow-Origin: *`, so direct browser fetch works), main column previews the slide replica, Download button rebuilds the PPTX via PptxGenJS CDN.
- Pure logic lives in `docs/lib.js` (no DOM) — testable with plain `node`: parseWhen, bucketEvents, seasonTitle must stay in sync with `pull.py`.
- Slide geometry is duplicated in three places that must stay consistent: EMU in `pull.py`, percentages in `docs/style.css`, inches in `docs/app.js` (`PPTX` const). All derive from `SCDS Slide template.pptx` (12192000 x 6858000 EMU).
- Serve locally over http (`cd docs && python -m http.server`) — opening index.html via file:// breaks fetch/CORS.

## Environment

- `.venv` works: Python 3.11 based, recreated 2026-08-21. Run scripts with `.venv\Scripts\python.exe`.
- No requirements file. Python packages: `requests`, `beautifulsoup4`, `pandas`, `python-pptx`. Node 22 is available for testing docs JS.
- Don't commit `.venv/`, `__pycache__/`, or `~$*` PowerPoint lock files. There is no `.gitignore`.

## Git layout

- This folder IS its own repository now: remote `https://github.com/scds/scds-event-slide-generator.git`, branch `master`. (Previously it shared a parent repo — that's gone.)
