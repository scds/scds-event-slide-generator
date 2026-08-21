# AGENTS.md

Personal automation scripts that scrape McMaster SCDS library events (LibCal) into a PowerPoint slide of upcoming workshops.

## Run

- `pull.py` is the pipeline: fetches the live LibCal feed (`https://libcal.mcmaster.ca/api_events.php?m=upc&cid=8132_7565&...`, URL in the script), parses events, buckets them into the four program-area boxes, and writes a dated copy of the deck (`SCDS Slide - YYYY-MM-DD.pptx`, auto-versioned `-v2`, `-v3`... if it already exists). It never modifies `SCDS Slide.pptx` itself.
- The feed is table-based: one `<table class="s-lc-ea-tb">` per event with rows `s-lc-ea-ttit` (title/link), `s-lc-ea-from` (date like "11:00am Thursday, September 24, 2026" — parsed with `%I:%M%p %A, %B %d, %Y`), and `s-lc-ea-tcat` (comma-separated category names). Changes there break parsing.
- Category → box mapping is the editable `BOXES` priority list in `pull.py`: an event tagged with several categories lands in the first matching box; unmatched events fall through to "Do More with Digital Scholarship". Current LibCal category names seen: "DR", "Research Data Management", "DASH", "DMDS".
- `scds_pull.py` + `slide.py` are legacy: they parse the OLD `<li>`-style feed from a locally saved HTML file at hardcoded `C:/Users/pratti/Downloads/api_events.php.htm`. Kept for CSV output only; don't model new work on them.
- Season in the slide title ("Fall/Winter/Spring/Summer YEAR") is derived from the earliest event date (Jan–Feb Winter, Mar–May Spring, Jun–Aug Summer, Sep–Dec Fall).

## Deck formatting (SCDS Slide.pptx)

- 16:9, Blank layout, two near-duplicate slides — `pull.py` keeps only the second (newer) revision.
- Four quadrant text boxes are found by their header text ("Data Analysis Support Hub", "Digital Research", "Research Data Management", "Do More with Digital Scholarship"); event lines are a bold "Month D: " run plus a regular title run, Aptos 16pt. White text on DASH/Do More boxes, #0F0F0F on Digital Research/RDM. New lines are built by deep-copying an existing event paragraph as a template — don't rebuild formatting from scratch.
- `TextBox 2` was a manual overflow box; `pull.py` clears it.

## Environment

- `.venv` works again: Python 3.11 based, recreated 2026-08-21. Run scripts with `.venv\Scripts\python.exe`.
- No requirements file. Needed packages: `requests`, `beautifulsoup4`, `pandas`, `python-pptx`.
- `~$SCDS Slide.pptx` is a PowerPoint lock file (deck open in PowerPoint), not data. There is no `.gitignore`; don't commit `.venv/`, `__pycache__/`, or `~$*` files.

## Git layout

- This folder is not its own repository. Git toplevel is the parent `C:/Users/pratti/Documents/Github` (branch `main`, zero commits), so `git status` lists all sibling projects as untracked. Scope any commits carefully.
