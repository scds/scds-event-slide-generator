# SCDS Event Slide

Generates the Sherman Centre for Digital Scholarship (SCDS) "Upcoming Workshops" PowerPoint slide from the McMaster Libraries LibCal event feed.

`pull.py` fetches upcoming events from [LibCal](https://libcal.mcmaster.ca/api_events.php?m=upc&cid=8132_7565&audience=&c=&d=&tags=&l=5&tar=0), groups them into the four program-area boxes, and writes a dated copy of `SCDS Slide.pptx` with refreshed event lines — fonts, colours, quadrant layout, QR code, and footer all come from the original deck.

## Setup

Requires Python 3.11+. No requirements file; install the four dependencies directly:

```
python -m venv .venv
.venv\Scripts\python -m pip install requests beautifulsoup4 pandas python-pptx
```

## Run

```
.venv\Scripts\python pull.py
```

Output: `SCDS Slide - YYYY-MM-DD.pptx` in the repo root. If that file already exists, a version suffix is appended (`-v2`, `-v3`, ...). The original `SCDS Slide.pptx` is never modified.

The console prints each event found with its date and categories.

## Configuration

All knobs are constants at the top of `pull.py`:

- `FEED_URL` — the LibCal feed. Returns table-style HTML: one `<table class="s-lc-ea-tb">` per event with title/link (`s-lc-ea-ttit`), start time (`s-lc-ea-from`), and comma-separated category names (`s-lc-ea-tcat`). Changes to that markup break parsing.
- `BOXES` — priority list mapping LibCal category names to slide quadrants ("Data Analysis Support Hub", "Research Data Management", "Digital Research", "Do More with Digital Scholarship"). An event tagged with several categories lands in the first matching box; events with no recognised category fall through to the last box. Add new category names as they appear in LibCal.
- Season in the slide title (Fall/Winter/Spring/Summer + year) is derived automatically from the earliest event date.

## Repository notes

- `scds_pull.py` and `slide.py` are legacy scripts that parse an older `<li>`-style feed from a locally saved HTML file; kept only for CSV output (`scds_events.csv`).
- `~$SCDS Slide.pptx` files are PowerPoint lock files (deck open in PowerPoint), not data.
