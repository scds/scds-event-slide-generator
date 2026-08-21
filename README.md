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

## Feed URL parameters

```
api_events.php?m=upc&cid=8132_7565&audience=&c=&d=&tags=&l=5&tar=0
```

| Param | Meaning |
|---|---|
| `m=` | Mode: =upc upcoming events | =today today's events | =month the next month | =fortnight 2 weeks | =week the next week |
| `cid=8132_7565` | Calendar ID(s); underscore-separate to merge calendars (this feed combines RPX + SCDS) |
| `audience=` | Audience filter by numeric audience ID; empty = all audiences | 2-3 digit numeric codes are used for different audience categories | 
| `c=` | Flter by Campus - not enabled here |
| `d=` | Category filter by numeric category ID, e.g. `33846` = Research Data Management, `36203` = DR; empty = all categories |
| `tags=` | Filter by event tag names; not used |
| `l=` | numberic limit on number of events returned, max 50 |
| `tar=0` | Link target flag for event links; leave as `0` |

Finding IDs without LibCal admin access:

- **Calendar IDs** appear in calendar page URLs/filters — e.g. `/calendar/scds` uses `cid=7565`.
- **Category IDs** appear in every feed response: each event table carries classes like `cat36203 cat33846`, and event pages link categories as `ct[]=36203`.

Only `m`, `cid`, and `l` are verified against live responses; the rest follow LibCal widget conventions. When wiring up pickers in a webtool, spot-check each filter once in a browser.

## Web tool

`docs/` contains a static web version of this pipeline (no server needed):

- Enter the full feed URL, or build one from mode / calendar IDs / category IDs / audience / limit.
- Events are fetched straight from LibCal by the browser (the feed allows cross-origin requests), parsed, bucketed into the four quadrants, and previewed on a replica of the slide.
- **Download .pptx** rebuilds the slide client-side via [PptxGenJS](https://github.com/gitbrent/PptxGenJS).

Run locally with any static server, e.g.:

```
cd docs && python -m http.server 8000
```

To publish, enable GitHub Pages for this repo (Settings → Pages) serving from the `main` branch `/docs` folder.

## Repository notes

- `$SCDS Slide.pptx` files are PowerPoint files, not data.
