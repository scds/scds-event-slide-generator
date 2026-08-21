"""Fetch the LibCal feed and rebuild the SCDS upcoming-workshops slide.

Reads the live LibCal feed, groups events into the four program-area boxes,
and writes a dated copy of SCDS Slide.pptx with the event lines refreshed.
Formatting (fonts, colours, boxes, QR code) comes from the original deck.
"""
import copy
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

FEED_URL = ("https://libcal.mcmaster.ca/api_events.php"
            "?m=upc&cid=8132_7565&audience=&c=&d=&tags=&l=5&tar=0")
DECK_PATH = Path("SCDS Slide.pptx")

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x0F, 0x0F, 0x0F)

# Priority order: an event tagged with several categories lands in the first
# box whose keywords match one of its LibCal category names (exact match).

BOXES = [
    {"header": "Data Analysis Support Hub", "color": WHITE,
     "keywords": ["DASH", "data analysis support hub"]},
    {"header": "Research Data Management", "color": DARK,
     "keywords": ["research data management", "rdm"]},
    {"header": "Digital Research", "color": DARK,
     "keywords": ["digital research", "dr"]},
    {"header": "Do More with Digital Scholarship", "color": WHITE,
     "keywords": ["DMDS", "Do More with Digital Scholarship"]},
]


def fetch_feed(url=FEED_URL):
    response = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    response.raise_for_status()
    response.encoding = "utf-8"
    return response.text


def parse_feed(html):
    """Parse the table-style api_events markup into sorted event dicts."""
    soup = BeautifulSoup(html, "html.parser")
    events = []
    for table in soup.find_all("table", class_="s-lc-ea-tb"):
        title_row = table.find("tr", class_="s-lc-ea-ttit")
        from_row = table.find("tr", class_="s-lc-ea-from")
        cat_row = table.find("tr", class_="s-lc-ea-tcat")
        if not (title_row and from_row):
            continue
        link = title_row.find("a")
        if not link:
            continue

        when_text = from_row.find_all("td")[1].get_text(strip=True)
        # e.g. "11:00am Thursday, September 24, 2026"
        when = datetime.strptime(when_text, "%I:%M%p %A, %B %d, %Y")

        categories = []
        if cat_row:
            cells = cat_row.find_all("td")
            if len(cells) > 1:
                categories = [c.strip() for c in cells[1].get_text().split(",")
                              if c.strip()]

        events.append({
            "title": link.get_text(strip=True),
            "link": link.get("href"),
            "when": when,
            "categories": categories,
        })
    events.sort(key=lambda e: e["when"])
    return events


def bucket_events(events):
    buckets = {box["header"]: [] for box in BOXES}
    for event in events:
        tokens = {c.casefold() for c in event["categories"]}
        for box in BOXES:
            if any(kw.casefold() in tokens for kw in box["keywords"]):
                buckets[box["header"]].append(event)
                break
        else:
            buckets[BOXES[-1]["header"]].append(event)
    return buckets


def term_title(events):
    when = events[0]["when"] if events else datetime.now()
    season = "Fall" if when.month >= 9 else "Summer" if when.month >= 6 else "Spring" if when.month >=3 else "Winter"
    return f"{season} {when.year}: Upcoming Workshops"


def find_shapes(slide):
    """Locate the title bar, the four quadrant boxes, and the overflow box."""
    title = None
    boxes = {}
    overflow = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if shape.name == "Title 4" and shape.top < 1_000_000:
            title = shape
        elif shape.name == "Subtitle 2":
            header = shape.text_frame.paragraphs[0].text.strip()
            if header in buckets_keys():
                boxes[header] = shape
        elif shape.name == "TextBox 2":
            overflow.append(shape)
    return title, boxes, overflow


def buckets_keys():
    return {box["header"] for box in BOXES}


def harvest_template(slide):
    """Deep-copy one existing event paragraph ('Month D: ' + title) as template."""
    for shape in slide.shapes:
        if not shape.has_text_frame or shape.name != "Subtitle 2":
            continue
        for para in shape.text_frame.paragraphs:
            runs = para.runs
            if len(runs) >= 2 and runs[0].text.strip().endswith(":"):
                return copy.deepcopy(para._p)
    raise RuntimeError("No event-line paragraph found to use as a template")


def fill_box(shape, events, template, color):
    tf = shape.text_frame
    for para in list(tf.paragraphs)[1:]:
        para._p.getparent().remove(para._p)
    for _ in events:
        tf._txBody.append(copy.deepcopy(template))
    for para, event in zip(list(tf.paragraphs)[1:], events):
        runs = para.runs
        runs[0].text = f"{event['when']:%B} {event['when'].day}: "
        runs[0].font.color.rgb = color
        runs[1].text = event["title"]
        runs[1].font.color.rgb = color
        runs[1].font.bold = False
        for extra in runs[2:]:
            extra._r.getparent().remove(extra._r)


def output_path():
    stamp = f"{datetime.now():%Y-%m-%d}"
    candidate = Path(f"SCDS Slide - {stamp}.pptx")
    version = 1
    while candidate.exists():
        version += 1
        candidate = Path(f"SCDS Slide - {stamp} v{version}.pptx")
    return candidate


def build_slide(events, deck_path=DECK_PATH):
    out_path = output_path()
    out_path.write_bytes(deck_path.read_bytes())
    prs = Presentation(str(out_path))

    # Keep only the second slide (the more current revision of the two).
    sld_id_lst = prs.slides._sldIdLst
    first = list(sld_id_lst)[0]
    prs.part.drop_rel(first.get(qn("r:id")))
    sld_id_lst.remove(first)

    slide = prs.slides[0]
    title_shape, boxes, overflow = find_shapes(slide)
    template = harvest_template(slide)

    title_shape.text_frame.paragraphs[0].runs[0].text = term_title(events)

    buckets = bucket_events(events)
    for box in BOXES:
        shape = boxes.get(box["header"])
        if shape is not None:
            fill_box(shape, buckets[box["header"]], template, box["color"])

    for shape in overflow:  # stale manual overflow lines
        shape.text_frame.clear()

    prs.save(str(out_path))
    return out_path


if __name__ == "__main__":
    html = fetch_feed()
    found = parse_feed(html)
    print(f"{len(found)} upcoming event(s) found")
    for event in found:
        print(f"  {event['when']:%b %d}  [{'; '.join(event['categories'])}]  {event['title']}")
    print(f"Wrote {build_slide(found)}")
