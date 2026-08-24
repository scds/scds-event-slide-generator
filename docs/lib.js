// Pure feed/slide logic shared by the webtool; no DOM access here.
// Priority order mirrors pull.py: first box whose keyword equals one of the
// event's LibCal category names wins; unmatched events fall into the last box.
const BOXES = [
  { header: "Data Analysis Support Hub", keywords: ["DASH", "data analysis support hub", "dash"] },
  { header: "Research Data Management", keywords: ["research data management", "rdm"] },
  { header: "Digital Research", keywords: ["digital research", "dr"] },
  { header: "Do More with Digital Scholarship", keywords: ["DMDS"] },
];

// Max *rendered text lines* per section (titles wrap); keep in sync with
// pull.py. DMDS shares its quadrant with an overflow textbox, hence 7.
const MAX_LINES_PER_BOX = {
  "Data Analysis Support Hub": 6,
  "Research Data Management": 8,
  "Digital Research": 6,
  "Do More with Digital Scholarship": 7,
};

// Approximate Aptos 16pt character capacity of the 5.67in-wide section
// boxes, calibrated so medium titles don't overcount wrapped lines.
const CHARS_PER_LINE = 45;

function estimateLines(text) {
  return Math.max(1, Math.ceil(text.length / CHARS_PER_LINE));
}

function dateLabel(when) {
  return `${when.toLocaleString("en-US", { month: "long" })} ${when.getDate()} - `;
}

// The Do More quadrant stacks a short main textbox over a narrow overflow
// textbox that stops short of the QR code: 2 lines up top, 5 below (7 total).
const DMDS_MAIN_LINES = 2;
const DMDS_EXTRA_LINES = 5;

function splitDmds(events) {
  const main = [];
  const extra = [];
  let mainUsed = 0;
  let extraUsed = 0;
  for (const event of events) {
    const lines = estimateLines(`${dateLabel(event.when)}${event.title}`);
    if (mainUsed + lines <= DMDS_MAIN_LINES) {
      main.push(event);
      mainUsed += lines;
    } else if (extraUsed + lines <= DMDS_EXTRA_LINES) {
      extra.push(event);
      extraUsed += lines;
    }
  }
  return { main, extra };
}

const MONTHS = ["january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december"];

function parseWhen(text) {
  const m = text.trim().match(/^(\d{1,2}):(\d{2})(am|pm)\s+\w+,\s+([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})$/i);
  if (!m) return null;
  const month = MONTHS.indexOf(m[4].toLowerCase());
  if (month < 0) return null;
  let hours = parseInt(m[1], 10) % 12;
  if (m[3].toLowerCase() === "pm") hours += 12;
  return new Date(+m[6], month, +m[5], hours, +m[2]);
}

function bucketEvents(events) {
  const buckets = new Map(BOXES.map((b) => [b.header, []]));
  for (const event of events) {
    const tokens = event.categories.map((c) => c.toLowerCase());
    const box = BOXES.find((b) =>
      b.keywords.some((k) => tokens.includes(k.toLowerCase())));
    buckets.get((box || BOXES[BOXES.length - 1]).header).push(event);
  }
  return buckets;
}

function capBuckets(buckets) {
  const capped = new Map();
  const dropped = [];
  for (const [header, events] of buckets) {
    const budget = MAX_LINES_PER_BOX[header] ?? Infinity;
    const kept = [];
    let used = 0;
    for (const event of events) {
      const lines = estimateLines(`${dateLabel(event.when)}${event.title}`);
      if (used + lines <= budget) {
        kept.push(event);
        used += lines;
      } else {
        dropped.push({ header, event });
      }
    }
    capped.set(header, kept);
  }
  return { capped, dropped };
}

function seasonTitle(events) {
  const when = events.length ? events[0].when : new Date();
  const month = when.getMonth() + 1;
  const season =
    month >= 9 ? "Fall" :
    month >= 6 ? "Summer" :
    month >= 3 ? "Spring" : "Winter";
  return `${season} ${when.getFullYear()}: Upcoming Workshops`;
}

if (typeof module !== "undefined") {
  module.exports = {
    BOXES, MAX_LINES_PER_BOX, CHARS_PER_LINE, DMDS_MAIN_LINES, DMDS_EXTRA_LINES,
    MONTHS, estimateLines, dateLabel, parseWhen, bucketEvents, capBuckets,
    splitDmds, seasonTitle,
  };
}
