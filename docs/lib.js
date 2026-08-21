// Pure feed/slide logic shared by the webtool; no DOM access here.
// Priority order mirrors pull.py: first box whose keyword equals one of the
// event's LibCal category names wins; unmatched events fall into the last box.
const BOXES = [
  { header: "Data Analysis Support Hub", keywords: ["DASH", "data analysis support hub", "dash"] },
  { header: "Research Data Management", keywords: ["research data management", "rdm"] },
  { header: "Digital Research", keywords: ["digital research", "dr"] },
  { header: "Do More with Digital Scholarship", keywords: ["DMDS"] },
];

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
  module.exports = { BOXES, MONTHS, parseWhen, bucketEvents, seasonTitle };
}
