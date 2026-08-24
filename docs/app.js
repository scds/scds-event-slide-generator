const FEED_BASE = "https://libcal.mcmaster.ca/api_events.php";

let currentEvents = [];

const $ = (id) => document.getElementById(id);

function buildFeedUrl() {
  const params = new URLSearchParams();
  params.set("m", $("opt-mode").value);
  params.set("cid", $("opt-cid").value.trim());
  params.set("audience", $("opt-audience").value.trim());
  params.set("c", "");
  params.set("d", $("opt-d").value.trim());
  params.set("tags", "");
  params.set("l", String(parseInt($("opt-l").value, 10) || 5));
  params.set("tar", "0");
  return `${FEED_BASE}?${params.toString()}`;
}

function parseFeed(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const events = [];
  for (const table of doc.querySelectorAll("table.s-lc-ea-tb")) {
    const link = table.querySelector("tr.s-lc-ea-ttit a");
    const fromCell = table.querySelector("tr.s-lc-ea-from td:last-child");
    const catCell = table.querySelector("tr.s-lc-ea-tcat td:last-child");
    if (!link || !fromCell) continue;
    const when = parseWhen(fromCell.textContent);
    if (!when) continue;
    const categories = catCell
      ? catCell.textContent.split(",").map((c) => c.trim()).filter(Boolean)
      : [];
    events.push({ title: link.textContent.trim(), when, categories });
  }
  events.sort((a, b) => a.when - b.when);
  return events;
}

function tagSections(events) {
  for (const [header, evs] of bucketEvents(events)) {
    for (const event of evs) event.section = header;
  }
}

function assignInitialSelection(events) {
  const { capped } = capBuckets(bucketEvents(events));
  for (const event of events) event.included = false;
  for (const evs of capped.values()) {
    for (const event of evs) event.included = true;
  }
}

function chosenEvents() {
  return currentEvents.filter((event) => event.included);
}

function renderLines(el, events) {
  el.replaceChildren(...events.map((event) => {
    const line = document.createElement("div");
    line.className = "line";
    const bold = document.createElement("b");
    bold.textContent = dateLabel(event.when);
    line.appendChild(bold);
    line.appendChild(document.createTextNode(event.title));
    return line;
  }));
}

function renderSlide() {
  const chosen = chosenEvents();
  $("slide-title").textContent = seasonTitle(chosen);
  const buckets = bucketEvents(chosen);
  const { main: dmdsMain, extra: dmdsExtra } =
    splitDmds(buckets.get("Do More with Digital Scholarship"));
  for (const box of BOXES) {
    const events = box.header === "Do More with Digital Scholarship"
      ? dmdsMain
      : buckets.get(box.header);
    renderLines(document.querySelector(`.box[data-box="${box.header}"] .lines`), events);
  }
  renderLines(document.querySelector(".box.dmds-extra .lines"), dmdsExtra);
  return chosen.length;
}

function eventRow(event) {
  const label = document.createElement("label");
  label.className = "pick";
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.checked = event.included;
  checkbox.addEventListener("change", () => {
    event.included = checkbox.checked;
    renderEventLists();
    renderSlide();
  });
  label.appendChild(checkbox);
  const bold = document.createElement("b");
  bold.textContent = dateLabel(event.when);
  label.appendChild(bold);
  label.appendChild(document.createTextNode(event.title));
  return label;
}

function fillGroupedList(containerId, events) {
  const wrap = $(containerId);
  const rows = wrap.querySelector(".rows");
  const children = [];
  for (const box of BOXES) {
    const group = events.filter((e) => e.section === box.header);
    if (!group.length) continue;
    const header = document.createElement("div");
    header.className = "group-header";
    header.textContent = box.header;
    children.push(header, ...group.map(eventRow));
  }
  rows.replaceChildren(...children);
  wrap.classList.toggle("hidden", events.length === 0);
}

function renderEventLists() {
  const included = [];
  const excluded = [];
  for (const event of currentEvents) {
    (event.included ? included : excluded).push(event);
  }
  fillGroupedList("picker-included", included);
  fillGroupedList("picker-dropped", excluded);
}

async function loadEvents(url) {
  const status = $("status");
  status.className = "";
  status.textContent = "Loading…";
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    currentEvents = parseFeed(await response.text());
    tagSections(currentEvents);
    assignInitialSelection(currentEvents);
    renderEventLists();
    const onSlide = renderSlide();
    $("download-btn").disabled = false;
    status.className = "ok";
    status.textContent = `${currentEvents.length} upcoming event(s) loaded, ` +
      `${onSlide} on the slide.`;
  } catch (err) {
    currentEvents = [];
    renderEventLists();
    renderSlide();
    status.className = "error";
    status.textContent = `Could not load feed: ${err.message}. ` +
      "If you opened this page from a local file, serve it over http instead.";
  }
}

let qrDataUrlCache = null;

function qrDataUrl() {
  if (!qrDataUrlCache) {
    qrDataUrlCache = fetch("assets/qr.png").then((res) => res.blob()).then((blob) =>
      new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      }));
    qrDataUrlCache.catch(() => { qrDataUrlCache = null; });
  }
  return qrDataUrlCache;
}

// Geometry copied from SCDS Slide template.pptx (EMU / 914400 = inches).
const PPTX = {
  slide: { width: 13.333, height: 7.5 },
  rects: [
    { x: 0.7087, y: 1.1024, w: 5.9049, h: 2.3611, color: "156082" },
    { x: 6.6146, y: 1.1024, w: 5.9049, h: 2.3611, color: "0F9ED5" },
    { x: 6.6146, y: 3.4646, w: 5.9049, h: 2.9528, color: "196B24" },
    { x: 0.7087, y: 3.4646, w: 5.9049, h: 2.9528, color: "E3E86A" },
  ],
  boxes: {
    "Data Analysis Support Hub": { x: 0.8268, y: 1.2205, w: 5.6693, h: 2.1250, color: "FFFFFF" },
    "Digital Research": { x: 6.7323, y: 1.2205, w: 5.6693, h: 2.1250, color: "0F0F0F" },
    "Research Data Management": { x: 0.8268, y: 3.5827, w: 5.6693, h: 2.7165, color: "0F0F0F" },
    "Do More with Digital Scholarship": { x: 6.7323, y: 3.5827, w: 5.6693, h: 0.9449, color: "FFFFFF" },
  },
  title: { x: 1.2917, y: 0.5408, w: 11.2444, h: 0.4915 },
  footer: { x: 0.75, y: 6.6097, w: 7.8239, h: 0.6176 },
  qr: { x: 10.6563, y: 4.8228, w: 2.6786, h: 2.6786 },
  dmdsExtra: { x: 6.7323, y: 4.5281, w: 3.7800, h: 1.7716 },
};

function addEventText(slide, events, geo, color, header) {
  const runs = [];
  if (header) {
    runs.push({
      text: header,
      options: { bold: true, fontSize: 20, color, breakLine: true },
    });
  }
  for (const event of events) {
    runs.push({ text: dateLabel(event.when), options: { bold: true, fontSize: 16, color } });
    runs.push({
      text: event.title,
      options: { fontSize: 16, color, breakLine: true, paraSpaceAfter: 5 },
    });
  }
  if (runs.length) {
    slide.addText(runs, { ...geo, fontFace: "Aptos", align: "left", valign: "top" });
  }
}

async function downloadPptx() {
  if (typeof PptxGenJS === "undefined") {
    alert("PptxGenJS failed to load; check your internet connection.");
    return;
  }
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "SCDS", width: PPTX.slide.width, height: PPTX.slide.height });
  pptx.layout = "SCDS";
  const slide = pptx.addSlide();

  for (const rect of PPTX.rects) {
    slide.addShape("rect", { ...rect, fill: { color: rect.color } });
  }

  slide.addText([{
    text: seasonTitle(chosenEvents()),
    options: { bold: true, fontSize: 24, color: "0F0F0F" },
  }], { ...PPTX.title, fontFace: "Aptos", align: "left", valign: "middle" });

  const buckets = bucketEvents(chosenEvents());
  for (const box of BOXES) {
    const geo = PPTX.boxes[box.header];
    let events = buckets.get(box.header);
    if (box.header === "Do More with Digital Scholarship") {
      const { main, extra } = splitDmds(events);
      addEventText(slide, extra, PPTX.dmdsExtra, geo.color);
      events = main; // main keeps its own geometry from PPTX.boxes
    }
    addEventText(slide, events, geo, geo.color, box.header);
  }

  try {
    slide.addImage({ data: await qrDataUrl(), ...PPTX.qr });
  } catch (err) {
    console.warn("QR code not embedded:", err);
  }

  slide.addText([{
    text: "Register for Upcoming Workshops: https://u.mcmaster.ca/scds-workshops",
    options: { fontSize: 17, color: "0F0F0F" },
  }], { ...PPTX.footer, fontFace: "Aptos", align: "left", valign: "middle" });

  await pptx.writeFile({ fileName: `SCDS Slide - ${new Date().toISOString().slice(0, 10)}.pptx` });
}

$("load-url-btn").addEventListener("click", () => loadEvents($("feed-url").value.trim()));
$("load-build-btn").addEventListener("click", () => loadEvents(buildFeedUrl()));
$("download-btn").addEventListener("click", downloadPptx);

loadEvents($("feed-url").value.trim());
