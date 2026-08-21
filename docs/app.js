const FEED_BASE = "https://libcal.mcmaster.ca/api_events.php";

let currentEvents = [];

const $ = (id) => document.getElementById(id);

function currentMode() {
  return document.querySelector('input[name="mode"]:checked').value;
}

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

function renderPreview(events) {
  $("slide-title").textContent = seasonTitle(events);
  const buckets = bucketEvents(events);
  for (const box of BOXES) {
    const el = document.querySelector(`.box[data-box="${box.header}"] .lines`);
    el.replaceChildren();
    for (const event of buckets.get(box.header)) {
      const line = document.createElement("div");
      line.className = "line";
      const bold = document.createElement("b");
      bold.textContent = `${event.when.toLocaleString("en-US", { month: "long" })} ${event.when.getDate()}: `;
      line.appendChild(bold);
      line.appendChild(document.createTextNode(event.title));
      el.appendChild(line);
    }
  }
}

async function loadEvents() {
  const status = $("status");
  status.className = "";
  status.textContent = "Loading…";
  const url = currentMode() === "url" ? $("feed-url").value.trim() : buildFeedUrl();
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    currentEvents = parseFeed(await response.text());
    renderPreview(currentEvents);
    $("download-btn").disabled = false;
    status.className = "ok";
    status.textContent = `${currentEvents.length} upcoming event(s) loaded.`;
  } catch (err) {
    status.className = "error";
    status.textContent = `Could not load feed: ${err.message}. ` +
      "If you opened this page from a local file, serve it over http instead.";
  }
}

async function qrDataUrl() {
  const blob = await (await fetch("assets/qr.png")).blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// Geometry copied from SCDS Slide template.pptx (EMU / 914400 = inches).
const PPTX = {
  slide: { width: 13.333, height: 7.5 },
  rects: [
    { x: 0.75, y: 1.1091, w: 11.8318, h: 5.3439, color: "E3E86A" },
    { x: 0.75, y: 1.1073, w: 5.9158, h: 2.3852, color: "156082" },
    { x: 6.6667, y: 1.1089, w: 5.9158, h: 2.3831, color: "0F9ED5" },
    { x: 6.6667, y: 3.4920, w: 5.9158, h: 2.9629, color: "196B24" },
  ],
  boxes: {
    "Data Analysis Support Hub": { x: 0.9329, y: 1.2882, w: 5.6644, h: 1.7733, color: "FFFFFF" },
    "Digital Research": { x: 6.8503, y: 1.2847, w: 5.3807, h: 1.7733, color: "0F0F0F" },
    "Research Data Management": { x: 0.9721, y: 3.6547, w: 5.3576, h: 2.5722, color: "0F0F0F" },
    "Do More with Digital Scholarship": { x: 6.9141, y: 3.6777, w: 5.3736, h: 2.5057, color: "FFFFFF" },
  },
  title: { x: 1.2917, y: 0.5408, w: 11.2444, h: 0.4915 },
  footer: { x: 0.75, y: 6.6097, w: 7.8239, h: 0.6176 },
  qr: { x: 10.6303, y: 4.7833, w: 2.7017, h: 2.7017 },
};

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
    text: seasonTitle(currentEvents),
    options: { bold: true, fontSize: 24, color: "0F0F0F" },
  }], { ...PPTX.title, fontFace: "Aptos", align: "left", valign: "middle" });

  const buckets = bucketEvents(currentEvents);
  for (const box of BOXES) {
    const geo = PPTX.boxes[box.header];
    const runs = [{
      text: box.header,
      options: { bold: true, fontSize: 20, color: geo.color, breakLine: true },
    }];
    for (const event of buckets.get(box.header)) {
      runs.push({
        text: `${event.when.toLocaleString("en-US", { month: "long" })} ${event.when.getDate()}: `,
        options: { bold: true, fontSize: 16, color: geo.color },
      });
      runs.push({
        text: event.title,
        options: { fontSize: 16, color: geo.color, breakLine: true },
      });
    }
    slide.addText(runs, { ...geo, fontFace: "Aptos", align: "left", valign: "top" });
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

document.querySelectorAll('input[name="mode"]').forEach((radio) => {
  radio.addEventListener("change", () => {
    $("fieldset-url").classList.toggle("hidden", currentMode() !== "url");
    $("fieldset-build").classList.toggle("hidden", currentMode() !== "build");
  });
});
$("load-btn").addEventListener("click", loadEvents);
$("download-btn").addEventListener("click", downloadPptx);

loadEvents();
