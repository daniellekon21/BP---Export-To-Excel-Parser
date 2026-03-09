// ─── Shared Utilities ──────────────────────────────────────────────────────────

export function parseTime(str) {
  if (!str) return null;
  const cleaned = str.replace(/\./g, ":").replace(/[^0-9:]/g, "").trim();
  const parts = cleaned.split(":");
  if (parts.length === 2 || parts.length === 3) {
    const h = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const s = parts.length === 3 ? parseInt(parts[2], 10) : 0;
    if (!isNaN(h) && !isNaN(m) && !isNaN(s)) return { h, m, s };
  }
  return null;
}

export function timeToDecimal(t) {
  if (!t) return 0;
  return t.h + t.m / 60;
}

export function formatTime(t) {
  if (!t) return "";
  return `${String(t.h).padStart(2, "0")}:${String(t.m).padStart(2, "0")}`;
}

export function dateToStr(d) {
  if (!d) return "";
  return `${String(d.day).padStart(2, "0")}/${String(d.month).padStart(2, "0")}/${d.year}`;
}

export function dateSortKey(d) {
  if (!d) return 0;
  return d.year * 10000 + d.month * 100 + d.day;
}

// ─── WhatsApp Message Splitter ─────────────────────────────────────────────────
// Saves the WhatsApp system timestamp date alongside each message.
// Baling/Cutting parsers may prefer body date and fall back to timestamp.

export function splitWhatsAppMessages(text) {
  const lines = text.split("\n");
  const messages = [];
  const tsRegex = /^(\d{4}\/\d{1,2}\/\d{1,2}),\s*\d{1,2}:\d{2}\s*-\s*(.+?):\s*([\s\S]*)$/;

  for (let i = 0; i < lines.length; i++) {
    const match = lines[i].match(tsRegex);
    if (match) {
      const sender = match[2].trim();
      let body = match[3];
      while (i + 1 < lines.length && !lines[i + 1].match(/^\d{4}\/\d{1,2}\/\d{1,2},/)) {
        i++;
        body += "\n" + lines[i];
      }
      const tsParts = match[1].split("/");
      const tsDate = {
        year: parseInt(tsParts[0], 10),
        month: parseInt(tsParts[1], 10),
        day: parseInt(tsParts[2], 10)
      };
      messages.push({ sender, body: body.trim(), tsDate });
    }
  }
  return messages;
}

// ─── Month Grouping ───────────────────────────────────────────────────────────

export const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

export function monthKey(date) {
  if (!date) return "Unknown";
  return `${date.year}-${String(date.month).padStart(2, "0")}`;
}

export function monthLabel(key) {
  if (key === "Unknown") return "Unknown";
  const [year, month] = key.split("-");
  return `${MONTH_NAMES[parseInt(month, 10) - 1]} ${year}`;
}

export function groupByMonth(records) {
  const map = {};
  for (const r of records) {
    const key = monthKey(r.date);
    if (!map[key]) map[key] = [];
    map[key].push(r);
  }
  return map;
}
