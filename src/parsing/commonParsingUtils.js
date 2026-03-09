import { parseTime, formatTime } from "../helpers.js";

export function normalizeWhitespace(text) {
  return String(text || "")
    .replace(/\r/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/[^\S\n]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

export function normalizeDashes(text) {
  return String(text || "").replace(/[‐‑‒–—−]/g, "-");
}

export function normalizeTimes(text) {
  let s = String(text || "");
  s = s.replace(/\b(\d{1,2})\.(\d{2})\b/g, "$1:$2");
  s = s.replace(/\b(\d{1,2})-(\d{2})\b/g, "$1:$2");
  return s;
}

export function normalizeBalingText(text) {
  let s = String(text || "");
  s = normalizeDashes(s);
  s = s.replace(/×/g, "x");
  s = s.replace(/\b4\s*x\s*4\b/gi, "4x4");
  s = s.replace(/\bpassengers\b/gi, "Passenger");
  s = s.replace(/\blight\s*commercials\b/gi, "Light Commercial");
  s = s.replace(/\bmotor\s*cycle\b/gi, "Motorcycle");
  s = s.replace(/\bkg\b/gi, "KG");
  s = s.replace(/\s*,\s*/g, ", ");
  s = normalizeTimes(s);
  s = normalizeWhitespace(s);
  return s;
}

export function parseLabelTime(body, label) {
  const re = new RegExp(`${label}\\s*(?:time)?\\s*[:\\-=]\\s*\\b(\\d{1,2}:\\d{2}(?::\\d{2})?)\\b`, "i");
  const m = String(body || "").match(re);
  return m ? parseTime(m[1]) : null;
}

export function minutesBetween(startTime, finishTime) {
  if (!startTime || !finishTime) return null;
  const start = startTime.h * 60 + startTime.m;
  const finish = finishTime.h * 60 + finishTime.m;
  if (finish < start) return null;
  return finish - start;
}

export function formatTimeSafe(timeObj) {
  return timeObj ? formatTime(timeObj) : "";
}

export function numericOrNull(value) {
  if (value === null || value === undefined || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

export function parseWeightKg(body) {
  const text = String(body || "");
  let m = text.match(/\bweight\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*KG\b/i);
  // Fallback for compact forms like "899kg" where "Weight -" label is omitted.
  if (!m) {
    const all = [...text.matchAll(/\b(\d{2,4}(?:\.\d+)?)\s*KG\b/gi)];
    if (all.length > 0) m = all[all.length - 1];
  }
  if (!m) return null;
  const n = Number(m[1]);
  return Number.isFinite(n) ? n : null;
}

export function parseTotalQty(body) {
  const text = String(body || "");
  let m = text.match(/\btotal\s*qty\s*[:\-]?\s*(\d+)\b/i);
  if (!m) m = text.match(/\btotal\s*[:\-]\s*(\d+)\b/i);
  if (!m) m = text.match(/\btotal\s+(\d+)\b/i);
  if (!m) return null;
  const n = Number(m[1]);
  return Number.isFinite(n) ? n : null;
}
