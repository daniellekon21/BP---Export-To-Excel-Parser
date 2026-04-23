import { splitWhatsAppMessages, dateToStr } from "../helpers.js";
import { normalizeBalingText, numericOrNull } from "./commonParsingUtils.js";

function isValidDateObject(d) {
  if (!d) return false;
  if (!Number.isInteger(d.year) || !Number.isInteger(d.month) || !Number.isInteger(d.day)) return false;
  if (d.month < 1 || d.month > 12 || d.day < 1 || d.day > 31) return false;
  return true;
}

function dateObjToSortable(d) {
  return d.year * 10000 + d.month * 100 + d.day;
}

function todayDateObj() {
  const now = new Date();
  return { year: now.getFullYear(), month: now.getMonth() + 1, day: now.getDate() };
}

function resolveDate(bodyDate, tsDate) {
  if (isValidDateObject(bodyDate)) {
    if (dateObjToSortable(bodyDate) > dateObjToSortable(todayDateObj())) {
      return tsDate || bodyDate;
    }
    return bodyDate;
  }
  return tsDate || null;
}

function extractBodyDate(body) {
  const m = String(body || "").match(/\b(\d{1,2})\/(\d{1,2})\/(\d{2,4})\b/);
  if (!m) return null;
  const day = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  let year = parseInt(m[3], 10);
  if (year < 100) year += 2000;
  return { year, month, day };
}

const TYRE_LABELS = [
  { key: "delHeavyCommercialSW", patterns: [/^\*?\s*heavy\s*commercial\s*sw\s*[-:]/i] },
  { key: "delHeavyCommercialT",  patterns: [/^\*?\s*heavy\s*commercial\s*t\s*[-:]/i] },
  { key: "delHeavyCommercial",   patterns: [/^\*?\s*heavy\s*commercial\s*[-:]/i] },
  { key: "delLightCommercial",   patterns: [/^\*?\s*light\s*commercial\s*[-:]/i] },
  { key: "delFourByFour",        patterns: [/^\*?\s*4x4\s*[-:]/i, /^\*?\s*4\s*x\s*4\s*[-:]/i] },
  { key: "delMotorcycle",        patterns: [/^\*?\s*motorcycle\s*[-:]/i, /^\*?\s*motor\s*cycle\s*[-:]/i] },
  { key: "delPassenger",         patterns: [/^\*?\s*passenger\s*[-:]/i] },
  { key: "delAgricultural",      patterns: [/^\*?\s*agricultural\s*[-:]/i, /^\*?\s*agri\s*[-:]/i] },
];

function matchTyreLine(line) {
  for (const entry of TYRE_LABELS) {
    for (const re of entry.patterns) {
      if (re.test(line)) {
        const numMatch = line.match(/(\d+)\s*\*?\s*$/);
        const count = numMatch ? numericOrNull(numMatch[1]) : null;
        return { key: entry.key, count };
      }
    }
  }
  return null;
}

function extractLabeledField(line, labelPatterns) {
  for (const re of labelPatterns) {
    const m = line.match(re);
    if (m) return String(m[1] || "").trim().replace(/\*+$/g, "").trim();
  }
  return null;
}

function extractTotalReported(body) {
  const m = String(body || "").match(/\btotal\s*[-:]?\s*\*?\s*(\d+)\s*\*?/i);
  if (!m) return null;
  return numericOrNull(m[1]);
}

function emptyRecord() {
  return {
    date: null,
    delTruckNo: null,
    delWbd: null,
    delGrv: null,
    delDepot: null,
    delDepotManager: null,
    delCollectionNo: null,
    delTransporter: null,
    delPassenger: null,
    delFourByFour: null,
    delMotorcycle: null,
    delLightCommercial: null,
    delHeavyCommercial: null,
    delHeavyCommercialSW: null,
    delHeavyCommercialT: null,
    delAgricultural: null,
    delTotalReported: null,
    rawMessage: "",
  };
}

function logValidation(validationLog, msg, record, issue, action = "Logged") {
  validationLog.push({
    severity: "WARNING",
    issueType: "DELIVERIES_PARSER_WARNING",
    date: record?.date ? dateToStr(record.date) : "",
    sourceMessageTimestamp: msg?.tsDate ? dateToStr(msg.tsDate) : "",
    sender: msg?.sender || "",
    issue,
    action,
    rawMessage: msg?.body || "",
  });
}

function isDeliveryMessage(body) {
  return /\btruck\s*(?:#|no\.?|number)\s*[:\-]/i.test(body);
}

export function parseDeliveriesMessages(text) {
  const messages = splitWhatsAppMessages(text);
  const records = [];
  const validationLog = [];

  for (const msg of messages) {
    if (!isDeliveryMessage(msg.body)) continue;

    const normalized = normalizeBalingText(msg.body);
    const lines = normalized.split("\n").map((l) => l.trim()).filter(Boolean);

    const record = emptyRecord();
    record.rawMessage = msg.body;

    const bodyDate = extractBodyDate(msg.body);
    record.date = resolveDate(bodyDate, msg.tsDate);

    let tyreCategoryCount = 0;

    for (const line of lines) {
      const truck = extractLabeledField(line, [/^truck\s*(?:#|no\.?|number)\s*[:\-]\s*(.+)$/i]);
      if (truck !== null) { record.delTruckNo = truck; continue; }

      const wbd = extractLabeledField(line, [/^wbd\s*[:\-]?\s*(.+)$/i]);
      if (wbd !== null) { record.delWbd = wbd; continue; }

      const grv = extractLabeledField(line, [/^grv\s*[:\-]?\s*(.+)$/i]);
      if (grv !== null) { record.delGrv = grv; continue; }

      const depotMgr = extractLabeledField(line, [/^depot\s*manager\s*[:\-]?\s*(.+)$/i]);
      if (depotMgr !== null) { record.delDepotManager = depotMgr; continue; }

      const depot = extractLabeledField(line, [/^depot\s*[:\-]?\s*(.+)$/i]);
      if (depot !== null) { record.delDepot = depot; continue; }

      const collection = extractLabeledField(line, [/^collection\s*(?:#|no\.?|number)\s*[:\-]?\s*(.+)$/i]);
      if (collection !== null) { record.delCollectionNo = collection; continue; }

      const transporter = extractLabeledField(line, [/^transporter\s*[:\-]?\s*(.+)$/i]);
      if (transporter !== null) { record.delTransporter = transporter; continue; }

      const tyre = matchTyreLine(line);
      if (tyre) {
        if (tyre.count != null) {
          record[tyre.key] = tyre.count;
          tyreCategoryCount += 1;
        }
        continue;
      }
    }

    record.delTotalReported = extractTotalReported(msg.body);

    if (!record.date) {
      logValidation(validationLog, msg, record, "Missing or unparseable date");
    }
    if (tyreCategoryCount === 0) {
      logValidation(validationLog, msg, record, "No tyre categories parsed");
    }
    if (record.delTotalReported != null) {
      const summed =
        (record.delPassenger || 0) +
        (record.delFourByFour || 0) +
        (record.delMotorcycle || 0) +
        (record.delLightCommercial || 0) +
        (record.delHeavyCommercial || 0) +
        (record.delHeavyCommercialSW || 0) +
        (record.delHeavyCommercialT || 0) +
        (record.delAgricultural || 0);
      if (summed !== record.delTotalReported) {
        logValidation(
          validationLog,
          msg,
          record,
          `Total mismatch: reported ${record.delTotalReported}, summed categories = ${summed}`,
        );
      }
    }

    records.push(record);
  }

  return { records, validationLog };
}
