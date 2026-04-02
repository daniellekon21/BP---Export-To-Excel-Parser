/**
 * Quick manual test for old-format cutting parser.
 * Run: node scripts/test_cutting_fmt.mjs
 */

import { parseCuttingMessages } from "../src/cuttingParser.js";

const SEPARATOR = "=".repeat(50);
const NO_RECORDS = "  (no records)";

// WhatsApp messages need the format: YYYY/MM/DD, HH:MM - Sender: body
function wrap(body) {
  return `2025/10/10, 13:00 - Sender: ${body}`;
}

function printHeader(label) {
  console.log(`\n${SEPARATOR}`);
  console.log(label);
  console.log(SEPARATOR);
}

function formatRow(r) {
  return `  ${String(r.cmNumber).padEnd(8)} | HC=${String(r.radialsHC ?? "-").padStart(3)} | LC=${String(r.radialsLC ?? "-").padStart(3)} | AGRI=${String(r.radialsAgri ?? "-").padStart(3)} | TREADS=${String(r.radialsAgriTreads ?? "-").padStart(3)}`;
}

const TEXT1 = wrap(
`10/10/2025
13:00-14:00
CM1-(HC) =17
(LC) =12
CM2-(HC)=06
(LC)=14
CM3-(LC)=24`
);

const TEXT2 = wrap(
`10/10/2025
13:00-14:00
CM1-(HC) =17
CM2-(HC)=06
CM3-(LC)=24`
);

function printRecords(label, text) {
  printHeader(label);
  const { records } = parseCuttingMessages(text);
  if (records.length === 0) {
    console.log(NO_RECORDS);
    return;
  }
  for (const r of records) console.log(formatRow(r));
}

const TEXT3 = wrap(
`10/10/2025
14:00-15:00
CM1-(HC) =12
(LC) =16
CM2-(HC)=04
(LC)=14
CM3-(HC)=10
(LC)=18`
);

const TEXT4 = wrap(
`10/10/2025
16:00-17:00
CM1-(HC) =02
(LC) =37
CM2-(HC)=07
(LC)=17
CM3-(HC)=}
(LC)=}`
);

const TEXT5 = wrap(
`14/10/2025
08:00-09:00
CM1-(HC)=16
CM2-(HC)=16
CM3-(HC)=18
(LC)=10`
);

const TEXT6 = wrap(
`22/10/2025
11:00-12:00
CM1-HC=29
CM2-HC=20
CM3-LC=20
-HC=03`
);

const TEXT7 = wrap(
`24/10/2025
08:00-09:00
CM1- AGRI=16
CM2-AGRI=15 (Threads)`
);

const TEXT8 = wrap(
`31/10/2025
09:00-10:00
CM3-HC=19
LC=02`
);

const TEXT9 = wrap(
`02/12/2024
10:01-11:00
Machine 2 HC-3
Machine 3 Truck Radials - 37
Reason for low output- cutting Light Commercial tyres`
);

const TEXT10 = wrap(
`13/09/2025
09:00-10:00
CM2-offloading truck
CM2-20 Agri
CM2-offloading truck`
);

const TEXT_CASES = [
  ["Text 1 — HC+LC dual-line", TEXT1],
  ["Text 2 — HC/LC single-line", TEXT2],
  ["Text 3 — all machines dual-line", TEXT3],
  ["Text 4 — } typo as zero", TEXT4],
  ["Text 5 — CM3 dual-line, CM1/CM2 HC only", TEXT5],
  ["Text 6 — no parens, CM3 continuation line", TEXT6],
  ["Text 7 — agri vs agri treads with parens", TEXT7],
  ["Text 8 — bare LC=02 continuation", TEXT8],
  ["Text 9 — Machine X format", TEXT9],
  ["Text 10 — offloading + real data, expect 1 CM2 row only", TEXT10],
];

// TEXT11: ambiguous "CM-13 Agri Treads" — no other machine does AgriTreads that day → orange + log
const TEXT11 = `2025/09/15, 11:00 - Sender: 15/09/2025
11:00-12:00
CM1-N/A
CM2-52 Agri
CM-13 Agri Treads`;

function printRecordsWithLog(label, text) {
  printHeader(label);
  const { records, validationLog } = parseCuttingMessages(text);
  if (records.length === 0) {
    console.log(NO_RECORDS);
    return;
  }
  for (const r of records) {
    const amb = r._ambiguousLine ? " [ORANGE]" : "";
    console.log(`${formatRow(r)}${amb}`);
  }
  if (validationLog.length > 0) {
    console.log("  Validation log:");
    for (const v of validationLog) console.log(`    ⚠ ${v.issue}`);
  }
}

for (const [label, text] of TEXT_CASES) {
  printRecords(label, text);
}

printRecordsWithLog("Text 11 — ambiguous CM- line, no inference possible → orange + log", TEXT11);
