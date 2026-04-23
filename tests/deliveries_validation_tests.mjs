import assert from "node:assert/strict";
import { parseDeliveriesMessages } from "../src/parsing/deliveriesParser.js";

function run(name, fn) {
  try {
    fn();
    console.log(`PASS ${name}`);
  } catch (err) {
    console.error(`FAIL ${name}`);
    console.error(err.message);
    process.exitCode = 1;
  }
}

run("Total mismatch produces a validation log entry", () => {
  const chat = [
    "2026/03/30, 08:15 - Driver: 30/03/2026",
    "Truck #: 1",
    "Depot: Thohoyandou",
    "Transporter: Lehari",
    "Passenger - 10",
    "Agricultural - 30",
    "*Total -*999*",
  ].join("\n");

  const { records, validationLog } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  assert.ok(validationLog.length >= 1);
  const mismatch = validationLog.find((v) => /Total mismatch/i.test(v.issue));
  assert.ok(mismatch, "expected a Total mismatch validation entry");
  assert.ok(mismatch.issue.includes("999"));
  assert.ok(mismatch.issue.includes("40"));
});

run("correct Total produces no mismatch entry", () => {
  const chat = [
    "2026/03/30, 08:15 - Driver: 30/03/2026",
    "Truck #: 1",
    "Depot: Thohoyandou",
    "Transporter: Lehari",
    "Passenger - 10",
    "Agricultural - 30",
    "*Total -*40*",
  ].join("\n");

  const { validationLog } = parseDeliveriesMessages(chat);
  const mismatch = validationLog.find((v) => /Total mismatch/i.test(v.issue));
  assert.equal(mismatch, undefined);
});

run("message with no tyre categories logs a validation entry", () => {
  const chat = [
    "2026/03/30, 08:15 - Driver: 30/03/2026",
    "Truck #: 7",
    "Depot: Polokwane",
    "Transporter: Lehari",
  ].join("\n");

  const { records, validationLog } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  const noCats = validationLog.find((v) => /No tyre categories/i.test(v.issue));
  assert.ok(noCats, "expected a 'No tyre categories parsed' validation entry");
});
