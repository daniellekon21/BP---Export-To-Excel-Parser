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

run("future body date falls back to WhatsApp timestamp date", () => {
  const chat = [
    "2024/10/25, 08:15 - Driver: 25/10/2034",
    "Truck #: 1",
    "Depot: Thohoyandou",
    "Agricultural - 50",
    "*Total -*50*",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  const r = records[0];
  assert.equal(r.date.year, 2024);
  assert.equal(r.date.month, 10);
  assert.equal(r.date.day, 25);
});

run("body date is used when in the past (not future)", () => {
  const chat = [
    "2024/10/25, 08:15 - Driver: 01/10/2024",
    "Truck #: 1",
    "Depot: Thohoyandou",
    "Agricultural - 50",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  const r = records[0];
  assert.equal(r.date.year, 2024);
  assert.equal(r.date.month, 10);
  assert.equal(r.date.day, 1);
});

run("missing body date falls back to WhatsApp timestamp", () => {
  const chat = [
    "2024/10/25, 08:15 - Driver: Truck #: 1",
    "Depot: Thohoyandou",
    "Agricultural - 50",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  const r = records[0];
  assert.equal(r.date.year, 2024);
  assert.equal(r.date.month, 10);
  assert.equal(r.date.day, 25);
});
