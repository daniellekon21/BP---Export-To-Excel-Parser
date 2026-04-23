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

run("Heavy Commercial SW does not leak into Heavy Commercial", () => {
  const chat = [
    "2026/03/31, 09:05 - Driver: 31/03/2026",
    "Truck #: 3",
    "Depot: Thohoyandou",
    "Heavy Commercial SW - 22",
    "*Total -*22*",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  const r = records[0];
  assert.equal(r.delHeavyCommercialSW, 22);
  assert.equal(r.delHeavyCommercial, null);
  assert.equal(r.delHeavyCommercialT, null);
});

run("Heavy Commercial T does not leak into Heavy Commercial", () => {
  const chat = [
    "2026/03/31, 09:05 - Driver: 31/03/2026",
    "Truck #: 3",
    "Depot: Thohoyandou",
    "Heavy Commercial T - 18",
    "*Total -*18*",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  const r = records[0];
  assert.equal(r.delHeavyCommercialT, 18);
  assert.equal(r.delHeavyCommercial, null);
  assert.equal(r.delHeavyCommercialSW, null);
});

run("plain Heavy Commercial parses independently when listed", () => {
  const chat = [
    "2026/03/31, 09:05 - Driver: 31/03/2026",
    "Truck #: 3",
    "Depot: Thohoyandou",
    "Heavy Commercial - 8",
    "Heavy Commercial SW - 22",
    "Heavy Commercial T - 18",
    "*Total -*48*",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  const r = records[0];
  assert.equal(r.delHeavyCommercial, 8);
  assert.equal(r.delHeavyCommercialSW, 22);
  assert.equal(r.delHeavyCommercialT, 18);
});

run("4x4 and 4 x 4 both map to delFourByFour", () => {
  const chat1 = [
    "2026/03/31, 09:05 - Driver: 31/03/2026",
    "Truck #: 1",
    "Depot: X",
    "4x4 - 6",
  ].join("\n");
  const chat2 = [
    "2026/03/31, 09:05 - Driver: 31/03/2026",
    "Truck #: 1",
    "Depot: X",
    "4 x 4 - 9",
  ].join("\n");

  const r1 = parseDeliveriesMessages(chat1).records[0];
  const r2 = parseDeliveriesMessages(chat2).records[0];
  assert.equal(r1.delFourByFour, 6);
  assert.equal(r2.delFourByFour, 9);
});
