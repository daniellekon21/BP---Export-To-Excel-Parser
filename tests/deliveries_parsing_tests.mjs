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

run("pasted sample: parses all metadata fields and Agricultural count", () => {
  const chat = [
    "2026/03/30, 08:15 - Driver: 30/03/2026",
    "Truck #: 1",
    "WBD: 2231",
    "GRV: 0936",
    "Depot: Thohoyandou",
    "Depot Manager: Adolf",
    "Collection #: 23781",
    "Transporter: Lehari",
    "Agricultural - 210",
    "*Total -*210*",
  ].join("\n");

  const { records, validationLog } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  const r = records[0];
  assert.deepEqual(r.date, { year: 2026, month: 3, day: 30 });
  assert.equal(r.delTruckNo, "1");
  assert.equal(r.delWbd, "2231");
  assert.equal(r.delGrv, "0936");
  assert.equal(r.delDepot, "Thohoyandou");
  assert.equal(r.delDepotManager, "Adolf");
  assert.equal(r.delCollectionNo, "23781");
  assert.equal(r.delTransporter, "Lehari");
  assert.equal(r.delAgricultural, 210);
  assert.equal(r.delTotalReported, 210);
  // Missing categories stay null (render as blank in Excel)
  assert.equal(r.delPassenger, null);
  assert.equal(r.delFourByFour, null);
  assert.equal(r.delMotorcycle, null);
  assert.equal(r.delLightCommercial, null);
  assert.equal(r.delHeavyCommercial, null);
  assert.equal(r.delHeavyCommercialSW, null);
  assert.equal(r.delHeavyCommercialT, null);
  // No validation issues expected for a well-formed message
  assert.equal(validationLog.length, 0);
});

run("multi-category message: extracts every tyre column", () => {
  const chat = [
    "2026/04/02, 10:30 - Driver: 02/04/2026",
    "Truck #: 4",
    "WBD: 2245",
    "GRV: 0948",
    "Depot: Giyani",
    "Depot Manager: Naledi",
    "Collection #: 23801",
    "Transporter: Lehari",
    "Passenger - 25",
    "4 x 4 - 10",
    "Motorcycle - 5",
    "Light Commercial - 60",
    "Heavy Commercial - 12",
    "Heavy Commercial SW - 7",
    "Heavy Commercial T - 3",
    "Agricultural - 20",
    "*Total -*142*",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  const r = records[0];
  assert.equal(r.delPassenger, 25);
  assert.equal(r.delFourByFour, 10);
  assert.equal(r.delMotorcycle, 5);
  assert.equal(r.delLightCommercial, 60);
  assert.equal(r.delHeavyCommercial, 12);
  assert.equal(r.delHeavyCommercialSW, 7);
  assert.equal(r.delHeavyCommercialT, 3);
  assert.equal(r.delAgricultural, 20);
  assert.equal(r.delTotalReported, 142);
});

run("non-delivery messages (no 'Truck #') are skipped", () => {
  const chat = [
    "2026/03/30, 08:15 - Admin: Good morning team",
    "2026/03/30, 09:00 - Driver: 30/03/2026",
    "Truck #: 1",
    "Depot: Thohoyandou",
    "Transporter: Lehari",
    "Agricultural - 50",
    "*Total -*50*",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  assert.equal(records[0].delTruckNo, "1");
});

run("rawMessage preserved on each record", () => {
  const chat = [
    "2026/03/30, 08:15 - Driver: 30/03/2026",
    "Truck #: 9",
    "Depot: Polokwane",
    "Agricultural - 10",
  ].join("\n");

  const { records } = parseDeliveriesMessages(chat);
  assert.equal(records.length, 1);
  assert.ok(records[0].rawMessage.includes("Truck #: 9"));
});
