import assert from "node:assert/strict";
import { parseBalingMessages } from "../src/parsing/balingParser.js";

function wrap(body, ts = "2024/10/18, 14:25") {
  return `${ts} - Tester: ${body}`;
}

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

function firstStandardMachine(body) {
  const parsed = parseBalingMessages(wrap(body));
  assert.equal(parsed.standardRecords.length, 1);
  return parsed.standardRecords[0].machine;
}

run("Machine 2 maps to BM - 2", () => {
  const m = firstStandardMachine("Machine 2 Date - 18/10/2024 B329 - Production Operator - A Assistant - B Item - Passenger Qty - 10 Total Qty - 10 Weight - 900kg");
  assert.equal(m, "BM - 2");
});

run("Machine-2 maps to BM - 2", () => {
  const m = firstStandardMachine("Machine-2 Date - 18/10/2024 B330 - Production Operator - A Assistant - B Item - Passenger Qty - 10 Total Qty - 10 Weight - 900kg");
  assert.equal(m, "BM - 2");
});

run("Machine2 maps to BM - 2", () => {
  const m = firstStandardMachine("Machine2 Date - 18/10/2024 B331 - Production Operator - A Assistant - B Item - Passenger Qty - 10 Total Qty - 10 Weight - 900kg");
  assert.equal(m, "BM - 2");
});

run("Machine One maps to BM - 1", () => {
  const m = firstStandardMachine("Machine One Date - 18/10/2024 B332 - Production Operator - A Assistant - B Item - Passenger Qty - 10 Total Qty - 10 Weight - 900kg");
  assert.equal(m, "BM - 1");
});
