// Run with: node --test extension/tests
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";
import {
  dayNumberOf,
  dayNumberOfIso,
  findCrossing,
  formatCents,
  isoOfDayNumber,
  parseMoneyToCents,
  projectDecay,
  timeframeSummary,
  windowTotalCents,
} from "../calc.js";

const D = dayNumberOfIso;

test("money parsing and formatting", () => {
  assert.equal(parseMoneyToCents("$1,234.56"), 123456);
  assert.equal(parseMoneyToCents("($12.00)"), -1200);
  assert.equal(parseMoneyToCents("-$0.99"), -99);
  assert.ok(Number.isNaN(parseMoneyToCents("—")));
  assert.equal(formatCents(650000), "$6,500.00");
  assert.equal(formatCents(-1250), "-$12.50");
  assert.equal(formatCents(1250, { sign: true }), "+$12.50");
});

test("day numbers use the local calendar day", () => {
  const ts = new Date(2026, 10, 17, 23, 59).getTime(); // Nov 17, 11:59 pm local
  assert.equal(isoOfDayNumber(dayNumberOf(ts)), "2026-11-17");
  // Across a daylight-saving change, days stay whole.
  assert.equal(D("2026-11-02") - D("2026-10-31"), 2);
});

test("window counts (end - window, end]", () => {
  const orders = [
    { day: D("2026-10-28"), cents: 100 }, // exactly 30 days before: excluded
    { day: D("2026-10-29"), cents: 200 },
    { day: D("2026-11-27"), cents: 400 },
  ];
  assert.equal(windowTotalCents(orders, D("2026-11-27"), 30), 600);
});

test("projection finds the first day under the limit", () => {
  const orders = [
    { day: D("2026-11-01"), cents: 400000 },
    { day: D("2026-11-10"), cents: 300000 },
  ];
  const base = D("2026-11-20");
  const { rows, crossing } = projectDecay(orders, base, 30, 30, 650000);
  assert.equal(rows[0].totalCents, 700000);
  // The Nov 1 order leaves the window on Dec 1.
  assert.equal(isoOfDayNumber(crossing.day), "2026-12-01");
  assert.equal(crossing.offset, 11);
  assert.equal(crossing.totalCents, 300000);
  assert.equal(rows[11].rollOffCents, -400000);
  const warn = findCrossing(orders, base, 30, 600000);
  assert.equal(isoOfDayNumber(warn.day), "2026-12-01");
  assert.equal(findCrossing(orders, base, 30, 800000).alreadyBelow, true);
});

test("timeframe summary", () => {
  const orders = [
    { day: D("2026-11-20"), cents: 100 },
    { day: D("2026-11-14"), cents: 200 },
    { day: D("2026-11-13"), cents: 400 },
  ];
  const [week] = timeframeSummary(orders, D("2026-11-20"), [7]);
  assert.equal(week.count, 2);
  assert.equal(week.totalCents, 300);
  assert.equal(isoOfDayNumber(week.startDay), "2026-11-14");
});

// Same orders through the Python app's engine (tests/fixtures/parity.json, made by
// tests/make_parity_fixture.py): every projected day must match to the cent.
test("matches the Python app's projection", () => {
  const fixture = JSON.parse(readFileSync(new URL("./fixtures/parity.json", import.meta.url), "utf8"));
  const orders = fixture.orders.map((o) => ({ day: D(o.date), cents: o.cents }));
  const { rows } = projectDecay(orders, D(fixture.base), fixture.window, fixture.horizon, 650000);
  assert.deepEqual(rows.map((r) => r.totalCents), fixture.totals);
});
