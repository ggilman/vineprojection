// Copyright (c) 2026 George Gilman
// SPDX-License-Identifier: MIT
//
// Rolling-window math for the Vine ETV extension. Pure functions with no browser
// APIs, so they run in the dashboard and under `node --test`.
//
// Dates are local calendar days, held as "day numbers": days since 1970-01-01
// counted from the local year/month/day, so daylight-saving shifts never move
// an order to another day. Money is held in whole cents to avoid float drift.

export const MS_PER_DAY = 86400000;

/** Local calendar day number of an epoch-milliseconds timestamp. */
export function dayNumberOf(ms) {
  const d = new Date(ms);
  return Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / MS_PER_DAY);
}

/** Day number of a local Date (its calendar day). */
export function dayNumberOfDate(date) {
  return Math.round(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / MS_PER_DAY);
}

/** Day number of a "YYYY-MM-DD" string. */
export function dayNumberOfIso(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return Math.round(Date.UTC(y, m - 1, d) / MS_PER_DAY);
}

/** A local Date at midnight for a day number. */
export function dateOfDayNumber(n) {
  const u = new Date(n * MS_PER_DAY);
  return new Date(u.getUTCFullYear(), u.getUTCMonth(), u.getUTCDate());
}

/** "YYYY-MM-DD" for a day number. */
export function isoOfDayNumber(n) {
  return new Date(n * MS_PER_DAY).toISOString().slice(0, 10);
}

export function toCents(dollars) {
  return Math.round(Number(dollars) * 100);
}

/** "$1,234.56" from cents. */
export function formatCents(cents, { sign = false } = {}) {
  const neg = cents < 0;
  const abs = Math.abs(cents);
  const text = "$" + (abs / 100).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  if (neg) return "-" + text;
  return sign ? "+" + text : text;
}

/** Parse "$1,234.56" (or "($12.00)" / "-$12.00") to cents; NaN if not money. */
export function parseMoneyToCents(text) {
  const raw = String(text ?? "").trim();
  const negative = /^\(.*\)$/.test(raw) || raw.startsWith("-");
  const digits = raw.replace(/[^0-9.]/g, "");
  if (!digits || !/\d/.test(digits)) return NaN;
  const cents = Math.round(parseFloat(digits) * 100);
  return negative ? -cents : cents;
}

/**
 * Orders as [{day, cents}] (day = day number). Sum of cents for orders whose
 * day is in (endDay - windowDays, endDay]: `windowDays` calendar days ending on endDay.
 */
export function windowTotalCents(orders, endDay, windowDays) {
  const startExclusive = endDay - windowDays;
  let total = 0;
  for (const o of orders) {
    if (o.day > startExclusive && o.day <= endDay) total += o.cents;
  }
  return total;
}

/**
 * Day-by-day rolling total from baseDay through baseDay + horizonDays, assuming
 * no new orders. Returns { rows, crossing } where crossing is the first day the
 * total is below `limitCents` (or null within the horizon).
 */
export function projectDecay(orders, baseDay, windowDays, horizonDays, limitCents) {
  const rows = [];
  let crossing = null;
  let previous = null;
  for (let offset = 0; offset <= horizonDays; offset++) {
    const day = baseDay + offset;
    const total = windowTotalCents(orders, day, windowDays);
    const below = total < limitCents;
    if (below && crossing === null) {
      crossing = { day, offset, totalCents: total, alreadyBelow: offset === 0 };
    }
    rows.push({
      offset,
      day,
      totalCents: total,
      varianceCents: total - limitCents,
      rollOffCents: previous === null ? 0 : total - previous,
      below,
    });
    previous = total;
  }
  return { rows, crossing };
}

/**
 * First day on or after baseDay that the total drops below limitCents, looking
 * as far as `searchDays` (an order rolls off at most windowDays later, so
 * windowDays always suffices with no new orders). Null if never.
 */
export function findCrossing(orders, baseDay, windowDays, limitCents, searchDays = windowDays) {
  return projectDecay(orders, baseDay, windowDays, searchDays, limitCents).crossing;
}

/** Rolling totals for each period length (days) ending on baseDay. */
export function timeframeSummary(orders, baseDay, periods) {
  return periods.map((days) => {
    const inWindow = orders.filter((o) => o.day > baseDay - days && o.day <= baseDay);
    return {
      days,
      startDay: baseDay - days + 1,
      endDay: baseDay,
      count: inWindow.length,
      totalCents: inWindow.reduce((sum, o) => sum + o.cents, 0),
    };
  });
}
