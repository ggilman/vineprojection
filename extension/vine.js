// Copyright (c) 2026 George Gilman
// SPDX-License-Identifier: MIT
//
// Reads Amazon Vine pages with the browser's own Amazon sign-in. Runs in the
// dashboard page (it needs DOMParser); the extension's host permission for
// www.amazon.com lets these requests carry the user's Amazon cookies.

import { dayNumberOf, dayNumberOfDate, parseMoneyToCents } from "./calc.js";

export const VINE_ORIGIN = "https://www.amazon.com";
export const ORDERS_PATH = "/vine/orders";
export const ITEMS_PATH = "/vine/vine-items";
export const MAX_ORDER_PAGES = 80; // 10 orders per page
export const PAGE_DELAY_MS = 700; // pause between page reads, to go easy on Amazon

export const CAP_OPEN = "OPEN";
export const CAP_CAPPED = "CAPPED";
export const CAP_NEEDS_LOGIN = "NEEDS_LOGIN";
export const CAP_BLOCKED = "BLOCKED";
export const CAP_UNKNOWN = "UNKNOWN";

// The messages Vine shows while claiming is paused.
const CAP_MESSAGE = /claiming\s+will\s+resume\s+soon|vine\s+limits\s+claiming\s+activity/i;
// Amazon's bot check. (Not the "automated access" footer: every page has that.)
const BLOCKED_MESSAGE = /enter the characters you see in this image|api-services-support@amazon\.com/i;
const LOGIN_URL = /\/ap\/|\/signin/i;
const CAPTCHA_URL = /\/errors\/validateCaptcha/i;

export class VineAccessError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status; // CAP_NEEDS_LOGIN or CAP_BLOCKED
  }
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function fetchVinePage(path) {
  let resp;
  try {
    resp = await fetch(VINE_ORIGIN + path, { credentials: "include", cache: "no-store" });
  } catch (err) {
    throw new Error(`Could not reach amazon.com (${err.message}).`);
  }
  const url = resp.url || "";
  const html = await resp.text();
  if (LOGIN_URL.test(new URL(url, VINE_ORIGIN).pathname)) {
    throw new VineAccessError(
      CAP_NEEDS_LOGIN,
      "Amazon wants you to sign in. Open amazon.com/vine in this browser, sign in, then press Refresh.",
    );
  }
  if (CAPTCHA_URL.test(url) || BLOCKED_MESSAGE.test(html)) {
    throw new VineAccessError(
      CAP_BLOCKED,
      "Amazon showed a bot check. Open amazon.com/vine in a tab, complete it, then press Refresh.",
    );
  }
  if (!resp.ok) throw new Error(`Amazon answered ${resp.status} for ${path}.`);
  return new DOMParser().parseFromString(html, "text/html");
}

/** Rows of one Vine Orders page: [{ts, cents, asin, name}]. */
export function parseOrdersPage(doc) {
  const rows = [];
  for (const tr of doc.querySelectorAll("table.vvp-orders-table tr.vvp-orders-table--row")) {
    // The visible date is filled in by Amazon's script; the cell carries epoch ms.
    const tsCell = tr.querySelector("[data-order-timestamp]");
    const cells = tr.querySelectorAll("td");
    const ts = tsCell ? Number(tsCell.getAttribute("data-order-timestamp")) : NaN;
    const cents = cells.length > 3 ? parseMoneyToCents(cells[3].textContent) : NaN;
    if (!Number.isFinite(ts) || !Number.isFinite(cents)) continue;
    const link = tr.querySelector('a[href*="/dp/"]');
    const asinMatch = link ? (link.getAttribute("href") || "").match(/\/dp\/([A-Z0-9]{10})/) : null;
    const nameEl = tr.querySelector(".a-truncate-full") || cells[1];
    rows.push({
      ts,
      cents,
      asin: asinMatch ? asinMatch[1] : "",
      name: nameEl ? nameEl.textContent.replace(/\s+/g, " ").trim() : "",
    });
  }
  return rows;
}

/**
 * Walk the Orders pages (newest first) and collect orders placed in the last
 * `lookbackDays` days, stopping at the first page that reaches older ones.
 */
export async function fetchRecentOrders(lookbackDays, onProgress = () => {}) {
  const cutoffDay = dayNumberOfDate(new Date()) - lookbackDays; // orders after this day count
  const orders = [];
  let pagesRead = 0;
  let complete = false;
  for (let page = 1; page <= MAX_ORDER_PAGES; page++) {
    if (page > 1) await sleep(PAGE_DELAY_MS);
    onProgress(`Reading Vine orders page ${page}…`);
    const doc = await fetchVinePage(`${ORDERS_PATH}?page=${page}`);
    const rows = parseOrdersPage(doc);
    pagesRead++;
    if (rows.length === 0) {
      if (page === 1 && !doc.querySelector("table.vvp-orders-table")) {
        throw new Error("The Vine Orders page did not show an orders table. Is this account in Vine?");
      }
      complete = true; // past the last page
      break;
    }
    let reachedCutoff = false;
    for (const row of rows) {
      if (dayNumberOf(row.ts) > cutoffDay) orders.push(row);
      else reachedCutoff = true;
    }
    if (reachedCutoff) {
      complete = true;
      break;
    }
  }
  return { orders, pagesRead, complete, lookbackDays, fetchedAt: Date.now() };
}

/**
 * Whether Vine claiming is paused. Vine pages carry a hidden copy of the cap
 * message (class aok-hidden), so hidden elements are dropped before matching.
 */
export async function fetchCapStatus() {
  const doc = await fetchVinePage(ITEMS_PATH);
  doc.querySelectorAll(".aok-hidden, [hidden], script, style, template, noscript").forEach((e) => e.remove());
  doc.querySelectorAll("[style]").forEach((e) => {
    if (/display\s*:\s*none|visibility\s*:\s*hidden/i.test(e.getAttribute("style"))) e.remove();
  });
  const text = doc.body ? doc.body.textContent : "";
  return CAP_MESSAGE.test(text) ? CAP_CAPPED : CAP_OPEN;
}
