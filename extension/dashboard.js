// Copyright (c) 2026 George Gilman
// SPDX-License-Identifier: MIT
//
// Dashboard page: shows the saved orders and claim status at once, and reads
// Amazon only when the user presses Refresh.

import {
  dateOfDayNumber,
  dayNumberOf,
  dayNumberOfDate,
  dayNumberOfIso,
  findCrossing,
  formatCents,
  isoOfDayNumber,
  projectDecay,
  timeframeSummary,
  toCents,
} from "./calc.js";
import {
  CAP_BLOCKED,
  CAP_CAPPED,
  CAP_NEEDS_LOGIN,
  CAP_OPEN,
  VineAccessError,
  fetchCapStatus,
  fetchRecentOrders,
} from "./vine.js";

const DEFAULT_SETTINGS = {
  windowDays: 30,
  blockLimit: 6500,
  warningLevel: 6000,
  matchHorizon: true,
  horizonDays: 30,
  baseMode: "today",
  customDate: "",
};
const MIN_LOOKBACK_DAYS = 30;

const $ = (id) => document.getElementById(id);
const state = { settings: { ...DEFAULT_SETTINGS }, ordersData: null, lastCheck: null, busy: false };

// ---------------------------------------------------------------- formatting

const shortDate = (day, refDay) => {
  const d = dateOfDayNumber(day);
  const text = d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
  const refYear = dateOfDayNumber(refDay ?? day).getFullYear();
  return d.getFullYear() === refYear ? text : `${text}, ${d.getFullYear()}`;
};
const longDate = (day) =>
  dateOfDayNumber(day).toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
const plural = (n, word) => `${n.toLocaleString("en-US")} ${word}${n === 1 ? "" : "s"}`;

function formatCheckedAt(ms) {
  const d = new Date(ms);
  const time = d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
  return dayNumberOfDate(d) === dayNumberOfDate(new Date())
    ? `today ${time}`
    : `${d.toLocaleDateString("en-US", { month: "short", day: "numeric" })}, ${time}`;
}

function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (key === "class") node.className = value;
    else node.setAttribute(key, value);
  }
  for (const child of children) node.append(child);
  return node;
}

// Banner text with **bold** runs, built without innerHTML (product names and
// Amazon messages are untrusted).
function banner(kind, text) {
  const box = el("div", { class: `banner banner-${kind}` });
  text.split(/\*\*/).forEach((part, i) => box.append(i % 2 ? el("strong", {}, part) : part));
  return box;
}

// ---------------------------------------------------------------- storage

async function loadState() {
  const saved = await chrome.storage.local.get(["settings", "ordersData", "lastCheck"]);
  state.settings = { ...DEFAULT_SETTINGS, ...(saved.settings || {}) };
  state.ordersData = saved.ordersData || null;
  state.lastCheck = saved.lastCheck || null;
}

const saveSettings = () => chrome.storage.local.set({ settings: state.settings });

// ---------------------------------------------------------------- settings UI

function readNumber(id, fallback, min) {
  const value = Number($(id).value);
  return Number.isFinite(value) && value >= min ? value : fallback;
}

function bindSettings() {
  const s = state.settings;
  $("window-days").value = s.windowDays;
  $("block-limit").value = s.blockLimit;
  $("warning-level").value = s.warningLevel;
  $("match-horizon").checked = s.matchHorizon;
  $("horizon-days").value = s.matchHorizon ? s.windowDays : s.horizonDays;
  $("horizon-days").disabled = s.matchHorizon;
  for (const radio of document.querySelectorAll('input[name="base-mode"]')) {
    radio.checked = radio.value === s.baseMode;
  }
  $("custom-date").hidden = s.baseMode !== "custom";
  $("custom-date").value = s.customDate || isoOfDayNumber(dayNumberOfDate(new Date()));
}

function onSettingsChange() {
  const s = state.settings;
  s.windowDays = Math.min(365, Math.round(readNumber("window-days", s.windowDays, 1)));
  s.blockLimit = readNumber("block-limit", s.blockLimit, 0);
  s.warningLevel = readNumber("warning-level", s.warningLevel, 0);
  s.matchHorizon = $("match-horizon").checked;
  s.horizonDays = Math.min(365, Math.round(readNumber("horizon-days", s.horizonDays, 1)));
  s.baseMode = document.querySelector('input[name="base-mode"]:checked')?.value || "today";
  s.customDate = $("custom-date").value;
  $("horizon-days").disabled = s.matchHorizon;
  if (s.matchHorizon) $("horizon-days").value = s.windowDays;
  $("custom-date").hidden = s.baseMode !== "custom";
  saveSettings();
  render();
}

// ---------------------------------------------------------------- refresh

function showProgress(text) {
  $("progress").hidden = !text;
  $("progress").textContent = text || "";
}

function showError(text) {
  $("error").hidden = !text;
  $("error").textContent = text || "";
}

async function refresh() {
  if (state.busy) return;
  state.busy = true;
  $("refresh").disabled = true;
  showError("");
  const check = { checkedAt: Date.now(), capStatus: null, error: null };
  try {
    showProgress("Checking whether Vine claiming is open…");
    check.capStatus = await fetchCapStatus();
    const lookback = Math.max(MIN_LOOKBACK_DAYS, state.settings.windowDays);
    const result = await fetchRecentOrders(lookback, showProgress);
    state.ordersData = {
      ...result,
      orders: result.orders.map((o) => ({ ...o })),
    };
    await chrome.storage.local.set({ ordersData: state.ordersData });
  } catch (err) {
    if (err instanceof VineAccessError) check.capStatus = err.status;
    check.error = err.message || String(err);
    showError(check.error);
  } finally {
    // Keep the last known claim status if this check never reached Vine.
    if (check.capStatus || !state.lastCheck) {
      state.lastCheck = check;
      await chrome.storage.local.set({ lastCheck: check });
    }
    state.busy = false;
    $("refresh").disabled = false;
    showProgress("");
    render();
  }
}

// ---------------------------------------------------------------- rendering

function renderStatus() {
  const pill = $("cap-pill");
  const check = state.lastCheck;
  const labels = {
    [CAP_OPEN]: ["Claiming open", "pill-open"],
    [CAP_CAPPED]: ["Claiming paused", "pill-capped"],
    [CAP_NEEDS_LOGIN]: ["Sign-in needed", "pill-login"],
    [CAP_BLOCKED]: ["Bot check", "pill-blocked"],
  };
  const [text, cls] = labels[check?.capStatus] || ["Not checked yet", "pill-unknown"];
  pill.textContent = text;
  pill.className = `pill ${cls}`;
  $("checked-at").textContent = check ? `checked ${formatCheckedAt(check.checkedAt)}` : "";
}

function render() {
  renderStatus();
  const data = state.ordersData;
  const hasData = !!data;
  $("empty").hidden = hasData;
  $("content").hidden = !hasData;
  if (!hasData) return;

  const s = state.settings;
  const windowDays = s.windowDays;
  const horizon = s.matchHorizon ? windowDays : s.horizonDays;
  const blockCents = toCents(s.blockLimit);
  const warnCents = toCents(s.warningLevel);
  const showWarning = warnCents > 0 && warnCents < blockCents;
  const orders = data.orders.map((o) => ({ ...o, day: dayNumberOf(o.ts) }));
  const today = dayNumberOfDate(new Date());
  const latestDay = orders.length ? Math.max(...orders.map((o) => o.day)) : today;
  let baseDay = today;
  if (s.baseMode === "latest") baseDay = latestDay;
  else if (s.baseMode === "custom" && s.customDate) baseDay = dayNumberOfIso(s.customDate);

  const { rows, crossing } = projectDecay(orders, baseDay, windowDays, horizon, blockCents);
  const blockCross = crossing ?? findCrossing(orders, baseDay, windowDays, blockCents);
  const warnCross = showWarning ? findCrossing(orders, baseDay, windowDays, warnCents) : null;
  const current = rows[0].totalCents;
  const variance = current - blockCents;

  // KPIs
  $("kpi-window-label").textContent = `Current ${windowDays}-day total`;
  $("kpi-window").textContent = formatCents(current);
  const sub = $("kpi-window-sub");
  sub.textContent = variance >= 0
    ? `${formatCents(variance)} over the block limit`
    : `${formatCents(-variance)} under the block limit`;
  sub.className = `kpi-sub ${variance >= 0 ? "bad" : "good"}`;
  $("kpi-block").textContent = formatCents(blockCents);
  $("kpi-block-sub").textContent = showWarning ? `Warning at ${formatCents(warnCents)}` : "No warning level";
  if (!blockCross) {
    $("kpi-cross").textContent = "Not reached";
    $("kpi-cross-sub").textContent = `Stays at or above it for ${plural(windowDays, "day")}`;
  } else if (blockCross.alreadyBelow) {
    $("kpi-cross").textContent = "Under now";
    $("kpi-cross-sub").textContent = `${formatCents(blockCross.totalCents)} as of ${shortDate(baseDay)}`;
  } else {
    $("kpi-cross").textContent = shortDate(blockCross.day, baseDay);
    $("kpi-cross-sub").textContent =
      `${formatCents(blockCross.totalCents)} · in ${plural(blockCross.offset, "day")}` +
      (blockCross.offset > horizon ? " (past horizon)" : "");
  }
  const inWindow = orders.filter((o) => o.day > baseDay - windowDays && o.day <= baseDay).length;
  $("kpi-orders-label").textContent = `Orders in the ${windowDays}-day window`;
  $("kpi-orders").textContent = inWindow.toLocaleString("en-US");
  $("kpi-orders-sub").textContent =
    `${plural(orders.length, "order")} read ${formatCheckedAt(data.fetchedAt)} (${plural(data.pagesRead, "page")})`;

  renderBanners({ baseDay, windowDays, horizon, current, variance, blockCents, warnCents, showWarning, blockCross, warnCross, today, data });
  renderChart(rows, { blockCents, warnCents, showWarning, crossing, warnCross, windowDays, horizon });
  renderTables(orders, rows, baseDay, data);
}

function renderBanners(ctx) {
  const box = $("banners");
  box.replaceChildren();
  const { baseDay, windowDays, current, variance, blockCents, warnCents, showWarning, blockCross, warnCross, today, data } = ctx;
  const check = state.lastCheck;
  const asOf = check ? ` (as of ${formatCheckedAt(check.checkedAt)})` : "";

  if (check?.capStatus === CAP_CAPPED) {
    box.append(banner("critical", `**Claiming is paused.** Vine has paused your claiming because your cap was reached${asOf}.`));
  } else if (check?.capStatus === CAP_OPEN) {
    box.append(banner("good", `**Claiming is open**${asOf}.`));
  } else if (check?.capStatus === CAP_NEEDS_LOGIN || check?.capStatus === CAP_BLOCKED) {
    box.append(banner("warning", `**Amazon check unavailable:** ${check.error || "sign-in needed."}`));
  }

  const inWarningZone = showWarning && current >= warnCents && current < blockCents;
  if (inWarningZone) {
    const when = warnCross
      ? `drops below the warning level on **${longDate(warnCross.day)}** (in ${plural(warnCross.offset, "day")})`
      : "stays at or above the warning level";
    box.append(banner("warning",
      `**In the warning zone:** your ${windowDays}-day total as of ${longDate(baseDay)} is **${formatCents(current)}**, ` +
      `at or above the ${formatCents(warnCents)} warning level and ${formatCents(-variance)} under the ` +
      `${formatCents(blockCents)} block limit. Assuming no new orders, it ${when}.`));
  } else if (current < blockCents) {
    box.append(banner("good",
      `**Under the block limit:** your ${windowDays}-day total as of ${longDate(baseDay)} is ` +
      `**${formatCents(current)}**, ${formatCents(-variance)} below ${formatCents(blockCents)}.`));
  } else if (blockCross) {
    const warnText = showWarning && warnCross
      ? ` It drops below the ${formatCents(warnCents)} warning level on **${longDate(warnCross.day)}**.`
      : "";
    box.append(banner("info",
      `**Over the block limit:** assuming no new orders, your ${windowDays}-day total drops below ` +
      `${formatCents(blockCents)} on **${longDate(blockCross.day)}** (in ${plural(blockCross.offset, "day")}) ` +
      `at ${formatCents(blockCross.totalCents)}.${warnText}`));
  }

  // Data coverage: the window can only count orders that were read.
  const readFromDay = dayNumberOfDate(new Date(data.fetchedAt)) - data.lookbackDays + 1;
  if (!data.complete) {
    box.append(banner("warning",
      `Stopped after ${plural(data.pagesRead, "page")} before reaching ${data.lookbackDays} days back, so totals may be low.`));
  }
  if (baseDay - windowDays + 1 < readFromDay) {
    box.append(banner("warning",
      `Orders were read back to ${longDate(readFromDay)} only, so this window misses older orders. ` +
      "Press Refresh to read further back."));
  }
  if (baseDay > today) {
    box.append(banner("info", "The base date is in the future, so the window counts only orders already placed."));
  }
}

// ---------------------------------------------------------------- chart

const SVG_NS = "http://www.w3.org/2000/svg";
function svg(tag, attrs = {}, text) {
  const node = document.createElementNS(SVG_NS, tag);
  for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value);
  if (text !== undefined) node.textContent = text;
  return node;
}

function niceStep(maxValue, targetTicks = 5) {
  const raw = maxValue / targetTicks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const norm = raw / mag;
  return (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
}

let lastChart = null;

function renderChart(rows, opts) {
  lastChart = { rows, opts };
  const { blockCents, warnCents, showWarning, crossing, warnCross, windowDays, horizon } = opts;
  const chart = $("chart");
  const width = Math.max(320, $("chart-wrap").clientWidth);
  const height = 380;
  const m = { top: 16, right: 110, bottom: 34, left: 64 };
  const plotW = width - m.left - m.right;
  const plotH = height - m.top - m.bottom;
  chart.setAttribute("viewBox", `0 0 ${width} ${height}`);
  chart.replaceChildren();

  const maxCents = Math.max(blockCents, ...rows.map((r) => r.totalCents), 100);
  const step = niceStep(maxCents / 100) * 100;
  const yMax = Math.ceil((maxCents * 1.05) / step) * step;
  const x = (offset) => m.left + (rows.length > 1 ? (offset / (rows.length - 1)) * plotW : plotW / 2);
  const y = (cents) => m.top + plotH - (cents / yMax) * plotH;

  $("chart-title").textContent = `Projected ${windowDays}-day rolling total`;
  $("chart-sub").textContent =
    `Fair market value of orders in the window, day by day for ${plural(horizon, "day")}, assuming no new orders.`;

  // Grid and y ticks
  for (let v = 0; v <= yMax; v += step) {
    chart.append(svg("line", { class: v === 0 ? "axis" : "grid", x1: m.left, x2: m.left + plotW, y1: y(v), y2: y(v) }));
    chart.append(svg("text", { x: m.left - 8, y: y(v) + 4, "text-anchor": "end" }, formatCents(v).replace(/\.00$/, "")));
  }
  // X ticks: about one per week, always the first and last day
  const every = Math.max(1, Math.round(rows.length / 7));
  rows.forEach((r, i) => {
    if (i % every !== 0 && i !== rows.length - 1) return;
    if (i !== rows.length - 1 && rows.length - 1 - i < every / 2) return;
    chart.append(svg("text", { x: x(r.offset), y: m.top + plotH + 20, "text-anchor": "middle" }, shortDate(r.day, rows[0].day)));
  });

  // Reference lines, labelled at the right edge
  const refLine = (cents, cls, label) => {
    chart.append(svg("line", { class: cls, x1: m.left, x2: m.left + plotW, y1: y(cents), y2: y(cents) }));
    chart.append(svg("text", { class: "ref-label", x: m.left + plotW + 8, y: y(cents) + 4 }, label));
  };
  refLine(blockCents, "block", `Block ${formatCents(blockCents).replace(/\.00$/, "")}`);
  if (showWarning) {
    // Keep the two labels from overlapping when the lines are close.
    refLine(warnCents, "warn", `Warning ${formatCents(warnCents).replace(/\.00$/, "")}`);
    const labels = chart.querySelectorAll(".ref-label");
    const gap = Number(labels[1].getAttribute("y")) - Number(labels[0].getAttribute("y"));
    if (gap < 14) labels[1].setAttribute("y", Number(labels[0].getAttribute("y")) + 14);
  }

  // Series line and markers
  const path = rows.map((r, i) => `${i ? "L" : "M"}${x(r.offset).toFixed(1)},${y(r.totalCents).toFixed(1)}`).join("");
  chart.append(svg("path", { class: "series", d: path }));
  if (rows.length <= 45) {
    for (const r of rows) chart.append(svg("circle", { class: "dot", cx: x(r.offset), cy: y(r.totalCents), r: 3.5 }));
  }
  if (crossing && !crossing.alreadyBelow) {
    chart.append(svg("circle", { class: "cross-block", cx: x(crossing.offset), cy: y(crossing.totalCents), r: 7 }));
  }
  if (showWarning && warnCross && !warnCross.alreadyBelow && warnCross.offset <= horizon) {
    const cx = x(warnCross.offset), cy = y(warnCross.totalCents), r = 7;
    chart.append(svg("path", { class: "cross-warn", d: `M${cx},${cy - r}L${cx + r},${cy}L${cx},${cy + r}L${cx - r},${cy}Z` }));
  }

  // Hover layer: crosshair snaps to the nearest day; one tooltip with every value.
  const crosshair = svg("line", { class: "crosshair", y1: m.top, y2: m.top + plotH, visibility: "hidden" });
  const hoverDot = svg("circle", { class: "hover-dot", r: 5, visibility: "hidden" });
  const hit = svg("rect", { x: m.left, y: m.top, width: plotW, height: plotH, fill: "transparent", tabindex: 0 });
  chart.append(crosshair, hoverDot, hit);
  const tooltip = $("tooltip");

  const showAt = (index) => {
    const r = rows[index];
    const cx = x(r.offset);
    crosshair.setAttribute("x1", cx);
    crosshair.setAttribute("x2", cx);
    crosshair.setAttribute("visibility", "visible");
    hoverDot.setAttribute("cx", cx);
    hoverDot.setAttribute("cy", y(r.totalCents));
    hoverDot.setAttribute("visibility", "visible");
    tooltip.replaceChildren(
      el("div", { class: "tt-date" }, `${longDate(r.day)} · day +${r.offset}`),
      el("div", { class: "tt-row" }, el("span", { class: "key" }), el("strong", {}, formatCents(r.totalCents)), "rolling total"),
      el("div", {}, `${formatCents(r.varianceCents, { sign: true })} vs. block limit`),
      el("div", {}, `Rolls off that day: ${formatCents(-r.rollOffCents)}`),
    );
    tooltip.hidden = false;
    const scale = $("chart-wrap").clientWidth / width;
    const left = cx * scale;
    const flip = left > $("chart-wrap").clientWidth - 200;
    tooltip.style.left = `${flip ? left - 190 : left + 12}px`;
    tooltip.style.top = `${Math.max(0, y(r.totalCents) * scale - 40)}px`;
  };
  const hide = () => {
    crosshair.setAttribute("visibility", "hidden");
    hoverDot.setAttribute("visibility", "hidden");
    tooltip.hidden = true;
  };
  let focusIndex = 0;
  hit.addEventListener("pointermove", (ev) => {
    const box = chart.getBoundingClientRect();
    const px = ((ev.clientX - box.left) / box.width) * width;
    focusIndex = Math.max(0, Math.min(rows.length - 1, Math.round(((px - m.left) / plotW) * (rows.length - 1))));
    showAt(focusIndex);
  });
  hit.addEventListener("pointerleave", hide);
  hit.addEventListener("focus", () => showAt(focusIndex));
  hit.addEventListener("blur", hide);
  hit.addEventListener("keydown", (ev) => {
    if (ev.key === "ArrowRight" || ev.key === "ArrowLeft") {
      focusIndex = Math.max(0, Math.min(rows.length - 1, focusIndex + (ev.key === "ArrowRight" ? 1 : -1)));
      showAt(focusIndex);
      ev.preventDefault();
    }
  });

  // Legend (line keys)
  const legend = $("legend");
  legend.replaceChildren();
  const key = (color, dash, label) => {
    const i = el("i");
    i.style.borderTopColor = color;
    i.style.borderTopStyle = dash;
    legend.append(el("span", {}, i, label));
  };
  key("var(--series)", "solid", "Rolling total");
  key("var(--critical)", "dashed", "Block limit");
  if (showWarning) key("var(--warning)", "dotted", "Warning level");
}

// ---------------------------------------------------------------- tables

function renderTables(orders, rows, baseDay, data) {
  const periods = [7, 30, 60, 90].filter((d) => d <= data.lookbackDays);
  if (!periods.includes(state.settings.windowDays) && state.settings.windowDays <= data.lookbackDays) {
    periods.push(state.settings.windowDays);
    periods.sort((a, b) => a - b);
  }
  $("summary-sub").textContent = `Ending ${longDate(baseDay)}, by order date.`;
  $("summary-body").replaceChildren(
    ...timeframeSummary(orders, baseDay, periods).map((p) =>
      el("tr", {},
        el("td", {}, `Last ${p.days} days`),
        el("td", {}, `${shortDate(p.startDay, baseDay)} – ${shortDate(p.endDay, baseDay)}`),
        el("td", { class: "num" }, p.count.toLocaleString("en-US")),
        el("td", { class: "num" }, formatCents(p.totalCents)),
      )),
  );

  $("schedule-body").replaceChildren(
    ...rows.map((r) =>
      el("tr", { class: r.below ? "below" : "" },
        el("td", {}, isoOfDayNumber(r.day)),
        el("td", { class: "num" }, formatCents(r.totalCents)),
        el("td", { class: "num" }, formatCents(r.varianceCents, { sign: true })),
        el("td", {}, r.below ? "Below limit" : "At or above limit"),
      )),
  );

  const sorted = [...orders].sort((a, b) => b.ts - a.ts);
  $("orders-summary").textContent =
    `Orders read from amazon.com/vine/orders (${plural(sorted.length, "order")}, ${formatCents(sorted.reduce((s, o) => s + o.cents, 0))})`;
  $("orders-body").replaceChildren(
    ...sorted.map((o) => {
      const asinCell = el("td", {});
      if (o.asin) {
        asinCell.append(el("a", { href: `https://www.amazon.com/dp/${o.asin}`, target: "_blank", rel: "noopener" }, o.asin));
      }
      const product = el("td", { class: "product", title: o.name }, o.name);
      return el("tr", {},
        el("td", {}, isoOfDayNumber(o.day)),
        product,
        asinCell,
        el("td", { class: "num" }, formatCents(o.cents)),
      );
    }),
  );
}

// ---------------------------------------------------------------- start

async function start() {
  await loadState();
  bindSettings();
  for (const id of ["window-days", "block-limit", "warning-level", "match-horizon", "horizon-days", "custom-date"]) {
    $(id).addEventListener("change", onSettingsChange);
  }
  for (const radio of document.querySelectorAll('input[name="base-mode"]')) {
    radio.addEventListener("change", onSettingsChange);
  }
  $("reset-settings").addEventListener("click", () => {
    state.settings = { ...DEFAULT_SETTINGS };
    bindSettings();
    saveSettings();
    render();
  });
  $("refresh").addEventListener("click", refresh);
  let resizeTimer = 0;
  window.addEventListener("resize", () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => lastChart && renderChart(lastChart.rows, lastChart.opts), 120);
  });
  render();
}

start();
