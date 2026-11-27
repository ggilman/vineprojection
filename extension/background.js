// Copyright (c) 2026 George Gilman
// SPDX-License-Identifier: MIT
//
// The toolbar button opens the dashboard in a tab, or focuses it if it is
// already open. Nothing runs in the background otherwise.

const DASHBOARD = "dashboard.html";

chrome.action.onClicked.addListener(async () => {
  const url = chrome.runtime.getURL(DASHBOARD);
  const contexts = await chrome.runtime.getContexts({
    contextTypes: ["TAB"],
    documentUrls: [url],
  });
  const open = contexts.find((c) => c.tabId >= 0);
  if (open) {
    await chrome.tabs.update(open.tabId, { active: true });
    if (open.windowId >= 0) await chrome.windows.update(open.windowId, { focused: true });
    return;
  }
  await chrome.tabs.create({ url });
});
