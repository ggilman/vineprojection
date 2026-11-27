# Vine ETV Analyzer: browser extension (Chrome & Edge)

The same rolling-window view as the Python app, as a Manifest V3 extension. It needs
no Python, no installer, no separate Chrome window, and there is nothing to quit:
close the tab and it's done.

## What it does

- Reads your **Vine Orders** pages (`amazon.com/vine/orders?page=1`, `2`, …, newest
  first) using the Amazon sign-in already in your browser, and adds up the
  **Fair market value** of every order placed in the rolling window (by order date).
  It stops at the first page older than the window.
- Checks `amazon.com/vine/vine-items` for the "claiming paused" message.
- Shows the current window total, the **block limit** ($6,500 by default) and
  **warning level** ($6,000 by default), the date you drop back under each, a
  day-by-day projection chart, rolling totals, and the orders it read.
- Saves the orders and the last claim status in the extension's storage, so it
  opens instantly on that data. Amazon is contacted only when you press
  **Refresh from Amazon**.

## Install (unpacked, for now)

**Chrome:** open `chrome://extensions`, turn on **Developer mode**, click
**Load unpacked**, and pick this `extension` folder.

**Edge:** open `edge://extensions`, turn on **Developer mode**, click
**Load unpacked**, and pick this `extension` folder.

Then click the **Vine ETV Analyzer** toolbar button (pin it from the puzzle-piece
menu) to open the dashboard in a tab, and press **Refresh from Amazon**. If it says
sign-in is needed, sign in at amazon.com in the same browser and refresh again.

## Privacy

Everything stays in your browser: the extension only talks to `www.amazon.com`
(its one host permission), and stores orders, the last check and your settings in
`chrome.storage.local`. Removing the extension deletes them.

## Differences from the Python app

- Orders-page route only: no itemized `.xlsx` report, so it counts by order date
  and shows no annual total.
- The window reads back `max(30, window length)` days on each refresh.

## Development

- `calc.js`: rolling-window math (pure; runs in Node too).
- `vine.js`: fetches and parses the Vine pages.
- `dashboard.*`: the page; `background.js` opens it from the toolbar button.
- Tests: `node --test extension/tests` from the repo root.
  `extension/tests/fixtures/parity.json` holds the Python app's projection for a set
  of random orders, and a test checks the JavaScript math matches it to the cent.
  Regenerate it with `python extension/tests/make_parity_fixture.py`.
