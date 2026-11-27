# Vine ETV Analyzer

**Vine ETV Analyzer** is a lightweight, user-friendly desktop and browser utility for Amazon Vine members to visualize their **floating window of Estimated Tax Value (ETV)** and see exactly when older orders roll off so they can acquire items again.

### Why This Utility Exists
Amazon Vine enforces a rolling ETV cap—most commonly **$6,000.00 over a trailing 30-day floating window** (though both the dollar cap and window length are configurable). When your trailing 30-day ETV reaches or exceeds that cap, you are temporarily blocked from requesting additional Vine items until enough older orders age out of the rolling window to bring your total back below the threshold.

While Amazon provides downloadable **Vine Itemized Reports (`.xlsx`)** for tax purposes, the spreadsheet only lists raw historical transactions—it does not show your current floating window total or when your rolling balance will drop back under the cap. **Vine ETV Analyzer** reads your itemized report directly and shows:
- Your **exact current rolling ETV total** (e.g., trailing 30 days) and how far above or below the cap you are right now.
- A **day-by-day decay curve and schedule** projecting how your rolling ETV drops each day as older orders fall out of the window (assuming no new orders are placed).
- The **exact date and projected dollar total** when your rolling ETV crosses back below your threshold so you know when you can order again.

> **Important Note on Unshipped Orders:** Amazon's downloadable Vine Itemized Report (`.xlsx`) only records items once they have shipped—it **does not include items you have ordered that have not yet shipped**. As a result, pending unshipped orders are not part of this analysis. Keep any unshipped orders in mind when comparing your projected total against your rolling cap.

---

## Quick Install & Launch (No Git Required)

You do **not** need `git` or prior Python setup. Choose the method below for your operating system.

### Option A: One-Line Terminal Installer (Recommended)

The installer will display what it plans to set up and **ask for your confirmation** before installing Python (if missing), downloading the application to your home folder (`~/VineProjection`), creating a **Desktop shortcut**, and installing the required Python libraries (`streamlit`, `pandas`, `openpyxl`, `plotly`). It **never modifies your system `PATH`**.

- **Windows (PowerShell):**
  1. Press **`Win + X`** and click **Terminal** or **PowerShell**.
  2. Paste the following command and press **Enter**:
     ```powershell
     irm https://raw.githubusercontent.com/ggilman/vineprojection/main/vineproject.cmd | iex
     ```

- **macOS / Linux (Terminal):**
  1. Open **Terminal**.
  2. Paste the following command and press **Enter**:
     ```bash
     curl -fsSL https://raw.githubusercontent.com/ggilman/vineprojection/main/vineproject | bash
     ```

---

### Option B: Browser ZIP Download (No Terminal Required on Windows)

1. [**Click here to download the latest ZIP from GitHub**](https://github.com/ggilman/vineprojection/archive/refs/heads/main.zip).
2. Extract (`Unzip`) the downloaded folder anywhere on your computer.
3. Double-click **`vineproject.cmd`** (on Windows) or **`vineproject`** (on macOS/Linux).
   - On first run, it will ask for your confirmation before installing Python (if needed), creating a **Desktop shortcut** (`Vine ETV Analyzer`), installing required Python packages, and opening the app in your web browser.

---

## Daily Usage & Automatic Updates

Once installed:

1. **Launch from your Desktop:**
   - Double-click the **`Vine ETV Analyzer`** shortcut on your Desktop, **or**
   - Run `.\vineproject` (Windows PowerShell), `vineproject` (Windows Command Prompt), or `./vineproject` (macOS/Linux/Git Bash) inside the project folder.

2. **Automatic GitHub Update Check on Startup:**
   - Each time you start the app, `vineproject` automatically checks [`github.com/ggilman/vineprojection`](https://github.com/ggilman/vineprojection) to see if a newer version is available.
   - If an update is found, it displays the commit summary and changed files and asks:
     `Download and install this update from GitHub now? [Y/n]`
   - Updating preserves all of your `.xlsx` tax report files in the folder and does not require `git`.

3. **Command-Line Flags:**
   All flags passed to `vineproject` / `vineproject.cmd` are forwarded to `install_deps.py`:
   ```powershell
   .\vineproject --check-only   # Check updates and verify/install dependencies without launching the UI
   .\vineproject --no-update    # Skip the GitHub update check on startup
   .\vineproject --upgrade      # Upgrade all required Python packages to their latest versions
   .\vineproject --user         # Install Python packages into the user site-packages directory
   .\vineproject --yes          # Automatically answer "Yes" to update and package confirmation prompts
   .\vineproject --install-browser  # (Re)download the Chromium browser used by the live Vine feature
   ```

---

## Manual Developer Setup (Optional)

If you prefer to manage your own Python 3.9+ virtual environment and run `pip` / `streamlit` directly:

```bash
pip install streamlit pandas openpyxl plotly
streamlit run app.py
```

---

## How to Use the App

Once the application opens in your browser (`http://localhost:8501`):

1. **Select a Data Source (Left Sidebar -> `1. Data Source`):**
   - Click **Browse files** to upload your Amazon Vine `.xlsx` report, **or**
   - Select a local `.xlsx` file from the workspace folder (if any are present), **or**
   - Check **Load built-in sample report** to test the dashboard immediately with realistic sample data.

2. **Configure Projection Parameters (Left Sidebar -> `2. Projection Parameters`):**
   - **Rolling Window Length (Days):** Number of trailing days in the rolling window (default: `30` days).
   - **Count Items By:** `Ship Date` *(default)* places each item in the window by its Shipped Date (cancellations, which never ship, use their Cancelled Date); `Order Date` uses the date you ordered it.
   - **Tax Threshold ($):** Target dollar threshold to monitor (default: `$6,000.00`).
   - **Projection Horizon (Days):** Number of future days to project (defaults to matching your window length, or uncheck **Match Projection Horizon to Window Length** to enter a custom horizon).

3. **Choose Base Date (Left Sidebar -> `3. Base Date Settings`):**
   - **Today's Date** *(default)* — Projects decay starting from today.
   - **Latest Ship Date / Latest Order Date** — Projects decay starting from the most recent ship (or order) date in your report, matching **Count Items By**.
   - **Custom Date** — Pick any reference date on the calendar.

4. **Exit the Application:**
   - Click the **🚪 Quit Application** button at the bottom of the left sidebar in your browser for a one-click clean shutdown, **or**
   - Press **`Ctrl+C`** in the terminal window.

---

## Optional: Live Amazon Vine Connection

**By default the app is fully manual and offline — nothing changes until you opt in.** Click **🔌 Connect to Amazon Vine** in the left sidebar to enable:

- **Live cap status** — the app loads `amazon.com/vine/vine-items` and looks for Amazon's *"Vine limits claiming activity… Your claiming will resume soon."* message, so you know **right now** whether you're capped instead of inferring it from the spreadsheet.
- **Live order totals** — instead of downloading the itemized spreadsheet, the app reads your **Vine Orders** pages (`amazon.com/vine/orders?page=1`, `page=2`, …, newest first) and adds up the **Fair market value** of every order placed in the last 30 days (by **Order date**). It stops at the first page older than that, saves the result to `vine_reports/vine_orders.json`, and offers it as the default data source for the chart. Because the Orders page lists order dates only, these totals are counted by Order Date. (The spreadsheet-download version of the app is tagged `0.1`.)

### How Sign-In Works (passwords are never read or stored)

1. **Silent session reuse (first try):** the app copies a tiny cookie snapshot (`Cookies` database only) from your **installed Chrome or Edge** profile into `%LOCALAPPDATA%\VineProjection\vine-session\` and lets the *real* Chrome/Edge open Amazon with it. Because your browser was already signed in, this usually works with **zero interaction** — but Chrome/Edge guard their cookie file against copying while running, so for this path **close the browser completely first** (the app detects the lock and tells you).
2. **One-time sign-in fallback:** if Amazon doesn't recognize that session (or you have no Chrome/Edge), a browser window opens once — sign in there like normal. Everything afterward is automatic.
3. **Later launches:** the app refreshes cap status and re-reads your recent Vine orders on startup. Expect a normal Chrome window to flash briefly during these checks — Amazon rejects headless and off-screen browsers, so the window has to be genuinely visible. If the saved session has expired you'll just see a **🟠 Sign-in needed** status until you click **🔑 Sign in** — no prompts, no surprises. (The app keeps a private backup of the session cookies inside the same folder and self-restores it if Amazon silently expires the live copy.)

**Control:** **🔄 Refresh** re-checks on demand, **⛔ Disconnect** deletes the saved session entirely.

**Privacy & storage:** See **[PRIVACY.md](PRIVACY.md)** for the full privacy notice. In short:
- Session snapshot (equivalent to an Amazon login cookie): `%LOCALAPPDATA%\VineProjection\vine-session\` — delete it or click **Disconnect** to revoke. Nothing is uploaded anywhere.
- Saved Vine orders (`vine_orders.json`) and troubleshooting snapshots: `vine_reports/` (git-ignored).
- The **Troubleshooting → Save page snapshot** button dumps page HTML/screenshots to `vine_reports/debug/` to help tune selectors if Amazon's page layout changes.

> **Note:** the live feature requires the `playwright` package (installed by the launcher, with a one-time Chromium download prompt). Automated page access can occasionally be challenged by Amazon's bot defenses — the **🔑 Sign in** button exists for that. Use this optional feature at your own discretion with respect to Amazon's Terms of Service.

---

## Features & Dashboard Overview

1. **Dynamic Excel Report Loading (`.xlsx`):**
   - Automatically locates the itemized transaction table (`Order Date`, `Shipped Date`, `Estimated Tax Value`, `Order Type`, `Cancelled Date`) regardless of how many summary or title rows appear at the top of the worksheet.
   - Automatically strips out footer notes, FAQ text, and blank rows at the bottom of the worksheet.
   - Normalizes currency strings (`$1,234.56`, `($45.00)`, `-$45.00`) and ensures `CANCELLATION` rows carry negative tax values.

2. **Rolling Window & Decay Projection Engine:**
   - **Current Rolling Window Total:** Sums `Estimated Tax Value` within `(Base Date - Window Length, Base Date]`.
   - **Future Daily Decay Projection:** Calculates the rolling window total for each day `D` from Day 0 through your projection horizon assuming **$0.00 in new orders**, showing how older orders roll off over time.
   - **Threshold Crossing Detection:** Pinpoints the exact date and dollar amount when your rolling total drops below your target threshold, or indicates if your total is already below the threshold.

3. **GUI Layout & Visual Elements:**
   - **Top KPI Scorecards:** Current Window Total (with variance vs. threshold), Threshold Target, Threshold Crossing Date & Projected Value, and Total Annual Tax Value.
   - **Projected Decay Curve Chart:** Interactive Plotly chart displaying the projected rolling total (solid blue line), the target threshold (dashed red line), and a highlighted callout marker on the exact crossing date.
   - **Timeframe Summary Table:** Gross orders, cancellations, and net tax totals for the **Last 7 Days**, **Last 30 Days**, **Last 60 Days**, **Last 90 Days**, and **Full Year Total**.
   - **Daily Projection Schedule Table:** Day-by-day breakdown of `Date`, `Projected Total`, `Variance to Threshold`, and `Status` (`Above Threshold` / `Below Threshold`).
   - **Cleaned Itemized Transactions Inspector:** Expandable table at the bottom to inspect all parsed rows and workbook metadata.

---

## Project Files

| File | Purpose |
| :--- | :--- |
| `vineproject.cmd` | Combined Windows one-line web installer, standalone bootstrapper, and daily launcher |
| `vineproject` | Combined macOS / Linux / Git Bash one-line web installer and daily launcher |
| `install_deps.py` | Startup GitHub auto-updater, dependency installer (with confirmation), and Streamlit launcher |
| `app.py` | Streamlit GUI application, Excel parser, decay calculation engine, and Plotly charts |
| `PRIVACY.md` | Full privacy notice (data touched, stored, and transmitted) |
| `icon.png` / `icon.ico` | Custom application icon for browser tab, Windows `.lnk` shortcut, and macOS/Linux desktop launchers |
| `requirements.txt` | Python package requirements (`streamlit`, `pandas`, `openpyxl`, `plotly`, `playwright`) |
| `LICENSE` | MIT License (`Copyright (c) 2026 George Gilman`) |

---

## License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

Copyright (c) 2026 George Gilman