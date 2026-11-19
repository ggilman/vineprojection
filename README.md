# Amazon Vine Itemized Tax Report & Rolling Decay Analyzer

A lightweight, user-friendly GUI application that analyzes **Amazon Vine Itemized Tax Reports (`.xlsx`)** and visualizes when your rolling Estimated Tax Value (ETV) total will drop below a target dollar threshold assuming no new orders are placed.

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
   - On first run, it will ask for your confirmation before installing Python (if needed), creating a **Desktop shortcut** (`Amazon Vine Tax Analyzer`), installing required Python packages, and opening the app in your web browser.

---

## Daily Usage & Automatic Updates

Once installed:

1. **Launch from your Desktop:**
   - Double-click the **`Amazon Vine Tax Analyzer`** shortcut on your Desktop, **or**
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
   - **Tax Threshold ($):** Target dollar threshold to monitor (default: `$6,000.00`).
   - **Projection Horizon (Days):** Number of future days to project (defaults to matching your window length, or uncheck **Match Projection Horizon to Window Length** to enter a custom horizon).

3. **Choose Base Date (Left Sidebar -> `3. Base Date Settings`):**
   - **Today's Date** *(default)* — Projects decay starting from today.
   - **Latest Order Date** — Projects decay starting from the most recent order date in your report.
   - **Custom Date** — Pick any reference date on the calendar.

4. **Exit the Application:**
   - Click the **🚪 Quit Application** button at the bottom of the left sidebar in your browser for a one-click clean shutdown, **or**
   - Press **`Ctrl+C`** in the terminal window.

---

## Features & Dashboard Overview

1. **Dynamic Excel Report Loading (`.xlsx`):**
   - Automatically locates the itemized transaction table (`Order Date`, `Estimated Tax Value`, `Order Type`, `Cancelled Date`) regardless of how many summary or title rows appear at the top of the worksheet.
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
| `requirements.txt` | Python package requirements (`streamlit`, `pandas`, `openpyxl`, `plotly`) |
| `LICENSE` | MIT License (`Copyright (c) 2026 George Gilman`) |

---

## License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.

Copyright (c) 2026 George Gilman