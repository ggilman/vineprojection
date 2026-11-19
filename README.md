# Amazon Vine Itemized Tax Report & Rolling Decay Analyzer

A lightweight, user-friendly GUI application that analyzes **Amazon Vine Itemized Tax Reports (`.xlsx`)** and visualizes when your rolling Estimated Tax Value (ETV) total will drop below a target dollar threshold assuming no new orders are placed.

---

## Prerequisites

- **Python 3.9 or newer** installed on your system ([python.org/downloads](https://www.python.org/downloads/)).

---

## Installation & Running the App

### Option 1: One-Step Launch (Recommended)

The project includes a cross-platform wrapper pair (`vineproject` and `vineproject.cmd`) that automatically verifies your Python environment, installs any missing dependencies (`streamlit`, `pandas`, `openpyxl`, `plotly`), and opens the app in your default web browser:

- **Windows Command Prompt (`cmd.exe`):**
  ```cmd
  vineproject
  ```
  *(Or double-click `vineproject.cmd` directly in File Explorer.)*

- **Windows PowerShell:**
  ```powershell
  .\vineproject
  ```

- **Linux / macOS / Git Bash:**
  ```bash
  chmod +x vineproject   # run once if needed
  ./vineproject
  ```

*(Tip: If you add this project directory to your system `PATH`, you can simply run `vineproject` from any terminal.)*

#### Launcher Flags
All command-line flags are forwarded to `install_deps.py`:
```powershell
.\vineproject --check-only   # Verify and install dependencies without launching the UI
.\vineproject --upgrade      # Upgrade all required packages to their latest versions
.\vineproject --user         # Install packages into the user site-packages directory
```

---

### Option 2: Manual Installation & Launch

If you prefer to manage your virtual environment or run `pip` and `streamlit` directly:

1. **(Optional) Create and activate a virtual environment:**
   - **Windows (PowerShell):**
     ```powershell
     python -m venv .venv
     .\.venv\Scripts\Activate.ps1
     ```
   - **Linux / macOS:**
     ```bash
     python3 -m venv .venv
     source .venv/bin/activate
     ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
   *(Or explicitly: `pip install streamlit pandas openpyxl plotly`)*

   Alternatively, run the dependency installer script directly:
   ```bash
   python install_deps.py --check-only
   ```

3. **Run the application:**
   ```bash
   streamlit run app.py
   ```
   *(Or run `python install_deps.py --run`)*

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
| `vineproject` | Extensionless Bash launcher for Linux, macOS, and Git Bash |
| `vineproject.cmd` | Windows launcher for Command Prompt, PowerShell, and File Explorer |
| `install_deps.py` | Dependency checker, auto-installer, and clean Streamlit process manager |
| `app.py` | Streamlit GUI application, Excel parser, decay calculation engine, and Plotly charts |
| `requirements.txt` | Python package requirements (`streamlit`, `pandas`, `openpyxl`, `plotly`) |