# Privacy Notice — Vine ETV Analyzer

**TL;DR**

- The analyzer is a **local, offline tool by default**. It only reads an `.xlsx` file you pick, computes on it, and shows charts in your own browser tab. No personal data leaves your machine.
- The optional **🔌 Amazon Vine Live** feature stores a **copy of your Amazon browser session** (login cookies) on disk. This is equivalent to being signed in to your Amazon account and is the single most sensitive thing the app touches.
- The app **never sees, reads, stores, or transmits your Amazon password**, and never sends any of your data to the developer or any third party.

---

## 1. Default (Manual) Operation

With no live connection configured:

| Action | Where data goes |
| :--- | :--- |
| Opening your Vine itemized `.xlsx` | Parsed in memory on your machine; never written elsewhere, never uploaded |
| Charts, tables, projections | Rendered by the local Streamlit server into your own browser tab |
| App UI server | Binds to `localhost` (default port 8501) on your machine only — see §1.2 |

### 1.1 The Excel report itself

Amazon Vine itemized reports contain purchase history: product names, order numbers, ASINs, dates, Estimated Tax Values, and sometimes your name/account holder details. Everything shown is derived from that file, and it stays on your machine. **Treat the `.xlsx` itself as financial data** — the app neither encrypts nor protects it; normal file-sharing caution applies.

### 1.2 Local web server caveat

Streamlit serves the UI over an unauthenticated HTTP server on `localhost`. Any program running under your Windows user account — and any *other user* on a shared machine with standard loopback access — could open the UI and see the loaded report while it is displayed. Don't use the app on shared/family machines with the report open, or lock your workstation.

### 1.3 Streamlit's own telemetry

Streamlit performs an anonymous "is a newer version available?" version check on start. It does not include your files or reports. Disable it with `STREAMLIT_GATHER_USAGE_STATS=false` or by setting `browser.gatherUsageStats = false` in Streamlit's config.

---

## 2. Optional Live Amazon Vine Connection

This feature is **strictly opt-in**. Nothing happens until you press **🔌 Connect to Amazon Vine**.

### 2.1 What is copied, and where

When you connect, the app copies a **minimal browser-session snapshot** out of your installed Chrome or Edge profile:

| File | Why it's needed |
| :--- | :--- |
| `Local State` (browser root) | Contains the key material Chrome/Edge uses to decrypt its cookies |
| `<profile>/Network/Cookies` (+ `-wal`, `-shm`) | Your browser cookies — including the Amazon sign-in session |

These are copied into:

```text
%LOCALAPPDATA%\VineProjection\vine-session\
```

(Windows; an equivalent per-user folder on macOS/Linux.) The folder is **outside the project directory**, never uploaded, and removed entirely by the sidebar **⛔ Disconnect** button or by deleting the folder yourself.

### 2.2 Why this snapshot is sensitive

A valid Amazon session cookie is **equivalent to a signed-in session** for that browser:

- Any program running under your Windows account can hand that folder to a real Chrome/Edge instance and be recognized by Amazon as *you*.
- Anyone with access to your powered-on, unlocked machine (or malware running as your user) can use it **until the session expires or you delete it**.
- Windows mitigations help but are not walls: the copy is decryptable only by genuine Chrome/Edge binaries under the same Windows account (DPAPI + app-bound encryption), but that is exactly what this app — and anything else running as you — can do.

**Mitigations the app applies:** the copy contains cookies only (no bookmarks, history, saved passwords, extensions), it is stored under your user profile directory (not readable by other standard Windows users), and **Disconnect** deletes it in one click. You can also revoke it server-side at Amazon: *Accounts → Login & security → "Forget devices" / sign out of all sessions*, or change your password.

### 2.3 Your password

- The app never reads saved passwords from your browser, never shows a password form, and never intercepts keyboard input.
- If silent session reuse fails, a **standard Chrome/Edge window** opens and you sign in directly on `https://www.amazon.com/...` — exactly as you normally would. The app only *observes the resulting page state* to know login succeeded.
- The login session is then persisted inside the `vine-session` folder described above (same sensitivity, same deletion). A private cookie backup is kept in the same folder so an Amazon-side session wipe can self-heal without a new sign-in; **Disconnect** deletes that too.
- After you connect, automatic status checks reuse that same `vine-session` profile. Because Amazon rejects headless or off-screen browsers, these checks run as a *genuine but short-lived* visible Chrome window.

### 2.4 What is downloaded to your machine

- **Recent Vine orders** → `vine_reports/vine_orders.json` in the app folder: order dates, Fair market values, ASINs and product names read from your Vine Orders pages for the last 30 days (same financial sensitivity as §1.1; the folder is git-ignored so it is not accidentally published, but it is not encrypted).
- **Itemized reports** (only with the *Itemized report file* data route) → `vine_reports/*.xlsx` in the app folder (same financial sensitivity as §1.1; git-ignored, not encrypted).
- **Settings** → `vine_reports/settings.json`: which data route you picked, nothing else.
- **Debug snapshots** (only if you press "Save page snapshot") → `vine_reports/debug/` — page **HTML and screenshots** that can include Amazon-internal identifiers such as your obfuscated customer ID and CSRF tokens. Treat these as sensitive: redact before sharing them in any bug report, and delete when done.

### 2.5 Network traffic while connected

All live traffic goes **only between your own machine and `amazon.com`**, using your own browser engine and your own cookies — to Amazon, the traffic is indistinguishable from you browsing the Vine pages yourself. The app adds no analytics, no "phone home", and no developer endpoints. Note the flip side: Amazon can detect, challenge (bot interstitials), or act against automated access under its Terms of Service — this optional feature is offered as-is for personal convenience; use at your own discretion.

### 2.6 Data lifecycle / cleanup checklist

| What | Location | Delete with |
| :--- | :--- | :--- |
| Session cookie snapshot (+ browser profile cache from the live sessions) | `%LOCALAPPDATA%\VineProjection\vine-session\` | **⛔ Disconnect** or delete the folder |
| Saved Vine orders | `vine_reportsine_orders.json` | delete file |
| Downloaded itemized reports | `vine_reports\*.xlsx` | delete files |
| Debug HTML/screenshots | `vine_reports\debug\` | delete files |
| Browser profile copy lock artifacts | same as session folder | same |
| Amazon-side sessions | Amazon's servers | Amazon: sign out everywhere / change password |

---

## 3. Updater & Installer (also network-active)

- On each launch, `vineproject` / `install_deps.py` queries **`api.github.com` and `github.com`** for the latest release commit and downloads the app ZIP. This exposes your IP to GitHub, plus a user-agent string identifying the updater — no report contents or account data are involved. Skip it with `--no-update`.
- Dependency installs contact **PyPI**; the one-time Chromium download contacts **Playwright's CDN** (or fails gracefully; the live feature can fall back to your installed Chrome/Edge).

---

## 4. What the application never does

- Never captures, requests, or stores Amazon credentials (username/password/2FA codes).
- Never uploads your Excel data, reports, charts, or the session snapshot anywhere — not to GitHub, not to the author, not to analytics.
- Never runs outside `localhost`, `amazon.com`, `github.com`/`api.github.com`, PyPI, and the Playwright CDN.
- Never installs system services, modifies your real browser profile (the profile is read-only-copied, never written), or changes system configuration beyond installing Python packages it asks you to confirm.

If you find behavior that contradicts this notice, that's a bug — please report it.

---

*This notice describes version `automatic`-branch behavior (Live Amazon Vine Connection added). Copyright (c) 2026 George Gilman · [MIT License](LICENSE).*
