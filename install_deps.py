#!/usr/bin/env python3
# Copyright (c) 2026 George Gilman
# SPDX-License-Identifier: MIT
"""
One-Step Startup Auto-Updater, Dependency Checker, Installer & App Launcher
for Vine ETV Analyzer.

By default, running this script will:
  1. Check GitHub (https://github.com/ggilman/vineprojection) for updates and
     prompt for confirmation before downloading/applying any update.
  2. Check if required Python packages (streamlit, pandas, openpyxl, plotly, playwright) are
     installed, prompt for confirmation before installing any missing or outdated packages,
     and offer a one-time (remembered) Chromium download for the optional Amazon Vine Live
     feature. Use --install-browser to (re)attempt that download.
  3. Launch the Streamlit GUI application (app.py) in your default web browser.

Usage:
    ./vineproject       (Linux / macOS / Git Bash)
    .\vineproject       (Windows PowerShell)
    vineproject         (Windows Command Prompt, or Desktop Shortcut)

Options:
    --check-only        Only check/install dependencies without launching the app
    --no-update         Skip the startup GitHub update check
    --upgrade           Upgrade all required packages to their latest versions
    --user              Install packages to the user site-packages directory
    --yes / -y          Automatically answer "Yes" to confirmation prompts
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import os
import re
import shutil
import signal
import stat
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import zipfile
from pathlib import Path
from typing import Any

GITHUB_OWNER = "ggilman"
GITHUB_REPO = "vineprojection"
GITHUB_BRANCH = "main"
GITHUB_COMMITS_API = (
    f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/commits/{GITHUB_BRANCH}"
)
GITHUB_ZIP_URL = (
    f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}/archive/refs/heads/{GITHUB_BRANCH}.zip"
)
VERSION_STATE_FILE = ".vineproject_version.json"
CHROMIUM_STATE_FILE = ".vineproject_chromium.json"

PLAYWRIGHT_CDN_HOST = "cdn.playwright.dev"

# Application files tracked and updated during self-update
MANAGED_APP_FILES: tuple[str, ...] = (
    "app.py",
    "install_deps.py",
    "requirements.txt",
    "README.md",
    "PRIVACY.md",
    "LICENSE",
    "vineproject",
    "vineproject.cmd",
    "icon.png",
    "icon.ico",
    ".gitignore",
    ".gitattributes",
)

# Default packages if requirements.txt is not found alongside the script
DEFAULT_REQUIREMENTS: list[str] = [
    "streamlit>=1.30.0",
    "pandas>=2.0.0",
    "openpyxl>=3.1.0",
    "plotly>=5.18.0",
    "playwright>=1.40.0",
]

REQ_PATTERN = re.compile(r"^([A-Za-z0-9_.\-]+)\s*(?:([><=!~]+)\s*([0-9A-Za-z_.\-]+))?")


def prompt_confirm(prompt_text: str, auto_yes: bool = False, default: bool = True) -> bool:
    """
    Prompt the user for a Yes/No confirmation.
    Returns `default` if running non-interactively and `auto_yes` is False,
    or True immediately when `auto_yes` is True.
    """
    if auto_yes:
        print(f"{prompt_text} [Y/n]: y (auto-confirmed)")
        return True

    suffix = "[Y/n]" if default else "[y/N]"
    full_prompt = f"{prompt_text} {suffix}: "

    try:
        if not sys.stdin or not sys.stdin.isatty():
            # Try reading from controlling terminal if stdin was redirected
            tty_path = "CONIN$" if os.name == "nt" else "/dev/tty"
            with open(tty_path, encoding="utf-8") as tty:
                sys.stdout.write(full_prompt)
                sys.stdout.flush()
                answer = tty.readline()
                if not answer:
                    return default
                answer = answer.strip().lower()
        else:
            answer = input(full_prompt).strip().lower()
    except (EOFError, OSError):
        print()
        return default
    except KeyboardInterrupt:
        print("\n[INFO] Cancelled by user.")
        return False

    if not answer:
        return default
    return answer in {"y", "yes"}


def _git_blob_sha1(file_path: Path, normalize_lf: bool = True) -> str | None:
    """
    Calculate the Git blob SHA-1 hash for a local file.
    Git stores text files normalized to LF in the repository tree (except *.cmd which is CRLF).
    """
    if not file_path.is_file():
        return None
    try:
        raw = file_path.read_bytes()
    except OSError:
        return None

    if normalize_lf:
        data = raw.replace(b"\r\n", b"\n")
    else:
        data = raw
    header = f"blob {len(data)}\0".encode("ascii")
    return hashlib.sha1(header + data).hexdigest()


def _read_local_commit_sha(script_dir: Path) -> str | None:
    """Read the locally recorded commit SHA from .vineproject_version.json or .git."""
    version_file = script_dir / VERSION_STATE_FILE
    if version_file.is_file():
        try:
            data = json.loads(version_file.read_text(encoding="utf-8"))
            sha = data.get("commit_sha")
            if isinstance(sha, str) and len(sha) >= 7:
                return sha.strip()
        except (ValueError, OSError):
            pass

    git_head_ref = script_dir / ".git" / "refs" / "heads" / GITHUB_BRANCH
    if git_head_ref.is_file():
        try:
            sha = git_head_ref.read_text(encoding="utf-8").strip()
            if sha:
                return sha
        except OSError:
            pass
    return None


def _write_local_commit_sha(script_dir: Path, commit_sha: str, commit_date: str = "") -> None:
    """Record the installed commit SHA in .vineproject_version.json."""
    version_file = script_dir / VERSION_STATE_FILE
    payload = {
        "commit_sha": commit_sha,
        "commit_date": commit_date,
        "branch": GITHUB_BRANCH,
        "repository": f"{GITHUB_OWNER}/{GITHUB_REPO}",
    }
    try:
        version_file.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    except OSError:
        pass


def _fetch_json(url: str, timeout: float = 3.5) -> dict[str, Any] | None:
    """Fetch JSON from a GitHub API endpoint with a short timeout."""
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "VineProjection-Updater/1.0",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status != 200:
                return None
            return json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, ValueError, OSError):
        return None


def _compare_tree_with_local(script_dir: Path, tree_url: str) -> list[str] | None:
    """
    Compare remote Git tree blob SHAs against local managed files.
    Returns a list of changed filenames, or None if the tree API could not be fetched.
    """
    tree_data = _fetch_json(tree_url, timeout=3.5)
    if not tree_data or "tree" not in tree_data:
        return None

    remote_blobs: dict[str, str] = {}
    for entry in tree_data.get("tree", []):
        if entry.get("type") == "blob" and "path" in entry and "sha" in entry:
            remote_blobs[str(entry["path"])] = str(entry["sha"])

    changed_files: list[str] = []
    for rel_name in MANAGED_APP_FILES:
        if rel_name not in remote_blobs:
            continue
        remote_sha = remote_blobs[rel_name]
        local_path = script_dir / rel_name
        if not local_path.is_file():
            changed_files.append(rel_name)
            continue

        # Git stores LF blobs in the repo tree; verify both LF-normalized and exact raw bytes
        sha_lf = _git_blob_sha1(local_path, normalize_lf=True)
        sha_raw = _git_blob_sha1(local_path, normalize_lf=False)
        if remote_sha not in {sha_lf, sha_raw}:
            changed_files.append(rel_name)

    return changed_files


def check_and_apply_updates(script_dir: Path, auto_yes: bool = False) -> bool:
    """
    Check GitHub for a newer commit on `main`. If an update is available, prompt the user
    for confirmation, download the latest archive from GitHub, update application files
    in place (preserving user .xlsx files), and return True if an update was applied.
    """
    print("[UPDATE] Checking GitHub for application updates...")
    commit_data = _fetch_json(GITHUB_COMMITS_API, timeout=3.5)
    if not commit_data or "sha" not in commit_data:
        print("[INFO] Could not reach GitHub (offline or rate-limited). Skipping update check.")
        return False

    remote_sha = str(commit_data["sha"]).strip()
    commit_info = commit_data.get("commit", {})
    commit_msg = str(commit_info.get("message", "Update available")).splitlines()[0].strip()
    commit_date = str(commit_info.get("committer", {}).get("date", "")).strip()
    tree_url = str(commit_info.get("tree", {}).get("url", "")).strip()

    local_sha = _read_local_commit_sha(script_dir)
    if local_sha and local_sha == remote_sha:
        print(f"[INFO] Application is up to date (commit {remote_sha[:7]}).")
        return False

    # Check if local managed files already match the remote commit tree
    changed_files: list[str] | None = None
    if tree_url:
        changed_files = _compare_tree_with_local(script_dir, tree_url)
        if changed_files is not None and len(changed_files) == 0:
            _write_local_commit_sha(script_dir, remote_sha, commit_date)
            print(f"[INFO] Application is up to date (commit {remote_sha[:7]}).")
            return False

    print("\n" + "=" * 64)
    print("  UPDATE AVAILABLE ON GITHUB")
    print("=" * 64)
    print(f"  Latest Commit : {remote_sha[:7]} ({commit_date or 'latest'})")
    print(f"  Summary       : {commit_msg}")
    if changed_files:
        print(f"  Updated Files : {', '.join(changed_files)}")
    print("  Note          : Your Excel (.xlsx) reports will not be modified.")
    print("=" * 64)

    if not prompt_confirm(
        "Download and install this update from GitHub now?",
        auto_yes=auto_yes,
        default=True,
    ):
        print("[INFO] Skipping update. Launching current version...")
        return False

    return perform_self_update(script_dir, remote_sha, commit_date)


def perform_self_update(script_dir: Path, remote_sha: str, commit_date: str = "") -> bool:
    """
    Download the latest repository ZIP from GitHub and update managed files in `script_dir`.
    Neverrequires `git` to be installed.
    """
    print(f"\n[UPDATE] Downloading latest version from {GITHUB_ZIP_URL}...")
    req = urllib.request.Request(
        GITHUB_ZIP_URL,
        headers={"User-Agent": "VineProjection-Updater/1.0"},
    )
    try:
        with tempfile.TemporaryDirectory(prefix="vineproj_update_") as tmp_dir_str:
            tmp_dir = Path(tmp_dir_str)
            zip_path = tmp_dir / "main.zip"
            with urllib.request.urlopen(req, timeout=30.0) as resp:
                zip_path.write_bytes(resp.read())

            extract_dir = tmp_dir / "extracted"
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(extract_dir)

            # Locate the top-level extracted folder (e.g., vineprojection-main)
            subdirs = [p for p in extract_dir.iterdir() if p.is_dir()]
            source_root = subdirs[0] if subdirs else extract_dir

            updated_count = 0
            for item in source_root.iterdir():
                if item.name == ".git" or item.name == VERSION_STATE_FILE:
                    continue
                dest = script_dir / item.name
                if item.is_file():
                    raw_bytes = item.read_bytes()
                    # Ensure correct line endings for shell/batch scripts
                    if item.name in {"vineproject", "install_deps.py"}:
                        raw_bytes = raw_bytes.replace(b"\r\n", b"\n")
                    elif item.name == "vineproject.cmd":
                        raw_bytes = raw_bytes.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
                    dest.write_bytes(raw_bytes)
                    if item.name in {"vineproject", "install_deps.py"} and os.name != "nt":
                        try:
                            dest.chmod(dest.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
                        except OSError:
                            pass
                    updated_count += 1
                elif item.is_dir():
                    if dest.exists():
                        shutil.rmtree(dest)
                    shutil.copytree(item, dest)
                    updated_count += 1

            # Also fast-forward local .git metadata if git and a clean working tree exist
            if (script_dir / ".git").is_dir() and shutil.which("git"):
                subprocess.run(
                    ["git", "fetch", "origin", GITHUB_BRANCH],
                    cwd=str(script_dir),
                    capture_output=True,
                    check=False,
                )
                subprocess.run(
                    ["git", "reset", "--hard", f"origin/{GITHUB_BRANCH}"],
                    cwd=str(script_dir),
                    capture_output=True,
                    check=False,
                )

            _write_local_commit_sha(script_dir, remote_sha, commit_date)
            print(
                f"[SUCCESS] Updated {updated_count} application files to commit {remote_sha[:7]}!"
            )
            return True
    except (urllib.error.URLError, TimeoutError, zipfile.BadZipFile, OSError) as exc:
        print(f"[WARN] Update failed ({exc}). Continuing with current installed version...")
        return False


def parse_version(version_str: str) -> tuple[int, ...]:
    """Convert a version string like '1.30.2rc1' into a comparable integer tuple."""
    numeric_parts: list[int] = []
    for part in version_str.split("."):
        match = re.match(r"^(\d+)", part)
        if match:
            numeric_parts.append(int(match.group(1)))
        else:
            numeric_parts.append(0)
    return tuple(numeric_parts)


def version_satisfies(installed: str, operator: str | None, required: str | None) -> bool:
    """Evaluate whether an installed version string satisfies a requirement specifier."""
    if not operator or not required:
        return True

    inst_tuple = parse_version(installed)
    req_tuple = parse_version(required)

    max_len = max(len(inst_tuple), len(req_tuple))
    inst_padded = inst_tuple + (0,) * (max_len - len(inst_tuple))
    req_padded = req_tuple + (0,) * (max_len - len(req_tuple))

    if operator == ">=":
        return inst_padded >= req_padded
    if operator == "==":
        return inst_padded == req_padded
    if operator == "<=":
        return inst_padded <= req_padded
    if operator == ">":
        return inst_padded > req_padded
    if operator == "<":
        return inst_padded < req_padded
    if operator == "!=":
        return inst_padded != req_padded
    return True


def load_requirements(req_file: Path) -> list[str]:
    """Read requirement specifiers from requirements.txt or fall back to defaults."""
    if not req_file.is_file():
        return list(DEFAULT_REQUIREMENTS)

    requirements: list[str] = []
    for raw_line in req_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.split("#", 1)[0].strip()
        if line and not line.startswith("-"):
            requirements.append(line)
    return requirements or list(DEFAULT_REQUIREMENTS)


def ensure_pip_available() -> bool:
    """Verify pip is installed for the active Python interpreter; bootstrap via ensurepip if missing."""
    check = subprocess.run(
        [sys.executable, "-m", "pip", "--version"],
        capture_output=True,
        text=True,
        check=False,
    )
    if check.returncode == 0:
        print(f"[INFO] Python : {sys.executable} (v{sys.version.split()[0]})")
        print(f"[INFO] Pip    : {check.stdout.strip()}")
        return True

    print("[WARN] pip was not found for this Python interpreter. Attempting bootstrap via ensurepip...")
    boot = subprocess.run(
        [sys.executable, "-m", "ensurepip", "--upgrade"],
        check=False,
    )
    return boot.returncode == 0


def check_requirements(requirements: list[str]) -> tuple[list[str], list[str]]:
    """
    Check each requirement against installed metadata.
    Returns (satisfied_packages, missing_or_outdated_specs).
    """
    satisfied: list[str] = []
    to_install: list[str] = []

    print("\nChecking required Python packages:")
    print("-" * 60)
    for spec in requirements:
        match = REQ_PATTERN.match(spec)
        if not match:
            to_install.append(spec)
            continue

        pkg_name, operator, req_version = match.groups()
        try:
            installed_ver = importlib.metadata.version(pkg_name)
        except importlib.metadata.PackageNotFoundError:
            print(f"  [MISSING]  {spec:<25} (not installed)")
            to_install.append(spec)
            continue

        if version_satisfies(installed_ver, operator, req_version):
            print(f"  [OK]       {spec:<25} (installed: {installed_ver})")
            satisfied.append(pkg_name)
        else:
            print(
                f"  [OUTDATED] {spec:<25} (installed: {installed_ver}, requires {operator}{req_version})"
            )
            to_install.append(spec)
    print("-" * 60)

    return satisfied, to_install


def install_packages(specs: list[str], upgrade: bool = False, user: bool = False) -> bool:
    """Run pip install for the specified package requirement strings."""
    cmd = [sys.executable, "-m", "pip", "install"]
    if upgrade:
        cmd.append("--upgrade")
    if user:
        cmd.append("--user")
    cmd.extend(specs)

    print(f"\n[INSTALL] Installing required packages: {' '.join(specs)}...")
    result = subprocess.run(cmd, check=False)
    if result.returncode == 0:
        return True

    # On modern Linux distributions outside a venv (PEP 668), retry with --user
    if not user and sys.platform.startswith("linux"):
        print("\n[WARN] Standard install failed on Linux. Retrying with --user...")
        retry_cmd = [sys.executable, "-m", "pip", "install", "--user"]
        if upgrade:
            retry_cmd.append("--upgrade")
        retry_cmd.extend(specs)
        retry = subprocess.run(retry_cmd, check=False)
        if retry.returncode == 0:
            return True

    return False


def _browsers_root() -> Path:
    home = Path.home()
    if os.name == "nt":
        base = Path(os.getenv("LOCALAPPDATA", str(home / "AppData" / "Local")))
    elif sys.platform == "darwin":
        base = home / "Library" / "Caches"
    else:
        base = Path(os.getenv("XDG_CACHE_HOME", str(home / ".cache")))
    return base / "ms-playwright"


def _chromium_present() -> bool:
    root = _browsers_root()
    if not root.is_dir():
        return False
    try:
        return any(child.name.startswith("chromium") for child in root.iterdir())
    except OSError:
        return False


def _read_json_state(path: Path) -> dict[str, Any]:
    if path.is_file():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, ValueError):
            pass
    return {}


def _write_json_state(path: Path, payload: dict[str, Any]) -> None:
    try:
        path.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    except OSError:
        pass


def _download_chromium(extra_env: dict[str, str] | None = None) -> tuple[bool, str]:
    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)
    try:
        result = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            check=False,
            timeout=1800,
            env=env,
        )
    except subprocess.TimeoutExpired:
        return False, "Download timed out."
    output = ((result.stdout or "") + "\n" + (result.stderr or "")).strip()
    return result.returncode == 0, output


def _is_certificate_error(output: str) -> bool:
    lowered = output.lower()
    return "self-signed certificate" in lowered or "unable to get local issuer" in lowered


def _parse_proxy_hostport() -> tuple[str, int] | None:
    proxy = os.getenv("HTTPS_PROXY") or os.getenv("https_proxy") or ""
    raw = proxy.split("://", 1)[-1].split("@", 1)[-1].split("/", 1)[0].strip()
    if not raw:
        return None
    host, colon, port = raw.rpartition(":")
    if colon and port.isdigit():
        return host, int(port)
    return host or raw, 80


def _export_proxy_ca_bundle() -> Path | None:
    """
    Some corporate/agency TLS proxies intercept HTTPS using an internal CA trusted by
    the Windows certificate store but unknown to Node.js (which powers the Playwright
    browser downloader). Export the chain Windows already trusts for the Playwright
    CDN into a PEM bundle so the Node download can verify it.
    """
    if os.name != "nt":
        return None
    proxy = _parse_proxy_hostport()
    proxy_host, proxy_port = proxy if proxy else ("", 80)
    bundle_path = Path(tempfile.gettempdir()) / "vineprojection_corp_ca.pem"
    script_path = Path(tempfile.gettempdir()) / "vineprojection_corp_ca.ps1"
    if bundle_path.is_file():
        try:
            bundle_path.unlink()
        except OSError:
            pass
    try:
        script_path.write_text(CORP_CA_EXPORT_PS, encoding="utf-8")
        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                "-TargetHost",
                PLAYWRIGHT_CDN_HOST,
                "-ProxyHost",
                proxy_host,
                "-ProxyPort",
                str(proxy_port),
                "-OutPath",
                str(bundle_path),
            ],
            capture_output=True,
            text=True,
            check=False,
            timeout=90,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    if result.returncode != 0 or not bundle_path.is_file():
        return None
    return bundle_path


CORP_CA_EXPORT_PS = r"""
param([string]$TargetHost, [string]$ProxyHost, [int]$ProxyPort, [string]$OutPath)
try {
  $tcp = New-Object System.Net.Sockets.TcpClient
  if ($ProxyHost) {
    $tcp.Connect($ProxyHost, $ProxyPort)
    $net = $tcp.GetStream()
    $req = "CONNECT ${TargetHost}:443 HTTP/1.1`r`n"
    $req += "Host: ${TargetHost}:443`r`n`r`n"
    $bytes = [Text.Encoding]::ASCII.GetBytes($req)
    $net.Write($bytes, 0, $bytes.Length)
    $net.Flush()
    $reader = New-Object System.IO.StreamReader($net)
    $status = $reader.ReadLine()
    while ($true) {
      $line = $reader.ReadLine()
      if ($null -eq $line -or $line -eq '') { break }
    }
    if ($status -notmatch '200') { Write-Output "ERR tunnel: $status"; exit 1 }
  } else {
    $tcp.Connect($TargetHost, 443)
    $net = $tcp.GetStream()
  }
  $callback = [System.Net.Security.RemoteCertificateValidationCallback] {
    param($s, $cert, $chain, $errors) $true
  }
  $ssl = New-Object System.Net.Security.SslStream($net, $false, $callback)
  $ssl.AuthenticateAsClient($TargetHost)
  $leaf = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate)
  $ssl.Close()
  $tcp.Close()
  $chainObj = New-Object System.Security.Cryptography.X509Certificates.X509Chain
  $chainObj.ChainPolicy.RevocationMode = 'NoCheck'
  [void]$chainObj.Build($leaf)
  $pem = @()
  foreach ($el in $chainObj.ChainElements) {
    $b64 = [Convert]::ToBase64String($el.Certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
    $pem += ('-----BEGIN CERTIFICATE-----' + [Environment]::NewLine + $b64 + [Environment]::NewLine + '-----END CERTIFICATE-----')
  }
  Set-Content -Path $OutPath -Value ($pem -join [Environment]::NewLine) -Encoding ascii
  Write-Output "BUNDLE chainlen:$($chainObj.ChainElements.Count)"
} catch {
  Write-Output "ERR:$_"
  exit 1
}
"""


def ensure_playwright_browsers(
    script_dir: Path, auto_yes: bool = False, force: bool = False
) -> bool:
    """
    If the optional Amazon Vine Live feature is available (playwright package installed),
    make sure the bundled Chromium browser has been downloaded exactly once. The outcome
    (installed / declined / failed) is remembered in CHROMIUM_STATE_FILE so the user is
    never nagged again; --install-browser forces a retry.
    """
    try:
        import importlib.util

        if importlib.util.find_spec("playwright") is None:
            return True
    except (ImportError, ValueError):
        return True

    if _chromium_present():
        return True

    state_path = script_dir / CHROMIUM_STATE_FILE
    status = str(_read_json_state(state_path).get("status", ""))
    if not force and status == "declined":
        print(
            "[INFO] Bundled Chromium download was skipped earlier. "
            "Retry with: vineproject --install-browser"
        )
        return True
    if not force and status == "failed":
        print(
            "[INFO] Bundled Chromium download previously failed; not retrying automatically. "
            "The live Vine connection still works via installed Chrome/Edge. "
            "Retry with: vineproject --install-browser"
        )
        return True

    print(
        "\n[NOTICE] The optional 'Amazon Vine Live' feature works best with the bundled "
        "Chromium browser (~150 MB, downloaded once). If skipped, your installed "
        "Chrome/Edge will be used instead."
    )
    if not force and not prompt_confirm(
        "Download the Chromium browser for the live Vine connection now?",
        auto_yes=auto_yes,
        default=True,
    ):
        _write_json_state(state_path, {"status": "declined"})
        print(
            "[INFO] Skipped. The live Vine connection will use an installed Chrome/Edge "
            "when possible. You can enable this later with: vineproject --install-browser"
        )
        return True

    print("[INSTALL] Downloading Playwright Chromium browser (this may take a few minutes)...")
    ok, output = _download_chromium()
    if not ok and _is_certificate_error(output):
        print(
            "[INFO] TLS-intercepting proxy detected; retrying with the CA chain your "
            "Windows certificate store already trusts..."
        )
        bundle = _export_proxy_ca_bundle()
        if bundle is not None:
            ok, output = _download_chromium(extra_env={"NODE_EXTRA_CA_CERTS": str(bundle)})

    if ok:
        _write_json_state(state_path, {"status": "installed"})
        print("[SUCCESS] Chromium browser is now available for the live Vine connection!")
        return True

    detail = next((ln.strip() for ln in reversed(output.splitlines()) if ln.strip()), "unknown error")
    _write_json_state(state_path, {"status": "failed", "detail": detail[:200]})
    print(f"[WARN] Chromium download failed ({detail[:200]}).")
    print(
        "[INFO] The live Vine connection can still use your installed Chrome/Edge. "
        "Retry the download later with: vineproject --install-browser"
    )
    return False


def main(default_run: bool = True) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Check for app updates, verify/install Python dependencies, "
            "and launch the Vine ETV Analyzer."
        )
    )
    parser.add_argument(
        "--upgrade",
        action="store_true",
        help="Upgrade all required packages to their latest available versions.",
    )
    parser.add_argument(
        "--user",
        action="store_true",
        help="Install packages into the user site-packages directory (--user).",
    )
    parser.add_argument(
        "--run",
        action="store_true",
        default=default_run,
        help="Launch the Streamlit application (app.py) after verifying dependencies (enabled by default).",
    )
    parser.add_argument(
        "--check-only",
        "--no-run",
        dest="check_only",
        action="store_true",
        help="Only check and install dependencies without launching the Streamlit app.",
    )
    parser.add_argument(
        "--no-update",
        "--skip-update-check",
        dest="no_update",
        action="store_true",
        help="Skip checking GitHub for application updates on startup.",
    )
    parser.add_argument(
        "-y",
        "--yes",
        dest="auto_yes",
        action="store_true",
        help="Automatically confirm update and package installation prompts.",
    )
    parser.add_argument(
        "--install-browser",
        dest="install_browser",
        action="store_true",
        help="(Re)attempt downloading the bundled Chromium browser for the live Vine feature.",
    )
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent

    # Step 1: Check GitHub for application updates on startup (unless --no-update)
    if not args.no_update:
        updated = check_and_apply_updates(script_dir, auto_yes=args.auto_yes)
        if updated:
            print("[UPDATE] Restarting launcher with updated version...\n")
            restart_cmd = [
                sys.executable,
                str(script_dir / "install_deps.py"),
                "--skip-update-check",
            ]
            for arg in sys.argv[1:]:
                if arg not in {"--no-update", "--skip-update-check"}:
                    restart_cmd.append(arg)
            return subprocess.call(restart_cmd)

    req_file = script_dir / "requirements.txt"

    if not ensure_pip_available():
        print(
            "[ERROR] pip is not available and could not be bootstrapped. "
            "Please install python3-pip or reinstall Python with pip enabled.",
            file=sys.stderr,
        )
        return 1

    requirements = load_requirements(req_file)
    _, to_install = check_requirements(requirements)

    target_specs = requirements if args.upgrade else to_install

    if target_specs:
        action_word = "upgrade" if args.upgrade else "install"
        print(
            f"\n[NOTICE] The following Python packages need to be {action_word}ed: "
            f"{', '.join(target_specs)}"
        )
        if not prompt_confirm(
            f"Do you want to {action_word} these Python packages now?",
            auto_yes=args.auto_yes,
            default=True,
        ):
            print(
                "[ERROR] Required packages are missing. Cannot launch application without dependencies.",
                file=sys.stderr,
            )
            return 1

        success = install_packages(target_specs, upgrade=args.upgrade, user=args.user)
        if not success:
            print("\n[ERROR] Package installation failed.", file=sys.stderr)
            return 1

        importlib.invalidate_caches()
        _, still_missing = check_requirements(requirements)
        if still_missing:
            print(
                f"\n[ERROR] Some packages could not be verified after install: {still_missing}",
                file=sys.stderr,
            )
            return 1
        print("\n[SUCCESS] All required packages are now installed and verified!")
    else:
        print("\n[SUCCESS] All required packages are already installed and up to date!")

    ensure_playwright_browsers(
        script_dir, auto_yes=args.auto_yes, force=args.install_browser
    )

    should_run = args.run and not args.check_only
    if should_run:
        app_path = script_dir / "app.py"
        print(f"\n[LAUNCH] Starting Vine ETV Analyzer ({app_path.name})...")
        print(
            "[INFO] Your web browser will open automatically. "
            "Use the 'Quit Application' button in the sidebar (or press Ctrl+C here) to exit.\n"
        )
        return launch_streamlit(app_path)

    return 0


def launch_streamlit(app_path: Path) -> int:
    """Launch Streamlit and handle Ctrl+C / termination cleanly without stack traces."""
    cmd = [sys.executable, "-m", "streamlit", "run", str(app_path)]
    proc = subprocess.Popen(cmd)

    original_sigint = signal.getsignal(signal.SIGINT)
    try:
        while proc.poll() is None:
            try:
                proc.wait(timeout=0.5)
            except subprocess.TimeoutExpired:
                continue
    except KeyboardInterrupt:
        print("\n[INFO] Stopping Vine ETV Analyzer...")
        try:
            proc.wait(timeout=3.0)
        except subprocess.TimeoutExpired:
            proc.terminate()
            try:
                proc.wait(timeout=2.0)
            except subprocess.TimeoutExpired:
                proc.kill()
        print("[INFO] Application stopped cleanly.")
        return 0
    finally:
        signal.signal(signal.SIGINT, original_sigint)

    print("[INFO] Application stopped cleanly.")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n[INFO] Interrupted by user. Exiting cleanly.")
        sys.exit(0)