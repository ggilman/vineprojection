#!/usr/bin/env python3
# Copyright (c) 2026 George Gilman
# SPDX-License-Identifier: MIT
"""
One-Step Startup Auto-Updater, Dependency Checker, Installer & App Launcher
for Vine ETV Analyzer.

By default, running this script will:
  1. Check GitHub (https://github.com/ggilman/vineprojection) for updates and
     prompt for confirmation before downloading/applying any update.
  2. Check if required Python packages (streamlit, pandas, openpyxl, plotly) are installed,
     and prompt for confirmation before installing any missing or outdated packages.
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

# Application files tracked and updated during self-update
MANAGED_APP_FILES: tuple[str, ...] = (
    "app.py",
    "install_deps.py",
    "requirements.txt",
    "README.md",
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