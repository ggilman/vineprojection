#!/usr/bin/env python3
"""
One-Step Dependency Checker, Installer & App Launcher for Amazon Vine Tax Decay Analyzer.

By default, running this script will:
  1. Check if required Python packages (streamlit, pandas, openpyxl, plotly) are installed.
  2. Automatically install any missing or outdated packages via pip.
  3. Launch the Streamlit GUI application (app.py) in your default web browser.

Usage:
    ./vineproject       (Linux / macOS / Git Bash)
    .\vineproject       (Windows PowerShell)
    vineproject         (Windows Command Prompt, or when directory is on PATH)

Options:
    --check-only   Only check/install dependencies without launching the app
    --upgrade      Upgrade all required packages to their latest versions
    --user         Install packages to the user site-packages directory
"""

from __future__ import annotations

import argparse
import importlib.metadata
import re
import signal
import subprocess
import sys
from pathlib import Path

# Default packages if requirements.txt is not found alongside the script
DEFAULT_REQUIREMENTS: list[str] = [
    "streamlit>=1.30.0",
    "pandas>=2.0.0",
    "openpyxl>=3.1.0",
    "plotly>=5.18.0",
]

REQ_PATTERN = re.compile(r"^([A-Za-z0-9_.\-]+)\s*(?:([><=!~]+)\s*([0-9A-Za-z_.\-]+))?")


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

    # On modern Linux distributions outside a venv (PEP 668), retry with --user or --break-system-packages
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
        description="Check dependencies, install any missing packages, and launch the Amazon Vine Tax Decay Analyzer."
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
    args = parser.parse_args()

    script_dir = Path(__file__).resolve().parent
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
        print(f"\n[LAUNCH] Starting Amazon Vine Tax Decay Analyzer ({app_path.name})...")
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

    # Allow the child process to handle SIGINT first while parent waits cleanly
    original_sigint = signal.getsignal(signal.SIGINT)
    try:
        while proc.poll() is None:
            try:
                proc.wait(timeout=0.5)
            except subprocess.TimeoutExpired:
                continue
    except KeyboardInterrupt:
        print("\n[INFO] Stopping Amazon Vine Tax Decay Analyzer...")
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