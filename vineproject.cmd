<# : batch portion
@echo off
REM Copyright (c) 2026 George Gilman
REM SPDX-License-Identifier: MIT
if exist "%~dp0install_deps.py" (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1
    if not errorlevel 1 (
        python "%~dp0install_deps.py" %*
        exit /b %ERRORLEVEL%
    )
    py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)" >nul 2>&1
    if not errorlevel 1 (
        py -3 "%~dp0install_deps.py" %*
        exit /b %ERRORLEVEL%
    )
)
set "VINE_SELF=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression ([System.IO.File]::ReadAllText($env:VINE_SELF))"
exit /b %ERRORLEVEL%
#>
# Copyright (c) 2026 George Gilman
# SPDX-License-Identifier: MIT
#
# One-Line PowerShell Installer & Bootstrapper:
#   irm https://raw.githubusercontent.com/ggilman/vineprojection/main/vineproject.cmd | iex

$ErrorActionPreference = 'Stop'
$RepoOwner = 'ggilman'
$RepoName = 'vineprojection'
$Branch = 'main'
$ZipUrl = "https://github.com/$RepoOwner/$RepoName/archive/refs/heads/$Branch.zip"

function Find-Python39Plus {
    foreach ($candidate in @('python', 'py')) {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($cmd) {
            try {
                if ($candidate -eq 'py') {
                    $ver = & py -3 -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}' if sys.version_info >= (3, 9) else '')" 2>$null
                    if ($ver) { return @{ Exe = 'py'; Args = @('-3'); Version = $ver.Trim() } }
                } else {
                    $ver = & python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}' if sys.version_info >= (3, 9) else '')" 2>$null
                    if ($ver) { return @{ Exe = $cmd.Source; Args = @(); Version = $ver.Trim() } }
                }
            } catch {}
        }
    }
    $localPyDirs = Get-ChildItem -Path "$env:LOCALAPPDATA\Programs\Python\Python3*" -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending
    foreach ($dir in $localPyDirs) {
        $pyExe = Join-Path $dir.FullName 'python.exe'
        if (Test-Path $pyExe) {
            try {
                $ver = & $pyExe -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}' if sys.version_info >= (3, 9) else '')" 2>$null
                if ($ver) { return @{ Exe = $pyExe; Args = @(); Version = $ver.Trim() } }
            } catch {}
        }
    }
    return $null
}

function Ask-Confirmation([string]$PromptText) {
    $ans = Read-Host "$PromptText [Y/n]"
    if ([string]::IsNullOrWhiteSpace($ans)) { return $true }
    return ($ans.Trim().ToLower() -in @('y', 'yes'))
}

# Determine if running from an existing folder containing vineproject.cmd or via web (irm | iex)
$LocalDir = $null
if ($env:VINE_SELF -and (Test-Path $env:VINE_SELF)) {
    $LocalDir = Split-Path -Parent $env:VINE_SELF
}

$PyInfo = Find-Python39Plus
if ($LocalDir -and (Test-Path (Join-Path $LocalDir 'install_deps.py')) -and $PyInfo) {
    & $PyInfo.Exe @($PyInfo.Args) (Join-Path $LocalDir 'install_deps.py') @args
    exit $LASTEXITCODE
}

$InstallDir = if ($LocalDir -and (Test-Path (Join-Path $LocalDir 'app.py'))) { $LocalDir } else { Join-Path $HOME 'VineProjection' }
$DesktopDir = [Environment]::GetFolderPath('Desktop')
$ShortcutPath = Join-Path $DesktopDir 'Vine ETV Analyzer.lnk'

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Vine ETV Analyzer - Setup Wizard" -ForegroundColor Cyan
Write-Host "  Copyright (c) 2026 George Gilman (MIT License)" -ForegroundColor DarkCyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Target Folder    : $InstallDir"
if ($PyInfo) {
    Write-Host "  Python Status    : Found Python $($PyInfo.Version) ($($PyInfo.Exe))" -ForegroundColor Green
} else {
    Write-Host "  Python Status    : Not found (Python 3.12 will be installed)" -ForegroundColor Yellow
}
Write-Host "  Python Packages  : streamlit, pandas, openpyxl, plotly"
Write-Host "  Desktop Shortcut : $ShortcutPath"
Write-Host "  System PATH      : Not modified"
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

if (-not (Ask-Confirmation "Proceed with installation?")) {
    Write-Host "[INFO] Installation cancelled by user." -ForegroundColor Yellow
    exit 0
}

# 1. Install Python 3.12 if Python 3.9+ is not present
if (-not $PyInfo) {
    Write-Host ""
    if (-not (Ask-Confirmation "Python 3.9+ is required. Download and install Python 3.12 now?")) {
        Write-Host "[ERROR] Python 3.9+ is required to run the application." -ForegroundColor Red
        exit 1
    }

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Host "[INSTALL] Installing Python 3.12 via Windows Package Manager (winget)..." -ForegroundColor Cyan
        & winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements
    } else {
        $pyInstallerUrl = "https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe"
        $tmpInstaller = Join-Path ([System.IO.Path]::GetTempPath()) "python-3.12.10-amd64.exe"
        Write-Host "[INSTALL] Downloading official Python 3.12 installer from python.org..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $pyInstallerUrl -OutFile $tmpInstaller -UseBasicParsing
        Write-Host "[INSTALL] Running Python 3.12 installer (per-user)..." -ForegroundColor Cyan
        Start-Process -FilePath $tmpInstaller -ArgumentList "/passive InstallAllUsers=0 PrependPath=1 Include_test=0" -Wait
        Remove-Item $tmpInstaller -Force -ErrorAction SilentlyContinue
    }

    $PyInfo = Find-Python39Plus
    if (-not $PyInfo) {
        Write-Host "[ERROR] Python installation completed, but Python 3.9+ could not be located. Please restart PowerShell and try again." -ForegroundColor Red
        exit 1
    }
    Write-Host "[SUCCESS] Python $($PyInfo.Version) is ready!" -ForegroundColor Green
}

# 2. Download application files from GitHub if not already running inside a complete local folder
if (-not (Test-Path (Join-Path $InstallDir 'app.py')) -or -not $LocalDir) {
    Write-Host ""
    Write-Host "[DOWNLOAD] Downloading latest Vine ETV Analyzer from GitHub..." -ForegroundColor Cyan
    $tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("vineproj_setup_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
    try {
        $zipFile = Join-Path $tmpRoot "main.zip"
        Invoke-WebRequest -Uri $ZipUrl -OutFile $zipFile -UseBasicParsing
        $extractDir = Join-Path $tmpRoot "extracted"
        Expand-Archive -Path $zipFile -DestinationPath $extractDir -Force
        $sourceFolder = Get-ChildItem -Path $extractDir -Directory | Select-Object -First 1

        if (-not (Test-Path $InstallDir)) {
            New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
        }
        Get-ChildItem -Path $sourceFolder.FullName -Force | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination (Join-Path $InstallDir $_.Name) -Recurse -Force
        }
        Write-Host "[SUCCESS] Application files installed to $InstallDir" -ForegroundColor Green
    } finally {
        Remove-Item -Path $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# 3. Create Desktop Shortcut (does not modify system PATH)
try {
    if ($DesktopDir -and (Test-Path $DesktopDir)) {
        $wsh = New-Object -ComObject WScript.Shell
        $shortcut = $wsh.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = Join-Path $InstallDir "vineproject.cmd"
        $shortcut.WorkingDirectory = $InstallDir
        $shortcut.Description = "Launch Vine ETV Analyzer"
        $iconPath = Join-Path $InstallDir "icon.ico"
        if (Test-Path $iconPath) {
            $shortcut.IconLocation = "$iconPath,0"
        } else {
            $shortcut.IconLocation = "shell32.dll,21"
        }
        $shortcut.Save()
        Write-Host "[SUCCESS] Created Desktop shortcut: $ShortcutPath" -ForegroundColor Green
    }
} catch {
    Write-Host "[WARN] Could not create Desktop shortcut automatically: $_" -ForegroundColor Yellow
}

# 4. Hand off to install_deps.py to confirm/install Python packages and launch the app
Write-Host ""
Write-Host "[LAUNCH] Starting dependency verification and application launcher..." -ForegroundColor Cyan
& $PyInfo.Exe @($PyInfo.Args) (Join-Path $InstallDir "install_deps.py") @args
exit $LASTEXITCODE