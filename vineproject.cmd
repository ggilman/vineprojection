@echo off
where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    python "%~dp0install_deps.py" %*
    exit /b %ERRORLEVEL%
)
where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    py -3 "%~dp0install_deps.py" %*
    exit /b %ERRORLEVEL%
)
python3 "%~dp0install_deps.py" %*