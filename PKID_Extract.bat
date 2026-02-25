@echo off
setlocal

:: -----------------------------------------------------------------------
:: PKID Extract Tool – Launcher
:: Checks for Python, offers to install via winget if missing, then runs
:: the GUI script.
:: -----------------------------------------------------------------------

:: -- Check if Python is available on PATH --------------------------------
where python >nul 2>&1
if %errorlevel%==0 (
    goto :run_script
)

echo.
echo  Python was not found on this system.
echo.

:: -- Check if winget is available ----------------------------------------
where winget >nul 2>&1
if not %errorlevel%==0 (
    echo  winget is not available on this system.
    echo  Please install Python manually from https://www.python.org/downloads/
    echo  Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

:: -- Offer to install Python via winget ----------------------------------
echo  Would you like to install Python now using winget?
echo.
choice /c YN /m "Install Python 3"
if %errorlevel%==2 (
    echo.
    echo  Installation cancelled.
    echo  Download Python manually from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo  Installing Python via winget (this may take a minute)...
echo.
winget install --id Python.Python.3.12 --source winget --accept-package-agreements --accept-source-agreements

if %errorlevel% neq 0 (
    echo.
    echo  Python installation failed.
    echo  Please install Python manually from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo.
echo  Python installed successfully.
echo  NOTE: You may need to close and reopen this window for PATH changes
echo        to take effect, then double-click this launcher again.
echo.
pause
exit /b 0

:: -- Run the script -------------------------------------------------------
:run_script
python "%~dp0extract_PKID.py"
if %errorlevel% neq 0 (
    echo.
    echo  The script exited with an error. Check the messages above.
    pause
)
