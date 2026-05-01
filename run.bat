@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv" (
    echo Creating virtual environment in .venv ...
    python -m venv .venv
    if errorlevel 1 exit /b 1
)

if not exist ".venv\.deps-installed" goto install_deps
for %%a in (requirements.txt) do set REQ_TIME=%%~ta
for %%a in (.venv\.deps-installed) do set DEP_TIME=%%~ta
if "%REQ_TIME%" GTR "%DEP_TIME%" goto install_deps
goto run

:install_deps
echo Installing dependencies ...
.venv\Scripts\pip.exe install --quiet --upgrade pip
.venv\Scripts\pip.exe install --quiet -r requirements.txt
if errorlevel 1 exit /b 1
type nul > ".venv\.deps-installed"

:run
.venv\Scripts\python.exe main.py %*
