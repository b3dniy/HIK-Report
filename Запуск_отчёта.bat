@echo off
chcp 65001 >nul
title NVR login report - Hikvision
cd /d "%~dp0"

set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY where python >nul 2>&1 && set "PY=python"
if not defined PY goto nopython

%PY% "%~dp0nvr_report.py"
goto end

:nopython
echo.
echo [ERROR] Python 3 not found.
echo Install Python 3 from https://www.python.org/downloads/
echo and enable "Add python.exe to PATH" during setup.
echo.
pause

:end
