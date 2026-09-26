@echo off
cd /d "%~dp0"
rem ASCII only, and no chcp. cmd re-reads a batch file by byte
rem offset, so switching codepage mid-file chops later commands.
rem Korean text is printed by Python, which writes Unicode to the
rem console directly and does not care about the codepage.
set PY=
where py >nul 2>&1 && set PY=py
if not defined PY (where python >nul 2>&1 && set PY=python)
if not defined PY goto nopython
%PY% budget_bridge.py
if errorlevel 1 pause
goto :eof
:nopython
echo Python not found. Install it from https://www.python.org
pause
