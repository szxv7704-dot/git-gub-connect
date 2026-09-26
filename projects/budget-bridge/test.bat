@echo off
cd /d "%~dp0"
rem ASCII only, no chcp. See run.bat for why.
set PY=
where py >nul 2>&1 && set PY=py
if not defined PY (where python >nul 2>&1 && set PY=python)
if not defined PY goto nopython
%PY% -m unittest test_budget_bridge -v
pause
goto :eof
:nopython
echo Python not found. Install it from https://www.python.org
pause
