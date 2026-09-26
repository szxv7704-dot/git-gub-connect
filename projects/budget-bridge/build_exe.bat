@echo off
cd /d "%~dp0"
rem ASCII only, no chcp. Every message comes from build.py.
rem A chcp call inside a batch file shifts cmd's byte offset and
rem chops the commands after it (openpyxl became enpyxl).
set PY=
where py >nul 2>&1 && set PY=py
if not defined PY (where python >nul 2>&1 && set PY=python)
if not defined PY goto nopython
%PY% build.py %*
goto :eof
:nopython
echo Python not found. Install it from https://www.python.org
pause
