@echo off
rem Build RunSheetEditor install package - see the README txt
cd /d "%~dp0"
set PY=
rem 1. WinPython extracted in this folder - no install needed
for /d %%d in ("%~dp0WPy*") do (
  for /f "delims=" %%p in ('dir /b /s "%%d\python.exe" 2^>nul') do (
    if not defined PY set PY="%%p"
  )
)
rem 2. installed Python
if not defined PY (
  python -c "import sys" >nul 2>nul && set "PY=python"
)
if not defined PY (
  py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
)
if not defined PY (
  echo.
  echo  Python not found.
  echo  Download WinPython - no install needed - from https://winpython.github.io/
  echo  and extract it into this folder, then run this file again.
  echo.
  pause
  exit /b 1
)
echo Using Python: %PY%
%PY% "%~dp0release.py"
