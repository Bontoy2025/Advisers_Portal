@echo off
setlocal
cd /d "%~dp0"

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is not installed or is not on PATH.
  echo Install Node.js 22.5 or newer, then reopen this file.
  pause
  exit /b 1
)

where npm >nul 2>&1
if errorlevel 1 (
  echo npm is not installed or is not on PATH.
  echo Repair your Node.js installation, then reopen this file.
  pause
  exit /b 1
)

node -e "const v=process.versions.node.split('.').map(Number);if(v[0]<22||(v[0]===22&&v[1]<5))process.exit(1);require('node:sqlite')" >nul 2>&1
if errorlevel 1 (
  echo This portal requires Node.js 22.5 or newer with built-in SQLite support.
  echo Update Node.js, reopen this terminal, and start the portal again.
  pause
  exit /b 1
)

if exist "node_modules\exceljs\package.json" if exist "node_modules\@e965\xlsx\package.json" goto dependenciesready

echo Installing portal dependencies from the lockfile...
if exist "package-lock.json" (
  call npm ci
) else (
  call npm install
)
if errorlevel 1 (
  echo Dependency installation failed. Check your internet connection and npm output.
  pause
  exit /b 1
)

:dependenciesready
echo Starting the portal. Keep this window open while you use it.
npm start
set "PORTAL_EXIT_CODE=%ERRORLEVEL%"
if not "%PORTAL_EXIT_CODE%"=="0" echo The portal stopped with exit code %PORTAL_EXIT_CODE%.
pause
exit /b %PORTAL_EXIT_CODE%
