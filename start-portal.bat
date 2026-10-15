@echo off
setlocal
cd /d "%~dp0"

if not exist "node_modules\exceljs\package.json" (
  echo Portal dependencies are not installed.
  echo Run npm install in this folder, then start this file again.
  pause
  exit /b 1
)

npm start
pause
