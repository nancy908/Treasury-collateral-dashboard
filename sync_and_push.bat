@echo off
echo === Step 1: Finding and normalizing latest files from Google Drive ===

REM 尋找 G 槽帶有日期的 Collateral calculator 檔案，並覆寫為本機固定檔名
for %%f in ("G:\Shared drives\Tibber – House of Business\Treasury\Dashboards\Data\Collateral calculator 2 - Summary *.csv") do (
    copy /Y "%%f" "%~dp0Data\Collateral calculator 2 - Summary.csv"
)

REM 尋找 G 槽帶有日期的 CFF Treasury 檔案，並覆寫為本機固定檔名
for %%f in ("G:\Shared drives\Tibber – House of Business\Treasury\Dashboards\Data\CFF Treasury *.csv") do (
    copy /Y "%%f" "%~dp0Data\CFF Treasury.csv"
)

REM 尋找 G 槽帶有日期的 CFF FP&A 檔案，並覆寫為本機固定檔名
for %%f in ("G:\Shared drives\Tibber – House of Business\Treasury\Dashboards\Data\CFF FP&A *.csv") do (
    copy /Y "%%f" "%~dp0Data\CFF FP&A.csv"
)

REM 複製當天的 index.html
copy /Y "G:\Shared drives\Tibber – House of Business\Treasury\Dashboards\Data\index.html" "%~dp0"

echo === Step 2: Committing and pushing to GitHub ===
cd /d "%~dp0"
git add .
git commit -m "Auto sync and normalize CSVs: %date% %time%"
git push origin main

echo === Done! Dashboard and normalized data updated ===