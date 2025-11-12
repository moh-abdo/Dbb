@echo off
setlocal enabledelayedexpansion
set REPO_URL=https://github.com/moh-abdo/Dbb/archive/refs/heads/main.zip
set BASEDIR=%~dp0
set ZIPFILE=%BASEDIR%repo_main.zip
powershell -NoProfile -Command "try { Invoke-WebRequest -Uri '%REPO_URL%' -OutFile '%ZIPFILE%'; Expand-Archive -LiteralPath '%ZIPFILE%' -DestinationPath '%BASEDIR%'; Remove-Item '%ZIPFILE%'; } catch { Write-Host 'Error downloading or extracting repository.'; exit 1 }"
if exist "%BASEDIR%Dbb-main\access\CarRental.accdb" (
  start "" "%BASEDIR%Dbb-main\access\CarRental.accdb"
) else if exist "%BASEDIR%Dbb-main\access\open_database.bat" (
  start "" "%BASEDIR%Dbb-main\access\open_database.bat"
) else (
  echo Could not find database or open script.
  pause
  exit /b 1
)
endlocal
