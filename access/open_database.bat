@echo off
REM Open Access database (CarRental.accdb) located in the same folder as this script
SET DB=%~dp0CarRental.accdb
IF NOT EXIST "%DB%" (
  echo Database file not found: "%DB%"
  pause
  exit /b 1
)
start "" "%DB%"