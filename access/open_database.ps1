$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$db = Join-Path $scriptDir 'CarRental.accdb'
if (-Not (Test-Path $db)) { Write-Host "Database not found: $db"; exit 1 }
Start-Process -FilePath $db