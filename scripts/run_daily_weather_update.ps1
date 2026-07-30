$ProjectRoot = Split-Path -Parent $PSScriptRoot
$PythonCandidates = @(
    (Join-Path $ProjectRoot ".conda\python.exe"),
    (Join-Path $ProjectRoot ".venv\Scripts\python.exe")
)

$PythonExe = $PythonCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $PythonExe) {
    Write-Error "Python executable not found in .conda or .venv under $ProjectRoot"
    exit 1
}

Set-Location $ProjectRoot
& $PythonExe "scripts/refresh_weather_maps.py" --history-days 21
if ($LASTEXITCODE -ne 0) {
    Write-Error "CoT weather map refresh failed."
    exit $LASTEXITCODE
}
