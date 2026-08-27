$ErrorActionPreference = 'Stop'
$ProjectPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (Test-Path -LiteralPath $ProjectPython) {
    $PythonExe = $ProjectPython
} else {
    $PyenvRoot = Join-Path $env:USERPROFILE '.pyenv\pyenv-win\versions'
    $PyenvPython = Get-ChildItem -LiteralPath $PyenvRoot -Filter python.exe -Recurse -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
    if ($PyenvPython) { $PythonExe = $PyenvPython } else { $PythonExe = (Get-Command python -ErrorAction Stop).Source }
}
$AppName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('64SI66y0IOyXkeyFgCDtjKHshZgg7IKs7Jqp7ZWY7KeAIOuniOyEuOyalA=='))
& $PythonExe -m PyInstaller --noconfirm --clean --onefile --windowed --name $AppName (Join-Path $PSScriptRoot 'app.py')
if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed. Exit code: $LASTEXITCODE" }
Write-Host "Built: dist\$AppName.exe"
