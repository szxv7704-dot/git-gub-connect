param([switch]$SkipInstall)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Resolve-Path (Join-Path $ProjectDir "..\..")

$PythonCommand = Get-Command py -ErrorAction SilentlyContinue
if ($PythonCommand) {
    $Python = $PythonCommand.Source
    $PythonArgs = @("-3")
} else {
    $PythonCommand = Get-Command python -ErrorAction Stop
    $Python = $PythonCommand.Source
    $PythonArgs = @()
}

if (-not $SkipInstall) {
    & $Python @PythonArgs -m pip install -r (Join-Path $ProjectDir "requirements.txt") pyinstaller
}

Push-Location $RepoRoot
try {
    & $Python @PythonArgs -m PyInstaller --noconfirm --clean (Join-Path $ProjectDir "SchoolDataWorkflow.spec")
} finally {
    Pop-Location
}

$Exe = Join-Path $RepoRoot "dist\학교자료취합도우미.exe"
if (-not (Test-Path $Exe)) {
    throw "실행 파일이 생성되지 않았습니다: $Exe"
}
Write-Host "빌드 완료: $Exe"
