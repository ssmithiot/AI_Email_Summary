param(
    [switch]$SkipPyInstaller
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $repoRoot

Write-Host "== AI Email Summary installer build =="

if (-not (Test-Path ".\venv\Scripts\python.exe")) {
    throw "Virtual environment not found at .\venv. Run setup first."
}

if (-not $SkipPyInstaller) {
    Write-Host "Installing PyInstaller..."
    & .\venv\Scripts\python.exe -m pip install pyinstaller | Out-Host

    Write-Host "Building bundled executable..."
    & .\venv\Scripts\python.exe -m PyInstaller --clean --noconfirm .\AI_Email_Summary.spec | Out-Host
}

$iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if (-not $iscc) {
    $fallbackIscc = Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'
    if (Test-Path $fallbackIscc) {
        $iscc = @{ Source = $fallbackIscc }
    }
}
if (-not $iscc) {
    Write-Warning "Inno Setup Compiler (ISCC.exe) was not found in PATH."
    Write-Host "PyInstaller output is ready at .\dist\AI_Email_Summary.exe"
    Write-Host "Install Inno Setup, then run:"
    Write-Host '  ISCC.exe ".\installer\AI_Email_Summary.iss"'
    exit 0
}

Write-Host "Compiling Setup.exe with Inno Setup..."
& $iscc.Source ".\installer\AI_Email_Summary.iss" | Out-Host

Write-Host "Done."
