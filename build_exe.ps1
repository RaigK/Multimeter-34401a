# ============================================================
# build_exe.ps1
# Builds a standalone Windows .exe from multimeter_34401A.py
# Run this script from the repo directory:
#   .\build_exe.ps1
# ============================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$MainPy    = Join-Path $ScriptDir "multimeter_34401A.py"
$DistDir   = Join-Path $ScriptDir "dist"

Write-Host "=== Multimeter 34401A – EXE Builder ===" -ForegroundColor Cyan

# ── 1. Check Python ──────────────────────────────────────────
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: Python not found. Install from python.org" -ForegroundColor Red
    exit 1
}
Write-Host "Python: $(python --version)" -ForegroundColor Green

# ── 2. Install / upgrade PyInstaller ────────────────────────
Write-Host "`nInstalling PyInstaller..." -ForegroundColor Cyan
pip install pyinstaller --upgrade -q
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: pip install failed" -ForegroundColor Red
    exit 1
}

# ── 3. Install required packages ────────────────────────────
Write-Host "Installing dependencies..." -ForegroundColor Cyan
pip install pyvisa matplotlib openpyxl numpy -q

# ── 4. Run PyInstaller ───────────────────────────────────────
Write-Host "`nBuilding EXE..." -ForegroundColor Cyan
Set-Location $ScriptDir

python -m PyInstaller `
    --onefile `
    --windowed `
    --name "Multimeter_34401A" `
    --hidden-import "pyvisa" `
    --hidden-import "pyvisa.resources" `
    --hidden-import "pyvisa.resources.gpib" `
    --hidden-import "pyvisa.resources.serial" `
    --hidden-import "pyvisa.resources.usb" `
    --hidden-import "openpyxl" `
    --hidden-import "openpyxl.styles" `
    --hidden-import "openpyxl.chart" `
    --hidden-import "matplotlib.backends.backend_tkagg" `
    --hidden-import "numpy" `
    --collect-submodules "matplotlib" `
    "$MainPy"

if ($LASTEXITCODE -ne 0) {
    Write-Host "`nERROR: Build failed" -ForegroundColor Red
    exit 1
}

# ── 5. Show result ───────────────────────────────────────────
$exe = Join-Path $DistDir "Multimeter_34401A.exe"
if (Test-Path $exe) {
    $size = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host "`n=== BUILD SUCCESSFUL ===" -ForegroundColor Green
    Write-Host "EXE: $exe" -ForegroundColor Green
    Write-Host "Size: $size MB" -ForegroundColor Green
    Write-Host "`nNote: settings.json will be created next to the EXE on first run." -ForegroundColor Yellow
} else {
    Write-Host "ERROR: EXE not found after build" -ForegroundColor Red
    exit 1
}
