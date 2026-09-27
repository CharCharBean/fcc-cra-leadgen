# setup.ps1 — one-time setup on Windows.
#   powershell -ExecutionPolicy Bypass -File .\setup.ps1
# Builds the virtual environment, verifies the scripts and the master ledger.
# See START-HERE.md for everything else.

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot
Write-Host "Folder: $PSScriptRoot`n"

# --- Python -----------------------------------------------------------------
$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) { $py = Get-Command py -ErrorAction SilentlyContinue }
if (-not $py) {
    Write-Host "ERROR: Python not found on PATH." -ForegroundColor Red
    Write-Host "  Install 3.9+ from https://www.python.org/downloads/windows/"
    Write-Host "  and tick 'Add python.exe to PATH' during setup."
    exit 1
}
Write-Host "== Python ==" ; & $py.Source --version

# --- venv -------------------------------------------------------------------
Write-Host "`n== Virtual environment =="
$venvPy = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (Test-Path $venvPy) {
    Write-Host "  .venv exists - leaving it. Delete the folder and rerun for a clean build."
} else {
    & $py.Source -m venv .venv
    Write-Host "  created .venv"
}
& $venvPy -m pip install --quiet --upgrade pip
& $venvPy -m pip install --quiet -r requirements.txt
& $venvPy -c "import openpyxl; print('  openpyxl', openpyxl.__version__)"

# --- scripts ----------------------------------------------------------------
Write-Host "`n== Pipeline smoke test =="
& $venvPy scripts\weekly_leads.py --help | Out-Null
if ($LASTEXITCODE -eq 0) { Write-Host "  weekly_leads.py OK" }

# --- ledger -----------------------------------------------------------------
Write-Host "`n== Master ledger =="
& $venvPy -c @"
from openpyxl import load_workbook
w = load_workbook('Nemko CRA Leads.xlsx')
print('  sheets:', w.sheetnames)
print('  All Leads rows (incl. header):', w['All Leads'].max_row)
cfg = {r[0].value: r[1].value for r in w['Config'].iter_rows(min_row=2, max_col=2) if r[0].value}
print('  last_run_date:', cfg.get('last_run_date'))
print('  internal_list_path:', cfg.get('internal_list_path') or '(blank)')
"@

# --- internal workbook ------------------------------------------------------
Write-Host "`n== Internal leads workbook (optional dedupe source) =="
if (Test-Path "Nemko_CRA_DATA_Transfer.xlsx") {
    Write-Host "  present."
} else {
    Write-Host "  NOT present - this is expected; it is not shipped with the handoff." -ForegroundColor Yellow
    Write-Host "  The pipeline runs fine without it (that dedupe layer is skipped)."
    Write-Host "  To wire yours in: drop the file here and set Config!internal_list_path."
}

# --- OneDrive warning -------------------------------------------------------
if ($PSScriptRoot -match "OneDrive") {
    Write-Host "`nWARNING: this folder is inside OneDrive." -ForegroundColor Yellow
    Write-Host "  Files-On-Demand can evict the .xlsx files and break runs."
    Write-Host "  Right-click the folder -> 'Always keep on this device', or move it to C:\Nemko\."
}

Write-Host "`nDone. Next:"
Write-Host "  1. Read START-HERE.md (sections 4 and 9)."
Write-Host "  2. Run the transport smoke test - nothing works until that passes."
Write-Host "  3. Set Config!last_run_date to your starting date."
