# Builds the client package in dist\ : the exe, the Word user guide and the input files.
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

& .venv\Scripts\pyinstaller --onefile --console --name pricetracker --noconfirm --clean `
    --distpath dist --workpath build pricetracker\__main__.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed" }

& .venv\Scripts\python tools\build_guide_docx.py
if ($LASTEXITCODE -ne 0) { throw "Building the Word guide failed" }

# Input files: the working copy if there is one, otherwise the sample data
$source = if (Test-Path input) { "input" } else { "sample-input" }
New-Item -ItemType Directory -Force dist\input | Out-Null
Copy-Item "$source\*.xlsx" dist\input\ -Force
Write-Host "Built dist\pricetracker.exe with input files from $source"
