$ErrorActionPreference = "Stop"

if (-not (Test-Path ".\.venv")) {
    python -m venv .venv
}

.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements-packaging.txt
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --windowed --name "DustDB_v1" .\dust_db_downloader.py

Write-Host "打包完成，exe 位于：.\dist\DustDB_v1.exe"
