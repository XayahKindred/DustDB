$ErrorActionPreference = "Stop"

if (-not (Test-Path ".\.venv")) {
    python -m venv .venv
}

.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

Write-Host "环境配置完成。运行工具："
Write-Host ".\.venv\Scripts\python.exe .\dust_db_downloader.py"
