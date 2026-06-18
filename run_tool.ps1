$ErrorActionPreference = "Stop"

if (Test-Path ".\.venv\Scripts\python.exe") {
    .\.venv\Scripts\python.exe .\dust_db_downloader.py
} else {
    python .\dust_db_downloader.py
}
