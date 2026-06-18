# DustDB v1

DustDB v1 is a Windows desktop tool for searching, filtering, exporting, and visualizing spectra from a dust-retention research database.

## Quick Start

For ordinary users, download and run:

```text
dist/DustDB_v1.exe
```

No Python installation is required for the packaged executable.

To try the software without the full research database, use the example files:

```text
sample_data/滞尘数据汇总_示例.xlsx
sample_data/光谱/
```

In the software, select the example workbook as the summary table, select `sample_data/光谱/` as the spectra folder, then click `加载汇总表`.

## Repository Scope

This public repository contains the packaged executable, source code, documentation, and a small example dataset. The complete research database belongs to the research group and is maintained locally by the group. It is not included in this repository.

## Main Features

- Search and filter database records by metadata fields.
- Export selected records to CSV/XLSX.
- Copy matching spectral `.txt` files during export.
- Visualize single-sample spectra.
- Apply simple wavelength range filtering, moving-average smoothing, and 0-1 normalization.

## Contents

- `dist/DustDB_v1.exe`: packaged Windows executable for ordinary users.
- `sample_data/`: small example dataset for testing the tool.
- `dust_db_downloader.py`: main desktop application source code.
- `requirements.txt`: runtime Python dependency list.
- `requirements-packaging.txt`: packaging dependency list.
- `setup_env.ps1`: create and prepare the local Python environment.
- `run_tool.ps1`: run the tool from source.
- `build_exe.ps1`: build a Windows executable.
- Chinese documentation files: usage notes, database structure notes, and version notes.

## Run From Source

```powershell
.\setup_env.ps1
.\run_tool.ps1
```

## Build Executable

```powershell
.\build_exe.ps1
```

The generated executable is placed under `dist/`.

## Data Policy

Do not commit the full database, original spectral files, export results, or any non-public research data to this repository.

Only approved small example data should be placed under `sample_data/`.
