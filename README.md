# DustDB v1

DustDB v1 is a local desktop tool for searching, filtering, and exporting a dust-retention research database.

## Repository Scope

This public repository contains the management tool, scripts, and documentation only. The complete research database belongs to the research group and is maintained locally by the group. It is not included in this repository.

Only small, public, desensitized examples should be placed under `sample_data/`.

## Contents

- `dust_db_downloader.py`: main desktop application source code.
- `requirements.txt`: runtime Python dependency list.
- `requirements-packaging.txt`: packaging dependency list.
- `setup_env.ps1`: create and prepare the local Python environment.
- `run_tool.ps1`: run the tool from source.
- `build_exe.ps1`: build a Windows executable.
- `sample_data/`: optional public example data only.
- Chinese documentation files: usage notes, database structure notes, packaging notes, and version notes.

## Run From Source

```powershell
.\setup_env.ps1
.\run_tool.ps1
```

## Build Executable

```powershell
.\build_exe.ps1
```

The generated executable is placed under `dist/`. Build outputs are intentionally ignored by Git. If an executable needs to be shared, publish it through GitHub Releases instead of committing it directly to the repository.

## Data Policy

Do not commit the full database, original spectral files, export results, or any non-public research data to this repository.

Before adding sample data, confirm that it is public, desensitized, and approved for release.
