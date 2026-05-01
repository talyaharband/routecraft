# Routecraft

This repository contains Python scripts for processing address data in Excel files.

For the full address-to-route workflow, see [PIPELINE.md](PIPELINE.md). For a clickable visual flow map, open [pipeline.html](pipeline.html).

The current VS Code run profiles are configured for:

- `data.cleanup.py`
- `cluster.py`
- `k-mean.py`
- `geocoding.py`
- `distance.matrix.py`
- `TSP.py`

## Prerequisites

Team members need:

- Windows
- VS Code
- Python 3.13 or later
- The VS Code Python extension

## First-Time Setup In VS Code

1. Clone the repository.
2. Open the repository folder in VS Code.
3. Open a new terminal in VS Code.
4. Create a virtual environment:

```powershell
python -m venv .venv
```

5. Activate the virtual environment:

```powershell
.\.venv\Scripts\Activate.ps1
```

6. Install dependencies:

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

7. In VS Code, select the interpreter:

- Press `Ctrl+Shift+P`
- Run `Python: Select Interpreter`
- Choose `.venv\Scripts\python.exe`

The repository already includes `.vscode/settings.json`, so VS Code should usually detect the local `.venv` automatically.

## Running A Script From VS Code

### Option 1: Run From VS Code

1. Open the Run and Debug view.
2. Select one of these profiles:

- `Python: data.cleanup.py`
- `Python: cluster.py`
- `Python: k-mean.py`
- `Python: geocoding.py`
- `Python: distance.matrix.py`
- `Python: TSP.py`

3. Press `F5`.

The script opens a file picker:

1. Select the input Excel file.
2. After processing, choose where to save the output Excel file.

## Running From The Terminal

Activate the virtual environment first:

```powershell
.\.venv\Scripts\Activate.ps1
python .\cluster.py
```

Examples for the other scripts:

```powershell
python .\data.cleanup.py
python .\geocoding.py
python .\cluster.py
python .\k-mean.py
python .\distance.matrix.py
python .\closesToEshtaol.py
python .\TSP.py
```

## What `cluster.py` Expects

The input Excel file must contain these columns:

- `Street_Name`
- `House_Number`
- `LAT`
- `LNG`

## What `cluster.py` Produces

The output is a new Excel file with one representative row per cluster, plus these generated columns:

- `cluster_id`
- `total_orders_in_cluster`
- `detailed_addresses`

## Rebuilding The Environment

If the environment gets into a bad state:

1. Delete `.venv`
2. Recreate it
3. Reinstall from `requirements.txt`

Commands:

```powershell
Remove-Item -LiteralPath .venv -Recurse -Force
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Should `.vscode` Be In Git?

Yes, in this repo it makes sense to commit `.vscode`.

Why:

- `launch.json` gives the team a shared run profile for `cluster.py`
- `launch.json` gives the team shared run profiles for the project scripts
- `settings.json` points VS Code at the project-local `.venv`
- these are workspace-level project settings, not personal machine preferences

What should stay out of Git:

- `.venv/`
- local temp files
- machine-specific secrets

That is already handled by `.gitignore`.
