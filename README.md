# ELC Packing Tool

Desktop application for evaluating shipper fit and generating Excel packing reports.

## Features

- Load shipper dimensions from an Excel file
- Optional shipper weight support from Excel (included in total weight checks)
- Calculate Single and Wrap configurations
- Apply fill threshold logic (default 80%)
- Apply line-type weight limits (Non-robot: 10 kg, Robot: 14 kg)
- Export formatted Excel report sheets
- View results in an interactive 3D viewer
- Build as a Windows onefile executable

## Project Files

- packing_tool.py: Main application used for build
- requirements.txt: Python dependencies
- build_windows_app.ps1: Builds onefile Windows executable with PyInstaller
- build_installer.ps1: Creates Windows installer using Inno Setup
- ELC_Packing_Tool_v1.0.iss: Inno Setup script
- HOW_TO_INSTALL_ON_OTHER_PC.md: End-user install guide

## Requirements

- Windows
- Python 3.x (for development/build)
- PowerShell
- Inno Setup 6 (only if creating installer)

## Shipper Excel Format

Required columns:

- A
- B
- C

Optional columns:

- Shipper name column (any non-A/B/C column is used as shipper name)
- Weight column (any header containing the word Weight)

Weight values can be provided in flexible formats, for example:

- 0.22
- 0,22
- 0.22 KG

If no weight column is present, shipper weight is treated as 0.0 kg.

Total weight used by the app is:

- total_weight = (total_pieces * piece_weight) + shipper_weight

## Setup (Development)

1. Open PowerShell in this project folder.
2. Install dependencies:

```powershell
py -3 -m pip install -r requirements.txt
```

3. Run the app:

```powershell
py -3 packing_tool.py
```

Fallback command:

```powershell
python packing_tool.py
```

## Build Onefile EXE

Run:

```powershell
.\build_windows_app.ps1
```

Output:

- dist\ELC_Packing_Tool_v1.0.exe

## Build Installer (Optional)

1. Install Inno Setup 6.
2. Build app first.
3. Run:

```powershell
.\build_installer.ps1
```

Output:

- dist_installer\ELC_Packing_Tool_v1.0_Setup.exe

## Distribute to Another PC

Option A:

- Copy dist\ELC_Packing_Tool_v1.0.exe and run it

Option B:

- Copy dist_installer\ELC_Packing_Tool_v1.0_Setup.exe and install

## Notes

- Target machine does not need Python.
- First launch in onefile mode can be slower because runtime files are extracted to a temp location.
- If Windows SmartScreen appears, click More info and then Run anyway.
- Weight pass/fail is evaluated against the selected line type limit.
- Report sheets include Piece Weight (kg), Shipper Weight (kg), and Total Weight (kg).
