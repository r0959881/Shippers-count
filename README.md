# ELC Packing Tool

Desktop application for evaluating shipper fit and generating Excel packing reports.

## Features

- Load shipper dimensions from an Excel file
- Calculate Single and Wrap configurations
- Apply fill threshold logic (default 80%)
- Export formatted Excel report sheets
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
