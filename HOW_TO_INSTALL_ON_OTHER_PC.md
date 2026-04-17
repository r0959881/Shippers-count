# ELC Packing Tool - Build and Install Guide

## 1) Build the standalone app

On your development machine (with Python installed), open PowerShell in this folder and run:

```powershell
.\build_windows_app.ps1
```

This creates:

- `dist\ELC_Packing_Tool_v1.0.exe`

You can already copy that single EXE to another machine and run it.

## 2) Optional: Create a setup installer (.exe)

Install **Inno Setup 6** once on your development machine, then run:

```powershell
.\build_installer.ps1
```

This creates:

- `dist_installer\ELC_Packing_Tool_v1.0_Setup.exe`

## 3) Install on another machine

### Option A (no installer)

- Copy `dist\ELC_Packing_Tool_v1.0.exe` to the target machine.
- Double-click `ELC_Packing_Tool_v1.0.exe`.

### Option B (with installer)

- Copy `dist_installer\ELC_Packing_Tool_v1.0_Setup.exe` to the target machine.
- Run it and follow setup.

## Notes

- The target machine does **not** need Python.
- First launch can be slower in onefile mode because files are extracted to a temp folder.
- If SmartScreen warns, click `More info` then `Run anyway` (common for internal unsigned apps).
