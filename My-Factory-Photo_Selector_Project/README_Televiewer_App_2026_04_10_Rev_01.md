# Televiewer App 2026-04-10 Rev 01

This file records the stable working snapshot of the Televiewer cropping/report app.

## Protected Runtime File

- App: `Televiewer_App_2026_04_10_Rev_01.py`
- Launcher: `run_televiewer_report_app.bat`

## What This Revision Protects

- Dynamic horizontal crop boundaries (prevents 90/180 clipping).
- First-page top band/header preservation improvements.
- Tick-based vertical crop sizing (regression-safe from fixed-pixel windowing).
- Page-to-page overlap support (`OVERLAP_FEET = 0.1`).

## How To Run

1. Double-click `run_televiewer_report_app.bat`, or run from terminal:

```powershell
./run_televiewer_report_app.bat
```

2. Select input image and output folder in the UI.
3. Generate cropped images and/or report.

## Environment Notes

- OS: Windows
- Python: virtual environment in `.venv`
- Main libraries used by app: `Pillow`, `python-docx`, `requests`

## Release Protection Steps

Use this sequence to protect this revision in Git:

```powershell
git add Televiewer_App_2026_04_10_Rev_01.py run_televiewer_report_app.bat README_Televiewer_App_2026_04_10_Rev_01.md SOP_Televiewer_App_2026_04_10_Rev_01.md
git commit -m "chore: protect Televiewer Rev 01 with SOP and README"
git push origin master
```

## Rollback Safety

A backup created before dynamic horizontal crop changes exists in this workspace:

- `Televiewer_App_2026_04_09_Rev_00_backup_before_dynamic_horizontal_crop_2026_04_10_021215.py`
