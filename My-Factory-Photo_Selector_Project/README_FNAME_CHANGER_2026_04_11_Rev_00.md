# FNAME_CHANGER 2026-04-11 Rev 00

This utility copies JPG images from an input folder to an output folder and assigns new filenames from a CSV file.

## Program Files

- `FNAME_CHANGER_2026_04_11_Rev_00.py`
- `run_fname_changer_rev00.bat`
- `run_fname_changer_rev00_cmd.bat`

## Safety Rules Enforced

- Output folder must be different from input folder.
- Input files must all be `.jpg` and follow `IMG_####.jpg`.
- CSV must contain required headers:
  - `PHOTO'S CAMERA NAME` (or tolerated legacy typo `PHOTO'S CAMERAS NAME`)
  - `DATE`
  - `NEW DESCRIPTIVE FILE NAME`
- Input JPG count must exactly match CSV row count.
- CSV source names must match input names.
- Duplicate source names or duplicate target names are blocked.
- Existing output files cause a hard stop (no overwrite).

## Run

From project root:

```cmd
run_fname_changer_rev00_cmd.bat
```

or

```powershell
./run_fname_changer_rev00.bat
```
