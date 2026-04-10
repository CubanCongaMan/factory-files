# SOP: Protect Televiewer App Revision

## Document Control

- SOP ID: SOP-TEL-CROP-REV01
- App Revision: `Televiewer_App_2026_04_10_Rev_01.py`
- Effective Date: 2026-04-10
- Scope: Protect and publish the validated Televiewer revision to Git/GitHub.

## Purpose

Create a repeatable release process that protects the validated Televiewer version and supporting documentation.

## Prerequisites

- Working directory: project root.
- Git is installed and authenticated to GitHub.
- Confirm app launch file exists:
  - `Televiewer_App_2026_04_10_Rev_01.py`
  - `run_televiewer_report_app.bat`

## Procedure

1. Verify current branch and status:

```powershell
git branch --show-current
git status --short
```

2. Stage only release files:

```powershell
git add Televiewer_App_2026_04_10_Rev_01.py
git add run_televiewer_report_app.bat
git add README_Televiewer_App_2026_04_10_Rev_01.md
git add SOP_Televiewer_App_2026_04_10_Rev_01.md
```

3. Validate staged content:

```powershell
git diff --staged --name-only
git diff --staged
```

4. Create release commit:

```powershell
git commit -m "chore: protect Televiewer Rev 01 with SOP and README"
```

5. Publish to GitHub:

```powershell
git push origin master
```

6. Confirm remote contains commit:

```powershell
git log --oneline -n 3
```

## Acceptance Criteria

- Commit contains exactly these files:
  - `Televiewer_App_2026_04_10_Rev_01.py`
  - `run_televiewer_report_app.bat`
  - `README_Televiewer_App_2026_04_10_Rev_01.md`
  - `SOP_Televiewer_App_2026_04_10_Rev_01.md`
- Push completes successfully.
- Team can launch app from batch file with no filename mismatch.

## Recovery

If validation fails after commit:

1. Create a corrective commit (preferred).
2. If not yet pushed, use:

```powershell
git reset --soft HEAD~1
```

Do not use destructive history rewrite after shared push unless team-approved.
