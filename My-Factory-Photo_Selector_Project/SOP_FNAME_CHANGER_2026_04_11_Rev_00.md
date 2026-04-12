# SOP: FNAME_CHANGER Release Protection

## Scope

Protect and publish the validated `FNAME_CHANGER_2026_04_11_Rev_00.py` revision.

## Pre-Run Checklist

1. Confirm test CSV and image folder are validated.
2. Confirm output folder is empty or intended for new copies.
3. Confirm launcher works from `cmd`.

## Git Protection Workflow

1. Stage only FNAME_CHANGER release files:

```powershell
git add FNAME_CHANGER_2026_04_11_Rev_00.py run_fname_changer_rev00.bat run_fname_changer_rev00_cmd.bat README_FNAME_CHANGER_2026_04_11_Rev_00.md SOP_FNAME_CHANGER_2026_04_11_Rev_00.md
```

2. Verify staged files:

```powershell
git diff --staged --name-only
```

3. Commit:

```powershell
git commit -m "feat(fname-changer): add Rev00 app, launchers, and protection docs"
```

4. Push to GitHub:

```powershell
git push
```

## Acceptance Criteria

- Commit contains only FNAME_CHANGER files for this release.
- Commit is visible on remote branch.
- App launches successfully from `run_fname_changer_rev00_cmd.bat`.
