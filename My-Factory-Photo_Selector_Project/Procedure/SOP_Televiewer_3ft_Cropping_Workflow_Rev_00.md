# Team SOP: Televiewer 3-ft Cropping Workflow

Document ID: SOP-TEL-3FT-CROP-REV-00  
Effective Date: 2026-04-09  
Owner: Televiewer Processing Team  
Status: Active

## 1. Purpose
Standardize televiewer image cropping so each output crop represents a true 3.0 ft interval using ruler-evidence validation (30 left-ruler ticks), preventing regressions to visually compressed 1.5 ft crops.

## 2. Scope
This SOP applies to all televiewer long-image cropping runs used for report generation and QA review.

## 3. Authoritative Tools
- Crop utility: `Televiewer_TickCrop_30Ticks.py`
- Launcher: `run_tick_crop_30ticks.bat`
- Output folder: `cropping_test`

## 4. Definitions
- Tick: One visible depth ruler mark on the left ruler.
- 3.0 ft validation rule: 30 left-ruler ticks per crop.
- Accepted crop: A crop saved only when left-ruler tick count is at least 30.

## 5. Required Inputs
- One source televiewer long image file (JPG/compatible image).
- Local workspace with the crop utility and launcher present.

## 6. Preconditions (Must Pass Before Run)
1. Confirm `Televiewer_TickCrop_30Ticks.py` exists in project root.
2. Confirm `run_tick_crop_30ticks.bat` exists in project root.
3. Confirm source image is final and not edited during run.
4. Confirm no other process is writing files into `cropping_test`.

## 7. Standard Run Procedure
1. Start in project root.
2. Run `run_tick_crop_30ticks.bat`.
3. Provide source image path when prompted (or pass path as argument).
4. Wait for completion and record console summary:
   - Detected tick spacing
   - Estimated 3-ft window height
   - Segments written
   - Tick criterion statement
5. Verify that new crops are created in `cropping_test`.

## 8. Acceptance Criteria (Release Gate)
A run is accepted only if all criteria pass:
1. Utility reports nonzero `Segments written`.
2. Utility reports `Tick criterion: >= 30 ticks on LEFT ruler`.
3. Random sample QA (minimum 5 crops or all if fewer than 5): each sampled crop shows approximately 30 left-ruler ticks.
4. Segment sequence is continuous with no obvious depth jumps or overlaps at transitions.
5. No stale prior run artifacts are mixed in output (the utility must have cleaned old JPG outputs before writing new ones).

## 9. QA Sampling Method
1. Open first, middle, and last crop, plus two additional random crops.
2. Count visible left-ruler ticks in each sampled crop.
3. If any sampled crop is less than 30 ticks, mark run as failed.
4. If failed, execute troubleshooting section and rerun.

## 10. Troubleshooting and Recovery
- ERROR: Image not found
  - Validate full image path and rerun.
- ERROR: Could not detect ruler tick spacing
  - Check image quality/contrast and rerun with correct source image.
  - Confirm ruler strip is visible and not clipped in source.
- WARNING: No 30-tick crops produced
  - Treat run as failed.
  - Inspect source image for unreadable or missing ruler ticks.
  - Escalate to engineering if repeated on valid source data.

## 11. Regression Prevention Rules
1. Do not replace left-ruler-only acceptance with combined left+right tick counting.
2. Do not accept crops using pixel-height-only logic without ruler-evidence confirmation.
3. Do not modify threshold from 30 left ticks unless approved via controlled change.
4. Any code change affecting crop geometry must include side-by-side QA against baseline outputs.

## 12. Change Control
1. Proposed algorithm changes require:
   - Rationale and expected impact.
   - Test set with before/after evidence.
   - QA sign-off by owner.
2. Increment SOP revision after approved process changes.
3. Archive sample outputs for traceability.

## 13. Operational Checklist (Quick Use)
- [ ] Source image verified
- [ ] Utility and launcher present
- [ ] Run completed without errors
- [ ] Segments written is nonzero
- [ ] 30-left-tick criterion confirmed
- [ ] QA sample completed and passed
- [ ] Outputs ready for downstream report workflow

## 14. Roles and Responsibilities
- Operator: Executes run and completes checklist.
- QA Reviewer: Performs sampling and acceptance decision.
- Engineering Owner: Handles algorithm changes and incidents.

## 15. Recordkeeping
For each run, store:
- Source image name
- Run date/time
- Segments written
- QA sample notes
- Final pass/fail decision

End of Document
