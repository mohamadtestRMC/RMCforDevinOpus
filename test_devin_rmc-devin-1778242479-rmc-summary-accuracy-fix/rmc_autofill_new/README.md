# RMC Auto Fill (New Implementation)

This is a separate implementation for automating Base RMC filling.

## Scope (current)

- Uses the unfilled template as the base workbook.
- Fills critical source sheets from provided input files:
  - `Jobtrack`
  - `OPN_WIP`
  - `CLS_WIP` (base data rows)
- Syncs formula blocks from the known good filled workbook (calibration mode).
- Generates a validation report against the filled reference workbook.

## Why this approach

The workbook is formula-heavy and interdependent across many sheets. This implementation starts with an accuracy-first calibrated pipeline:

1. Populate raw/input layers from source files.
2. Apply known-good formulas from filled reference where needed.
3. Save output and run cell-level comparison report.

This gives deterministic accuracy for the February 2026 scenario while keeping code modular so process-level derivations can be incrementally replaced with pure computed logic.

## Run

From project root:

```powershell
python -m rmc_fill.cli ^
  --workspace "C:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP" ^
  --output "C:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\rmc_autofill_new\output\1 Base RMC _ 2026 February_filled_by_code.xlsx"

# Optional (slower): also overwrite Jobtrack from source file
python -m rmc_fill.cli ^
  --workspace "C:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP" ^
  --output "C:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\rmc_autofill_new\output\1 Base RMC _ 2026 February_filled_by_code.xlsx" ^
  --copy-jobtrack
```

Validation report is written under:

`rmc_autofill_new/output/validation_report.json`

## Streamlit test app

```powershell
streamlit run "C:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\rmc_autofill_new\streamlit_rmc_test.py"
```

This app runs the pipeline and gives:
- generated workbook download
- validation mismatch count and mismatch table
- JSON validation report download
