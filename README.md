# Payroll and Provisions Consolidator

A desktop tool for standardizing payroll and provision spreadsheets into consolidated output bases.

**Python, pandas, openpyxl, Tkinter**

## What it does

- Normalizes input columns and combines workbook records.
- Handles payroll and provision flows with spreadsheet outputs.
- Provides execution feedback and backups before replacing generated files.

## Workflow

```mermaid
flowchart LR
  N0["Input spreadsheets"] --> N1
  N1["Normalization"] --> N2
  N2["Payroll and provisions consolidation"] --> N3
  N3["Backup"] --> N4
  N4["Output bases"]
```

## Local setup and use

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python main.py
```
Python 3.11+ is recommended. Tkinter requires a desktop Python installation.
Configure `PAYROLL_BASE_DIR`, `PROVISIONS_INPUT_DIR` and `PAYROLL_OUTPUT_DIR` as process environment variables. Default paths refer to local `input` and `outputs` folders.

## Configuration

Environment names declared by the included examples (values must be supplied locally):

`PAYROLL_BASE_DIR`, `PAYROLL_OUTPUT_DIR`, `PROVISIONS_INPUT_DIR`.

## Boundaries and limitations

Workbook layouts and financial rules are specific to this implementation and require review before reuse. No payroll, employee records or operational spreadsheets are included.

This is a source snapshot, not a claim of production readiness. Business results, dates and unmeasured performance are intentionally not asserted. Application UI language is preserved.

## License

No new distribution license is granted for this corporate source edition. Public visibility does not grant reuse rights; obtain permission from the relevant rights holder. Existing third-party notices remain applicable.

## Verification record

See [snapshot checks and remaining limitations](docs/VERIFICATION.md).
