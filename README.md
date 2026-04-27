# HB Relationship Review Studio — Mockups Based on V1

This package uses the uploaded V1 calculation and export logic as the base, while redesigning the front-end flow.

## What is included

- **Homepage**
  - Portfolio metrics
  - Search by group name or group number
  - Matching group result list
  - Open a client group into the RBR page
  - Open, download, and delete archived versions for future reference

- **RBR page**
  - Snapshot-style relationship summary at the top
  - Relationship-level database fields above manual inputs
  - **Loan Tabs**
    - One tab per loan
    - Loan overview with database-driven loan data above manual inputs
    - Rent Roll Analysis
    - Investor CRE Stress Test
    - Rollover Risk
    - Construction / Bridge Loan Stress Test

- **Final Review / Export**
  - Select one loan for final review and export
  - Edit final commentary
  - Validate all sections
  - Download Word document
  - Archive the current relationship version with a dated filename

## Notes

- Data is synthetic.
- **COGNOS / Database fields** are locked.
- **Manual fields** are editable.
- Archived code versions are stored under `archives/`.
- Archived app data snapshots are stored under `archives/data_snapshots/`.
- Rent roll is fully manual and supports dynamic rows.
- Word export reuses the uploaded `export_credit_report.py` and `export_helpers.py`.

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m streamlit run app.py
```
