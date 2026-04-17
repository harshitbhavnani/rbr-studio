# HB Relationship Review Studio — Mockups Based on V1

This package uses the uploaded V1 calculation and export logic as the base, while redesigning the front-end flow.

## What is included

- **Homepage**
  - Portfolio metrics
  - Client group list
  - Open a client group into the RBR page

- **RBR page**
  - Relationship-level common information at the top
  - Two second-page mockups:
    1. **Mock 1 — Loan Tabs**
       - One tab per loan
       - Loan overview
       - Rent Roll Analysis
       - Investor CRE Stress Test
       - Rollover Risk
       - Construction / Bridge Loan Stress Test
    2. **Mock 2 — Consolidated Table**
       - One compact table for all loans
       - Important locked COGNOS fields
       - Important editable manual fields
       - Detail view for the selected loan
       - Detailed workspace for the selected loan

- **Final Review / Export**
  - Select one loan for final review and export
  - Edit final commentary
  - Validate all sections
  - Download Word document

## Notes

- Data is synthetic.
- **COGNOS / Database fields** are locked.
- **Manual fields** are editable.
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
