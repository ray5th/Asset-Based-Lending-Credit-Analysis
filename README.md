# Asset-Based Lending Credit Analysis

This project is an interview-ready ABL analyst case study for **Genuine Parts Company (GPC)**. It uses public SEC financial statement data to build a borrowing-base and cash-flow stress-test model, then summarizes the credit view in a short lender-style memo.

## Deliverables

- `outputs/019f1f2b-982d-79f1-aa91-d9d43d8deb5a/GPC_ABL_Borrowing_Base_Credit_Model.xlsx`
- `outputs/019f1f2b-982d-79f1-aa91-d9d43d8deb5a/GPC_ABL_Credit_Memo.md`

## What The Model Shows

- 5-year historical financial analysis using SEC companyfacts data
- AR and inventory collateral schedules scaled to reported FY2025 balances
- AR aging eligibility controls and ineligible receivables
- Inventory eligibility controls for slow-moving, obsolete, consigned, and offsite inventory
- Customer concentration risk and concentration reserve mechanics
- Borrowing-base calculation by base, downside, and stress scenarios
- Liquidity, leverage, interest coverage, and free-cash-flow stress testing
- Dashboard with key ABL credit monitoring outputs
- Checks tab for model tie-outs and formula QA

## Source Data

The model pulls SEC data for Genuine Parts Company, CIK `0000040987`.

- Latest annual source used: FY2025 Form 10-K filed February 20, 2026
- Latest quarterly context noted: Q1 2026 Form 10-Q filed April 21, 2026
- Extracted source files:
  - `data/gpc_sec_financials.json`
  - `data/gpc_historical_financials.csv`

Public filings do not provide borrowing-base certificate detail such as invoice-level AR aging, customer-level AR balances, inventory appraisal categories, or field-exam results. Those ABL control schedules are illustrative and clearly separated from SEC-reported financial statement inputs.

## Rebuild Instructions

```bash
python3 scripts/fetch_gpc_sec_data.py
python3 scripts/build_abl_credit_model.py
```

The fetch script uses the SEC companyfacts and submissions APIs. The build script creates the Excel workbook and Markdown credit memo in `outputs/$CODEX_THREAD_ID/`.

## Interview Framing

I built this project to answer three ABL credit questions:

1. How much can the lender advance against eligible AR and inventory collateral?
2. Can the borrower support debt service through cash flow?
3. What happens to availability, liquidity, and coverage under downside and stress cases?

This positions the project as a credit monitoring and collateral valuation tool rather than a generic credit dashboard.
