# BHEL — Integrated Financial Model & Equity Valuation

An institutional-grade, fully-linked three-statement financial model for **Bharat Heavy Electricals Limited (BHEL)** (NSE: BHEL | BSE: 500103), built to investment-banking / equity-research conventions.

- **Workbook:** [`BHEL_Financial_Model.xlsx`](BHEL_Financial_Model.xlsx) — 22 sheets, fully dynamic Excel formulas
- **Generator:** [`build_bhel_model.py`](build_bhel_model.py) — Python/openpyxl script that produces the workbook

## Scope

| | |
|---|---|
| Reporting basis | Consolidated, Ind-AS |
| Historical period | FY2020 – FY2026 (reported) |
| Forecast period | FY2027 – FY2033 (7 years) |
| Units | INR Crore (per-share in INR) |

## Workbook structure (22 sheets)

Cover Page · Assumptions · Historical Financial Statements · Historical Ratio Analysis · Operational Drivers · Revenue Build-up · Cost Forecast · Working Capital Schedule · Fixed Asset Schedule · Depreciation Schedule · Debt Schedule · Equity Schedule · Other Assets & Liabilities · Cash Flow Statement · Forecast Financial Statements · Ratio Analysis · DCF Valuation · Relative Valuation · Sensitivity Analysis · Scenario Analysis · Dashboard · Error Checks.

## Methodology

- **Order-book-driven revenue:** `Revenue = Execution rate × Opening order book`; `Closing OB = Opening + Inflow − Revenue`. No CAGR shortcuts.
- **Cost build-up** as % of revenue (raw materials as residual) so EBITDA = revenue × target margin.
- **Fully integrated:** working-capital, fixed-asset/depreciation, debt and equity schedules feed the income statement, an indirect cash-flow statement, and the balance sheet. Cash is derived from the cash flow; the balance sheet ties to zero every forecast year (no plugs).
- **DCF (FCFF)** with a dynamically computed WACC (CAPM cost of equity; target capital structure). Interest is charged on opening debt, so there are no circular references.
- **Relative valuation** against capital-goods peers; **scenario analysis** (bull/base/bear); two-way **sensitivity** (WACC × terminal growth); automated **error checks**.

### Formatting conventions
- **Black** = hardcoded inputs · **Blue** = in-sheet formulas · **Green** = cross-sheet links · **Gold fill** = editable assumptions.

## Key base-case outputs

- FY2033E: revenue ~INR 80,800 Cr, EBITDA ~INR 9,700 Cr (12% margin), PAT ~INR 7,800 Cr, ROE ~16%, ROCE ~18%.
- Valuation spread: DCF ~INR 111 · scenario prob-weighted ~INR 226 · relative-valuation avg ~INR 453 (current price ~INR 403). On a moderate base case the DCF sits below market, reflecting that BHEL trades at rich multiples (~87x P/E).

All drivers are editable on the **Assumptions** sheet; the workbook recalculates end-to-end.

## Data provenance & caveats

- **Headline financials FY2020–FY2026** (revenue, EBITDA, depreciation, finance cost, PAT, the aggregated balance sheet, reported cash-flow summary) are sourced from BHEL's BSE/NSE filings (aggregated via Screener.in). Order book, order inflow and segment data are from BHEL results releases and the Directors' Report.
- **Detailed expense splits** (raw material / employee / subcontracting) and **detailed balance-sheet sub-lines** (inventory, receivables, payables, cash, contract assets/liabilities) are **not separately disclosed** in the aggregated sources used. They are shown as **clearly-flagged analytical estimates anchored to the audited reported totals**.
- **Peer trading multiples** on the Relative Valuation sheet are **indicative placeholders** and must be refreshed with live market data.
- Reconcile against the company's audited Annual Report notes before any transactional use.

## Regenerating the workbook

```bash
pip install openpyxl
python build_bhel_model.py
```

---
*This model is for research/educational purposes and is not investment advice.*
