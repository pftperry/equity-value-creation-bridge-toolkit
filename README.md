# Equity Value Creation Bridge Toolkit

Post Fiat Analytical Infrastructure Library, v1.0. This is a reusable Excel workbook for lower-middle-market private equity. It bridges sponsor **entry equity value to exit equity value** across seven levers using live formulas: organic revenue growth, add-on revenue, margin change, multiple change, add-on purchase price, debt paydown and cash generation. It ranks the levers, tests how much the return relies on multiple expansion, runs a downside case and returns one underwriting recommendation.

## Files

| File | Purpose |
|---|---|
| `Equity_Value_Creation_Bridge_Toolkit.xlsx` | The workbook: 7 tabs, values recalculated and cached in Microsoft Excel |
| `build_value_creation_bridge.py` | openpyxl script that regenerates the workbook: `python build_value_creation_bridge.py out.xlsx` |
| `excel_recalc_and_screenshots.ps1` | Excel COM script that runs CalculateFull, scans for formula errors, saves and exports the evidence screenshots |
| `evidence/*.png` | Screenshots rendered in Excel (bridge, summary, worked example, chart) |

Workbook SHA-256: `8C3F037CC885626C0E03D0FDE6FF24EA2E8A55E9956B6B98F068954291B90C66` (31901 bytes)

## Tabs

1. ReadMe - Validator Instructions (Excel forbids `/` in tab names, so "Read Me / Validator Instructions" is named this way)
2. Entry and Exit Assumptions
3. Value Driver Inputs (EBITDA growth, margin, multiple, leverage and cash conversion, plus the annual EBITDA/FCF schedule)
4. Equity Value Creation Bridge
5. Lever Impact Ranking
6. Summary Recommendation
7. Anonymized Worked Example

Colour legend: yellow cells are inputs, blue cells are formulas and green cells are key outputs. Inputs exist only in `Entry and Exit Assumptions!B5:B24` and `Value Driver Inputs!B4:B18`. Everything else is a formula.

## Bridge math

`Exit equity - Entry equity = (EV1 - EV0) - (ND1 - ND0)`, decomposed exactly as:

- Organic growth = (Organic exit revenue - Entry revenue) x Entry margin x Entry multiple
- Add-on revenue = Add-on revenue x Entry margin x Entry multiple
- Margin change = Exit revenue x (Exit margin - Entry margin) x Entry multiple
- Multiple change = Exit EBITDA x (Exit multiple - Entry multiple)
- Add-on purchase price = -(debt-funded add-on price)
- Debt paydown + cash generation = cumulative FCF, split by the sweep %

The bridge is checked against exit equity computed directly (`Bridge!D13:D16`). `Bridge!E16` must read `OK - bridge reconciles`.

## Validator instructions: criterion -> tab -> cell range

| # | Acceptance criterion | Tab | Cell range | Expected (worked example) |
|---|---|---|---|---|
| 1 | Required tab structure | all | 7 tabs listed above; `ReadMe!A15:C21` | 7 tabs |
| 2 | Entry and exit assumptions | Entry and Exit Assumptions | `B5:B24` (inputs), `B7`, `B9`, `B11`, `B12` (formulas) | Entry equity 32.00 |
| 3 | Value driver inputs: EBITDA growth, margin, multiple, leverage, cash conversion | Value Driver Inputs | `B4:B18`; schedule `A22:F29`; check `B30` | `B30` = OK |
| 4 | Bridge entry -> exit equity with live visible formulas | Equity Value Creation Bridge | `D4:D12`; FORMULATEXT `F4:F19` | Exit equity 110.13 |
| 5 | Bridge reconciles to direct calculation | Equity Value Creation Bridge | `D13:D16`, status `E16` | `D16` = 0.00, `E16` = OK |
| 6 | Lever impact ranking | Lever Impact Ranking | `A4:F10`; reliance tests `A14:D20` | Largest lever: Debt paydown (33.3%) |
| 7 | Clear recommendation | Summary Recommendation | `B27`; reasoning `B28`; formula text `B29`; downside `B16:B23` | UNDERWRITE - VALUE CREATION PLAN SUPPORTS RETURNS |
| 8 | Anonymized worked example with visible bridge output | Anonymized Worked Example | `A14:C22` plus bar chart; result `A33:B38` | MOIC 3.44x, IRR 28.0% |
| 9 | At least 3 named judgment calls | Anonymized Worked Example | `A26:F29` (JC1-JC4, with live impact formulas in `E26:E29`) | 4 judgment calls, 54.06 $M excluded |
| 10 | Zero formula errors | all | `excel_recalc_and_screenshots.ps1` scan | 0 error cells |

## Worked example: HoldCo B (anonymized, scaled, illustrative)

Entry: 9.00 EBITDA at 8.0x, with 40.0 net debt and 32.00 equity. The model runs a 5-year hold to 15.69 exit EBITDA at 8.0x.

| Bridge step | $M | % of value created |
|---|---:|---:|
| Entry equity | 32.00 | |
| Organic revenue growth | 24.35 | 31.2% |
| Add-on revenue | 14.40 | 18.4% |
| Margin change | 14.77 | 18.9% |
| Multiple change | 0.00 | 0.0% |
| Add-on purchase price (debt-funded) | -7.92 | -10.1% |
| Debt paydown | 26.03 | 33.3% |
| Cash generation (retained) | 6.51 | 8.3% |
| **Exit equity** | **110.13** | |

Result: gross MOIC 3.44x and gross IRR 28.0%. Operational levers account for 68.5% of value created (minimum 50%) and multiple change for 0% (maximum 25%). The downside case gives a MOIC of 2.03x against a 1.25x floor. Recommendation: **UNDERWRITE - VALUE CREATION PLAN SUPPORTS RETURNS**.

Named judgment calls (equity excluded relative to the management case): JC1 exit multiple held at 8.0x rather than 9.5x (23.53); JC2 +200 bps margin credited rather than +400 bps (14.77); JC3 unsigned $20M add-on pipeline excluded (6.00); JC4 cash conversion of 50% rather than 65% (9.76).

## Verification

Recalculated in Microsoft Excel through COM (`CalculateFull`). The `SpecialCells(xlCellTypeFormulas, xlErrors)` scan found **0 error cells on all 7 tabs** (212 formula cells). Cached values were saved.

All data is anonymized, scaled and illustrative. No company identifiers are used.
