# Changelog

All notable pricing decisions and strategy updates for QC 2026.

## [1.1.0] — 2026-02-21

### Estimate Template Rebuilt (`QC_Estimate_Template_2026_v2.xlsx`)

#### Client Setup
- Replaced single flat markup field (C19 = 54%) with 11-category markup table (rows 19–31)
- Each category has its own editable markup percentage with rationale note
- Added "Absolute Floor: 54%" reference row
- Created 11 named ranges (`Markup_Catering`, `Markup_Venues`, `Markup_AV`, etc.) for readable cross-sheet formulas
- Shifted CC Processing, Commission, GDP, Restaurant Fees, How-to-Use, and Tax Table sections down 12 rows
- Updated `LocationList` named range to match shifted tax table position

#### Venue Estimate
- Replaced all flat-markup formulas with category-specific named range references:
  - F&B (rows 21–23) → `Markup_Catering` (54%)
  - Staffing (row 25) → `Markup_Staffing` (90%)
  - Catering/Additional Equipment (rows 26, 28) → `Markup_AV` (65%)
  - Rental Equipment (row 27) → `Markup_Decor` (85%)
  - Venue Fees (rows 30–32) → `Markup_Venues` (60%)
  - QC Staff (rows 35–37) → `Markup_Staffing` (90%)
- Added Margin Health Indicator (✓ STRONG / → ON TARGET / ⚠ REVIEW / ✗ BELOW FLOOR)
- Added Estimated OpEx row (hours × $90/hr)
- Added True Net Profit and True Net Margin % calculations
- Added True Net Health Indicator (✓ STRONG / → ON TARGET / ⚠ THIN / ✗ LOSING MONEY)
- Updated all cross-sheet references for shifted Client Setup rows

#### Decor Estimate
- Replaced all flat-markup formulas with category-specific named range references:
  - Floral Product, Seating, Lounge, Tables, Rugs/Accessories (various rows) → `Markup_Decor` (85%)
  - Floral Fees, Rental Fees (rows 24–26, 66–71) → `Markup_Delivery` (85%)
  - AV Equipment (rows 77–86) → `Markup_AV` (65%)
  - AV Fees & Labor (rows 88–92) → `Markup_Staffing` (90%)
- Added same margin analysis rows as Venue (Health Indicator, OpEx, True Net)
- Updated all cross-sheet references

#### SAMPLE Decor Estimate
- Applied same category-based formula updates as Decor Estimate
- All sample data preserved for team reference
- Added margin analysis rows

## [1.0.0] — 2026-02-21

### Established
- Full category-based markup schedule (11 categories)
- Hours-based OpEx allocation model at $90/hr internal rate
- Three commission scenarios (Direct, GDP, GDP + Client)
- Program P&L evaluation framework with margin targets
- Color-coded margin health indicators
- Decision rules for management fees, minimum fees, and walk-away triggers
- Standard rate cards for tours, staffing, transportation
- Fee defaults (CC processing, service charges, gratuities)

### Config Files Created
- `config/markups.json` — Category markup schedule
- `config/rates.json` — Standard rates and OpEx tiers
- `config/fees.json` — Fee defaults and thresholds
- `config/commissions.json` — Commission scenarios
- `config/thresholds.json` — Margin targets and health indicators
