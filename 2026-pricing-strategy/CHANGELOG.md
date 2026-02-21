# Changelog

All notable pricing decisions and strategy updates for QC 2026.

## [1.2.0] — 2026-02-21

### Travel Expense Calculator (`QC_Estimate_Template_2026_v2.xlsx`)

#### Client Setup
- Added **QC Travel Rate Table** (rows 52–63) — 10 destinations with drive mileage, flight cost, hotel/night, per diem/day rates
- Added **Vehicle Rate Table** (rows 65–72) — None, SUV ($450), Sprinter ($850), Mini Bus ($1,200), Motor Coach ($1,800), Custom
- Created named ranges: `TravelRates`, `DestinationList`, `VehicleRates`, `VehicleList`
- All rates are editable by the team

#### Venue Estimate
- Added **QC Travel Expenses** section (rows 65–81) between Client Pricing and Profit & Margin
- Supports up to 3 staff trips per program with per-trip inputs: destination, travel type, staff count, nights, on-site vehicle, custom vehicle cost
- Travel Type auto-suggests Drive/Flight based on destination (NYC → Flight, all others → Drive) with manual override dropdown
- Calculated fields: Travel Cost, Hotel Cost, Per Diem Cost, Vehicle Cost, Trip Total
- Total Travel Expenses feeds into Profit & Margin Analysis
- **True Net Profit** now = QC Revenue − OpEx − Travel Expenses
- Rewrote all Profit & Margin formulas (rows 82–93) after 18-row insert

#### Decor Estimate
- Added same **QC Travel Expenses** section (rows 110–126)
- Same 3-trip structure, formulas, and dropdowns as Venue
- **True Net Profit** updated to deduct travel expenses
- Rewrote all Profit & Margin formulas (rows 127–138)

#### SAMPLE Decor Estimate
- Applied all Decor travel changes — formulas active, fields empty

## [1.1.0] — 2026-02-21

### Estimate Template Rebuilt (`QC_Estimate_Template_2026_v2.xlsx`)

#### Client Setup
- Replaced single flat markup (C19 = 54%) with 11-category markup table (rows 22–34)
- Each markup cell has a data validation dropdown (5% increments: 50%–200%)
- Added "Absolute Floor: 54%" reference row
- Created `CategoryTable` and `CategoryList` named ranges for VLOOKUP-based formulas
- Added new fields: **Company Name** (B11), **Program Name** (B12), **Client Hotel** (B13)
- Formatted Event Time (B10) as h:mm AM/PM
- Ensured dropdowns on all fee overrides (Service Charge, Gratuity, Admin, GDP, Commission)
- Updated `LocationList` named range and tax lookup VLOOKUPs for shifted rows

#### Venue Estimate
- Added **Category column (G)** with dropdown on every line item — auto-guesses category by section
- Client Cost formulas use `VLOOKUP(G{row}, CategoryTable, 2, FALSE)` — team can override category per row
- NA Beverages (row 26) auto-populates quantity from guest count
- Default qty = 1 pre-filled on equipment, venue, and staffing rows
- QC Staff auto-calculates: 0–50 guests = 1, 51–100 = 2, 101+ = 3 (with override dropdown)
- QC Staff unit price = $500
- Added Company Name, Program Name, Client Hotel propagated from Client Setup
- Added Margin Health Indicator, Estimated OpEx, True Net Profit, True Net Margin %, True Net Health
- Updated all summary formulas, client-ready table, and cross-sheet references

#### Decor Estimate
- Added **Category column (I)** with dropdown on every line item — auto-guesses by section
- Client Cost formulas use VLOOKUP against CategoryTable (same pattern as Venue)
- Added Company Name, Program Name, Client Hotel from Client Setup
- Added margin analysis rows (Health, OpEx, True Net Profit/Margin/Health)
- Updated CC Processing, Commission, GDP, and tax rate cross-sheet references

#### SAMPLE Decor Estimate
- Applied all Decor Estimate changes — VLOOKUP formulas, category dropdowns, margin analysis
- All sample data preserved (florals, furniture, AV, delivery fees)

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
