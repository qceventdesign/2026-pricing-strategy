# Claude Code Prompt — Rebuild QC Estimate Template with Category-Based Pricing

## Project Context

You are working inside the `2026-pricing-strategy` GitHub repo for **Quill Creative Event Design**, a boutique DMC specializing in corporate events for Fortune 500 companies.

**Start by reading these files:**
- `CONTEXT.md` — Full pricing framework (markups, rates, commissions, margin targets, decision rules)
- `config/markups.json` — Category-based markup schedule (11 categories with specific percentages)
- `config/rates.json` — Standard rates for staffing, transportation, OpEx tiers
- `config/fees.json` — Fee defaults (CC processing, service charges, gratuities, management fees)
- `config/commissions.json` — Three commission scenarios (Direct, GDP, GDP + Client)
- `config/thresholds.json` — Margin targets and health indicator thresholds

**Current template to rebuild:** `docs/templates/QC_Estimate_Template_2026.xlsx`
(Copy the uploaded file to this location first.)

---

## What Needs to Change

The current template uses a **single flat markup percentage** (cell C19 on Client Setup, currently 54%) applied to every line item identically. The 2026 pricing strategy requires **category-based markups** — each vendor category has its own markup percentage based on how transparent that cost is to the client.

The template needs to be reworked so the team sets all category markups once on the Client Setup tab, and every estimate tab automatically pulls the correct markup per line item category.

---

## Detailed Rebuild Specification

### SHEET 1: Client Setup

**Keep everything that currently exists** (event details, location/tax, restaurant fee defaults, how-to-use, tax lookup table rows 59–81). Modify the PRICING & COMMISSION section as follows:

**Remove:**
- Row 19: Single "Markup Percentage" field (C19 = 0.54)

**Replace with: CATEGORY MARKUP SCHEDULE (starting at row 19)**

Build a category markup table with these columns:
- Column A: Category label
- Column B: Markup percentage (editable input — blue font per financial model conventions)  
- Column C: Brief rationale/note

Pre-populate with these values from `config/markups.json`:

| Category | Markup | Note |
|---|---|---|
| Catering & F&B | 54% | Floor — clients compare per-head pricing |
| Venues & Room Rentals | 60% | Full venue relationship management |
| AV & Production | 65% | Technical coordination, hard to price-shop |
| Décor & Design | 85% | Custom creative — no comparison exists |
| Entertainment | 75% | Talent sourcing, contracts, riders |
| Activities & Experiences | 75% | Curation, coordination, permitting |
| Transportation | 75% | Route planning, high-liability |
| Staffing & Labor | 90% | Recruitment, scheduling, management |
| Purchased / Sourced Items | 200% | Small items, disproportionate time (3×) |
| Delivery & Logistics | 85% | Timing risk, damage liability |
| Tours & Guided Experiences | 65% | Guide sourcing, route design |

**Add a row below the table:**
- "Absolute Floor: 54%" with a note: "No line item below this without owner approval"

**Add a named range or clear cell references** so every estimate tab can look up the correct markup by category. Suggested approach: use named ranges like `Markup_Catering`, `Markup_Venues`, `Markup_AV`, `Markup_Decor`, etc. — OR use a VLOOKUP/INDEX-MATCH against the category table.

**Shift down the remaining sections** (CC Processing Fee, Third-Party Commission, GDP toggle, Restaurant Fee Defaults, How to Use, Tax Table) to accommodate the new rows. **Update all cross-sheet references** that previously pointed to specific row numbers.

---

### SHEET 2: Venue Estimate

**Current behavior:** Every line item uses `'Client Setup'!C19` (single flat markup) in the Client Cost formula:
```
=IFERROR(D21*(1+'Client Setup'!C19),0)
```

**New behavior:** Each line item section pulls its **category-specific markup** from the Client Setup category table.

Map the existing sections to categories:

| Venue Estimate Section | Rows | Pricing Category | Markup |
|---|---|---|---|
| FOOD & BEVERAGE (Per Person Food, Bar Package, NA Beverages) | 21–23 | Catering & F&B | 54% |
| TAXABLE EQUIPMENT, RENTALS & STAFFING (Staffing, Catering Equipment, Rental Equipment, Additional Equipment) | 25–28 | Split — see below | Varies |
| — Staffing (row 25) | 25 | Staffing & Labor | 90% |
| — Catering Equipment (row 26) | 26 | AV & Production | 65% |
| — Rental Equipment (row 27) | 27 | Décor & Design | 85% |
| — Additional Equipment (row 28) | 28 | AV & Production | 65% |
| VENUE FEES (Venue/Room Rental, Additional Venue Space, Additional Services) | 30–32 | Venues & Room Rentals | 60% |
| NON-TAXABLE QC STAFFING (QC Staff rows) | 35–37 | Staffing & Labor | 90% |

**Update each Client Cost formula** to reference the correct category markup instead of the flat rate. Example for F&B:
```
=IFERROR(D21*(1+'Client Setup'!$B$[Catering_row]),0)
```

**Add a "Category" column (or use the existing Notes column G)** so each line item's category assignment is visible and auditable.

**Update the Profit & Margin Analysis section (rows 61–66):**
- Keep: Third-Party Commission, GDP Fee, Total Vendor Costs, QC Revenue, QC Margin %
- **Add: Margin Health Indicator** — use the thresholds from `config/thresholds.json`:
  - ✓ STRONG (≥45%), → ON TARGET (≥35%), ⚠ REVIEW (≥25%), ✗ BELOW FLOOR (<25%)
  - Use conditional formatting or an IF formula to display the indicator
- **Add: Estimated OpEx row** — input field for estimated team hours, multiplied by $90/hr
- **Add: True Net Profit** = QC Revenue − OpEx
- **Add: True Net Margin %** = True Net Profit / Client Invoice Total
- **Add: True Net Health Indicator** — ✓ STRONG (≥15%), → ON TARGET (≥7%), ⚠ THIN (≥0%), ✗ LOSING MONEY (<0%)

**Update the Client-Ready Table (columns I–K):** No changes needed to the structure — it already rounds and aggregates. Just verify formulas still work after row shifts.

---

### SHEET 3: Decor Estimate

**Same pattern as Venue.** Current behavior uses `'Client Setup'!C19` for all line items.

Map sections to categories:

| Decor Estimate Section | Rows | Pricing Category | Markup |
|---|---|---|---|
| TAXABLE FLORAL PRODUCT | 13–22 | Décor & Design | 85% |
| NON-TAXABLE FLORAL FEES (Delivery/Setup) | 24–26 | Delivery & Logistics | 85% |
| SEATING (chairs, barstools, banquettes) | 32–39 | Décor & Design | 85% |
| LOUNGE (sofas, sectionals, club chairs) | 41–48 | Décor & Design | 85% |
| TABLES (cocktail, coffee, side, dining) | 50–57 | Décor & Design | 85% |
| RUGS, DECOR & ACCESSORIES | 59–64 | Décor & Design | 85% |
| NON-TAXABLE RENTAL FEES | 66–71 | Delivery & Logistics | 85% |
| TAXABLE AV & PRODUCTION EQUIPMENT | 77–86 | AV & Production | 65% |
| NON-TAXABLE AV FEES & LABOR | 88–92 | Staffing & Labor | 90% |

**Update each formula** from `'Client Setup'!C19` to the correct category markup reference.

**Update the Profit & Margin section (rows 106–111):**
- Same additions as Venue: Health indicators, OpEx, True Net Profit, True Net Margin

---

### SHEET 4: SAMPLE Decor Estimate

Update to use the new category-based formulas (same as Decor Estimate). Keep the sample data intact so the team can see the new markup structure in action with real numbers.

---

## Technical Requirements

1. **All markups must be Excel formulas referencing the Client Setup category table** — never hardcode markup values in estimate sheets.
2. **Use Excel formulas, not Python calculations** — the spreadsheet must be dynamic and recalculate when inputs change.
3. **Preserve all existing formatting, colors, and layout conventions** from the original template.
4. **Yellow-highlighted cells** = user input fields. Keep this convention.
5. **Blue font** for editable assumption cells (markup percentages on Client Setup).
6. **Conditional formatting** for margin health indicators (green = STRONG, blue/gray = ON TARGET, orange = REVIEW, red = BELOW FLOOR/LOSING MONEY).
7. **Data validation dropdowns** where appropriate (commission scenario, location, fee overrides).
8. **Named ranges preferred** for category markups to keep formulas readable.
9. **Zero formula errors** — validate with recalc after building.
10. **Update CHANGELOG.md** with what was changed.

## File Locations

- **Input template:** `docs/templates/QC_Estimate_Template_2026.xlsx`
- **Output:** Overwrite in place (same file) or save as `docs/templates/QC_Estimate_Template_2026_v2.xlsx`
- **Config reference:** `config/` directory (markups.json, rates.json, fees.json, commissions.json, thresholds.json)

## Validation Checklist

After building, verify:
- [ ] Each estimate tab line item pulls the correct category markup (not the old flat C19)
- [ ] Changing a markup on Client Setup instantly updates all affected line items across all tabs
- [ ] Margin health indicators display correctly at each threshold
- [ ] True Net Profit/Margin calculates correctly with OpEx hours input
- [ ] Client-Ready Tables still generate correct rounded totals
- [ ] Tax lookups still work after row shifts
- [ ] No #REF!, #DIV/0!, #VALUE!, or #NAME? errors anywhere
- [ ] SAMPLE Decor tab shows realistic margin numbers with the new markup structure
