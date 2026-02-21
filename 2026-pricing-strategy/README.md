# QC 2026 Pricing Strategy — Project Repository

**Quill Creative Event Design**
Internal pricing framework, tools, and templates for 2026 program estimation.

---

## Purpose

This repo contains Quill Creative's pricing infrastructure for 2026 — the strategy document, markup schedules, calculation tools, and templates that power accurate, profitable program estimates across all markets (Charlotte, DC, Philadelphia, Mid-Atlantic).

It is also structured as a **working context for Claude** to quickly build estimates, evaluate program profitability, and generate client-ready deliverables.

---

## Repository Structure

```
2026-pricing-strategy/
│
├── docs/
│   ├── strategy/
│   │   └── QC_2026_Pricing_Strategy.md      # Full strategy (markdown version)
│   ├── templates/
│   │   └── (estimate templates, P&L templates)
│   └── references/
│       └── (rate cards, fee schedules, vendor benchmarks)
│
├── tools/
│   ├── calculators/
│   │   └── (margin calculator, OpEx estimator, P&L evaluator)
│   └── generators/
│       └── (estimate generators, client-ready output builders)
│
├── config/
│   ├── markups.json                          # Category markup schedule
│   ├── rates.json                            # Standard rates (tours, transport, staffing)
│   ├── fees.json                             # Fee defaults (CC, service charge, gratuity)
│   ├── commissions.json                      # Commission scenarios
│   └── thresholds.json                       # Margin health indicators & targets
│
├── examples/
│   └── (sample program P&Ls, estimate walkthroughs)
│
├── CONTEXT.md                                # Quick-reference brief for Claude
├── CHANGELOG.md                              # Version history of pricing decisions
└── README.md                                 # This file
```

---

## Key Principles

1. **Category-based markups** — not flat rates. Mark up higher where clients lack price benchmarks.
2. **Hours-based OpEx** — estimate real team hours at $90/hr, not a blanket percentage.
3. **Commission-aware pricing** — margin floors shift based on commission load (0%, 6.5%, 21.5%).
4. **P&L evaluation on every estimate** — True Net Margin must be positive before sending to client.
5. **Absolute markup floor: 54%** — no exceptions without owner approval.

---

## Quick Start for Claude

1. Read `CONTEXT.md` for the condensed pricing brief.
2. Reference `config/` JSON files for all current rates, markups, and thresholds.
3. Use `docs/strategy/QC_2026_Pricing_Strategy.md` for full rationale and decision rules.

---

*Quill Creative Event Design · 2025 DMC of the Year · quillcreative.co*
