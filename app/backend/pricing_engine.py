"""
QC Pricing Engine — replaces all Excel formula logic with auditable Python.

This module contains the core calculation logic for:
- Line item cost calculation (vendor cost -> client cost via category markup)
- Section subtotals (taxable vs non-taxable)
- Tax calculation
- Fee calculation (CC processing, service charge, admin fee)
- Commission deductions
- P&L evaluation (QC Margin, True Net Profit/Margin)
- Margin health indicators
"""

from dataclasses import dataclass, field
from enum import Enum
from . import config_loader


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

class LineItemType(str, Enum):
    QUANTITY = "quantity"   # cost = qty * unit_price
    FLAT_FEE = "flat_fee"  # cost = unit_price (no qty multiplier)
    PER_PERSON = "per_person"  # cost = guest_count * unit_price


@dataclass
class LineItem:
    name: str
    category_key: str          # maps to markups.json key
    item_type: LineItemType
    unit_price: float = 0.0
    quantity: int = 0
    is_taxable: bool = True
    notes: str = ""
    section: str = ""          # grouping label (e.g. "Florals", "Seating")

    @property
    def vendor_cost(self) -> float:
        if self.item_type == LineItemType.FLAT_FEE:
            return self.unit_price
        return self.quantity * self.unit_price

    @property
    def markup_pct(self) -> float:
        return config_loader.get_markup_for_category(self.category_key)

    @property
    def client_cost(self) -> float:
        return round(self.vendor_cost * (1 + self.markup_pct), 2)

    @property
    def markup_amount(self) -> float:
        return round(self.client_cost - self.vendor_cost, 2)

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "category_key": self.category_key,
            "category_label": config_loader.get_markups()["categories"].get(
                self.category_key, {}
            ).get("label", self.category_key),
            "item_type": self.item_type.value,
            "unit_price": self.unit_price,
            "quantity": self.quantity,
            "is_taxable": self.is_taxable,
            "section": self.section,
            "notes": self.notes,
            "vendor_cost": self.vendor_cost,
            "markup_pct": self.markup_pct,
            "client_cost": self.client_cost,
            "markup_amount": self.markup_amount,
        }


@dataclass
class EstimateInput:
    """All inputs needed to calculate an estimate."""
    estimate_name: str
    estimate_type: str                     # "venue", "decor", "tour", etc.
    client_name: str = ""
    event_name: str = ""
    event_date: str = ""
    event_time: str = ""
    guest_count: int = 0
    location: str = ""
    tax_rate: float = 0.0                  # e.g. 0.07 for 7%
    commission_scenario: str = "direct_client"
    cc_processing_pct: float = 0.035
    gratuity_pct: float = 0.20
    admin_fee_pct: float = 0.05
    gdp_enabled: bool = False
    opex_hours: float = 0.0
    line_items: list[LineItem] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Calculation engine
# ---------------------------------------------------------------------------

def calculate_estimate(inp: EstimateInput) -> dict:
    """
    Run the full P&L calculation for an estimate.
    Returns a dict with all calculated values, line item details, and health indicators.
    """
    rates = config_loader.get_rates()
    commissions = config_loader.get_commissions()
    thresholds = config_loader.get_thresholds()

    # -- Line item calculations --
    items = [li.to_dict() for li in inp.line_items]

    # -- Subtotals --
    taxable_vendor = sum(li.vendor_cost for li in inp.line_items if li.is_taxable)
    taxable_client = sum(li.client_cost for li in inp.line_items if li.is_taxable)
    nontaxable_vendor = sum(li.vendor_cost for li in inp.line_items if not li.is_taxable)
    nontaxable_client = sum(li.client_cost for li in inp.line_items if not li.is_taxable)

    total_vendor = taxable_vendor + nontaxable_vendor
    total_client_pretax = taxable_client + nontaxable_client

    # -- Tax --
    sales_tax = round(taxable_client * inp.tax_rate, 2)

    # -- Subtotal after tax --
    subtotal_after_tax = total_client_pretax + sales_tax

    # -- Fees --
    cc_fee = round(subtotal_after_tax * inp.cc_processing_pct, 2)

    # -- Commission --
    scenario = commissions["scenarios"].get(inp.commission_scenario, commissions["scenarios"]["direct_client"])
    commission_pct = scenario["commission_pct"] / 100
    commission_amount = round(subtotal_after_tax * commission_pct, 2)

    # -- Client invoice total --
    client_invoice_total = round(subtotal_after_tax + cc_fee, 2)

    # -- P&L --
    gdp_fee = round(client_invoice_total * 0.065, 2) if inp.gdp_enabled else 0.0
    qc_gross_revenue = round(client_invoice_total - commission_amount - gdp_fee - total_vendor - sales_tax, 2)
    qc_margin_pct = round(qc_gross_revenue / client_invoice_total * 100, 2) if client_invoice_total > 0 else 0.0

    # -- OpEx --
    opex_rate = rates["internal_hourly_rate"]
    opex_cost = round(inp.opex_hours * opex_rate, 2)

    # -- True Net --
    true_net_profit = round(qc_gross_revenue - opex_cost, 2)
    true_net_margin_pct = round(true_net_profit / client_invoice_total * 100, 2) if client_invoice_total > 0 else 0.0

    # -- Health indicators --
    qc_health = _evaluate_health(qc_margin_pct, thresholds["health_indicators"]["qc_gross_margin"])
    net_health = _evaluate_health(true_net_margin_pct, thresholds["health_indicators"]["true_net_margin"])

    # -- Section subtotals --
    sections = {}
    for li in inp.line_items:
        sec = li.section or "Other"
        if sec not in sections:
            sections[sec] = {"vendor_cost": 0, "client_cost": 0, "items": []}
        sections[sec]["vendor_cost"] += li.vendor_cost
        sections[sec]["client_cost"] += li.client_cost
        sections[sec]["items"].append(li.to_dict())

    # -- Client-ready summary (rounded to nearest $50) --
    client_ready_sections = {}
    for sec_name, sec_data in sections.items():
        client_ready_sections[sec_name] = _round_up(sec_data["client_cost"], 50)
    client_ready_tax = _round_up(sales_tax, 50)
    client_ready_total = sum(client_ready_sections.values()) + client_ready_tax

    return {
        "estimate_name": inp.estimate_name,
        "estimate_type": inp.estimate_type,
        "client_name": inp.client_name,
        "event_name": inp.event_name,
        "event_date": inp.event_date,
        "event_time": inp.event_time,
        "guest_count": inp.guest_count,
        "location": inp.location,
        "commission_scenario": {
            "key": inp.commission_scenario,
            "label": scenario["label"],
            "commission_pct": scenario["commission_pct"],
        },
        "line_items": items,
        "sections": sections,
        "subtotals": {
            "taxable_vendor": taxable_vendor,
            "taxable_client": taxable_client,
            "nontaxable_vendor": nontaxable_vendor,
            "nontaxable_client": nontaxable_client,
            "total_vendor": total_vendor,
            "total_client_pretax": total_client_pretax,
            "sales_tax": sales_tax,
            "cc_fee": cc_fee,
            "commission_amount": commission_amount,
            "client_invoice_total": client_invoice_total,
        },
        "pnl": {
            "client_invoice_total": client_invoice_total,
            "commission_amount": commission_amount,
            "gdp_fee": gdp_fee,
            "total_vendor_costs": total_vendor + sales_tax,
            "qc_gross_revenue": qc_gross_revenue,
            "qc_margin_pct": qc_margin_pct,
            "qc_margin_health": qc_health,
            "opex_hours": inp.opex_hours,
            "opex_rate": opex_rate,
            "opex_cost": opex_cost,
            "true_net_profit": true_net_profit,
            "true_net_margin_pct": true_net_margin_pct,
            "true_net_health": net_health,
        },
        "client_ready": {
            "sections": client_ready_sections,
            "tax": client_ready_tax,
            "total": client_ready_total,
        },
    }


def _evaluate_health(margin_pct: float, thresholds: dict) -> dict:
    """Evaluate margin health against threshold config."""
    # thresholds are ordered from highest to lowest
    ordered = sorted(thresholds.items(), key=lambda x: x[1]["threshold_pct"], reverse=True)
    for level, info in ordered:
        if margin_pct >= info["threshold_pct"]:
            return {"level": level, "label": info["label"], "meaning": info["meaning"]}
    # fallback to lowest
    last_level, last_info = ordered[-1]
    return {"level": last_level, "label": last_info["label"], "meaning": last_info["meaning"]}


def _round_up(value: float, nearest: int) -> float:
    """Round up to nearest increment (e.g. nearest $50)."""
    import math
    if value <= 0:
        return 0
    return math.ceil(value / nearest) * nearest


# ---------------------------------------------------------------------------
# Estimate type templates — defines default sections/line items per type
# ---------------------------------------------------------------------------

ESTIMATE_TEMPLATES = {
    "venue": {
        "label": "Venue Estimate",
        "sections": [
            {
                "name": "Food & Beverage",
                "category_key": "catering_fnb",
                "is_taxable": True,
                "default_items": [
                    {"name": "Per Person Food", "item_type": "per_person"},
                    {"name": "Bar Package", "item_type": "per_person"},
                    {"name": "Non-Alcoholic Beverages", "item_type": "per_person"},
                ],
            },
            {
                "name": "Staffing",
                "category_key": "staffing_labor",
                "is_taxable": True,
                "default_items": [
                    {"name": "Event Staff", "item_type": "quantity"},
                ],
            },
            {
                "name": "Equipment & Rentals",
                "category_key": "av_production",
                "is_taxable": True,
                "default_items": [
                    {"name": "Catering Equipment", "item_type": "quantity"},
                    {"name": "Rental Equipment", "item_type": "quantity"},
                    {"name": "Additional Equipment", "item_type": "quantity"},
                ],
            },
            {
                "name": "Venue Fees",
                "category_key": "venues_room_rentals",
                "is_taxable": True,
                "default_items": [
                    {"name": "Venue / Room Rental", "item_type": "flat_fee"},
                    {"name": "Additional Venue Space", "item_type": "flat_fee"},
                    {"name": "Additional Services", "item_type": "flat_fee"},
                ],
            },
            {
                "name": "QC Staffing (Non-Taxable)",
                "category_key": "staffing_labor",
                "is_taxable": False,
                "default_items": [
                    {"name": "QC On-Site Lead", "item_type": "flat_fee"},
                    {"name": "QC Event Coordinator", "item_type": "flat_fee"},
                    {"name": "Additional QC Staff", "item_type": "quantity"},
                ],
            },
        ],
    },
    "decor": {
        "label": "Decor Estimate",
        "sections": [
            {
                "name": "Florals (Taxable)",
                "category_key": "decor_design",
                "is_taxable": True,
                "default_items": [
                    {"name": "Centerpieces", "item_type": "quantity"},
                    {"name": "Arrangement 2", "item_type": "quantity"},
                    {"name": "Arrangement 3", "item_type": "quantity"},
                    {"name": "Arrangement 4", "item_type": "quantity"},
                    {"name": "Arrangement 5", "item_type": "quantity"},
                    {"name": "Arrangement 6", "item_type": "quantity"},
                    {"name": "Arrangement 7", "item_type": "quantity"},
                    {"name": "Arrangement 8", "item_type": "quantity"},
                    {"name": "Arrangement 9", "item_type": "quantity"},
                    {"name": "Arrangement 10", "item_type": "quantity"},
                ],
            },
            {
                "name": "Floral Fees (Non-Taxable)",
                "category_key": "delivery_logistics",
                "is_taxable": False,
                "default_items": [
                    {"name": "Delivery Fee", "item_type": "flat_fee"},
                    {"name": "Setup Fee", "item_type": "flat_fee"},
                    {"name": "Breakdown Fee", "item_type": "flat_fee"},
                ],
            },
            {
                "name": "Seating",
                "category_key": "decor_design",
                "is_taxable": True,
                "default_items": [
                    {"name": "Chairs", "item_type": "quantity"},
                    {"name": "Barstools", "item_type": "quantity"},
                    {"name": "Banquettes", "item_type": "quantity"},
                ],
            },
            {
                "name": "Lounge",
                "category_key": "decor_design",
                "is_taxable": True,
                "default_items": [
                    {"name": "Sofas", "item_type": "quantity"},
                    {"name": "Club Chairs", "item_type": "quantity"},
                    {"name": "Sectionals", "item_type": "quantity"},
                ],
            },
            {
                "name": "Tables",
                "category_key": "decor_design",
                "is_taxable": True,
                "default_items": [
                    {"name": "Cocktail Tables", "item_type": "quantity"},
                    {"name": "Coffee Tables", "item_type": "quantity"},
                    {"name": "Side Tables", "item_type": "quantity"},
                    {"name": "Dining Tables", "item_type": "quantity"},
                ],
            },
            {
                "name": "Rugs, Decor & Accessories",
                "category_key": "decor_design",
                "is_taxable": True,
                "default_items": [
                    {"name": "Rugs", "item_type": "quantity"},
                    {"name": "Decor Item 2", "item_type": "quantity"},
                    {"name": "Decor Item 3", "item_type": "quantity"},
                ],
            },
            {
                "name": "Rental Fees (Non-Taxable)",
                "category_key": "delivery_logistics",
                "is_taxable": False,
                "default_items": [
                    {"name": "Delivery Fee", "item_type": "flat_fee"},
                    {"name": "Setup Fee", "item_type": "flat_fee"},
                    {"name": "Breakdown Fee", "item_type": "flat_fee"},
                ],
            },
            {
                "name": "AV & Production (Taxable)",
                "category_key": "av_production",
                "is_taxable": True,
                "default_items": [
                    {"name": "AV Equipment 1", "item_type": "quantity"},
                    {"name": "AV Equipment 2", "item_type": "quantity"},
                    {"name": "AV Equipment 3", "item_type": "quantity"},
                ],
            },
            {
                "name": "AV Fees & Labor (Non-Taxable)",
                "category_key": "staffing_labor",
                "is_taxable": False,
                "default_items": [
                    {"name": "AV Technician", "item_type": "flat_fee"},
                    {"name": "Setup Labor", "item_type": "flat_fee"},
                ],
            },
        ],
    },
    "transportation": {
        "label": "Transportation Estimate",
        "sections": [
            {
                "name": "Vehicles",
                "category_key": "transportation",
                "is_taxable": True,
                "default_items": [
                    {"name": "SUV (6 pax)", "item_type": "quantity"},
                    {"name": "Sprinter Van (12-14 pax)", "item_type": "quantity"},
                    {"name": "Mini Bus (24-30 pax)", "item_type": "quantity"},
                    {"name": "Motor Coach (50-56 pax)", "item_type": "quantity"},
                ],
            },
            {
                "name": "Logistics Fees (Non-Taxable)",
                "category_key": "delivery_logistics",
                "is_taxable": False,
                "default_items": [
                    {"name": "Route Planning", "item_type": "flat_fee"},
                    {"name": "Coordination Fee", "item_type": "flat_fee"},
                ],
            },
        ],
    },
    "entertainment": {
        "label": "Entertainment Estimate",
        "sections": [
            {
                "name": "Talent & Performance",
                "category_key": "entertainment",
                "is_taxable": True,
                "default_items": [
                    {"name": "Performer / Act", "item_type": "flat_fee"},
                    {"name": "DJ / MC", "item_type": "flat_fee"},
                    {"name": "Band / Ensemble", "item_type": "flat_fee"},
                ],
            },
            {
                "name": "Production Support",
                "category_key": "av_production",
                "is_taxable": True,
                "default_items": [
                    {"name": "Sound System", "item_type": "flat_fee"},
                    {"name": "Lighting", "item_type": "flat_fee"},
                    {"name": "Stage / Riser", "item_type": "flat_fee"},
                ],
            },
        ],
    },
    "tour": {
        "label": "Tour & Experience Estimate",
        "sections": [
            {
                "name": "Tour Services",
                "category_key": "tours_guided_experiences",
                "is_taxable": True,
                "default_items": [
                    {"name": "Licensed Tour Guide", "item_type": "quantity"},
                    {"name": "Experience / Activity", "item_type": "per_person"},
                ],
            },
            {
                "name": "Staffing",
                "category_key": "staffing_labor",
                "is_taxable": False,
                "default_items": [
                    {"name": "QC Event Coordinator", "item_type": "flat_fee"},
                    {"name": "Photographer", "item_type": "flat_fee"},
                ],
            },
            {
                "name": "Transportation",
                "category_key": "transportation",
                "is_taxable": True,
                "default_items": [
                    {"name": "Vehicle", "item_type": "quantity"},
                ],
            },
        ],
    },
}


def get_estimate_template(estimate_type: str) -> dict:
    """Return the template definition for an estimate type."""
    if estimate_type not in ESTIMATE_TEMPLATES:
        raise ValueError(f"Unknown estimate type: {estimate_type}. Available: {list(ESTIMATE_TEMPLATES.keys())}")
    return ESTIMATE_TEMPLATES[estimate_type]


def get_available_estimate_types() -> list[dict]:
    """Return list of available estimate types."""
    return [{"key": k, "label": v["label"]} for k, v in ESTIMATE_TEMPLATES.items()]
