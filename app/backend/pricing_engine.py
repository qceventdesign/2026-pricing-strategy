"""Core pricing engine — replaces all Excel formula logic."""

from __future__ import annotations
from .config_loader import get_markups, get_fees, get_commissions, get_thresholds, get_rates


# ── Estimate type templates ──────────────────────────────────────────────

ESTIMATE_TEMPLATES: dict[str, dict] = {
    "venue": {
        "label": "Venue Estimate",
        "default_categories": [
            "venues_room_rentals",
            "catering_fnb",
            "av_production",
            "entertainment",
            "staffing_labor",
            "delivery_logistics",
        ],
    },
    "decor": {
        "label": "Decor Estimate",
        "default_categories": [
            "decor_design",
            "purchased_sourced_items",
            "delivery_logistics",
            "staffing_labor",
        ],
    },
    "transportation": {
        "label": "Transportation Estimate",
        "default_categories": [
            "transportation",
            "staffing_labor",
        ],
    },
    "entertainment": {
        "label": "Entertainment Estimate",
        "default_categories": [
            "entertainment",
            "av_production",
            "staffing_labor",
        ],
    },
    "tour": {
        "label": "Tour Estimate",
        "default_categories": [
            "tours_guided_experiences",
            "catering_fnb",
            "transportation",
            "activities_experiences",
            "staffing_labor",
        ],
    },
}


# ── Line item calculation ────────────────────────────────────────────────

def calculate_line_item(
    description: str,
    category_key: str,
    vendor_cost: float,
    quantity: int = 1,
    is_taxable: bool = False,
    tax_rate_pct: float = 0.0,
) -> dict:
    """Calculate a single line item with category-based markup."""
    markups = get_markups()
    cat = markups["categories"].get(category_key)
    if not cat:
        raise ValueError(f"Unknown category: {category_key}")

    markup_pct = cat["markup_pct"]
    vendor_total = vendor_cost * quantity
    markup_amount = vendor_total * (markup_pct / 100)
    client_cost = vendor_total + markup_amount

    tax = client_cost * (tax_rate_pct / 100) if is_taxable else 0.0

    return {
        "description": description,
        "category_key": category_key,
        "category_label": cat["label"],
        "vendor_cost": round(vendor_cost, 2),
        "quantity": quantity,
        "vendor_total": round(vendor_total, 2),
        "markup_pct": markup_pct,
        "markup_amount": round(markup_amount, 2),
        "client_cost": round(client_cost, 2),
        "is_taxable": is_taxable,
        "tax_rate_pct": tax_rate_pct,
        "tax_amount": round(tax, 2),
        "client_total": round(client_cost + tax, 2),
    }


# ── Full estimate P&L ───────────────────────────────────────────────────

def calculate_estimate(
    line_items: list[dict],
    commission_scenario: str = "direct_client",
    opex_hours: float = 0.0,
    management_fee: float = 0.0,
    travel_expenses: float = 0.0,
    tax_rate_pct: float = 0.0,
) -> dict:
    """Run the full P&L calculation for an estimate.

    Returns a dict with all summary financials and health indicators.
    """
    fees_cfg = get_fees()
    commissions_cfg = get_commissions()
    rates_cfg = get_rates()
    thresholds_cfg = get_thresholds()

    scenario = commissions_cfg["scenarios"].get(commission_scenario)
    if not scenario:
        raise ValueError(f"Unknown commission scenario: {commission_scenario}")

    # Calculate every line item
    calculated_items = []
    for item in line_items:
        calc = calculate_line_item(
            description=item["description"],
            category_key=item["category_key"],
            vendor_cost=item["vendor_cost"],
            quantity=item.get("quantity", 1),
            is_taxable=item.get("is_taxable", False),
            tax_rate_pct=tax_rate_pct,
        )
        calculated_items.append(calc)

    # Totals
    total_vendor = sum(i["vendor_total"] for i in calculated_items)
    total_client = sum(i["client_cost"] for i in calculated_items)
    total_tax = sum(i["tax_amount"] for i in calculated_items)

    # Client invoice total (what the client pays)
    client_invoice_subtotal = total_client + management_fee
    client_invoice_total = client_invoice_subtotal + total_tax

    # Commission deduction
    commission_pct = scenario["commission_pct"]
    commission_amount = client_invoice_subtotal * (commission_pct / 100)

    # CC processing fee
    cc_pct = fees_cfg["cc_processing_pct"]
    cc_fee = client_invoice_total * (cc_pct / 100)

    # QC revenue (what QC actually receives)
    qc_gross_revenue = client_invoice_subtotal - commission_amount - total_vendor
    qc_net_after_cc = client_invoice_subtotal - commission_amount - cc_fee - total_vendor

    # OpEx
    internal_rate = rates_cfg["internal_hourly_rate"]
    opex_cost = opex_hours * internal_rate

    # True net profit
    true_net_profit = qc_net_after_cc - opex_cost - travel_expenses

    # Margins
    qc_gross_margin_pct = (qc_gross_revenue / client_invoice_subtotal * 100) if client_invoice_subtotal > 0 else 0
    true_net_margin_pct = (true_net_profit / client_invoice_subtotal * 100) if client_invoice_subtotal > 0 else 0

    # Health indicators
    gross_health = _evaluate_health(qc_gross_margin_pct, thresholds_cfg["health_indicators"]["qc_gross_margin"])
    net_health = _evaluate_health(true_net_margin_pct, thresholds_cfg["health_indicators"]["true_net_margin"])

    return {
        "line_items": calculated_items,
        "summary": {
            "total_vendor_cost": round(total_vendor, 2),
            "total_client_cost": round(total_client, 2),
            "management_fee": round(management_fee, 2),
            "client_invoice_subtotal": round(client_invoice_subtotal, 2),
            "total_tax": round(total_tax, 2),
            "client_invoice_total": round(client_invoice_total, 2),
            "commission_scenario": commission_scenario,
            "commission_label": scenario["label"],
            "commission_pct": commission_pct,
            "commission_amount": round(commission_amount, 2),
            "cc_fee_pct": cc_pct,
            "cc_fee_amount": round(cc_fee, 2),
            "qc_gross_revenue": round(qc_gross_revenue, 2),
            "qc_gross_margin_pct": round(qc_gross_margin_pct, 2),
            "opex_hours": opex_hours,
            "opex_rate": internal_rate,
            "opex_cost": round(opex_cost, 2),
            "travel_expenses": round(travel_expenses, 2),
            "true_net_profit": round(true_net_profit, 2),
            "true_net_margin_pct": round(true_net_margin_pct, 2),
        },
        "health": {
            "qc_gross_margin": gross_health,
            "true_net_margin": net_health,
        },
    }


def _evaluate_health(margin_pct: float, indicator_cfg: dict) -> dict:
    """Determine health label and meaning for a given margin percentage."""
    # Indicators are ordered from highest threshold to lowest
    levels = sorted(indicator_cfg.values(), key=lambda x: x["threshold_pct"], reverse=True)
    for level in levels:
        if margin_pct >= level["threshold_pct"]:
            return {
                "label": level["label"],
                "meaning": level["meaning"],
                "margin_pct": round(margin_pct, 2),
            }
    # Fallback to the lowest level
    last = levels[-1]
    return {
        "label": last["label"],
        "meaning": last["meaning"],
        "margin_pct": round(margin_pct, 2),
    }


def get_category_options() -> list[dict]:
    """Return all markup categories for dropdowns."""
    markups = get_markups()
    return [
        {"key": k, "label": v["label"], "markup_pct": v["markup_pct"], "rationale": v["rationale"]}
        for k, v in markups["categories"].items()
    ]


def get_estimate_templates() -> dict:
    """Return all available estimate type templates."""
    return ESTIMATE_TEMPLATES
