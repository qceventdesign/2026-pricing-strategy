"""Load and provide access to QC pricing configuration files."""

import json
from pathlib import Path
from functools import lru_cache

CONFIG_DIR = Path(__file__).resolve().parent.parent.parent / "2026-pricing-strategy" / "config"


@lru_cache()
def _load(filename: str) -> dict:
    with open(CONFIG_DIR / filename) as f:
        return json.load(f)


def get_markups() -> dict:
    return _load("markups.json")


def get_rates() -> dict:
    return _load("rates.json")


def get_fees() -> dict:
    return _load("fees.json")


def get_commissions() -> dict:
    return _load("commissions.json")


def get_thresholds() -> dict:
    return _load("thresholds.json")


def get_all_config() -> dict:
    return {
        "markups": get_markups(),
        "rates": get_rates(),
        "fees": get_fees(),
        "commissions": get_commissions(),
        "thresholds": get_thresholds(),
    }


def get_markup_for_category(category_key: str) -> float:
    """Return the markup percentage (as decimal, e.g. 0.85) for a category key."""
    categories = get_markups()["categories"]
    if category_key not in categories:
        raise ValueError(f"Unknown category: {category_key}")
    return categories[category_key]["markup_pct"] / 100


def get_category_list() -> list[dict]:
    """Return list of {key, label, markup_pct, rationale} for all categories."""
    cats = get_markups()["categories"]
    return [
        {"key": k, "label": v["label"], "markup_pct": v["markup_pct"], "rationale": v["rationale"]}
        for k, v in cats.items()
    ]
