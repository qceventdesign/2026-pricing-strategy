"""Load pricing configuration from JSON files."""

import json
from pathlib import Path
from functools import lru_cache

CONFIG_DIR = Path(__file__).resolve().parent.parent.parent / "2026-pricing-strategy" / "config"


@lru_cache(maxsize=1)
def load_all_config() -> dict:
    """Load all config files into a single dict."""
    config = {}
    for name in ("markups", "rates", "fees", "commissions", "thresholds"):
        path = CONFIG_DIR / f"{name}.json"
        with open(path) as f:
            config[name] = json.load(f)
    return config


def get_markups() -> dict:
    return load_all_config()["markups"]


def get_rates() -> dict:
    return load_all_config()["rates"]


def get_fees() -> dict:
    return load_all_config()["fees"]


def get_commissions() -> dict:
    return load_all_config()["commissions"]


def get_thresholds() -> dict:
    return load_all_config()["thresholds"]


def reload_config():
    """Clear cache and reload config files."""
    load_all_config.cache_clear()
    return load_all_config()
