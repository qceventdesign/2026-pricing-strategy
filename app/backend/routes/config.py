"""Config and reference data routes."""

from fastapi import APIRouter
from .. import config_loader

router = APIRouter(prefix="/api/config", tags=["config"])


@router.get("/")
def get_all_config():
    """Return all pricing configuration."""
    return config_loader.get_all_config()


@router.get("/markups")
def get_markups():
    """Return category markup schedule."""
    return config_loader.get_markups()


@router.get("/markups/categories")
def get_categories():
    """Return flat list of categories with markup info."""
    return config_loader.get_category_list()


@router.get("/rates")
def get_rates():
    """Return staffing and transportation rates."""
    return config_loader.get_rates()


@router.get("/fees")
def get_fees():
    """Return fee defaults."""
    return config_loader.get_fees()


@router.get("/commissions")
def get_commissions():
    """Return commission scenarios."""
    return config_loader.get_commissions()


@router.get("/thresholds")
def get_thresholds():
    """Return margin targets and health indicators."""
    return config_loader.get_thresholds()
