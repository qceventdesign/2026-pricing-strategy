"""Pydantic schemas for API request/response validation."""

from __future__ import annotations
from pydantic import BaseModel, Field
from typing import Optional
from datetime import datetime


class LineItemInput(BaseModel):
    description: str
    category_key: str
    vendor_cost: float = Field(ge=0)
    quantity: int = Field(default=1, ge=1)
    is_taxable: bool = False


class EstimateCreate(BaseModel):
    name: str
    estimate_type: str  # venue, decor, transportation, entertainment, tour
    client_name: str = ""
    event_name: str = ""
    event_date: str = ""
    guest_count: int = Field(default=0, ge=0)
    commission_scenario: str = "direct_client"
    opex_hours: float = Field(default=0.0, ge=0)
    management_fee: float = Field(default=0.0, ge=0)
    travel_expenses: float = Field(default=0.0, ge=0)
    tax_rate_pct: float = Field(default=0.0, ge=0)
    line_items: list[LineItemInput] = []
    notes: str = ""


class EstimateUpdate(BaseModel):
    name: Optional[str] = None
    client_name: Optional[str] = None
    event_name: Optional[str] = None
    event_date: Optional[str] = None
    guest_count: Optional[int] = None
    commission_scenario: Optional[str] = None
    opex_hours: Optional[float] = None
    management_fee: Optional[float] = None
    travel_expenses: Optional[float] = None
    tax_rate_pct: Optional[float] = None
    line_items: Optional[list[LineItemInput]] = None
    notes: Optional[str] = None


class CalculateRequest(BaseModel):
    line_items: list[LineItemInput]
    commission_scenario: str = "direct_client"
    opex_hours: float = Field(default=0.0, ge=0)
    management_fee: float = Field(default=0.0, ge=0)
    travel_expenses: float = Field(default=0.0, ge=0)
    tax_rate_pct: float = Field(default=0.0, ge=0)
