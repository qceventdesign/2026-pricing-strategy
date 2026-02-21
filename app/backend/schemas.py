"""Pydantic schemas for API request/response validation."""

from pydantic import BaseModel, Field
from datetime import datetime


class LineItemCreate(BaseModel):
    section: str = ""
    name: str
    category_key: str
    item_type: str = "quantity"
    unit_price: float = 0.0
    quantity: int = 0
    is_taxable: bool = True
    notes: str = ""
    sort_order: int = 0


class LineItemResponse(LineItemCreate):
    id: int
    estimate_id: int
    vendor_cost: float = 0.0
    markup_pct: float = 0.0
    client_cost: float = 0.0
    category_label: str = ""

    class Config:
        from_attributes = True


class EstimateCreate(BaseModel):
    estimate_name: str
    estimate_type: str
    client_name: str = ""
    event_name: str = ""
    event_date: str = ""
    event_time: str = ""
    guest_count: int = 0
    location: str = ""
    tax_rate: float = 0.0
    commission_scenario: str = "direct_client"
    cc_processing_pct: float = 0.035
    gratuity_pct: float = 0.20
    admin_fee_pct: float = 0.05
    gdp_enabled: bool = False
    opex_hours: float = 0.0
    line_items: list[LineItemCreate] = Field(default_factory=list)


class EstimateUpdate(BaseModel):
    estimate_name: str | None = None
    client_name: str | None = None
    event_name: str | None = None
    event_date: str | None = None
    event_time: str | None = None
    guest_count: int | None = None
    location: str | None = None
    tax_rate: float | None = None
    commission_scenario: str | None = None
    cc_processing_pct: float | None = None
    gratuity_pct: float | None = None
    admin_fee_pct: float | None = None
    gdp_enabled: bool | None = None
    opex_hours: float | None = None
    status: str | None = None
    line_items: list[LineItemCreate] | None = None


class EstimateSummary(BaseModel):
    id: int
    estimate_name: str
    estimate_type: str
    client_name: str
    event_name: str
    event_date: str
    status: str
    created_at: datetime
    updated_at: datetime

    class Config:
        from_attributes = True


class CalculationResult(BaseModel):
    """Full calculation output from the pricing engine."""
    estimate_name: str
    estimate_type: str
    client_name: str
    event_name: str
    line_items: list[dict]
    sections: dict
    subtotals: dict
    pnl: dict
    client_ready: dict
