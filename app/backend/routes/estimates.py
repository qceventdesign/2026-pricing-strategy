"""Estimate CRUD and calculation routes."""

import json
from datetime import datetime, timezone
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from ..database import get_db, Estimate, EstimateLineItem, AuditLog
from ..schemas import EstimateCreate, EstimateUpdate, EstimateSummary
from ..pricing_engine import (
    calculate_estimate,
    EstimateInput,
    LineItem,
    LineItemType,
    get_estimate_template,
    get_available_estimate_types,
)
from .. import config_loader

router = APIRouter(prefix="/api/estimates", tags=["estimates"])


def _db_to_engine_input(estimate: Estimate) -> EstimateInput:
    """Convert DB estimate to pricing engine input."""
    line_items = []
    for li in estimate.line_items:
        line_items.append(LineItem(
            name=li.name,
            category_key=li.category_key,
            item_type=LineItemType(li.item_type),
            unit_price=li.unit_price,
            quantity=li.quantity,
            is_taxable=li.is_taxable,
            notes=li.notes,
            section=li.section,
        ))

    return EstimateInput(
        estimate_name=estimate.estimate_name,
        estimate_type=estimate.estimate_type,
        client_name=estimate.client_name,
        event_name=estimate.event_name,
        event_date=estimate.event_date,
        event_time=estimate.event_time,
        guest_count=estimate.guest_count,
        location=estimate.location,
        tax_rate=estimate.tax_rate,
        commission_scenario=estimate.commission_scenario,
        cc_processing_pct=estimate.cc_processing_pct,
        gratuity_pct=estimate.gratuity_pct,
        admin_fee_pct=estimate.admin_fee_pct,
        gdp_enabled=estimate.gdp_enabled,
        opex_hours=estimate.opex_hours,
        line_items=line_items,
    )


@router.get("/types")
def list_estimate_types():
    """List available estimate types (venue, decor, tour, etc.)."""
    return get_available_estimate_types()


@router.get("/types/{estimate_type}/template")
def get_template(estimate_type: str):
    """Get the template definition for an estimate type (sections + default line items)."""
    try:
        return get_estimate_template(estimate_type)
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))


@router.get("/")
def list_estimates(db: Session = Depends(get_db)):
    """List all estimates (summary view)."""
    estimates = db.query(Estimate).order_by(Estimate.updated_at.desc()).all()
    return [
        {
            "id": e.id,
            "estimate_name": e.estimate_name,
            "estimate_type": e.estimate_type,
            "client_name": e.client_name,
            "event_name": e.event_name,
            "event_date": e.event_date,
            "status": e.status,
            "created_at": e.created_at.isoformat() if e.created_at else None,
            "updated_at": e.updated_at.isoformat() if e.updated_at else None,
        }
        for e in estimates
    ]


@router.post("/")
def create_estimate(data: EstimateCreate, db: Session = Depends(get_db)):
    """Create a new estimate with line items."""
    estimate = Estimate(
        estimate_name=data.estimate_name,
        estimate_type=data.estimate_type,
        client_name=data.client_name,
        event_name=data.event_name,
        event_date=data.event_date,
        event_time=data.event_time,
        guest_count=data.guest_count,
        location=data.location,
        tax_rate=data.tax_rate,
        commission_scenario=data.commission_scenario,
        cc_processing_pct=data.cc_processing_pct,
        gratuity_pct=data.gratuity_pct,
        admin_fee_pct=data.admin_fee_pct,
        gdp_enabled=data.gdp_enabled,
        opex_hours=data.opex_hours,
    )
    db.add(estimate)
    db.flush()

    for i, li in enumerate(data.line_items):
        item = EstimateLineItem(
            estimate_id=estimate.id,
            section=li.section,
            name=li.name,
            category_key=li.category_key,
            item_type=li.item_type,
            unit_price=li.unit_price,
            quantity=li.quantity,
            is_taxable=li.is_taxable,
            notes=li.notes,
            sort_order=li.sort_order or i,
        )
        db.add(item)

    audit = AuditLog(
        estimate_id=estimate.id,
        action="created",
        details=json.dumps({"estimate_type": data.estimate_type, "client": data.client_name}),
    )
    db.add(audit)
    db.commit()
    db.refresh(estimate)

    # Return the calculated result
    engine_input = _db_to_engine_input(estimate)
    result = calculate_estimate(engine_input)
    result["id"] = estimate.id
    result["status"] = estimate.status
    return result


@router.get("/{estimate_id}")
def get_estimate(estimate_id: int, db: Session = Depends(get_db)):
    """Get a single estimate with full calculation."""
    estimate = db.query(Estimate).filter(Estimate.id == estimate_id).first()
    if not estimate:
        raise HTTPException(status_code=404, detail="Estimate not found")

    engine_input = _db_to_engine_input(estimate)
    result = calculate_estimate(engine_input)
    result["id"] = estimate.id
    result["status"] = estimate.status
    result["created_at"] = estimate.created_at.isoformat() if estimate.created_at else None
    result["updated_at"] = estimate.updated_at.isoformat() if estimate.updated_at else None
    return result


@router.put("/{estimate_id}")
def update_estimate(estimate_id: int, data: EstimateUpdate, db: Session = Depends(get_db)):
    """Update an estimate (fields and/or line items)."""
    estimate = db.query(Estimate).filter(Estimate.id == estimate_id).first()
    if not estimate:
        raise HTTPException(status_code=404, detail="Estimate not found")

    changes = {}
    for field_name in [
        "estimate_name", "client_name", "event_name", "event_date", "event_time",
        "guest_count", "location", "tax_rate", "commission_scenario",
        "cc_processing_pct", "gratuity_pct", "admin_fee_pct", "gdp_enabled",
        "opex_hours", "status",
    ]:
        new_val = getattr(data, field_name)
        if new_val is not None:
            old_val = getattr(estimate, field_name)
            if old_val != new_val:
                changes[field_name] = {"old": old_val, "new": new_val}
                setattr(estimate, field_name, new_val)

    if data.line_items is not None:
        # Replace all line items
        db.query(EstimateLineItem).filter(EstimateLineItem.estimate_id == estimate_id).delete()
        for i, li in enumerate(data.line_items):
            item = EstimateLineItem(
                estimate_id=estimate_id,
                section=li.section,
                name=li.name,
                category_key=li.category_key,
                item_type=li.item_type,
                unit_price=li.unit_price,
                quantity=li.quantity,
                is_taxable=li.is_taxable,
                notes=li.notes,
                sort_order=li.sort_order or i,
            )
            db.add(item)
        changes["line_items"] = "replaced"

    estimate.updated_at = datetime.now(timezone.utc)

    audit = AuditLog(
        estimate_id=estimate_id,
        action="updated",
        details=json.dumps(changes, default=str),
    )
    db.add(audit)
    db.commit()
    db.refresh(estimate)

    engine_input = _db_to_engine_input(estimate)
    result = calculate_estimate(engine_input)
    result["id"] = estimate.id
    result["status"] = estimate.status
    return result


@router.delete("/{estimate_id}")
def delete_estimate(estimate_id: int, db: Session = Depends(get_db)):
    """Delete an estimate."""
    estimate = db.query(Estimate).filter(Estimate.id == estimate_id).first()
    if not estimate:
        raise HTTPException(status_code=404, detail="Estimate not found")
    db.delete(estimate)
    db.commit()
    return {"deleted": True}


@router.get("/{estimate_id}/audit")
def get_audit_log(estimate_id: int, db: Session = Depends(get_db)):
    """Get the audit log for an estimate."""
    logs = db.query(AuditLog).filter(AuditLog.estimate_id == estimate_id).order_by(AuditLog.created_at.desc()).all()
    return [
        {
            "id": log.id,
            "action": log.action,
            "details": json.loads(log.details) if log.details else {},
            "created_at": log.created_at.isoformat() if log.created_at else None,
        }
        for log in logs
    ]
