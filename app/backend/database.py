"""SQLite database setup with SQLAlchemy."""

import json
from datetime import datetime, timezone
from sqlalchemy import create_engine, Column, Integer, String, Float, Boolean, Text, DateTime, ForeignKey
from sqlalchemy.orm import DeclarativeBase, sessionmaker, relationship

DATABASE_URL = "sqlite:///./qc_estimates.db"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class Base(DeclarativeBase):
    pass


class Estimate(Base):
    __tablename__ = "estimates"

    id = Column(Integer, primary_key=True, index=True)
    estimate_name = Column(String, nullable=False)
    estimate_type = Column(String, nullable=False)  # venue, decor, tour, etc.
    client_name = Column(String, default="")
    event_name = Column(String, default="")
    event_date = Column(String, default="")
    event_time = Column(String, default="")
    guest_count = Column(Integer, default=0)
    location = Column(String, default="")
    tax_rate = Column(Float, default=0.0)
    commission_scenario = Column(String, default="direct_client")
    cc_processing_pct = Column(Float, default=0.035)
    gratuity_pct = Column(Float, default=0.20)
    admin_fee_pct = Column(Float, default=0.05)
    gdp_enabled = Column(Boolean, default=False)
    opex_hours = Column(Float, default=0.0)
    status = Column(String, default="draft")  # draft, sent, approved, archived
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    updated_at = Column(DateTime, default=lambda: datetime.now(timezone.utc), onupdate=lambda: datetime.now(timezone.utc))

    line_items = relationship("EstimateLineItem", back_populates="estimate", cascade="all, delete-orphan")
    audit_logs = relationship("AuditLog", back_populates="estimate", cascade="all, delete-orphan")


class EstimateLineItem(Base):
    __tablename__ = "estimate_line_items"

    id = Column(Integer, primary_key=True, index=True)
    estimate_id = Column(Integer, ForeignKey("estimates.id"), nullable=False)
    section = Column(String, default="")
    name = Column(String, nullable=False)
    category_key = Column(String, nullable=False)
    item_type = Column(String, default="quantity")  # quantity, flat_fee, per_person
    unit_price = Column(Float, default=0.0)
    quantity = Column(Integer, default=0)
    is_taxable = Column(Boolean, default=True)
    notes = Column(String, default="")
    sort_order = Column(Integer, default=0)

    estimate = relationship("Estimate", back_populates="line_items")


class AuditLog(Base):
    __tablename__ = "audit_logs"

    id = Column(Integer, primary_key=True, index=True)
    estimate_id = Column(Integer, ForeignKey("estimates.id"), nullable=False)
    action = Column(String, nullable=False)  # created, updated, status_changed, exported
    details = Column(Text, default="")  # JSON string with change details
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

    estimate = relationship("Estimate", back_populates="audit_logs")


def init_db():
    Base.metadata.create_all(bind=engine)


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
