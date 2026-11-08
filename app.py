from __future__ import annotations

import os
from contextlib import asynccontextmanager
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Generator

from fastapi import Depends, FastAPI, HTTPException, Query
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, ConfigDict, Field, field_validator
from sqlalchemy import Date, DateTime, Float, ForeignKey, Integer, String, Text, UniqueConstraint, create_engine, func, select
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, sessionmaker


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
DATABASE_URL = os.getenv("SUPPLIER_DATABASE_URL", f"sqlite:///{ROOT / 'data' / 'suppliers.db'}")


def utc_now() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


class Base(DeclarativeBase):
    pass


class Supplier(Base):
    __tablename__ = "suppliers"

    id: Mapped[int] = mapped_column(primary_key=True)
    code: Mapped[str] = mapped_column(String(24), unique=True, index=True)
    name: Mapped[str] = mapped_column(String(160), index=True)
    category: Mapped[str] = mapped_column(String(80))
    country: Mapped[str] = mapped_column(String(80))
    contact_name: Mapped[str] = mapped_column(String(120), default="")
    email: Mapped[str] = mapped_column(String(180), default="")
    phone: Mapped[str] = mapped_column(String(48), default="")
    status: Mapped[str] = mapped_column(String(32), default="Pending review", index=True)
    risk: Mapped[str] = mapped_column(String(24), default="Low")
    on_time_rate: Mapped[float] = mapped_column(Float, default=92)
    quality_rate: Mapped[float] = mapped_column(Float, default=97)
    response_rate: Mapped[float] = mapped_column(Float, default=88)
    scored_deliveries: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utc_now)


class RFQ(Base):
    __tablename__ = "rfqs"

    id: Mapped[int] = mapped_column(primary_key=True)
    code: Mapped[str] = mapped_column(String(24), unique=True, index=True)
    title: Mapped[str] = mapped_column(String(180))
    category: Mapped[str] = mapped_column(String(80))
    description: Mapped[str] = mapped_column(Text, default="")
    quantity: Mapped[int] = mapped_column(Integer)
    due_date: Mapped[date] = mapped_column(Date)
    status: Mapped[str] = mapped_column(String(24), default="Open", index=True)
    awarded_supplier_id: Mapped[int | None] = mapped_column(ForeignKey("suppliers.id"), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utc_now)


class RFQInvitation(Base):
    __tablename__ = "rfq_invitations"
    __table_args__ = (UniqueConstraint("rfq_id", "supplier_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    rfq_id: Mapped[int] = mapped_column(ForeignKey("rfqs.id", ondelete="CASCADE"), index=True)
    supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)


class Quote(Base):
    __tablename__ = "quotes"
    __table_args__ = (UniqueConstraint("rfq_id", "supplier_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    rfq_id: Mapped[int] = mapped_column(ForeignKey("rfqs.id", ondelete="CASCADE"), index=True)
    supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
    unit_price: Mapped[float] = mapped_column(Float)
    currency: Mapped[str] = mapped_column(String(3), default="USD")
    lead_days: Mapped[int] = mapped_column(Integer)
    payment_terms: Mapped[str] = mapped_column(String(120), default="Net 30")
    note: Mapped[str] = mapped_column(Text, default="")
    submitted_at: Mapped[datetime] = mapped_column(DateTime, default=utc_now)


class Item(Base):
    __tablename__ = "items"

    id: Mapped[int] = mapped_column(primary_key=True)
    code: Mapped[str] = mapped_column(String(32), unique=True, index=True)
    name: Mapped[str] = mapped_column(String(180), index=True)
    category: Mapped[str] = mapped_column(String(80))
    uom: Mapped[str] = mapped_column(String(24), default="ea")
    stock_on_hand: Mapped[float] = mapped_column(Float, default=0)
    reorder_level: Mapped[float] = mapped_column(Float, default=0)
    buffer_days: Mapped[int] = mapped_column(Integer, default=0)
    unit_cost: Mapped[float] = mapped_column(Float, default=0)
    currency: Mapped[str] = mapped_column(String(3), default="USD")
    preferred_supplier_id: Mapped[int | None] = mapped_column(ForeignKey("suppliers.id"), nullable=True)
    lead_time_days: Mapped[int] = mapped_column(Integer, default=14)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utc_now)


class BOMComponent(Base):
    __tablename__ = "bom_components"
    __table_args__ = (UniqueConstraint("parent_item_id", "child_item_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    parent_item_id: Mapped[int] = mapped_column(ForeignKey("items.id", ondelete="CASCADE"), index=True)
    child_item_id: Mapped[int] = mapped_column(ForeignKey("items.id"), index=True)
    quantity: Mapped[float] = mapped_column(Float)


class MaterialRequest(Base):
    __tablename__ = "material_requests"

    id: Mapped[int] = mapped_column(primary_key=True)
    number: Mapped[str] = mapped_column(String(32), unique=True, index=True)
    item_id: Mapped[int] = mapped_column(ForeignKey("items.id"), index=True)
    quantity_requested: Mapped[float] = mapped_column(Float)
    quantity_received: Mapped[float] = mapped_column(Float, default=0)
    needed_by: Mapped[date] = mapped_column(Date)
    status: Mapped[str] = mapped_column(String(32), default="Open", index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utc_now)


class StockPurchaseOrder(Base):
    __tablename__ = "stock_purchase_orders"

    id: Mapped[int] = mapped_column(primary_key=True)
    po_number: Mapped[str] = mapped_column(String(32), unique=True, index=True)
    request_id: Mapped[int] = mapped_column(ForeignKey("material_requests.id"), unique=True, index=True)
    supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
    quantity_ordered: Mapped[float] = mapped_column(Float)
    quantity_received: Mapped[float] = mapped_column(Float, default=0)
    unit_price: Mapped[float] = mapped_column(Float)
    currency: Mapped[str] = mapped_column(String(3))
    status: Mapped[str] = mapped_column(String(24), default="Issued", index=True)
    ordered_at: Mapped[datetime] = mapped_column(DateTime, default=utc_now)
    expected_date: Mapped[date] = mapped_column(Date)


class StockDelivery(Base):
    __tablename__ = "stock_deliveries"

    id: Mapped[int] = mapped_column(primary_key=True)
    po_id: Mapped[int] = mapped_column(ForeignKey("stock_purchase_orders.id", ondelete="CASCADE"), index=True)
    quantity_received: Mapped[float] = mapped_column(Float)
    delivered_at: Mapped[date] = mapped_column(Date)
    note: Mapped[str] = mapped_column(Text, default="")


class PurchaseOrder(Base):
    __tablename__ = "purchase_orders"

    id: Mapped[int] = mapped_column(primary_key=True)
    po_number: Mapped[str] = mapped_column(String(32), unique=True, index=True)
    rfq_id: Mapped[int] = mapped_column(ForeignKey("rfqs.id"))
    supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
    quantity_ordered: Mapped[int] = mapped_column(Integer)
    quantity_received: Mapped[int] = mapped_column(Integer, default=0)
    unit_price: Mapped[float] = mapped_column(Float)
    currency: Mapped[str] = mapped_column(String(3))
    status: Mapped[str] = mapped_column(String(24), default="Issued", index=True)
    ordered_at: Mapped[datetime] = mapped_column(DateTime, default=utc_now)
    expected_date: Mapped[date] = mapped_column(Date)


class Delivery(Base):
    __tablename__ = "deliveries"

    id: Mapped[int] = mapped_column(primary_key=True)
    po_id: Mapped[int] = mapped_column(ForeignKey("purchase_orders.id", ondelete="CASCADE"), index=True)
    quantity_received: Mapped[int] = mapped_column(Integer)
    delivered_at: Mapped[date] = mapped_column(Date)
    note: Mapped[str] = mapped_column(Text, default="")


class Activity(Base):
    __tablename__ = "activities"

    id: Mapped[int] = mapped_column(primary_key=True)
    message: Mapped[str] = mapped_column(String(240))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=utc_now, index=True)


class QuoteDecision(BaseModel):
    supplier_id: int


class SupplierCreate(BaseModel):
    name: str = Field(min_length=2, max_length=160)
    category: str = Field(min_length=2, max_length=80)
    country: str = Field(min_length=2, max_length=80)
    contact_name: str = Field(default="", max_length=120)
    email: str = Field(default="", max_length=180)
    phone: str = Field(default="", max_length=48)


class SupplierUpdate(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str | None = Field(default=None, min_length=2, max_length=160)
    category: str | None = Field(default=None, min_length=2, max_length=80)
    country: str | None = Field(default=None, min_length=2, max_length=80)
    contact_name: str | None = Field(default=None, max_length=120)
    email: str | None = Field(default=None, max_length=180)
    phone: str | None = Field(default=None, max_length=48)
    status: str | None = Field(default=None, pattern="^(Pending review|Approved|On hold)$")
    response_rate: float | None = Field(default=None, ge=0, le=100)


class RFQCreate(BaseModel):
    title: str = Field(min_length=3, max_length=180)
    category: str = Field(min_length=2, max_length=80)
    description: str = Field(default="", max_length=4000)
    quantity: int = Field(gt=0, le=10_000_000)
    due_date: date
    supplier_ids: list[int] = Field(min_length=1, max_length=100)


class QuoteCreate(BaseModel):
    supplier_id: int
    unit_price: float = Field(gt=0, le=1_000_000_000)
    currency: str = Field(default="USD", pattern="^[A-Z]{3}$")
    lead_days: int = Field(gt=0, le=3650)
    payment_terms: str = Field(default="Net 30", max_length=120)
    note: str = Field(default="", max_length=2000)


class DeliveryCreate(BaseModel):
    quantity_received: int = Field(gt=0, le=10_000_000)
    delivered_at: date = Field(default_factory=date.today)
    note: str = Field(default="", max_length=1000)
    quality_accepted: bool = True


class ItemCreate(BaseModel):
    code: str = Field(min_length=2, max_length=32, pattern="^[A-Za-z0-9._-]+$")
    name: str = Field(min_length=2, max_length=180)
    category: str = Field(min_length=2, max_length=80)
    uom: str = Field(default="ea", min_length=1, max_length=24)
    stock_on_hand: float = Field(default=0, ge=0, le=1_000_000_000)
    reorder_level: float = Field(default=0, ge=0, le=1_000_000_000)
    buffer_days: int = Field(default=0, ge=0, le=5000)
    unit_cost: float = Field(default=0, ge=0, le=1_000_000_000)
    currency: str = Field(default="USD", pattern="^[A-Z]{3}$")
    preferred_supplier_id: int | None = None
    lead_time_days: int = Field(default=14, ge=0, le=3650)


class ItemUpdate(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str | None = Field(default=None, min_length=2, max_length=180)
    category: str | None = Field(default=None, min_length=2, max_length=80)
    uom: str | None = Field(default=None, min_length=1, max_length=24)
    stock_on_hand: float | None = Field(default=None, ge=0, le=1_000_000_000)
    reorder_level: float | None = Field(default=None, ge=0, le=1_000_000_000)
    buffer_days: int | None = Field(default=None, ge=0, le=5000)
    unit_cost: float | None = Field(default=None, ge=0, le=1_000_000_000)
    currency: str | None = Field(default=None, pattern="^[A-Z]{3}$")
    preferred_supplier_id: int | None = None
    lead_time_days: int | None = Field(default=None, ge=0, le=3650)


class BOMComponentCreate(BaseModel):
    child_item_id: int
    quantity: float = Field(ge=-1_000_000_000, le=1_000_000_000, allow_inf_nan=False)

    @field_validator("quantity")
    @classmethod
    def quantity_must_be_nonzero(cls, value: float) -> float:
        if value == 0:
            raise ValueError("BOM quantity must be non-zero")
        return value


class MaterialRequestCreate(BaseModel):
    item_id: int
    quantity: float = Field(gt=0, le=1_000_000_000)
    needed_by: date


class StockOrderCreate(BaseModel):
    supplier_id: int | None = None
    unit_price: float | None = Field(default=None, gt=0, le=1_000_000_000)


class StockDeliveryCreate(BaseModel):
    quantity_received: float = Field(gt=0, le=1_000_000_000)
    delivered_at: date = Field(default_factory=date.today)
    note: str = Field(default="", max_length=1000)
    quality_accepted: bool = True


def _supplier_dict(supplier: Supplier) -> dict:
    score = round(supplier.on_time_rate * 0.4 + supplier.quality_rate * 0.4 + supplier.response_rate * 0.2)
    return {
        "id": supplier.id,
        "code": supplier.code,
        "name": supplier.name,
        "category": supplier.category,
        "country": supplier.country,
        "contact_name": supplier.contact_name,
        "email": supplier.email,
        "phone": supplier.phone,
        "status": supplier.status,
        "risk": supplier.risk,
        "on_time_rate": round(supplier.on_time_rate, 1),
        "quality_rate": round(supplier.quality_rate, 1),
        "response_rate": round(supplier.response_rate, 1),
        "score": score,
    }


def _record_activity(db: Session, message: str) -> None:
    db.add(Activity(message=message))


def _seed_inventory_demo(db: Session) -> None:
    if db.scalar(select(func.count(Item.id))):
        return
    by_code = {supplier.code: supplier for supplier in db.scalars(select(Supplier)).all()}
    if not {"SUP-1001", "SUP-1002", "SUP-1003", "SUP-1004"}.issubset(by_code):
        return
    items = [
        Item(code="ASSY-001", name="Hydraulic valve assembly", category="Finished assembly", uom="ea", stock_on_hand=24, reorder_level=30, unit_cost=0, preferred_supplier_id=by_code["SUP-1002"].id, lead_time_days=14),
        Item(code="SUB-001", name="Valve body subassembly", category="Subassembly", uom="ea", stock_on_hand=18, reorder_level=25, unit_cost=42.80, preferred_supplier_id=by_code["SUP-1001"].id, lead_time_days=18),
        Item(code="MAT-101", name="Machined valve housing", category="CNC machining", uom="ea", stock_on_hand=12, reorder_level=40, unit_cost=42.80, preferred_supplier_id=by_code["SUP-1001"].id, lead_time_days=18),
        Item(code="MAT-102", name="Sealing O-ring, NBR 70", category="Plastics", uom="ea", stock_on_hand=420, reorder_level=200, unit_cost=0.35, currency="USD", preferred_supplier_id=by_code["SUP-1004"].id, lead_time_days=12),
        Item(code="MAT-103", name="M8 stainless fastener kit", category="Fasteners", uom="kit", stock_on_hand=85, reorder_level=120, unit_cost=3.10, preferred_supplier_id=by_code["SUP-1001"].id, lead_time_days=12),
        Item(code="SUB-002", name="Control connector subassembly", category="Electrical components", uom="ea", stock_on_hand=72, reorder_level=20, unit_cost=8.90, preferred_supplier_id=by_code["SUP-1003"].id, lead_time_days=22),
        Item(code="MAT-104", name="Copper connector contacts", category="Electrical components", uom="ea", stock_on_hand=160, reorder_level=300, unit_cost=1.15, preferred_supplier_id=by_code["SUP-1003"].id, lead_time_days=22),
        Item(code="MAT-105", name="Aluminium 6061 billet", category="Raw material", uom="kg", stock_on_hand=55, reorder_level=50, unit_cost=6.25, preferred_supplier_id=by_code["SUP-1001"].id, lead_time_days=9),
    ]
    db.add_all(items)
    db.flush()
    item_by_code = {item.code: item for item in items}
    db.add_all([
        BOMComponent(parent_item_id=item_by_code["ASSY-001"].id, child_item_id=item_by_code["SUB-001"].id, quantity=1),
        BOMComponent(parent_item_id=item_by_code["ASSY-001"].id, child_item_id=item_by_code["SUB-002"].id, quantity=1),
        BOMComponent(parent_item_id=item_by_code["ASSY-001"].id, child_item_id=item_by_code["MAT-103"].id, quantity=2),
        BOMComponent(parent_item_id=item_by_code["SUB-001"].id, child_item_id=item_by_code["MAT-101"].id, quantity=1),
        BOMComponent(parent_item_id=item_by_code["SUB-001"].id, child_item_id=item_by_code["MAT-102"].id, quantity=2),
        BOMComponent(parent_item_id=item_by_code["SUB-002"].id, child_item_id=item_by_code["MAT-104"].id, quantity=4),
        BOMComponent(parent_item_id=item_by_code["MAT-101"].id, child_item_id=item_by_code["MAT-105"].id, quantity=1.8),
    ])
    db.add(MaterialRequest(
        number="MR-26012",
        item_id=item_by_code["MAT-101"].id,
        quantity_requested=28,
        needed_by=date.today() + timedelta(days=24),
        status="Open",
    ))
    db.add(Activity(message="Reorder needed: machined valve housing is below its minimum stock level"))
    db.commit()


def _seed(db: Session) -> None:
    if db.scalar(select(func.count(Supplier.id))):
        _seed_inventory_demo(db)
        return

    suppliers = [
        Supplier(code="SUP-1001", name="Apex Precision Works", category="CNC machining", country="India", contact_name="Neha Rao", email="neha@apexprecision.example", phone="+91 80 5550 0182", status="Approved", risk="Low", on_time_rate=96, quality_rate=98, response_rate=91, scored_deliveries=14),
        Supplier(code="SUP-1002", name="Northstar Components", category="Metal fabrication", country="United States", contact_name="Jordan Lee", email="jordan@northstar.example", phone="+1 312 555 0146", status="Approved", risk="Low", on_time_rate=92, quality_rate=97, response_rate=94, scored_deliveries=10),
        Supplier(code="SUP-1003", name="Koyo Industrial Supply", category="Electrical components", country="Japan", contact_name="Aiko Tanaka", email="aiko@koyo.example", phone="+81 3 5555 0118", status="Approved", risk="Medium", on_time_rate=88, quality_rate=99, response_rate=81, scored_deliveries=8),
        Supplier(code="SUP-1004", name="Metro Polymer Group", category="Plastics", country="India", contact_name="Arjun Mehta", email="arjun@metropolymer.example", phone="+91 22 5550 0261", status="Pending review", risk="Medium", on_time_rate=90, quality_rate=95, response_rate=84, scored_deliveries=5),
        Supplier(code="SUP-1005", name="Blue Ridge Fasteners", category="Fasteners", country="Germany", contact_name="Mia Fischer", email="mia@blueridge.example", phone="+49 30 5550 0159", status="On hold", risk="High", on_time_rate=74, quality_rate=89, response_rate=72, scored_deliveries=9),
    ]
    db.add_all(suppliers)
    db.flush()
    by_code = {supplier.code: supplier for supplier in suppliers}

    first_rfq = RFQ(code="RFQ-26014", title="CNC machined valve housings", category="CNC machining", description="Aluminium 6061-T6 housings, deburred and dimensionally inspected.", quantity=500, due_date=date.today() + timedelta(days=6), status="Open")
    second_rfq = RFQ(code="RFQ-26013", title="M8 stainless fastener kit", category="Fasteners", description="A2-70 stainless steel, packed in kits of 20.", quantity=2400, due_date=date.today() + timedelta(days=2), status="Open")
    awarded_rfq = RFQ(code="RFQ-26011", title="Laser-cut mounting brackets", category="Metal fabrication", description="Powder-coated steel mounting brackets.", quantity=1200, due_date=date.today() - timedelta(days=8), status="Awarded", awarded_supplier_id=by_code["SUP-1002"].id)
    second_awarded_rfq = RFQ(code="RFQ-26010", title="Turned steel spacers", category="CNC machining", description="Zinc-plated turned spacers, batch inspected.", quantity=500, due_date=date.today() - timedelta(days=12), status="Awarded", awarded_supplier_id=by_code["SUP-1001"].id)
    third_awarded_rfq = RFQ(code="RFQ-26009", title="Electrical connector contacts", category="Electrical components", description="Tin-plated copper contacts with incoming inspection.", quantity=800, due_date=date.today() - timedelta(days=16), status="Awarded", awarded_supplier_id=by_code["SUP-1003"].id)
    db.add_all([first_rfq, second_rfq, awarded_rfq, second_awarded_rfq, third_awarded_rfq])
    db.flush()

    db.add_all([
        RFQInvitation(rfq_id=first_rfq.id, supplier_id=by_code["SUP-1001"].id),
        RFQInvitation(rfq_id=first_rfq.id, supplier_id=by_code["SUP-1002"].id),
        RFQInvitation(rfq_id=first_rfq.id, supplier_id=by_code["SUP-1003"].id),
        RFQInvitation(rfq_id=second_rfq.id, supplier_id=by_code["SUP-1001"].id),
        RFQInvitation(rfq_id=second_rfq.id, supplier_id=by_code["SUP-1005"].id),
        RFQInvitation(rfq_id=awarded_rfq.id, supplier_id=by_code["SUP-1002"].id),
        RFQInvitation(rfq_id=second_awarded_rfq.id, supplier_id=by_code["SUP-1001"].id),
        RFQInvitation(rfq_id=third_awarded_rfq.id, supplier_id=by_code["SUP-1003"].id),
        Quote(rfq_id=first_rfq.id, supplier_id=by_code["SUP-1001"].id, unit_price=42.80, currency="USD", lead_days=18, payment_terms="Net 30"),
        Quote(rfq_id=first_rfq.id, supplier_id=by_code["SUP-1002"].id, unit_price=45.25, currency="USD", lead_days=14, payment_terms="Net 45"),
        Quote(rfq_id=second_rfq.id, supplier_id=by_code["SUP-1001"].id, unit_price=3.10, currency="USD", lead_days=12, payment_terms="Net 30"),
        Quote(rfq_id=awarded_rfq.id, supplier_id=by_code["SUP-1002"].id, unit_price=8.40, currency="USD", lead_days=10),
        Quote(rfq_id=second_awarded_rfq.id, supplier_id=by_code["SUP-1001"].id, unit_price=8.65, currency="USD", lead_days=13),
        Quote(rfq_id=third_awarded_rfq.id, supplier_id=by_code["SUP-1003"].id, unit_price=8.90, currency="USD", lead_days=22),
    ])
    db.flush()

    purchase_orders = [
        PurchaseOrder(po_number="PO-26038", rfq_id=awarded_rfq.id, supplier_id=by_code["SUP-1002"].id, quantity_ordered=1200, quantity_received=1200, unit_price=8.40, currency="USD", status="Received", ordered_at=utc_now() - timedelta(days=18), expected_date=date.today() - timedelta(days=2)),
        PurchaseOrder(po_number="PO-26041", rfq_id=second_awarded_rfq.id, supplier_id=by_code["SUP-1001"].id, quantity_ordered=500, quantity_received=200, unit_price=8.65, currency="USD", status="Partially received", ordered_at=utc_now() - timedelta(days=9), expected_date=date.today() + timedelta(days=3)),
        PurchaseOrder(po_number="PO-26043", rfq_id=third_awarded_rfq.id, supplier_id=by_code["SUP-1003"].id, quantity_ordered=800, quantity_received=0, unit_price=8.90, currency="USD", status="Issued", ordered_at=utc_now() - timedelta(days=2), expected_date=date.today() + timedelta(days=11)),
    ]
    db.add_all(purchase_orders)
    db.flush()
    db.add_all([
        Delivery(po_id=purchase_orders[0].id, quantity_received=1200, delivered_at=date.today() - timedelta(days=2), note="Received in full; inspection passed."),
        Delivery(po_id=purchase_orders[1].id, quantity_received=200, delivered_at=date.today(), note="First shipment received; remaining balance pending."),
    ])
    _seed_inventory_demo(db)
    db.add_all([
        Activity(message="PO-26038 delivery received and quality approved"),
        Activity(message="Metro Polymer Group submitted onboarding details"),
        Activity(message="RFQ-26014 sent to 3 approved suppliers"),
    ])
    db.commit()


def create_app(database_url: str = DATABASE_URL) -> FastAPI:
    from planning import install_planning

    if database_url.startswith("sqlite:///"):
        Path(database_url.removeprefix("sqlite:///")).parent.mkdir(parents=True, exist_ok=True)
    engine = create_engine(
        database_url,
        connect_args={"check_same_thread": False} if database_url.startswith("sqlite") else {},
    )
    Base.metadata.create_all(engine)
    if "buffer_days" not in {column["name"] for column in __import__("sqlalchemy").inspect(engine).get_columns("items")}:
        with engine.begin() as connection:
            connection.exec_driver_sql("ALTER TABLE items ADD COLUMN buffer_days INTEGER NOT NULL DEFAULT 0")
    SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)

    @asynccontextmanager
    async def lifespan(_: FastAPI):
        with SessionLocal() as db:
            _seed(db)
        yield
        engine.dispose()

    app = FastAPI(title="Design to Cost | Supplier Hub", version="1.0.0", lifespan=lifespan)
    install_planning(app, SessionLocal)
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

    def get_db() -> Generator[Session, None, None]:
        db = SessionLocal()
        try:
            yield db
        finally:
            db.close()

    @app.get("/", include_in_schema=False)
    def home():
        return FileResponse(STATIC_DIR / "index.html")

    @app.get("/api/health")
    def health():
        return {"status": "ok"}

    @app.get("/api/dashboard")
    def dashboard(db: Session = Depends(get_db)):
        suppliers = db.scalars(select(Supplier)).all()
        orders = db.scalars(select(PurchaseOrder).order_by(PurchaseOrder.ordered_at.desc())).all()
        stock_orders = db.scalars(select(StockPurchaseOrder)).all()
        items = db.scalars(select(Item)).all()
        rfqs = db.scalars(select(RFQ)).all()
        spend_by_currency: dict[str, float] = {}
        for order in [*orders, *stock_orders]:
            spend_by_currency[order.currency] = spend_by_currency.get(order.currency, 0) + order.quantity_ordered * order.unit_price
        activities = db.scalars(select(Activity).order_by(Activity.created_at.desc(), Activity.id.desc()).limit(6)).all()
        active_orders = [order for order in [*orders, *stock_orders] if order.status != "Received"]
        open_requests = db.scalars(select(MaterialRequest).where(MaterialRequest.status.in_(["Open", "Ordered", "Partially received"]))).all()
        return {
            "metrics": {
                "suppliers": sum(s.status == "Approved" for s in suppliers),
                "open_rfqs": sum(r.status == "Open" for r in rfqs),
                "active_orders": len(active_orders),
                "items": len(items),
                "reorder_items": sum(item.stock_on_hand <= item.reorder_level for item in items),
                "open_material_requests": len(open_requests),
                "spend_by_currency": {currency: round(amount, 2) for currency, amount in spend_by_currency.items()},
            },
            "supplier_status": {
                "approved": sum(s.status == "Approved" for s in suppliers),
                "pending": sum(s.status == "Pending review" for s in suppliers),
                "on_hold": sum(s.status == "On hold" for s in suppliers),
            },
            "orders": sorted(
                [_order_dict(db, order) for order in orders]
                + [_stock_order_dict(db, order) for order in stock_orders],
                key=lambda order: order["ordered_at"],
                reverse=True,
            )[:5],
            "activities": [{"message": a.message, "created_at": a.created_at.isoformat()} for a in activities],
            "top_suppliers": sorted((_supplier_dict(s) for s in suppliers), key=lambda s: s["score"], reverse=True)[:4],
        }

    @app.get("/api/suppliers")
    def list_suppliers(search: str = Query(default="", max_length=120), status: str = Query(default="", max_length=32), db: Session = Depends(get_db)):
        statement = select(Supplier).order_by(Supplier.name)
        if search.strip():
            like = f"%{search.strip()}%"
            statement = statement.where((Supplier.name.ilike(like)) | (Supplier.code.ilike(like)) | (Supplier.category.ilike(like)))
        if status:
            statement = statement.where(Supplier.status == status)
        return [_supplier_dict(s) for s in db.scalars(statement).all()]

    @app.post("/api/suppliers", status_code=201)
    def create_supplier(payload: SupplierCreate, db: Session = Depends(get_db)):
        sequence = int(db.scalar(select(func.count(Supplier.id))) or 0) + 1001
        supplier = Supplier(
            code=f"SUP-{sequence:04d}",
            name=payload.name.strip(),
            category=payload.category.strip(),
            country=payload.country.strip(),
            contact_name=payload.contact_name.strip(),
            email=payload.email.strip(),
            phone=payload.phone.strip(),
            status="Pending review",
            risk="Medium",
        )
        db.add(supplier)
        _record_activity(db, f"{supplier.name} added to the onboarding review queue")
        db.commit()
        db.refresh(supplier)
        return _supplier_dict(supplier)

    @app.patch("/api/suppliers/{supplier_id}")
    def update_supplier(supplier_id: int, payload: SupplierUpdate, db: Session = Depends(get_db)):
        supplier = db.get(Supplier, supplier_id)
        if not supplier:
            raise HTTPException(status_code=404, detail="Supplier not found")
        for key, value in payload.model_dump(exclude_unset=True).items():
            if value is not None:
                setattr(supplier, key, value.strip() if isinstance(value, str) else value)
        _record_activity(db, f"{supplier.name} supplier profile updated")
        db.commit()
        db.refresh(supplier)
        return _supplier_dict(supplier)

    @app.get("/api/items")
    def list_items(search: str = Query(default="", max_length=120), db: Session = Depends(get_db)):
        statement = select(Item).order_by(Item.code)
        if search.strip():
            like = f"%{search.strip()}%"
            statement = statement.where((Item.name.ilike(like)) | (Item.code.ilike(like)) | (Item.category.ilike(like)))
        return [_item_dict(db, item) for item in db.scalars(statement).all()]

    @app.post("/api/items", status_code=201)
    def create_item(payload: ItemCreate, db: Session = Depends(get_db)):
        if db.scalar(select(Item.id).where(Item.code == payload.code.strip())):
            raise HTTPException(status_code=409, detail="An item with this code already exists")
        _validate_item_supplier(db, payload.preferred_supplier_id)
        item = Item(**payload.model_dump() | {"code": payload.code.strip(), "name": payload.name.strip(), "category": payload.category.strip(), "uom": payload.uom.strip()})
        db.add(item)
        _record_activity(db, f"Inventory item {item.code} added")
        db.commit()
        db.refresh(item)
        return _item_dict(db, item)

    @app.patch("/api/items/{item_id}")
    def update_item(item_id: int, payload: ItemUpdate, db: Session = Depends(get_db)):
        item = db.get(Item, item_id)
        if not item:
            raise HTTPException(status_code=404, detail="Inventory item not found")
        updates = payload.model_dump(exclude_unset=True)
        if "preferred_supplier_id" in updates:
            _validate_item_supplier(db, updates["preferred_supplier_id"])
        for key, value in updates.items():
            setattr(item, key, value.strip() if isinstance(value, str) else value)
        _record_activity(db, f"Inventory item {item.code} updated")
        db.commit()
        db.refresh(item)
        return _item_dict(db, item)

    @app.get("/api/items/{item_id}/bom")
    def get_bom(item_id: int, db: Session = Depends(get_db)):
        item = db.get(Item, item_id)
        if not item:
            raise HTTPException(status_code=404, detail="Inventory item not found")
        return _bom_tree(db, item, 0, {item.id}, 1)

    @app.post("/api/items/{item_id}/bom", status_code=201)
    def add_bom_component(item_id: int, payload: BOMComponentCreate, db: Session = Depends(get_db)):
        parent = db.get(Item, item_id)
        child = db.get(Item, payload.child_item_id)
        if not parent or not child:
            raise HTTPException(status_code=404, detail="Parent or component item not found")
        if parent.id == child.id:
            raise HTTPException(status_code=422, detail="An item cannot contain itself")
        if db.scalar(select(BOMComponent.id).where(BOMComponent.parent_item_id == parent.id, BOMComponent.child_item_id == child.id)):
            raise HTTPException(status_code=409, detail="This component is already in the BOM")
        if _bom_reaches(db, child.id, parent.id):
            raise HTTPException(status_code=422, detail="This component would create a circular BOM")
        db.add(BOMComponent(parent_item_id=parent.id, child_item_id=child.id, quantity=payload.quantity))
        _record_activity(db, f"{child.code} added to {parent.code} BOM")
        db.commit()
        return _bom_tree(db, parent, 0, {parent.id}, 1)

    @app.patch("/api/items/{item_id}/bom/{component_id}")
    def update_bom_component(item_id: int, component_id: int, payload: BOMComponentCreate, db: Session = Depends(get_db)):
        parent = db.get(Item, item_id)
        child = db.get(Item, payload.child_item_id)
        component = db.get(BOMComponent, component_id)
        if not parent or not child or not component or component.parent_item_id != item_id:
            raise HTTPException(status_code=404, detail="BOM component not found for this parent item")
        if parent.id == child.id or _bom_reaches(db, child.id, parent.id):
            raise HTTPException(status_code=422, detail="This component would create a circular BOM")
        duplicate = db.scalar(select(BOMComponent.id).where(
            BOMComponent.parent_item_id == item_id,
            BOMComponent.child_item_id == child.id,
            BOMComponent.id != component.id,
        ))
        if duplicate:
            raise HTTPException(status_code=409, detail="This item is already listed in the BOM")
        component.child_item_id = child.id
        component.quantity = payload.quantity
        _record_activity(db, f"{child.code} BOM quantity updated on {parent.code}")
        db.commit()
        return _bom_tree(db, parent, 0, {parent.id}, 1)

    @app.delete("/api/items/{item_id}/bom/{component_id}", status_code=204)
    def remove_bom_component(item_id: int, component_id: int, db: Session = Depends(get_db)):
        component = db.get(BOMComponent, component_id)
        parent = db.get(Item, item_id)
        if not parent or not component or component.parent_item_id != item_id:
            raise HTTPException(status_code=404, detail="BOM component not found for this parent item")
        child = db.get(Item, component.child_item_id)
        db.delete(component)
        _record_activity(db, f"{child.code if child else 'Component'} removed from {parent.code} BOM")
        db.commit()

    @app.get("/api/material-requests")
    def list_material_requests(db: Session = Depends(get_db)):
        requests = db.scalars(select(MaterialRequest).order_by(MaterialRequest.created_at.desc(), MaterialRequest.id.desc())).all()
        return [_material_request_dict(db, request) for request in requests]

    @app.post("/api/material-requests", status_code=201)
    def create_material_request(payload: MaterialRequestCreate, db: Session = Depends(get_db)):
        item = db.get(Item, payload.item_id)
        if not item:
            raise HTTPException(status_code=404, detail="Inventory item not found")
        if payload.needed_by < date.today():
            raise HTTPException(status_code=422, detail="The required date must be today or later")
        number = _next_number(db, MaterialRequest, "MR")
        request = MaterialRequest(number=number, item_id=item.id, quantity_requested=payload.quantity, needed_by=payload.needed_by)
        db.add(request)
        _record_activity(db, f"{number} requested {payload.quantity:g} {item.uom} of {item.name}")
        db.commit()
        db.refresh(request)
        return _material_request_dict(db, request)

    @app.post("/api/material-requests/reorder", status_code=201)
    def create_reorder_requests(db: Session = Depends(get_db)):
        items = db.scalars(select(Item).where(Item.stock_on_hand <= Item.reorder_level).order_by(Item.code)).all()
        created = []
        for item in items:
            active = db.scalar(select(MaterialRequest.id).where(
                MaterialRequest.item_id == item.id,
                MaterialRequest.status.in_(["Open", "Ordered", "Partially received"]),
            ))
            if active:
                continue
            number = _next_number(db, MaterialRequest, "MR")
            request = MaterialRequest(
                number=number,
                item_id=item.id,
                quantity_requested=max(item.reorder_level - item.stock_on_hand, 1),
                needed_by=date.today() + timedelta(days=item.lead_time_days),
            )
            db.add(request)
            db.flush()
            created.append(request)
        if created:
            db.flush()
            _record_activity(db, f"{len(created)} material request(s) created from reorder levels")
            db.commit()
            for request in created:
                db.refresh(request)
        return [_material_request_dict(db, request) for request in created]

    @app.post("/api/material-requests/{request_id}/purchase-order", status_code=201)
    def create_stock_purchase_order(request_id: int, payload: StockOrderCreate, db: Session = Depends(get_db)):
        request = db.get(MaterialRequest, request_id)
        if not request:
            raise HTTPException(status_code=404, detail="Material request not found")
        if request.status != "Open":
            raise HTTPException(status_code=409, detail="Only an open material request can be ordered")
        item = db.get(Item, request.item_id)
        supplier_id = payload.supplier_id or item.preferred_supplier_id
        supplier = db.get(Supplier, supplier_id) if supplier_id else None
        if not supplier or supplier.status != "Approved":
            raise HTTPException(status_code=422, detail="Choose an approved supplier for this item")
        price = payload.unit_price if payload.unit_price is not None else item.unit_cost
        if price <= 0:
            raise HTTPException(status_code=422, detail="Enter a purchase price before creating the order")
        number = _next_purchase_order_number(db)
        order = StockPurchaseOrder(
            po_number=number,
            request_id=request.id,
            supplier_id=supplier.id,
            quantity_ordered=request.quantity_requested,
            unit_price=price,
            currency=item.currency,
            expected_date=date.today() + timedelta(days=item.lead_time_days),
        )
        request.status = "Ordered"
        db.add(order)
        _record_activity(db, f"{number} issued to {supplier.name} for {item.name}")
        db.commit()
        db.refresh(order)
        return _stock_order_dict(db, order)

    @app.post("/api/stock-purchase-orders/{po_id}/deliveries", status_code=201)
    def receive_stock_delivery(po_id: int, payload: StockDeliveryCreate, db: Session = Depends(get_db)):
        order = db.get(StockPurchaseOrder, po_id)
        if not order:
            raise HTTPException(status_code=404, detail="Stock purchase order not found")
        remaining = order.quantity_ordered - order.quantity_received
        if payload.quantity_received > remaining:
            raise HTTPException(status_code=422, detail=f"Quantity exceeds the remaining {remaining:g} units")
        request = db.get(MaterialRequest, order.request_id)
        item = db.get(Item, request.item_id)
        supplier = db.get(Supplier, order.supplier_id)
        db.add(StockDelivery(po_id=order.id, quantity_received=payload.quantity_received, delivered_at=payload.delivered_at, note=payload.note.strip()))
        order.quantity_received += payload.quantity_received
        order.status = "Received" if order.quantity_received == order.quantity_ordered else "Partially received"
        item.stock_on_hand += payload.quantity_received
        request.quantity_received += payload.quantity_received
        request.status = "Received" if request.quantity_received >= request.quantity_requested else "Partially received"
        on_time = payload.delivered_at <= order.expected_date
        supplier.on_time_rate = round((supplier.on_time_rate * supplier.scored_deliveries + (100 if on_time else 0)) / (supplier.scored_deliveries + 1), 1)
        supplier.quality_rate = round((supplier.quality_rate * supplier.scored_deliveries + (100 if payload.quality_accepted else 0)) / (supplier.scored_deliveries + 1), 1)
        supplier.scored_deliveries += 1
        _record_activity(db, f"{payload.quantity_received:g} {item.uom} of {item.name} received into stock against {order.po_number}")
        db.commit()
        return _stock_order_dict(db, order)

    @app.get("/api/rfqs")
    def list_rfqs(db: Session = Depends(get_db)):
        rfqs = db.scalars(select(RFQ).order_by(RFQ.created_at.desc(), RFQ.id.desc())).all()
        return [_rfq_dict(db, rfq) for rfq in rfqs]

    @app.post("/api/rfqs", status_code=201)
    def create_rfq(payload: RFQCreate, db: Session = Depends(get_db)):
        if payload.due_date < date.today():
            raise HTTPException(status_code=422, detail="The response deadline must be today or later")
        suppliers = db.scalars(select(Supplier).where(Supplier.id.in_(set(payload.supplier_ids)))).all()
        if len(suppliers) != len(set(payload.supplier_ids)):
            raise HTTPException(status_code=422, detail="One or more selected suppliers do not exist")
        if any(s.status != "Approved" for s in suppliers):
            raise HTTPException(status_code=422, detail="RFQs can only be sent to approved suppliers")
        number = int(db.scalar(select(func.count(RFQ.id))) or 0) + 26001
        rfq = RFQ(code=f"RFQ-{number}", title=payload.title.strip(), category=payload.category.strip(), description=payload.description.strip(), quantity=payload.quantity, due_date=payload.due_date)
        db.add(rfq)
        db.flush()
        db.add_all(RFQInvitation(rfq_id=rfq.id, supplier_id=s.id) for s in suppliers)
        _record_activity(db, f"{rfq.code} sent to {len(suppliers)} suppliers")
        db.commit()
        db.refresh(rfq)
        return _rfq_dict(db, rfq)

    @app.post("/api/rfqs/{rfq_id}/quotes", status_code=201)
    def submit_quote(rfq_id: int, payload: QuoteCreate, db: Session = Depends(get_db)):
        rfq = db.get(RFQ, rfq_id)
        supplier = db.get(Supplier, payload.supplier_id)
        if not rfq or not supplier:
            raise HTTPException(status_code=404, detail="RFQ or supplier not found")
        if rfq.status != "Open":
            raise HTTPException(status_code=409, detail="Quotes can only be recorded for an open RFQ")
        invitation = db.scalar(select(RFQInvitation).where(RFQInvitation.rfq_id == rfq_id, RFQInvitation.supplier_id == supplier.id))
        if not invitation:
            raise HTTPException(status_code=422, detail="This supplier was not invited to the RFQ")
        quote = db.scalar(select(Quote).where(Quote.rfq_id == rfq_id, Quote.supplier_id == supplier.id))
        if quote:
            for key, value in payload.model_dump(exclude={"supplier_id"}).items():
                setattr(quote, key, value)
            quote.submitted_at = utc_now()
        else:
            quote = Quote(rfq_id=rfq_id, **payload.model_dump())
            db.add(quote)
        _record_activity(db, f"{supplier.name} quote recorded for {rfq.code}")
        db.commit()
        return _rfq_dict(db, rfq)

    @app.post("/api/rfqs/{rfq_id}/award")
    def award_rfq(rfq_id: int, payload: QuoteDecision, db: Session = Depends(get_db)):
        rfq = db.get(RFQ, rfq_id)
        if not rfq:
            raise HTTPException(status_code=404, detail="RFQ not found")
        if rfq.status != "Open":
            raise HTTPException(status_code=409, detail="This RFQ has already been awarded")
        quote = db.scalar(select(Quote).where(Quote.rfq_id == rfq_id, Quote.supplier_id == payload.supplier_id))
        supplier = db.get(Supplier, payload.supplier_id)
        if not quote or not supplier:
            raise HTTPException(status_code=422, detail="The selected supplier has no quote for this RFQ")
        rfq.status = "Awarded"
        rfq.awarded_supplier_id = supplier.id
        order = PurchaseOrder(
            po_number=_next_purchase_order_number(db),
            rfq_id=rfq.id,
            supplier_id=supplier.id,
            quantity_ordered=rfq.quantity,
            unit_price=quote.unit_price,
            currency=quote.currency,
            expected_date=date.today() + timedelta(days=quote.lead_days),
        )
        db.add(order)
        _record_activity(db, f"{rfq.code} awarded to {supplier.name}; {order.po_number} issued")
        db.commit()
        return {"rfq": _rfq_dict(db, rfq), "purchase_order": _order_dict(db, order)}

    @app.get("/api/purchase-orders")
    def list_purchase_orders(db: Session = Depends(get_db)):
        orders = db.scalars(select(PurchaseOrder).order_by(PurchaseOrder.ordered_at.desc(), PurchaseOrder.id.desc())).all()
        stock_orders = db.scalars(select(StockPurchaseOrder).order_by(StockPurchaseOrder.ordered_at.desc(), StockPurchaseOrder.id.desc())).all()
        return sorted(
            [_order_dict(db, order) for order in orders]
            + [_stock_order_dict(db, order) for order in stock_orders],
            key=lambda order: order["ordered_at"],
            reverse=True,
        )

    @app.post("/api/purchase-orders/{po_id}/deliveries", status_code=201)
    def record_delivery(po_id: int, payload: DeliveryCreate, db: Session = Depends(get_db)):
        order = db.get(PurchaseOrder, po_id)
        if not order:
            raise HTTPException(status_code=404, detail="Purchase order not found")
        if order.status == "Received":
            raise HTTPException(status_code=409, detail="This purchase order has already been received in full")
        remaining = order.quantity_ordered - order.quantity_received
        if payload.quantity_received > remaining:
            raise HTTPException(status_code=422, detail=f"Quantity exceeds the remaining {remaining} units")
        db.add(Delivery(po_id=order.id, quantity_received=payload.quantity_received, delivered_at=payload.delivered_at, note=payload.note.strip()))
        order.quantity_received += payload.quantity_received
        order.status = "Received" if order.quantity_received == order.quantity_ordered else "Partially received"
        supplier = db.get(Supplier, order.supplier_id)
        on_time = payload.delivered_at <= order.expected_date
        supplier.on_time_rate = round((supplier.on_time_rate * supplier.scored_deliveries + (100 if on_time else 0)) / (supplier.scored_deliveries + 1), 1)
        supplier.quality_rate = round((supplier.quality_rate * supplier.scored_deliveries + (100 if payload.quality_accepted else 0)) / (supplier.scored_deliveries + 1), 1)
        supplier.scored_deliveries += 1
        _record_activity(db, f"{payload.quantity_received:,} units received against {order.po_number}")
        db.commit()
        return _order_dict(db, order)

    @app.get("/api/performance")
    def performance(db: Session = Depends(get_db)):
        suppliers = db.scalars(select(Supplier).order_by(Supplier.name)).all()
        return {
            "suppliers": [_supplier_dict(s) for s in suppliers],
            "summary": {
                "average_score": round(sum(_supplier_dict(s)["score"] for s in suppliers) / len(suppliers)) if suppliers else 0,
                "on_time_rate": round(sum(s.on_time_rate for s in suppliers) / len(suppliers)) if suppliers else 0,
                "quality_rate": round(sum(s.quality_rate for s in suppliers) / len(suppliers)) if suppliers else 0,
                "response_rate": round(sum(s.response_rate for s in suppliers) / len(suppliers)) if suppliers else 0,
            },
        }

    return app


def _rfq_dict(db: Session, rfq: RFQ) -> dict:
    invitations = db.scalars(select(RFQInvitation).where(RFQInvitation.rfq_id == rfq.id)).all()
    supplier_ids = [invite.supplier_id for invite in invitations]
    suppliers = {s.id: s for s in db.scalars(select(Supplier).where(Supplier.id.in_(supplier_ids))).all()} if supplier_ids else {}
    quotes = db.scalars(select(Quote).where(Quote.rfq_id == rfq.id).order_by(Quote.unit_price)).all()
    return {
        "id": rfq.id,
        "code": rfq.code,
        "title": rfq.title,
        "category": rfq.category,
        "description": rfq.description,
        "quantity": rfq.quantity,
        "due_date": rfq.due_date.isoformat(),
        "status": rfq.status,
        "supplier_count": len(invitations),
        "quote_count": len(quotes),
        "invited_suppliers": [_supplier_dict(suppliers[sid]) for sid in supplier_ids if sid in suppliers],
        "quotes": [
            {
                "id": q.id,
                "supplier_id": q.supplier_id,
                "supplier_name": suppliers[q.supplier_id].name,
                "unit_price": q.unit_price,
                "currency": q.currency,
                "lead_days": q.lead_days,
                "payment_terms": q.payment_terms,
                "note": q.note,
                "total": round(q.unit_price * rfq.quantity, 2),
                "submitted_at": q.submitted_at.isoformat(),
            }
            for q in quotes if q.supplier_id in suppliers
        ],
        "awarded_supplier_id": rfq.awarded_supplier_id,
    }


def _order_dict(db: Session, order: PurchaseOrder) -> dict:
    supplier = db.get(Supplier, order.supplier_id)
    rfq = db.get(RFQ, order.rfq_id)
    return {
        "id": order.id,
        "po_number": order.po_number,
        "rfq_code": rfq.code if rfq else "—",
        "title": rfq.title if rfq else "Purchase order",
        "supplier_id": order.supplier_id,
        "supplier_name": supplier.name if supplier else "Unknown supplier",
        "quantity_ordered": order.quantity_ordered,
        "quantity_received": order.quantity_received,
        "unit_price": order.unit_price,
        "total": round(order.quantity_ordered * order.unit_price, 2),
        "currency": order.currency,
        "status": order.status,
        "expected_date": order.expected_date.isoformat(),
        "ordered_at": order.ordered_at.isoformat(),
        "progress": round(order.quantity_received * 100 / order.quantity_ordered) if order.quantity_ordered else 0,
        "order_type": "rfq",
    }


def _validate_item_supplier(db: Session, supplier_id: int | None) -> None:
    if supplier_id is not None and not db.get(Supplier, supplier_id):
        raise HTTPException(status_code=422, detail="Preferred supplier does not exist")


def _next_number(db: Session, model: type, prefix: str) -> str:
    existing = db.scalars(select(model.number)).all()
    highest = max((int(value.removeprefix(f"{prefix}-")) for value in existing if value.startswith(f"{prefix}-") and value.removeprefix(f"{prefix}-").isdigit()), default=26000)
    return f"{prefix}-{highest + 1}"


def _next_purchase_order_number(db: Session) -> str:
    numbers = (
        db.scalars(select(PurchaseOrder.po_number)).all()
        + db.scalars(select(StockPurchaseOrder.po_number)).all()
    )
    highest = max((int(value.removeprefix("PO-")) for value in numbers if value.startswith("PO-") and value.removeprefix("PO-").isdigit()), default=26000)
    return f"PO-{highest + 1}"


def _item_dict(db: Session, item: Item) -> dict:
    supplier = db.get(Supplier, item.preferred_supplier_id) if item.preferred_supplier_id else None
    component_count = int(db.scalar(select(func.count(BOMComponent.id)).where(BOMComponent.parent_item_id == item.id)) or 0)
    return {
        "id": item.id,
        "code": item.code,
        "name": item.name,
        "category": item.category,
        "uom": item.uom,
        "stock_on_hand": item.stock_on_hand,
        "reorder_level": item.reorder_level,
        "buffer_days": item.buffer_days,
        "unit_cost": item.unit_cost,
        "currency": item.currency,
        "preferred_supplier_id": item.preferred_supplier_id,
        "preferred_supplier_name": supplier.name if supplier else None,
        "lead_time_days": item.lead_time_days,
        "needs_reorder": item.stock_on_hand <= item.reorder_level,
        "bom_component_count": component_count,
    }


def _bom_reaches(db: Session, start_id: int, target_id: int) -> bool:
    pending = [start_id]
    visited: set[int] = set()
    while pending:
        current = pending.pop()
        if current == target_id:
            return True
        if current in visited:
            continue
        visited.add(current)
        pending.extend(db.scalars(select(BOMComponent.child_item_id).where(BOMComponent.parent_item_id == current)).all())
    return False


def _bom_tree(db: Session, item: Item, level: int, ancestors: set[int], multiplier: float) -> dict:
    children = []
    components = db.scalars(
        select(BOMComponent).where(BOMComponent.parent_item_id == item.id).order_by(BOMComponent.id)
    ).all()
    for component in components:
        child = db.get(Item, component.child_item_id)
        if not child or child.id in ancestors:
            continue
        children.append({
            "component_id": component.id,
            "quantity_per_parent": component.quantity,
            "quantity_in_assembly": component.quantity * multiplier,
            "level": level + 1,
            "item": _item_dict(db, child),
            "children": _bom_tree(db, child, level + 1, ancestors | {child.id}, multiplier * component.quantity)["children"],
        })
    return {"item": _item_dict(db, item), "level": level, "quantity_in_assembly": multiplier, "children": children}


def _material_request_dict(db: Session, request: MaterialRequest) -> dict:
    item = db.get(Item, request.item_id)
    supplier = db.get(Supplier, item.preferred_supplier_id) if item and item.preferred_supplier_id else None
    order = db.scalar(select(StockPurchaseOrder).where(StockPurchaseOrder.request_id == request.id))
    return {
        "id": request.id,
        "number": request.number,
        "item_id": request.item_id,
        "item_code": item.code if item else "—",
        "item_name": item.name if item else "Unknown item",
        "uom": item.uom if item else "ea",
        "quantity_requested": request.quantity_requested,
        "quantity_received": request.quantity_received,
        "needed_by": request.needed_by.isoformat(),
        "status": request.status,
        "preferred_supplier_id": supplier.id if supplier else None,
        "preferred_supplier_name": supplier.name if supplier else None,
        "lead_time_days": item.lead_time_days if item else 0,
        "unit_cost": item.unit_cost if item else 0,
        "currency": item.currency if item else "USD",
        "purchase_order_id": order.id if order else None,
        "purchase_order_number": order.po_number if order else None,
    }


def _stock_order_dict(db: Session, order: StockPurchaseOrder) -> dict:
    request = db.get(MaterialRequest, order.request_id)
    item = db.get(Item, request.item_id) if request else None
    supplier = db.get(Supplier, order.supplier_id)
    return {
        "id": order.id,
        "po_number": order.po_number,
        "rfq_code": request.number if request else "—",
        "title": item.name if item else "Stock purchase",
        "supplier_id": order.supplier_id,
        "supplier_name": supplier.name if supplier else "Unknown supplier",
        "quantity_ordered": order.quantity_ordered,
        "quantity_received": order.quantity_received,
        "unit_price": order.unit_price,
        "total": round(order.quantity_ordered * order.unit_price, 2),
        "currency": order.currency,
        "status": order.status,
        "expected_date": order.expected_date.isoformat(),
        "ordered_at": order.ordered_at.isoformat(),
        "progress": round(order.quantity_received * 100 / order.quantity_ordered) if order.quantity_ordered else 0,
        "order_type": "stock",
        "item_id": item.id if item else None,
        "item_code": item.code if item else None,
        "request_id": request.id if request else None,
    }


app = create_app()
