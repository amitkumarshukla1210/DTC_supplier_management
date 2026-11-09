from __future__ import annotations

import hashlib
import hmac
import io
import json
import math
import re
import secrets
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from fastapi import FastAPI, HTTPException, Query, Request, Response
from fastapi.responses import StreamingResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from pydantic import BaseModel, Field, field_validator
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from sqlalchemy import Boolean, Date, DateTime, Float, ForeignKey, Index, Integer, String, Text, UniqueConstraint, delete, func, or_, select
from sqlalchemy.orm import Mapped, Session, mapped_column, sessionmaker
from starlette.responses import JSONResponse
import xlsxwriter

from app import Base, BOMComponent, Item, MaterialRequest, Supplier


MONTH_RE = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
MODEL_CATALOG_PRODUCT_CODE = "__MODEL_CATALOG__"
WORKSPACES = {"trial", "production"}
ROLES = {"Admin", "HOD", "Sourcing", "Supply Chain", "User"}
SUPPLY_ROLES = {"Admin", "HOD", "Sourcing", "Supply Chain"}
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
REPORT_HEADERS = [
    "Month", "Material code", "Material name", "Supplier", "Upstream parent", "Tier", "SOB %",
    "Plan demand", "Actual demand", "Own stock", "Own stock allocated", "Supplier stock", "Incoming PO",
    "Projected balance", "Coverage days", "Buffer target", "Shortage", "Recommended PO",
    "Excess", "Unit cost", "Shortage value", "Excess value", "Lead time days", "Criticality", "Plan revision",
]


def _month(value: Any) -> date:
    if isinstance(value, datetime):
        return date(value.year, value.month, 1)
    if isinstance(value, date):
        return date(value.year, value.month, 1)
    if isinstance(value, (int, float)) and 30_000 <= value <= 100_000:
        from openpyxl.utils.datetime import from_excel
        parsed = from_excel(value)
        if isinstance(parsed, (date, datetime)):
            return date(parsed.year, parsed.month, 1)
    text = str(value or "").strip()
    if MONTH_RE.fullmatch(text):
        return date.fromisoformat(f"{text}-01")
    formats = (
        "%b-%y", "%B-%y", "%b-%Y", "%B-%Y", "%b %y", "%B %y", "%b %Y", "%B %Y",
        "%Y-%b", "%Y-%B", "%Y %b", "%Y %B", "%m/%Y", "%m-%Y", "%m.%Y",
        "%Y/%m", "%Y.%m", "%Y%m", "%Y-%m-%d", "%d-%b-%y", "%d-%B-%Y",
        "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%m/%d/%y", "%d/%m/%y",
    )
    for month_format in formats:
        try:
            parsed = datetime.strptime(text, month_format)
            return date(parsed.year, parsed.month, 1)
        except ValueError:
            continue
    raise HTTPException(status_code=422, detail=f"Invalid month value {text!r}; use a month such as 2026-04, Apr-26, or a valid Excel date")


def _month_text(value: date) -> str:
    return value.strftime("%Y-%m")


def _plus_months(value: date, count: int) -> date:
    index = value.year * 12 + value.month - 1 + count
    return date(index // 12, index % 12 + 1, 1)


def _password_hash(password: str, salt: bytes | None = None) -> str:
    actual_salt = salt or secrets.token_bytes(16)
    derived = hashlib.scrypt(password.encode("utf-8"), salt=actual_salt, n=2**14, r=8, p=1)
    return f"{actual_salt.hex()}:{derived.hex()}"


def _password_matches(password: str, encoded: str) -> bool:
    salt_hex, digest_hex = encoded.split(":", 1)
    return hmac.compare_digest(_password_hash(password, bytes.fromhex(salt_hex)).split(":", 1)[1], digest_hex)


class AppUser(Base):
    __tablename__ = "app_users"

    id: Mapped[int] = mapped_column(primary_key=True)
    username: Mapped[str] = mapped_column(String(80), unique=True, index=True)
    full_name: Mapped[str] = mapped_column(String(160))
    password_hash: Mapped[str] = mapped_column(String(200))
    role: Mapped[str] = mapped_column(String(32), default="User", index=True)
    can_view_prices: Mapped[bool] = mapped_column(Boolean, default=False)
    enabled: Mapped[bool] = mapped_column(Boolean, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class LoginSession(Base):
    __tablename__ = "login_sessions"

    id: Mapped[int] = mapped_column(primary_key=True)
    token_hash: Mapped[str] = mapped_column(String(64), unique=True, index=True)
    user_id: Mapped[int] = mapped_column(ForeignKey("app_users.id", ondelete="CASCADE"), index=True)
    expires_at: Mapped[datetime] = mapped_column(DateTime, index=True)


class Product(Base):
    __tablename__ = "products"

    id: Mapped[int] = mapped_column(primary_key=True)
    code: Mapped[str] = mapped_column(String(40), unique=True, index=True)
    name: Mapped[str] = mapped_column(String(180), index=True)


class ProductModel(Base):
    __tablename__ = "product_models"
    __table_args__ = (UniqueConstraint("product_id", "code"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    product_id: Mapped[int] = mapped_column(ForeignKey("products.id", ondelete="CASCADE"), index=True)
    code: Mapped[str] = mapped_column(String(40), index=True)
    name: Mapped[str] = mapped_column(String(180), index=True)


class ProductModelProduct(Base):
    __tablename__ = "product_model_products"
    __table_args__ = (UniqueConstraint("product_id", "model_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    product_id: Mapped[int] = mapped_column(ForeignKey("products.id", ondelete="CASCADE"), index=True)
    model_id: Mapped[int] = mapped_column(ForeignKey("product_models.id", ondelete="CASCADE"), index=True)


class ModelBOM(Base):
    __tablename__ = "model_bom"
    __table_args__ = (UniqueConstraint("model_id", "material_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    model_id: Mapped[int] = mapped_column(ForeignKey("product_models.id", ondelete="CASCADE"), index=True)
    material_id: Mapped[int] = mapped_column(ForeignKey("items.id"), index=True)
    quantity: Mapped[float] = mapped_column(Float)
    uom: Mapped[str] = mapped_column(String(24), default="ea")


class SupplierPart(Base):
    __tablename__ = "supplier_parts"
    __table_args__ = (UniqueConstraint("material_id", "supplier_id", "tier", "parent_supplier_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    material_id: Mapped[int] = mapped_column(ForeignKey("items.id", ondelete="CASCADE"), index=True)
    supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
    tier: Mapped[str] = mapped_column(String(24), default="Direct", index=True)
    parent_supplier_id: Mapped[int | None] = mapped_column(ForeignKey("suppliers.id"), nullable=True, index=True)
    sob_percent: Mapped[float] = mapped_column(Float, default=100)
    lead_time_days: Mapped[int] = mapped_column(Integer, default=0)
    moq: Mapped[float] = mapped_column(Float, default=0)
    buffer_days: Mapped[int] = mapped_column(Integer, default=0)
    unit_price: Mapped[float] = mapped_column(Float, default=0)
    currency: Mapped[str] = mapped_column(String(3), default="USD")
    price_mode: Mapped[str] = mapped_column(String(24), default="spot")
    price_interval_months: Mapped[int] = mapped_column(Integer, default=1)
    price_updated_month: Mapped[str] = mapped_column(String(7), default=lambda: _month_text(date.today()))


class PlanningCycle(Base):
    __tablename__ = "planning_cycles"
    __table_args__ = (
        UniqueConstraint("workspace", "issue_month", "revision"),
        Index("ix_planning_cycles_lookup", "workspace", "issue_month", "revision"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    workspace: Mapped[str] = mapped_column(String(16), index=True)
    issue_month: Mapped[str] = mapped_column(String(7), index=True)
    revision: Mapped[int] = mapped_column(Integer)
    created_by: Mapped[int] = mapped_column(ForeignKey("app_users.id"))
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


class PlanLine(Base):
    __tablename__ = "plan_lines"
    __table_args__ = (UniqueConstraint("cycle_id", "target_month", "model_id"), Index("ix_plan_lines_period", "target_month", "model_id"))

    id: Mapped[int] = mapped_column(primary_key=True)
    cycle_id: Mapped[int] = mapped_column(ForeignKey("planning_cycles.id", ondelete="CASCADE"), index=True)
    target_month: Mapped[str] = mapped_column(String(7), index=True)
    model_id: Mapped[int] = mapped_column(ForeignKey("product_models.id"), index=True)
    quantity: Mapped[float] = mapped_column(Float, default=0)


class ActualProduction(Base):
    __tablename__ = "actual_production"
    __table_args__ = (UniqueConstraint("workspace", "period", "model_id"), Index("ix_actual_period_workspace", "workspace", "period"))

    id: Mapped[int] = mapped_column(primary_key=True)
    workspace: Mapped[str] = mapped_column(String(16), index=True)
    period: Mapped[str] = mapped_column(String(7), index=True)
    model_id: Mapped[int] = mapped_column(ForeignKey("product_models.id"), index=True)
    actual_quantity: Mapped[float] = mapped_column(Float, default=0)
    oem_stock: Mapped[float] = mapped_column(Float, default=0)
    updated_by: Mapped[int] = mapped_column(ForeignKey("app_users.id"))
    updated_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class OwnStockSnapshot(Base):
    __tablename__ = "own_stock_snapshots"
    __table_args__ = (UniqueConstraint("workspace", "period", "material_id"), Index("ix_own_stock_period", "workspace", "period"))

    id: Mapped[int] = mapped_column(primary_key=True)
    workspace: Mapped[str] = mapped_column(String(16), index=True)
    period: Mapped[str] = mapped_column(String(7), index=True)
    material_id: Mapped[int] = mapped_column(ForeignKey("items.id"), index=True)
    quantity: Mapped[float] = mapped_column(Float, default=0)


class SupplierStockSnapshot(Base):
    __tablename__ = "supplier_stock_snapshots"
    __table_args__ = (UniqueConstraint("workspace", "report_month", "supplier_part_id"), Index("ix_supplier_stock_period", "workspace", "report_month", "supplier_part_id"))

    id: Mapped[int] = mapped_column(primary_key=True)
    workspace: Mapped[str] = mapped_column(String(16), index=True)
    report_month: Mapped[str] = mapped_column(String(7), index=True)
    supplier_part_id: Mapped[int] = mapped_column(ForeignKey("supplier_parts.id", ondelete="CASCADE"), index=True)
    quantity: Mapped[float] = mapped_column(Float, default=0)
    available_quantity: Mapped[float] = mapped_column(Float, default=0)
    arrival_month: Mapped[str | None] = mapped_column(String(7), nullable=True)


class PlanningPO(Base):
    __tablename__ = "planning_purchase_orders"

    id: Mapped[int] = mapped_column(primary_key=True)
    po_number: Mapped[str] = mapped_column(String(32), unique=True, index=True)
    workspace: Mapped[str] = mapped_column(String(16), index=True)
    supplier_part_id: Mapped[int] = mapped_column(ForeignKey("supplier_parts.id"), index=True)
    quantity_ordered: Mapped[float] = mapped_column(Float)
    quantity_received: Mapped[float] = mapped_column(Float, default=0)
    unit_price: Mapped[float] = mapped_column(Float)
    currency: Mapped[str] = mapped_column(String(3))
    ordered_month: Mapped[str] = mapped_column(String(7), index=True)
    expected_month: Mapped[str] = mapped_column(String(7), index=True)
    status: Mapped[str] = mapped_column(String(24), default="Planned", index=True)
    ncnr: Mapped[bool] = mapped_column(Boolean, default=False)
    change_note: Mapped[str] = mapped_column(Text, default="")


class TierMaterial(Base):
    """T2 / T3 Material Master: codes used further down the supply chain (by your
    suppliers' suppliers). Not planned by you; picked into supplier material masters."""
    __tablename__ = "tier_materials"

    id: Mapped[int] = mapped_column(primary_key=True)
    tier: Mapped[str] = mapped_column(String(2), index=True)  # "T2" or "T3"
    code: Mapped[str] = mapped_column(String(48), unique=True, index=True)
    name: Mapped[str] = mapped_column(String(180))
    category: Mapped[str] = mapped_column(String(80), default="Raw material")
    uom: Mapped[str] = mapped_column(String(24), default="ea")
    unit_price: Mapped[float] = mapped_column(Float, default=0)
    currency: Mapped[str] = mapped_column(String(3), default="INR")


class SupplierMaterial(Base):
    """A supplier's own material code (e.g. Iljin's solder wire), used only in that
    supplier's BOMs for cost breakdown."""
    __tablename__ = "supplier_materials"
    __table_args__ = (UniqueConstraint("supplier_id", "code"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id", ondelete="CASCADE"), index=True)
    code: Mapped[str] = mapped_column(String(48), index=True)
    name: Mapped[str] = mapped_column(String(180))
    uom: Mapped[str] = mapped_column(String(24), default="ea")
    unit_price: Mapped[float] = mapped_column(Float, default=0)
    currency: Mapped[str] = mapped_column(String(3), default="INR")
    # Picked from your T1 Material Master (item_id) or T2/T3 Material Master (tier_material_id);
    # both empty = the supplier's own code. Code and name follow the picked master entry.
    item_id: Mapped[int | None] = mapped_column(ForeignKey("items.id"), nullable=True, index=True)
    tier_material_id: Mapped[int | None] = mapped_column(ForeignKey("tier_materials.id"), nullable=True, index=True)
    # The supplier's own terms for this code (may differ from yours, e.g. your discounted steel rate)
    lead_time_days: Mapped[int] = mapped_column(Integer, default=0)
    buffer_days: Mapped[int] = mapped_column(Integer, default=0)


class SupplierBOMLine(Base):
    """One line of a supplier's BOM for a part you buy from them (cost breakdown only).
    parent_line_id empty = L1 line under the bought part; L2 / L3 lines hang below."""
    __tablename__ = "supplier_bom_lines"

    id: Mapped[int] = mapped_column(primary_key=True)
    supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id", ondelete="CASCADE"), index=True)
    product_material_id: Mapped[int] = mapped_column(ForeignKey("items.id", ondelete="CASCADE"), index=True)
    parent_line_id: Mapped[int | None] = mapped_column(ForeignKey("supplier_bom_lines.id", ondelete="CASCADE"), nullable=True, index=True)
    item_id: Mapped[int | None] = mapped_column(ForeignKey("items.id"), nullable=True)
    supplier_material_id: Mapped[int | None] = mapped_column(ForeignKey("supplier_materials.id"), nullable=True)
    quantity: Mapped[float] = mapped_column(Float)
    unit_price: Mapped[float | None] = mapped_column(Float, nullable=True)
    currency: Mapped[str | None] = mapped_column(String(3), nullable=True)


class SupplierBOMLineSource(Base):
    """Who supplies a supplier BOM line to the supplier (a sub-supplier from the Supplier
    register) and under which code (same code or an alternate), at what price and SOB."""
    __tablename__ = "supplier_bom_line_sources"

    id: Mapped[int] = mapped_column(primary_key=True)
    line_id: Mapped[int] = mapped_column(ForeignKey("supplier_bom_lines.id", ondelete="CASCADE"), index=True)
    sub_supplier_id: Mapped[int | None] = mapped_column(ForeignKey("suppliers.id"), nullable=True)
    item_id: Mapped[int | None] = mapped_column(ForeignKey("items.id"), nullable=True)
    supplier_material_id: Mapped[int | None] = mapped_column(ForeignKey("supplier_materials.id"), nullable=True)
    unit_price: Mapped[float] = mapped_column(Float, default=0)
    currency: Mapped[str] = mapped_column(String(3), default="INR")
    sob_percent: Mapped[float] = mapped_column(Float, default=100)


class MaterialTransfer(Base):
    __tablename__ = "material_transfers"

    id: Mapped[int] = mapped_column(primary_key=True)
    workspace: Mapped[str] = mapped_column(String(16), index=True)
    material_id: Mapped[int] = mapped_column(ForeignKey("items.id"), index=True)
    from_supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
    to_supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
    quantity: Mapped[float] = mapped_column(Float)
    transfer_month: Mapped[str] = mapped_column(String(7), index=True)
    status: Mapped[str] = mapped_column(String(24), default="Planned")
    note: Mapped[str] = mapped_column(Text, default="")


class UserCreate(BaseModel):
    username: str = Field(min_length=3, max_length=80, pattern=r"^[A-Za-z0-9._-]+$")
    full_name: str = Field(min_length=2, max_length=160)
    password: str = Field(min_length=12, max_length=200)
    role: str = "User"
    can_view_prices: bool = False


class SetupAdmin(UserCreate):
    role: str = "Admin"


class ProductCreate(BaseModel):
    code: str = Field(min_length=1, max_length=40)
    name: str = Field(min_length=1, max_length=180)
    model_ids: list[int] = Field(default_factory=list)


class ModelCreate(BaseModel):
    code: str = Field(min_length=1, max_length=40)
    name: str = Field(min_length=1, max_length=180)
    product_id: int | None = None
    product_ids: list[int] = Field(default_factory=list)


class ModelUpdate(BaseModel):
    code: str = Field(min_length=1, max_length=40)
    name: str = Field(min_length=1, max_length=180)
    product_ids: list[int] = Field(default_factory=list)


class MaterialInput(BaseModel):
    code: str = Field(min_length=2, max_length=32, pattern=r"^[A-Za-z0-9._-]+$")
    name: str = Field(min_length=2, max_length=180)
    category: str = Field(default="Raw material", min_length=2, max_length=80)
    uom: str = Field(default="ea", min_length=1, max_length=24)
    stock_on_hand: float = Field(default=0, ge=0)
    buffer_days: int = Field(default=0, ge=0, le=5000)
    unit_cost: float = Field(default=0, ge=0)
    currency: str = Field(default="USD", pattern=r"^[A-Z]{3}$")


class ModelBOMInput(BaseModel):
    material_id: int
    quantity: float = Field(gt=0, le=1_000_000_000)
    uom: str = Field(default="ea", min_length=1, max_length=24)


class TierMaterialInput(BaseModel):
    tier: str = Field(pattern=r"^T[23]$")
    code: str = Field(min_length=1, max_length=48)
    name: str = Field(min_length=1, max_length=180)
    category: str = Field(default="Raw material", min_length=1, max_length=80)
    uom: str = Field(default="ea", min_length=1, max_length=24)
    unit_price: float = Field(default=0, ge=0, allow_inf_nan=False)
    currency: str = Field(default="INR", pattern=r"^[A-Z]{3}$")


class SupplierMaterialInput(BaseModel):
    supplier_id: int
    item_id: int | None = None
    tier_material_id: int | None = None
    code: str = Field(default="", max_length=48)
    name: str = Field(default="", max_length=180)
    lead_time_days: int = Field(default=0, ge=0, le=5000)
    buffer_days: int = Field(default=0, ge=0, le=5000)
    uom: str = Field(default="ea", min_length=1, max_length=24)
    unit_price: float = Field(default=0, ge=0, allow_inf_nan=False)
    currency: str = Field(default="INR", pattern=r"^[A-Z]{3}$")


class SupplierBOMLineInput(BaseModel):
    supplier_id: int
    material_id: int
    parent_line_id: int | None = None
    item_id: int | None = None
    supplier_material_id: int | None = None
    quantity: float = Field(ge=-1_000_000_000, le=1_000_000_000, allow_inf_nan=False)
    unit_price: float | None = Field(default=None, ge=0, allow_inf_nan=False)
    currency: str | None = Field(default=None, pattern=r"^[A-Z]{3}$")


class SupplierBOMLineUpdate(BaseModel):
    quantity: float = Field(ge=-1_000_000_000, le=1_000_000_000, allow_inf_nan=False)
    unit_price: float | None = Field(default=None, ge=0, allow_inf_nan=False)
    currency: str | None = Field(default=None, pattern=r"^[A-Z]{3}$")


class SupplierBOMSourceInput(BaseModel):
    sub_supplier_id: int | None = None
    item_id: int | None = None
    supplier_material_id: int | None = None
    unit_price: float = Field(default=0, ge=0, allow_inf_nan=False)
    currency: str = Field(default="INR", pattern=r"^[A-Z]{3}$")
    sob_percent: float = Field(gt=0, le=100, allow_inf_nan=False)


class NewPartBOMLine(BaseModel):
    child_item_id: int
    quantity: float = Field(ge=-1_000_000_000, le=1_000_000_000, allow_inf_nan=False)

    @field_validator("quantity")
    @classmethod
    def quantity_must_be_nonzero(cls, value: float) -> float:
        if value == 0:
            raise ValueError("BOM quantity must be non-zero")
        return value


class NewPartBOMInput(BaseModel):
    tab: str = Field(pattern=r"^Level [23] Part BOM$")
    lines: list[NewPartBOMLine] = Field(min_length=1, max_length=500)


class SupplierPartInput(BaseModel):
    material_id: int
    supplier_id: int
    tier: str = Field(default="Direct", min_length=2, max_length=24)
    parent_supplier_id: int | None = None
    sob_percent: float = Field(default=100, ge=0, le=100)
    lead_time_days: int = Field(default=0, ge=0, le=5000)
    moq: float = Field(default=0, ge=0)
    buffer_days: int = Field(default=0, ge=0, le=5000)
    unit_price: float = Field(default=0, ge=0)
    currency: str = Field(default="USD", pattern=r"^[A-Z]{3}$")
    price_mode: str = Field(default="spot", pattern=r"^(spot|fixed_monthly|every_n_months)$")
    price_interval_months: int = Field(default=1, ge=1, le=120)


class PlanUpload(BaseModel):
    issue_month: str
    workspace: str = "trial"
    rows: list[dict[str, Any]] = Field(min_length=1, max_length=250_000)


class PlanRevisionLine(BaseModel):
    model_code: str = Field(min_length=1, max_length=40)
    target_month: str
    quantity: float = Field(ge=0, le=1_000_000_000)


class PlanRevisionUpdate(BaseModel):
    rows: list[PlanRevisionLine] = Field(min_length=1, max_length=250_000)


class ActualUpload(BaseModel):
    period: str
    workspace: str = "trial"
    rows: list[dict[str, Any]] = Field(min_length=1, max_length=250_000)


class StockUpload(BaseModel):
    period: str
    workspace: str = "trial"
    rows: list[dict[str, Any]] = Field(min_length=1, max_length=250_000)


class TransferInput(BaseModel):
    workspace: str = "trial"
    material_id: int
    from_supplier_id: int
    to_supplier_id: int
    quantity: float = Field(gt=0)
    transfer_month: str
    note: str = Field(default="", max_length=2000)


class POInput(BaseModel):
    workspace: str = "trial"
    supplier_part_id: int
    quantity_ordered: float = Field(gt=0)
    expected_month: str
    ncnr: bool = False


class POReceipt(BaseModel):
    quantity_received: float = Field(gt=0)
    received_month: str


def _validate_workspace(value: str) -> str:
    if value not in WORKSPACES:
        raise HTTPException(status_code=422, detail="Workspace must be trial or production")
    return value


def _user_json(user: AppUser) -> dict:
    return {
        "id": user.id, "username": user.username, "full_name": user.full_name,
        "role": user.role, "can_view_prices": user.can_view_prices, "enabled": user.enabled,
    }


def _read_xlsx(raw: bytes) -> list[dict[str, Any]]:
    if not raw or len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413 if len(raw) > MAX_UPLOAD_BYTES else 422, detail="Upload must be a non-empty Excel file smaller than 20 MB")
    try:
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        sheet = workbook.active
        values = sheet.iter_rows(values_only=True)
        headers = []
        for value in next(values, ()):
            try:
                headers.append(_month_text(_month(value)))
                continue
            except HTTPException:
                pass
            headers.append(str(value or "").strip().lower().replace(" ", "_"))
        if not headers or len(headers) > 40:
            raise HTTPException(status_code=422, detail="The first worksheet must have column headers in row 1")
        if len(set(headers)) != len(headers):
            raise HTTPException(status_code=422, detail="The Excel sheet has duplicate column headers")
        rows = []
        for values_row in values:
            if not any(value is not None for value in values_row):
                continue
            rows.append({headers[i]: values_row[i] if i < len(values_row) else None for i in range(len(headers))})
        workbook.close()
        if not rows:
            raise HTTPException(status_code=422, detail="The Excel sheet has no data rows")
        return rows
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=422, detail="Could not read the Excel workbook; use a valid .xlsx file") from exc


SUPPLIER_PART_COLUMN_ALIASES = {
    "material_code": "material_code", "part_code": "material_code", "item_code": "material_code",
    "supplier_name": "supplier_name", "supplier": "supplier_name", "vendor": "supplier_name", "vendor_name": "supplier_name",
    "tier": "tier", "parent_supplier_name": "parent_supplier_name", "parent_supplier": "parent_supplier_name",
    "sob_percent": "sob_percent", "sob": "sob_percent", "sob_%": "sob_percent", "sob_(%)": "sob_percent",
    "unit_price": "unit_price", "unit_cost": "unit_price", "price": "unit_price",
    "currency": "currency",
    "lead_time_days": "lead_time_days", "lead_time": "lead_time_days", "lead_time_(days)": "lead_time_days",
    "moq": "moq", "buffer_days": "buffer_days", "buffer": "buffer_days",
    "price_mode": "price_mode", "price_interval_months": "price_interval_months",
}


SUPPLIER_MATERIAL_COLUMN_ALIASES = {
    "code": "code", "material_code": "code", "part_code": "code", "item_code": "code",
    "name": "name", "material_name": "name", "part_name": "name", "description": "name",
    "unit": "uom", "uom": "uom", "unit_price": "unit_price", "unit_cost": "unit_price", "price": "unit_price", "currency": "currency",
    "lead_time_days": "lead_time_days", "lead_time": "lead_time_days", "lead_time_(days)": "lead_time_days",
    "buffer_days": "buffer_days", "buffer": "buffer_days", "buffer_stock_days": "buffer_days",
}
SUPPLIER_BOM_COLUMN_ALIASES = {
    "parent_code": "parent_code", "parent": "parent_code", "parent_part_code": "parent_code",
    "code": "code", "part_code": "code", "material_code": "code", "component_code": "code",
    "quantity": "quantity", "qty": "quantity", "quantity_per_parent": "quantity",
    "unit_price": "unit_price", "unit_cost": "unit_price", "price": "unit_price", "source": "source",
}


TIER_MATERIAL_COLUMN_ALIASES = {
    "code": "code", "material_code": "code", "part_code": "code", "item_code": "code",
    "name": "name", "material_name": "name", "part_name": "name", "description": "name", "category": "category",
    "unit": "uom", "uom": "uom", "unit_price": "unit_price", "unit_cost": "unit_price", "price": "unit_price",
    "reference_price": "unit_price", "currency": "currency",
}


def _price_cells(row: dict, row_number: int, can_price: bool) -> tuple[float, str]:
    try:
        price = float(row.get("unit_price") or 0)
    except (TypeError, ValueError) as exc:
        raise HTTPException(status_code=422, detail=f"Row {row_number}: unit price must be a number") from exc
    if price < 0 or not math.isfinite(price):
        raise HTTPException(status_code=422, detail=f"Row {row_number}: unit price must be zero or more")
    if price and not can_price:
        raise HTTPException(status_code=403, detail="Price access is required to upload prices")
    currency = str(row.get("currency") or "INR").strip().upper()
    if not re.fullmatch(r"[A-Z]{3}", currency):
        raise HTTPException(status_code=422, detail=f"Row {row_number}: currency must be a 3-letter code such as INR")
    return price, currency


def _read_flexible_sheet(raw: bytes, aliases: dict[str, str], required: set[str]) -> list[dict[str, Any]]:
    return _read_supplier_part_sheet(raw, aliases, required)


def _read_supplier_part_sheet(raw: bytes, aliases: dict[str, str] | None = None, required: set[str] | None = None) -> list[dict[str, Any]]:
    """Relationship sheet as the user keeps it: heading row anywhere in the first 30 rows
    (title rows above are skipped), column names like "Material code" / "Supplier Name" /
    "Unit cost" accepted, other columns (Material name, Category, On hand ...) ignored."""
    if not raw or len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413 if len(raw) > MAX_UPLOAD_BYTES else 422, detail="Upload must be a non-empty Excel file smaller than 20 MB")
    try:
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        lines = list(workbook.active.iter_rows(values_only=True))
        workbook.close()
    except Exception as exc:
        raise HTTPException(status_code=422, detail="Could not read the Excel workbook; use a valid .xlsx file") from exc

    def key(value: Any) -> str:
        return re.sub(r"\s+", "_", str(value or "").strip().lower())

    aliases = aliases or SUPPLIER_PART_COLUMN_ALIASES
    required = required or {"material_code", "supplier_name"}
    for header_index, line in enumerate(lines[:30]):
        columns = {index: aliases[key(value)] for index, value in enumerate(line) if key(value) in aliases}
        if required <= set(columns.values()):
            break
    else:
        names = " and ".join(f"'{name.replace('_', ' ').capitalize()}'" for name in sorted(required))
        raise HTTPException(status_code=422, detail=f"Could not find the heading row: the sheet needs {names} columns")
    if len(set(columns.values())) != len(columns):
        raise HTTPException(status_code=422, detail="The Excel sheet has the same column twice (e.g. 'Unit cost' and 'Unit price')")
    rows = []
    for offset, line in enumerate(lines[header_index + 1:]):
        row = {name: (line[index] if index < len(line) else None) for index, name in columns.items()}
        if any(value is not None and str(value).strip() != "" for value in row.values()):
            rows.append({**row, "_excel_row": header_index + 2 + offset})
    if not rows:
        raise HTTPException(status_code=422, detail="The Excel sheet has no data rows")
    return rows


def _required_columns(rows: list[dict[str, Any]], required: set[str]) -> None:
    missing = required - set(rows[0])
    if missing:
        raise HTTPException(status_code=422, detail=f"Excel sheet is missing required columns: {', '.join(sorted(missing))}")


def _plan_rows_from_excel(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    aliases = {"model_month": "model_name", "model": "model_name", "qty": "quantity"}
    normalized_rows = [
        {aliases.get(key, key): value for key, value in row.items()}
        for row in rows
    ]
    _required_columns(normalized_rows, {"model_code"})
    if "target_month" in normalized_rows[0]:
        _required_columns(normalized_rows, {"target_month"})
        if "quantity" not in normalized_rows[0]:
            raise HTTPException(status_code=422, detail="Plan workbook needs a quantity column")
        for row in normalized_rows:
            row["target_month"] = _month_text(_month(row.get("target_month")))
        return normalized_rows
    month_columns: list[tuple[str, str]] = []
    for column in normalized_rows[0]:
        try:
            month_columns.append((column, _month_text(_month(column))))
        except HTTPException:
            continue
    if not month_columns:
        raise HTTPException(status_code=422, detail="Plan workbook needs target_month and quantity columns, or one column per month")
    if len({month for _, month in month_columns}) != len(month_columns):
        raise HTTPException(status_code=422, detail="Plan workbook has duplicate month columns")
    expanded = []
    for row in normalized_rows:
        for column, month in month_columns:
            quantity = row.get(column)
            if quantity is None or str(quantity).strip() == "":
                continue
            expanded.append({"model_code": row.get("model_code"), "target_month": month, "quantity": quantity})
    if not expanded:
        raise HTTPException(status_code=422, detail="Plan workbook has no monthly quantity values")
    return expanded


def _check_price_access(request: Request) -> bool:
    user = request.state.user
    return user["role"] == "Admin" or bool(user["can_view_prices"])


def _check_supply_role(request: Request) -> None:
    if request.state.user["role"] not in SUPPLY_ROLES:
        raise HTTPException(status_code=403, detail="Supplier-control access is required")


def _criticality(days: int) -> str:
    if days <= 15:
        return "Very Low"
    if days <= 30:
        return "Medium"
    if days <= 60:
        return "High"
    return "Very High"


def _price_status(link: SupplierPart, current_month: str) -> str:
    if link.price_mode == "spot":
        return "Spot"
    last = _month(link.price_updated_month)
    current = _month(current_month)
    months_since_update = (current.year - last.year) * 12 + current.month - last.month
    interval = 1 if link.price_mode == "fixed_monthly" else link.price_interval_months
    return "Update due" if months_since_update >= interval else "Current"


def _item_map(db: Session) -> tuple[dict[int, Item], dict[int, list[tuple[int, float]]]]:
    items = {item.id: item for item in db.scalars(select(Item)).all()}
    children: dict[int, list[tuple[int, float]]] = {}
    for row in db.execute(select(BOMComponent.parent_item_id, BOMComponent.child_item_id, BOMComponent.quantity)):
        children.setdefault(row.parent_item_id, []).append((row.child_item_id, row.quantity))
    return items, children


def _rollup_model_bom(
    roots: list[ModelBOM], children: dict[int, list[tuple[int, float]]]
) -> dict[int, float]:
    result: dict[int, float] = {}

    def walk(item_id: int, qty: float, ancestors: frozenset[int]) -> None:
        if item_id in ancestors:
            raise HTTPException(status_code=422, detail="Circular BOM detected")
        result[item_id] = result.get(item_id, 0) + qty
        for child_id, child_qty in children.get(item_id, ()):
            walk(child_id, qty * child_qty, ancestors | {item_id})

    for root in roots:
        walk(root.material_id, root.quantity, frozenset())
    return result


def _planning_report(db: Session, workspace: str, start_month: str, end_month: str) -> list[dict]:
    start = _month(start_month)
    end = _month(end_month)
    if end < start or (end.year - start.year) * 12 + end.month - start.month > 240:
        raise HTTPException(status_code=422, detail="Choose a valid range no longer than 240 months")
    models = {m.id: m for m in db.scalars(select(ProductModel)).all()}
    suppliers = {s.id: s for s in db.scalars(select(Supplier)).all()}
    items, children = _item_map(db)
    roots_by_model: dict[int, list[ModelBOM]] = {}
    for root in db.scalars(select(ModelBOM)).all():
        roots_by_model.setdefault(root.model_id, []).append(root)
    plan_by_cycle_month: dict[tuple[int, str], dict[int, float]] = {}
    cycles_all = db.scalars(
        select(PlanningCycle).where(PlanningCycle.workspace == workspace, PlanningCycle.issue_month <= end_month)
        .order_by(PlanningCycle.issue_month, PlanningCycle.revision)
    ).all()
    plan_by_cycle_month.clear()
    cycle_ids = [cycle.id for cycle in cycles_all]
    if cycle_ids:
        for row in db.scalars(select(PlanLine).where(PlanLine.cycle_id.in_(cycle_ids))).all():
            plan_by_cycle_month.setdefault((row.cycle_id, row.target_month), {})[row.model_id] = row.quantity
    material_usage_by_model = {
        model_id: _rollup_model_bom(roots_by_model.get(model_id, []), children)
        for model_id in models
    }
    actuals = db.scalars(select(ActualProduction).where(
        ActualProduction.workspace == workspace,
        ActualProduction.period >= start_month,
        ActualProduction.period <= end_month,
    )).all()
    actual_by_key = {(a.period, a.model_id): a for a in actuals}
    supplier_parts = db.scalars(select(SupplierPart).order_by(SupplierPart.material_id, SupplierPart.tier, SupplierPart.id)).all()
    supplier_parts_by_material: dict[int, list[SupplierPart]] = {}
    for link in supplier_parts:
        supplier_parts_by_material.setdefault(link.material_id, []).append(link)
    stock_rows = db.scalars(select(SupplierStockSnapshot).where(
        SupplierStockSnapshot.workspace == workspace,
        SupplierStockSnapshot.report_month <= end_month,
    ).order_by(SupplierStockSnapshot.report_month)).all()
    stock_by_month_link = {(s.report_month, s.supplier_part_id): s for s in stock_rows}
    stock_history_by_link: dict[int, list[SupplierStockSnapshot]] = {}
    for stock in stock_rows:
        stock_history_by_link.setdefault(stock.supplier_part_id, []).append(stock)
    own_stock_rows = db.scalars(select(OwnStockSnapshot).where(
        OwnStockSnapshot.workspace == workspace,
        OwnStockSnapshot.period <= end_month,
    ).order_by(OwnStockSnapshot.period)).all()
    pos = db.scalars(select(PlanningPO).where(
        PlanningPO.workspace == workspace,
        PlanningPO.status.in_(["Planned", "Issued", "Partially received"]),
        PlanningPO.expected_month <= end_month,
    )).all()
    po_by_month_link: dict[tuple[str, int], float] = {}
    for po in pos:
        po_by_month_link[(po.expected_month, po.supplier_part_id)] = po_by_month_link.get((po.expected_month, po.supplier_part_id), 0) + max(0, po.quantity_ordered - po.quantity_received)
    transfers = db.scalars(select(MaterialTransfer).where(
        MaterialTransfer.workspace == workspace,
        MaterialTransfer.status == "Planned",
        MaterialTransfer.transfer_month <= end_month,
    )).all()
    link_ids_by_material_supplier: dict[tuple[int, int], list[int]] = {}
    for link in supplier_parts:
        link_ids_by_material_supplier.setdefault((link.material_id, link.supplier_id), []).append(link.id)
    transfer_delta_by_month_link: dict[tuple[str, int], float] = {}
    for transfer in transfers:
        for link_id in link_ids_by_material_supplier.get((transfer.material_id, transfer.from_supplier_id), []):
            key = (transfer.transfer_month, link_id)
            transfer_delta_by_month_link[key] = transfer_delta_by_month_link.get(key, 0) - transfer.quantity
        for link_id in link_ids_by_material_supplier.get((transfer.material_id, transfer.to_supplier_id), []):
            key = (transfer.transfer_month, link_id)
            transfer_delta_by_month_link[key] = transfer_delta_by_month_link.get(key, 0) + transfer.quantity
    report: list[dict] = []
    projected: dict[int, float] = {}
    own_stock_current: dict[int, float] = {}
    own_stock_by_month: dict[str, dict[int, float]] = {}
    for stock in own_stock_rows:
        own_stock_by_month.setdefault(stock.period, {})[stock.material_id] = stock.quantity
        if stock.period <= start_month:
            own_stock_current[stock.material_id] = stock.quantity
    material_models: dict[int, set[int]] = {}
    for model_id, usage in material_usage_by_model.items():
        for material_id in usage:
            material_models.setdefault(material_id, set()).add(model_id)
    period = start
    while period <= end:
        month = _month_text(period)
        own_stock_current.update(own_stock_by_month.get(month, {}))
        own_stock_opening = dict(own_stock_current)
        applicable_cycles = [c for c in cycles_all if c.issue_month <= month]
        cycle = max(applicable_cycles, key=lambda c: (c.issue_month, c.revision), default=None)
        plan_demand_by_material: dict[int, float] = {}
        actual_demand_by_material: dict[int, float] = {}
        actual_models_by_material: dict[int, set[int]] = {}
        for model_id, model in models.items():
            planned_qty = plan_by_cycle_month.get((cycle.id, month), {}).get(model_id, 0) if cycle else 0
            actual = actual_by_key.get((month, model_id))
            rolled = material_usage_by_model.get(model_id, {})
            for material_id, unit_qty in rolled.items():
                plan_demand_by_material[material_id] = plan_demand_by_material.get(material_id, 0) + planned_qty * unit_qty
                if actual:
                    actual_demand_by_material[material_id] = actual_demand_by_material.get(material_id, 0) + actual.actual_quantity * unit_qty
                    actual_models_by_material.setdefault(material_id, set()).add(model_id)
        for material_id in set(plan_demand_by_material) | set(actual_demand_by_material):
            plan_demand = plan_demand_by_material.get(material_id, 0)
            expected_actual_models = material_models.get(material_id, set())
            actual_demand = (
                actual_demand_by_material.get(material_id, 0)
                if expected_actual_models and actual_models_by_material.get(material_id, set()) >= expected_actual_models
                else None
            )
            material_demand = actual_demand if actual_demand is not None else plan_demand
            links = supplier_parts_by_material.get(material_id, [])
            groups: dict[tuple[str, int | None], list[SupplierPart]] = {}
            for link in links:
                groups.setdefault((link.tier, link.parent_supplier_id), []).append(link)
            direct_links = groups.get(("Direct", None), [])
            if not groups:
                groups[("Unassigned", None)] = []
            own_stock_used = 0.0
            material = items.get(material_id)
            for (tier, parent_id), group_links in groups.items():
                shares = sum(link.sob_percent for link in group_links)
                for link in group_links:
                    snapshot = stock_by_month_link.get((month, link.id))
                    if snapshot:
                        projected[link.id] = (
                            snapshot.available_quantity
                            if not snapshot.arrival_month or snapshot.arrival_month <= month
                            else 0
                        )
                    elif link.id not in projected:
                        prior = [row for row in stock_history_by_link.get(link.id, []) if row.report_month <= month]
                        latest = prior[-1] if prior else None
                        projected[link.id] = (
                            latest.available_quantity
                            if latest and (not latest.arrival_month or latest.arrival_month <= month)
                            else 0
                        )
                    incoming = po_by_month_link.get((month, link.id), 0)
                    transfer_delta = transfer_delta_by_month_link.get((month, link.id), 0)
                    allocated_demand = material_demand * link.sob_percent / 100
                    allocated_plan = plan_demand * link.sob_percent / 100
                    allocated_actual = actual_demand * link.sob_percent / 100 if actual_demand is not None else None
                    own_stock_allocated = (
                        own_stock_opening.get(material_id, 0) * link.sob_percent / 100
                        if tier == "Direct" and parent_id is None else 0
                    )
                    supplier_balance = projected[link.id] + incoming + transfer_delta - allocated_demand
                    balance = supplier_balance + own_stock_allocated
                    own_stock_used += min(own_stock_allocated, max(0, -supplier_balance))
                    daily_rate = allocated_demand / 30.4375
                    buffer_target = daily_rate * link.buffer_days
                    shortage = max(0, buffer_target - balance)
                    excess = max(0, balance - buffer_target)
                    recommended_po_qty = math.ceil(shortage / link.moq) * link.moq if shortage > 0 and link.moq > 0 else shortage
                    supplier = suppliers.get(link.supplier_id)
                    if material and supplier:
                        coverage_days = max(0, balance) / daily_rate if daily_rate > 0 else 0
                        report.append({
                            "month": month,
                            "material_id": material_id,
                            "material_code": material.code,
                            "material_name": material.name,
                            "uom": material.uom,
                            "supplier_id": supplier.id,
                            "supplier_name": supplier.name,
                            "supplier_code": supplier.code,
                            "parent_supplier_name": suppliers[parent_id].name if parent_id in suppliers else None,
                            "tier": link.tier,
                            "sob_percent": link.sob_percent,
                            "plan_demand": allocated_plan,
                            "actual_demand": allocated_actual,
                            "own_stock": own_stock_opening.get(material_id),
                            "own_stock_allocated": own_stock_allocated,
                            "supplier_stock": projected[link.id],
                            "incoming_po": incoming,
                            "projected_balance": balance,
                            "coverage_days": coverage_days,
                            "buffer_target": buffer_target,
                            "shortage_qty": shortage,
                            "recommended_po_qty": recommended_po_qty,
                            "excess_qty": excess,
                            "unit_price": link.unit_price,
                            "currency": link.currency,
                            "shortage_value": shortage * link.unit_price,
                            "excess_value": excess * link.unit_price,
                            "lead_time_days": link.lead_time_days,
                            "criticality": _criticality(link.lead_time_days),
                            "plan_revision": f"Rev.{cycle.revision:02d}" if cycle else "No plan",
                        })
                    projected[link.id] = max(0, supplier_balance)
                if shares < 100 and material:
                    unassigned_share = 100 - shares
                    unassigned = material_demand * unassigned_share / 100
                    own_stock_allocated = (
                        own_stock_opening.get(material_id, 0) * unassigned_share / 100
                        if (tier == "Direct" and parent_id is None) or not links else 0
                    )
                    unassigned_shortage = max(0, unassigned - own_stock_allocated)
                    own_stock_used += min(own_stock_allocated, unassigned)
                    report.append({
                        "month": month, "material_id": material_id, "material_code": material.code,
                        "material_name": material.name, "uom": material.uom, "supplier_id": None,
                        "supplier_name": "Unassigned SOB", "supplier_code": "—", "tier": tier,
                        "parent_supplier_name": suppliers[parent_id].name if parent_id in suppliers else None,
                        "sob_percent": unassigned_share, "plan_demand": plan_demand * unassigned_share / 100,
                        "supplier_stock": 0,
                        "actual_demand": (actual_demand * unassigned_share / 100) if actual_demand is not None else None,
                        "own_stock": own_stock_opening.get(material_id),
                        "own_stock_allocated": own_stock_allocated,
                        "incoming_po": 0, "projected_balance": own_stock_allocated - unassigned,
                        "coverage_days": 0, "buffer_target": 0,
                        "shortage_qty": unassigned_shortage, "recommended_po_qty": unassigned_shortage,
                        "excess_qty": 0, "unit_price": material.unit_cost,
                        "currency": material.currency, "shortage_value": unassigned_shortage * material.unit_cost,
                        "excess_value": 0, "lead_time_days": 0, "criticality": "Very Low",
                        "plan_revision": f"Rev.{cycle.revision:02d}" if cycle else "No plan",
                    })
            own_stock_current[material_id] = max(0, own_stock_opening.get(material_id, 0) - own_stock_used)
        period = _plus_months(period, 1)
    report.sort(key=lambda row: (
        -(row["shortage_qty"] > 0),
        -{"Very High": 4, "High": 3, "Medium": 2, "Very Low": 1}.get(row["criticality"], 0),
        -row["shortage_qty"], row["month"], row["material_code"], row["supplier_name"],
    ))
    return report


def _branded_workbook(title: str, headers: list[str], rows: Iterable[Iterable[Any]]) -> io.BytesIO:
    logo_path = Path(__file__).resolve().parent / "static" / "logo.png"
    with tempfile.TemporaryDirectory(prefix="dtc-export-") as temp_dir:
        workbook_path = Path(temp_dir) / "report.xlsx"
        workbook = xlsxwriter.Workbook(str(workbook_path), {"constant_memory": True})
        sheet = workbook.add_worksheet("DTC report")
        title_format = workbook.add_format({"bold": True, "font_name": "Aptos Display", "font_size": 16, "font_color": "#12263A"})
        subtitle_format = workbook.add_format({"font_name": "Aptos", "font_size": 11, "font_color": "#425466"})
        header_format = workbook.add_format({"bold": True, "font_name": "Aptos", "font_color": "#FFFFFF", "bg_color": "#12263A", "text_wrap": True})
        if logo_path.is_file():
            sheet.insert_image("A1", str(logo_path), {"x_scale": 0.45, "y_scale": 0.45, "object_position": 1})
        sheet.write_string("D1", "Design to Cost", title_format)
        sheet.write_string("D2", f"Supply Chain Management · {title}", subtitle_format)
        sheet.write_row("A4", headers, header_format)
        row_index = 4
        for row in rows:
            for column_index, value in enumerate(row):
                if value is None:
                    sheet.write_blank(row_index, column_index, None)
                elif isinstance(value, bool):
                    sheet.write_boolean(row_index, column_index, value)
                elif isinstance(value, (int, float)):
                    sheet.write_number(row_index, column_index, value)
                else:
                    sheet.write_string(row_index, column_index, str(value))
            row_index += 1
        sheet.freeze_panes(4, 0)
        sheet.autofilter(3, 0, max(3, row_index - 1), len(headers) - 1)
        sheet.set_default_row(18)
        workbook.close()
        return io.BytesIO(workbook_path.read_bytes())


def _report_workbook(rows: list[dict], title: str) -> io.BytesIO:
    return _branded_workbook(title, REPORT_HEADERS, ([
        row.get("month"), row.get("material_code"), row.get("material_name"), row.get("supplier_name"),
        row.get("parent_supplier_name"), row.get("tier"), row.get("sob_percent"), row.get("plan_demand"), row.get("actual_demand"),
        row.get("own_stock"), row.get("own_stock_allocated"), row.get("supplier_stock"), row.get("incoming_po"), row.get("projected_balance"),
        row.get("coverage_days"), row.get("buffer_target"), row.get("shortage_qty"), row.get("recommended_po_qty"),
        row.get("excess_qty"), row.get("unit_price"), row.get("shortage_value"), row.get("excess_value"),
        row.get("lead_time_days"), row.get("criticality"), row.get("plan_revision"),
    ] for row in rows))


def _without_prices(rows: list[dict], request: Request) -> list[dict]:
    if _check_price_access(request):
        return rows
    fields = {"unit_price", "unit_cost", "shortage_value", "excess_value"}
    return [{key: value for key, value in row.items() if key not in fields} for row in rows]


def _generic_workbook(title: str, headers: list[str], rows: list[list[Any]]) -> io.BytesIO:
    return _branded_workbook(title, headers, rows)


def _bom_tree_workbook(title: str, headers: list[str], rows: list[dict]) -> io.BytesIO:
    """BOM tree as an Excel outline: each row's outline level = its tree depth, so Excel's
    own +/- buttons open and close the levels; rows that have children are bold."""
    logo_path = Path(__file__).resolve().parent / "static" / "logo.png"
    with tempfile.TemporaryDirectory(prefix="dtc-export-") as temp_dir:
        workbook_path = Path(temp_dir) / "bom-tree.xlsx"
        workbook = xlsxwriter.Workbook(str(workbook_path))
        sheet = workbook.add_worksheet("BOM tree")
        title_format = workbook.add_format({"bold": True, "font_name": "Aptos Display", "font_size": 16, "font_color": "#12263A"})
        subtitle_format = workbook.add_format({"font_name": "Aptos", "font_size": 11, "font_color": "#425466"})
        header_format = workbook.add_format({"bold": True, "font_name": "Aptos", "font_color": "#FFFFFF", "bg_color": "#12263A", "text_wrap": True, "valign": "vcenter"})
        parent_format = workbook.add_format({"bold": True, "font_name": "Aptos", "font_color": "#12263A", "bottom": 1, "bottom_color": "#C9D6E0"})
        leaf_format = workbook.add_format({"font_name": "Aptos", "font_color": "#4F6275", "bottom": 1, "bottom_color": "#E4EBF0"})
        if logo_path.is_file():
            sheet.insert_image("A1", str(logo_path), {"x_scale": 0.45, "y_scale": 0.45, "object_position": 1})
        sheet.write_string("D1", "Design to Cost", title_format)
        sheet.write_string("D2", f"Supply Chain Management · {title}", subtitle_format)
        sheet.write_row(3, 0, headers, header_format)
        sheet.outline_settings(True, False, False, False)  # +/- button on the parent row above its lines
        for offset, row in enumerate(rows):
            row_index = 4 + offset
            cell_format = parent_format if row["has_children"] else leaf_format
            # Opens collapsed: only top-level rows show, each with "+".
            sheet.set_row(row_index, None, cell_format, {
                "level": min(row["depth"], 7), "hidden": row["depth"] > 0, "collapsed": row["has_children"],
            })
            for column_index, value in enumerate(row["values"]):
                if value is None or value == "":
                    sheet.write_blank(row_index, column_index, None, cell_format)
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    sheet.write_number(row_index, column_index, value, cell_format)
                else:
                    sheet.write_string(row_index, column_index, str(value), cell_format)
        sheet.set_column(0, 0, 9)
        sheet.set_column(1, len(headers) - 1, 16)
        sheet.freeze_panes(4, 0)
        sheet.autofilter(3, 0, max(3, 3 + len(rows)), len(headers) - 1)
        workbook.close()
        return io.BytesIO(workbook_path.read_bytes())


def _bom_tree_pdf_layout(headers: list[str], rows: list[dict]) -> dict:
    """BOM tree PDF layout: every row on one line. Column widths fit the longest text; if
    the table is wider than A4 landscape the font shrinks (down to 5 pt), then A3 is used;
    an extremely wide table gets a page as wide as the table (print with "fit to page")."""
    def text(header: str, value: Any) -> str:
        if value is None or value == "":
            return ""
        if isinstance(value, float) and header == "Part price":
            return f"{value:,.2f}"
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    data = [list(headers)] + [[text(header, value) for header, value in zip(headers, row["values"])] for row in rows]
    bold_rows = [0] + [index + 1 for index, row in enumerate(rows) if row["has_children"]]
    padding = 8  # points, left + right cell padding

    def natural_widths(size: float) -> list[float]:
        return [max(pdfmetrics.stringWidth(line[column], "Helvetica-Bold", size) for line in data) + padding
                for column in range(len(headers))]

    font_size = 7.0
    for pagesize in (landscape(A4), landscape(A3)):
        available = pagesize[0] - 20 * mm
        widths = natural_widths(font_size)
        if sum(widths) > available:
            font_size = max(5.0, font_size * (available - padding * len(headers)) / (sum(widths) - padding * len(headers)))
            widths = natural_widths(font_size)
        if sum(widths) <= available or pagesize == landscape(A3):
            if sum(widths) > available:
                pagesize = (sum(widths) + 20 * mm, pagesize[1])
                available = sum(widths)
            spare = max(0.0, available - sum(widths)) / len(headers)
            return {"data": data, "widths": [width + spare for width in widths], "font_size": font_size,
                    "bold_rows": bold_rows, "pagesize": pagesize}
        font_size = 7.0
    raise AssertionError("unreachable")


def _bom_tree_pdf(title: str, headers: list[str], rows: list[dict]) -> io.BytesIO:
    layout = _bom_tree_pdf_layout(headers, rows)
    output = io.BytesIO()
    document = SimpleDocTemplate(output, pagesize=layout["pagesize"], rightMargin=10 * mm, leftMargin=10 * mm, topMargin=10 * mm, bottomMargin=10 * mm)
    styles = getSampleStyleSheet()
    story = []
    logo_path = Path(__file__).resolve().parent / "static" / "logo.png"
    try:
        story.append(Image(str(logo_path), width=16 * mm, height=16 * mm))
    except OSError:
        pass
    story.extend([Paragraph("Design to Cost", styles["Title"]), Paragraph(f"Supply Chain Management · {title}", styles["Heading2"]), Spacer(1, 4 * mm)])
    table = Table(layout["data"], repeatRows=1, colWidths=layout["widths"])
    commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#12263a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), layout["font_size"]),
        ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#4f6275")),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), .25, colors.HexColor("#dce4ea")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f8fa")]),
    ]
    for row_index in layout["bold_rows"]:
        commands.append(("FONTNAME", (0, row_index), (-1, row_index), "Helvetica-Bold"))
        if row_index:
            commands.append(("TEXTCOLOR", (0, row_index), (-1, row_index), colors.HexColor("#12263a")))
    table.setStyle(TableStyle(commands))
    story.append(table)
    document.build(story)
    output.seek(0)
    return output


def _generic_pdf(title: str, headers: list[str], rows: list[list[Any]]) -> io.BytesIO:
    output = io.BytesIO()
    document = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=10 * mm, leftMargin=10 * mm, topMargin=10 * mm, bottomMargin=10 * mm)
    styles = getSampleStyleSheet()
    story = []
    logo_path = __import__("pathlib").Path(__file__).resolve().parent / "static" / "logo.png"
    try:
        story.append(Image(str(logo_path), width=16 * mm, height=16 * mm))
    except OSError:
        pass
    story.extend([Paragraph("Design to Cost", styles["Title"]), Paragraph(f"Supply Chain Management · {title}", styles["Heading2"]), Spacer(1, 4 * mm)])
    data = [headers, *[[str(value if value is not None else "") for value in row] for row in rows]]
    width = 277 * mm / max(1, len(headers))
    table = Table(data, repeatRows=1, colWidths=[width] * len(headers))
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#12263a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), .25, colors.HexColor("#dce4ea")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f8fa")]),
    ]))
    story.append(table)
    document.build(story)
    output.seek(0)
    return output


def _report_pdf(rows: list[dict], title: str) -> io.BytesIO:
    output = io.BytesIO()
    document = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=10 * mm, leftMargin=10 * mm, topMargin=10 * mm, bottomMargin=10 * mm)
    styles = getSampleStyleSheet()
    story = []
    logo_path = __import__("pathlib").Path(__file__).resolve().parent / "static" / "logo.png"
    try:
        story.append(Image(str(logo_path), width=18 * mm, height=18 * mm))
    except OSError:
        pass
    story.extend([Paragraph("Design to Cost", styles["Title"]), Paragraph(f"Supply Chain Management · {title}", styles["Heading2"]), Spacer(1, 4 * mm)])
    data = [["Month", "Part code / name", "Supplier", "Tier", "Plan demand", "Actual demand", "Balance", "Shortage", "Excess", "Lead", "Criticality"]]
    for row in rows:
        data.append([
            row.get("month", ""), f"{row.get('material_code', '')} / {row.get('material_name', '')}",
            row.get("supplier_name", ""), row.get("tier", ""), f"{row.get('plan_demand', 0):,.2f}",
            "" if row.get("actual_demand") is None else f"{row['actual_demand']:,.2f}",
            f"{row.get('projected_balance', 0):,.2f}", f"{row.get('shortage_qty', 0):,.2f}",
            f"{row.get('excess_qty', 0):,.2f}", str(row.get("lead_time_days", 0)), row.get("criticality", ""),
        ])
    table = Table(data, repeatRows=1, colWidths=[16 * mm, 42 * mm, 35 * mm, 17 * mm, 20 * mm, 20 * mm, 20 * mm, 18 * mm, 18 * mm, 13 * mm, 22 * mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#12263a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), .25, colors.HexColor("#dce4ea")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f8fa")]),
        ("ALIGN", (4, 1), (9, -1), "RIGHT"),
    ]))
    story.append(table)
    document.build(story)
    output.seek(0)
    return output


def install_planning(app: FastAPI, session_factory: sessionmaker) -> None:
    def open_db() -> Session:
        return session_factory()

    with open_db() as db:
        legacy_models = db.scalars(select(ProductModel).where(ProductModel.product_id.is_not(None))).all()
        existing_links = {
            (link.product_id, link.model_id)
            for link in db.scalars(select(ProductModelProduct)).all()
        }
        for model in legacy_models:
            product = db.get(Product, model.product_id)
            link = (model.product_id, model.id)
            if product and product.code != MODEL_CATALOG_PRODUCT_CODE and link not in existing_links:
                db.add(ProductModelProduct(product_id=link[0], model_id=link[1]))
        db.commit()
        # supplier_materials gained item_id / tier_material_id (T1/T2/T3 links); add them to
        # databases created before this version. Existing rows stay the supplier's own codes.
        from sqlalchemy import inspect as sa_inspect, text as sa_text
        present = {column["name"] for column in sa_inspect(db.get_bind()).get_columns("supplier_materials")}
        for column, definition in (("item_id", "INTEGER"), ("tier_material_id", "INTEGER"),
                                   ("lead_time_days", "INTEGER NOT NULL DEFAULT 0"), ("buffer_days", "INTEGER NOT NULL DEFAULT 0")):
            if column not in present:
                db.execute(sa_text(f"ALTER TABLE supplier_materials ADD COLUMN {column} {definition}"))
        db.commit()

    @app.middleware("http")
    async def require_session_and_hide_prices(request: Request, call_next):
        path = request.url.path
        if not path.startswith("/api/") or path in {"/api/auth/setup", "/api/auth/setup/status", "/api/auth/login", "/api/health"}:
            return await call_next(request)
        if path.startswith("/api/rfqs"):
            return JSONResponse({"detail": "RFQ workflows have been retired"}, status_code=410)
        raw_token = request.cookies.get("dtc_session")
        if not raw_token:
            return JSONResponse({"detail": "Sign in is required"}, status_code=401)
        with open_db() as db:
            token_hash = hashlib.sha256(raw_token.encode()).hexdigest()
            login = db.scalar(select(LoginSession).where(LoginSession.token_hash == token_hash, LoginSession.expires_at > datetime.utcnow()))
            user = db.get(AppUser, login.user_id) if login else None
            if not login or not user or not user.enabled:
                return JSONResponse({"detail": "Session expired; sign in again"}, status_code=401)
            request.state.user = _user_json(user)
        supplier_control = (
            path.startswith("/api/suppliers")
            or path.startswith("/api/purchase-orders")
            or path.startswith("/api/stock-purchase-orders")
            or path.startswith("/api/material-requests")
            or (path.startswith("/api/items") and request.method != "GET")
        )
        if supplier_control and request.state.user["role"] not in SUPPLY_ROLES:
            return JSONResponse({"detail": "Supplier-control access is required"}, status_code=403)
        response = await call_next(request)
        if path.startswith("/api/") and response.headers.get("content-type", "").startswith("application/json") and not _check_price_access(request):
            body = b"".join([chunk async for chunk in response.body_iterator])
            try:
                payload = json.loads(body)
                hidden_keys = {"unit_cost", "unit_price", "total", "spend_by_currency", "shortage_value", "excess_value"}

                def hide(value):
                    if isinstance(value, dict):
                        return {k: hide(v) for k, v in value.items() if k not in hidden_keys}
                    if isinstance(value, list):
                        return [hide(v) for v in value]
                    return value

                body = json.dumps(hide(payload), separators=(",", ":")).encode()
            except (ValueError, TypeError):
                pass
            headers = dict(response.headers)
            headers.pop("content-length", None)
            response = Response(content=body, status_code=response.status_code, headers=headers, media_type="application/json")
        return response

    def current_user(request: Request) -> dict:
        return request.state.user

    def admin_only(request: Request) -> dict:
        user = current_user(request)
        if user["role"] != "Admin":
            raise HTTPException(status_code=403, detail="Admin access is required")
        return user

    @app.post("/api/auth/setup", status_code=201)
    def setup_admin(payload: SetupAdmin, request: Request, response: Response):
        with open_db() as db:
            if db.scalar(select(func.count(AppUser.id))):
                raise HTTPException(status_code=409, detail="Initial admin is already set up")
            user = AppUser(
                username=payload.username.lower(), full_name=payload.full_name.strip(),
                password_hash=_password_hash(payload.password), role="Admin", can_view_prices=True,
            )
            db.add(user)
            db.commit()
            db.refresh(user)
            token = secrets.token_urlsafe(32)
            db.add(LoginSession(token_hash=hashlib.sha256(token.encode()).hexdigest(), user_id=user.id, expires_at=datetime.utcnow() + timedelta(hours=12)))
            db.commit()
            response.set_cookie("dtc_session", token, httponly=True, secure=request.url.scheme == "https", samesite="strict", max_age=43200, path="/")
            return {"user": _user_json(user)}

    @app.get("/api/auth/setup/status")
    def setup_status():
        with open_db() as db:
            return {"setup_required": not bool(db.scalar(select(func.count(AppUser.id))))}

    @app.post("/api/auth/login")
    def login(payload: dict, request: Request, response: Response):
        username = str(payload.get("username", "")).strip().lower()
        password = str(payload.get("password", ""))
        with open_db() as db:
            user = db.scalar(select(AppUser).where(AppUser.username == username))
            if not user or not user.enabled or not _password_matches(password, user.password_hash):
                raise HTTPException(status_code=401, detail="Incorrect username or password")
            token = secrets.token_urlsafe(32)
            db.add(LoginSession(token_hash=hashlib.sha256(token.encode()).hexdigest(), user_id=user.id, expires_at=datetime.utcnow() + timedelta(hours=12)))
            db.commit()
            response.set_cookie("dtc_session", token, httponly=True, secure=request.url.scheme == "https", samesite="strict", max_age=43200, path="/")
            return {"user": _user_json(user)}

    @app.post("/api/auth/logout")
    def logout(request: Request, response: Response):
        raw_token = request.cookies.get("dtc_session")
        if raw_token:
            with open_db() as db:
                login_session = db.scalar(select(LoginSession).where(LoginSession.token_hash == hashlib.sha256(raw_token.encode()).hexdigest()))
                if login_session:
                    db.delete(login_session)
                    db.commit()
        response.delete_cookie("dtc_session", path="/", samesite="strict")
        return {"status": "ok"}

    @app.get("/api/auth/me")
    def who_am_i(request: Request):
        with open_db() as db:
            return {"user": current_user(request), "setup_required": not bool(db.scalar(select(func.count(AppUser.id))))}

    @app.get("/api/users")
    def list_users(request: Request):
        admin_only(request)
        with open_db() as db:
            return [_user_json(user) for user in db.scalars(select(AppUser).order_by(AppUser.username)).all()]

    @app.post("/api/users", status_code=201)
    def create_user(payload: UserCreate, request: Request):
        admin_only(request)
        if payload.role not in ROLES:
            raise HTTPException(status_code=422, detail="Choose a valid user role")
        with open_db() as db:
            if db.scalar(select(AppUser.id).where(AppUser.username == payload.username.lower())):
                raise HTTPException(status_code=409, detail="That username already exists")
            user = AppUser(
                username=payload.username.lower(), full_name=payload.full_name.strip(),
                password_hash=_password_hash(payload.password), role=payload.role,
                can_view_prices=payload.can_view_prices,
            )
            db.add(user)
            db.commit()
            db.refresh(user)
            return _user_json(user)

    @app.patch("/api/users/{user_id}")
    def update_user(user_id: int, payload: dict, request: Request):
        actor = admin_only(request)
        allowed = {"role", "can_view_prices", "enabled", "password", "full_name"}
        if set(payload) - allowed:
            raise HTTPException(status_code=422, detail="Unsupported user update field")
        if payload.get("role") is not None and payload["role"] not in ROLES:
            raise HTTPException(status_code=422, detail="Choose a valid user role")
        if payload.get("password") is not None and len(payload["password"]) < 12:
            raise HTTPException(status_code=422, detail="Passwords must be at least 12 characters")
        with open_db() as db:
            user = db.get(AppUser, user_id)
            if not user:
                raise HTTPException(status_code=404, detail="User not found")
            if actor["id"] == user.id and payload.get("enabled") is False:
                raise HTTPException(status_code=422, detail="You cannot disable your own account")
            demotes_or_disables_admin = user.role == "Admin" and (
                payload.get("role") not in (None, "Admin") or payload.get("enabled") is False
            )
            if demotes_or_disables_admin and db.scalar(select(func.count(AppUser.id)).where(AppUser.role == "Admin", AppUser.enabled.is_(True))) <= 1:
                raise HTTPException(status_code=409, detail="At least one enabled Admin must remain")
            for key, value in payload.items():
                setattr(user, key if key != "password" else "password_hash", _password_hash(value) if key == "password" else value)
            db.commit()
            db.refresh(user)
            return _user_json(user)

    @app.get("/api/planning/products")
    def products(request: Request):
        current_user(request)
        with open_db() as db:
            models_by_product: dict[int, list[int]] = {}
            for link in db.scalars(select(ProductModelProduct)).all():
                models_by_product.setdefault(link.product_id, []).append(link.model_id)
            return [
                {"id": p.id, "code": p.code, "name": p.name, "model_ids": models_by_product.get(p.id, [])}
                for p in db.scalars(select(Product).where(Product.code != MODEL_CATALOG_PRODUCT_CODE).order_by(Product.code)).all()
            ]

    @app.post("/api/planning/products", status_code=201)
    def add_product(payload: ProductCreate, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            if payload.code.strip() == MODEL_CATALOG_PRODUCT_CODE:
                raise HTTPException(status_code=422, detail="This product code is reserved for the model catalog")
            if db.scalar(select(Product.id).where(func.lower(Product.code) == payload.code.strip().lower())):
                raise HTTPException(status_code=409, detail="Product code already exists")
            model_ids = list(dict.fromkeys(payload.model_ids))
            models = db.scalars(select(ProductModel).where(ProductModel.id.in_(model_ids))).all() if model_ids else []
            if len(models) != len(model_ids):
                raise HTTPException(status_code=422, detail="One or more selected models do not exist")
            product = Product(code=payload.code.strip(), name=payload.name.strip())
            db.add(product)
            db.flush()
            db.add_all(ProductModelProduct(product_id=product.id, model_id=model_id) for model_id in model_ids)
            db.commit()
            db.refresh(product)
            return {"id": product.id, "code": product.code, "name": product.name, "model_ids": model_ids}

    @app.patch("/api/planning/products/{product_id}")
    def update_product(product_id: int, payload: ProductCreate, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            product = db.get(Product, product_id)
            if not product or product.code == MODEL_CATALOG_PRODUCT_CODE:
                raise HTTPException(status_code=404, detail="Product not found")
            code = payload.code.strip()
            if code == MODEL_CATALOG_PRODUCT_CODE or db.scalar(select(Product.id).where(func.lower(Product.code) == code.lower(), Product.id != product_id)):
                raise HTTPException(status_code=409, detail="Product code already exists or is reserved")
            model_ids = list(dict.fromkeys(payload.model_ids))
            found = db.scalars(select(ProductModel.id).where(ProductModel.id.in_(model_ids))).all() if model_ids else []
            if len(found) != len(model_ids):
                raise HTTPException(status_code=422, detail="One or more selected models do not exist")
            product.code, product.name = code, payload.name.strip()
            db.execute(delete(ProductModelProduct).where(ProductModelProduct.product_id == product_id))
            db.add_all(ProductModelProduct(product_id=product_id, model_id=model_id) for model_id in model_ids)
            db.commit()
            return {"id": product.id, "code": product.code, "name": product.name, "model_ids": model_ids}

    @app.delete("/api/planning/products/{product_id}", status_code=204)
    def delete_product(product_id: int, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            product = db.get(Product, product_id)
            if not product or product.code == MODEL_CATALOG_PRODUCT_CODE:
                raise HTTPException(status_code=404, detail="Product not found")
            if db.scalar(select(ProductModel.id).where(ProductModel.product_id == product_id)):
                raise HTTPException(status_code=409, detail="This product owns legacy models; reassign those models before deleting it")
            db.execute(delete(ProductModelProduct).where(ProductModelProduct.product_id == product_id))
            db.delete(product)
            db.commit()
            return Response(status_code=204)

    @app.get("/api/planning/models")
    def models(request: Request):
        current_user(request)
        with open_db() as db:
            products_by_id = {p.id: p for p in db.scalars(select(Product)).all()}
            products_by_model: dict[int, list[int]] = {}
            for link in db.scalars(select(ProductModelProduct)).all():
                products_by_model.setdefault(link.model_id, []).append(link.product_id)
            models_with_bom = set(db.scalars(select(ModelBOM.model_id).distinct()).all())
            return [
                {
                    "id": m.id,
                    "has_bom": m.id in models_with_bom,
                    "product_id": (products_by_model.get(m.id) or [None])[0],
                    "product_ids": products_by_model.get(m.id, []),
                    "product_codes": [products_by_id[product_id].code for product_id in products_by_model.get(m.id, []) if product_id in products_by_id and products_by_id[product_id].code != MODEL_CATALOG_PRODUCT_CODE],
                    "product_names": [products_by_id[product_id].name for product_id in products_by_model.get(m.id, []) if product_id in products_by_id and products_by_id[product_id].code != MODEL_CATALOG_PRODUCT_CODE],
                    "product_code": ", ".join(products_by_id[product_id].code for product_id in products_by_model.get(m.id, []) if product_id in products_by_id and products_by_id[product_id].code != MODEL_CATALOG_PRODUCT_CODE),
                    "product_name": ", ".join(products_by_id[product_id].name for product_id in products_by_model.get(m.id, []) if product_id in products_by_id and products_by_id[product_id].code != MODEL_CATALOG_PRODUCT_CODE),
                    "code": m.code, "name": m.name
                }
                for m in db.scalars(select(ProductModel).order_by(ProductModel.code)).all()
            ]

    @app.post("/api/planning/models", status_code=201)
    def add_model(payload: ModelCreate, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            code = payload.code.strip()
            if db.scalar(select(ProductModel.id).where(func.lower(ProductModel.code) == code.lower())):
                raise HTTPException(status_code=409, detail="Model code already exists")
            if payload.product_id is not None and not db.get(Product, payload.product_id):
                raise HTTPException(status_code=404, detail="Product not found")
            catalog = db.scalar(select(Product).where(Product.code == MODEL_CATALOG_PRODUCT_CODE))
            if payload.product_id is None and catalog is None:
                catalog = Product(code=MODEL_CATALOG_PRODUCT_CODE, name="Model catalog")
                db.add(catalog)
                db.flush()
            model = ProductModel(product_id=payload.product_id or catalog.id, code=code, name=payload.name.strip())
            db.add(model)
            db.flush()
            product_ids = list(dict.fromkeys([*payload.product_ids, *([payload.product_id] if payload.product_id else [])]))
            if product_ids:
                found = db.scalars(select(Product.id).where(Product.id.in_(product_ids), Product.code != MODEL_CATALOG_PRODUCT_CODE)).all()
                if len(found) != len(product_ids):
                    raise HTTPException(status_code=422, detail="One or more selected products do not exist")
                db.add_all(ProductModelProduct(product_id=product_id, model_id=model.id) for product_id in product_ids)
            db.commit()
            db.refresh(model)
            return {"id": model.id, "product_id": product_ids[0] if product_ids else None, "product_ids": product_ids, "code": model.code, "name": model.name}

    @app.patch("/api/planning/models/{model_id}")
    def update_model(model_id: int, payload: ModelUpdate, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            model = db.get(ProductModel, model_id)
            if not model:
                raise HTTPException(status_code=404, detail="Model not found")
            code = payload.code.strip()
            if db.scalar(select(ProductModel.id).where(func.lower(ProductModel.code) == code.lower(), ProductModel.id != model_id)):
                raise HTTPException(status_code=409, detail="Model code already exists")
            product_ids = list(dict.fromkeys(payload.product_ids))
            found = db.scalars(select(Product.id).where(Product.id.in_(product_ids), Product.code != MODEL_CATALOG_PRODUCT_CODE)).all() if product_ids else []
            if len(found) != len(product_ids):
                raise HTTPException(status_code=422, detail="One or more selected products do not exist")
            model.code, model.name = code, payload.name.strip()
            db.execute(delete(ProductModelProduct).where(ProductModelProduct.model_id == model_id))
            db.add_all(ProductModelProduct(product_id=product_id, model_id=model_id) for product_id in product_ids)
            db.commit()
            return {"id": model.id, "code": model.code, "name": model.name, "product_ids": product_ids}

    @app.delete("/api/planning/models/{model_id}", status_code=204)
    def delete_model(model_id: int, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            model = db.get(ProductModel, model_id)
            if not model:
                raise HTTPException(status_code=404, detail="Model not found")
            if db.scalar(select(PlanLine.id).where(PlanLine.model_id == model_id)) or db.scalar(select(ActualProduction.id).where(ActualProduction.model_id == model_id)):
                raise HTTPException(status_code=409, detail="This model has plan or actual history and cannot be deleted")
            db.execute(delete(ModelBOM).where(ModelBOM.model_id == model_id))
            db.execute(delete(ProductModelProduct).where(ProductModelProduct.model_id == model_id))
            db.delete(model)
            db.commit()
            return Response(status_code=204)

    @app.get("/api/planning/materials")
    def materials(request: Request, offset: int = 0, limit: int = 100, search: str = "", purchased: bool = False):
        current_user(request)
        if limit < 1 or limit > 500 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–500 and a non-negative offset")
        with open_db() as db:
            statement = select(Item).order_by(Item.code)
            if purchased:
                statement = purchased_filter(statement)
            if search.strip():
                like = f"%{search.strip()}%"
                statement = statement.where((Item.code.ilike(like)) | (Item.name.ilike(like)))
            total = db.scalar(select(func.count()).select_from(statement.subquery())) or 0
            rows = db.scalars(statement.offset(offset).limit(limit)).all()
            can_price = _check_price_access(request)
            return {
                "total": total, "offset": offset, "limit": limit,
                "items": [{
                    "id": item.id, "code": item.code, "name": item.name, "category": item.category,
                    "uom": item.uom, "stock_on_hand": item.stock_on_hand, "reorder_level": item.reorder_level,
                    "buffer_days": getattr(item, "buffer_days", 0),
                    **({"unit_cost": item.unit_cost, "currency": item.currency} if can_price else {}),
                } for item in rows],
            }

    @app.get("/api/planning/material-categories")
    def material_categories(request: Request):
        current_user(request)
        with open_db() as db:
            names = db.scalars(select(Item.category).distinct()).all()
        unique: dict[str, str] = {}
        for name in names:
            text = (name or "").strip()
            if text:
                unique.setdefault(text.casefold(), text)
        return {"items": sorted(unique.values(), key=str.casefold)}

    @app.post("/api/planning/materials", status_code=201)
    def add_material(payload: MaterialInput, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        if payload.unit_cost and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to set a material price")
        with open_db() as db:
            code = payload.code.strip()
            if db.scalar(select(Item.id).where(func.lower(Item.code) == code.lower())):
                raise HTTPException(status_code=409, detail="Material code already exists")
            owner = code_owner(db, code)
            if owner:
                raise HTTPException(status_code=409, detail=f"{code} already exists in {owner} Material Master")
            item = Item(
                code=code, name=payload.name.strip(), category=payload.category.strip(),
                uom=payload.uom.strip(), stock_on_hand=payload.stock_on_hand,
                reorder_level=0, buffer_days=payload.buffer_days, unit_cost=payload.unit_cost, currency=payload.currency,
            )
            db.add(item)
            try:
                db.commit()
            except Exception as exc:
                db.rollback()
                raise HTTPException(status_code=409, detail="Material code already exists") from exc
            db.refresh(item)
            response_data = {"id": item.id, "code": item.code, "name": item.name, "uom": item.uom, "buffer_days": item.buffer_days}
            if _check_price_access(request):
                response_data.update(unit_cost=item.unit_cost, currency=item.currency)
            return response_data

    @app.patch("/api/planning/materials/{material_id}")
    def update_material(material_id: int, payload: MaterialInput, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        if payload.unit_cost and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to set a material price")
        with open_db() as db:
            item = db.get(Item, material_id)
            if not item:
                raise HTTPException(status_code=404, detail="Material not found")
            code = payload.code.strip()
            if db.scalar(select(Item.id).where(func.lower(Item.code) == code.lower(), Item.id != material_id)):
                raise HTTPException(status_code=409, detail="Material code already exists")
            if db.scalar(select(TierMaterial.tier).where(func.lower(TierMaterial.code) == code.lower()).limit(1)):
                raise HTTPException(status_code=409, detail=f"{code} already exists in {code_owner(db, code)} Material Master")
            item.code, item.name, item.category, item.uom = code, payload.name.strip(), payload.category.strip(), payload.uom.strip()
            item.stock_on_hand, item.buffer_days = payload.stock_on_hand, payload.buffer_days
            if _check_price_access(request):
                item.unit_cost, item.currency = payload.unit_cost, payload.currency
            sync_linked_supplier_materials(db, item=item)
            db.commit()
            response_data = {"id": item.id, "code": item.code, "name": item.name, "category": item.category, "uom": item.uom, "stock_on_hand": item.stock_on_hand, "buffer_days": item.buffer_days}
            if _check_price_access(request):
                response_data.update(unit_cost=item.unit_cost, currency=item.currency)
            return response_data

    @app.delete("/api/planning/materials/{material_id}", status_code=204)
    def delete_material(material_id: int, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            item = db.get(Item, material_id)
            if not item:
                raise HTTPException(status_code=404, detail="Material not found")
            referenced = (
                db.scalar(select(ModelBOM.id).where(ModelBOM.material_id == material_id))
                or db.scalar(select(BOMComponent.id).where(or_(BOMComponent.parent_item_id == material_id, BOMComponent.child_item_id == material_id)))
                or db.scalar(select(SupplierPart.id).where(SupplierPart.material_id == material_id))
                or db.scalar(select(OwnStockSnapshot.id).where(OwnStockSnapshot.material_id == material_id))
                or db.scalar(select(MaterialTransfer.id).where(MaterialTransfer.material_id == material_id))
                or db.scalar(select(MaterialRequest.id).where(MaterialRequest.item_id == material_id))
                or db.scalar(select(SupplierMaterial.id).where(SupplierMaterial.item_id == material_id).limit(1))
                or db.scalar(select(SupplierBOMLine.id).where(or_(SupplierBOMLine.item_id == material_id, SupplierBOMLine.product_material_id == material_id)).limit(1))
            )
            if referenced:
                raise HTTPException(status_code=409, detail="This material is used by a BOM, supplier, supplier material master, stock snapshot or transfer; remove those references first")
            db.delete(item)
            db.commit()
            return Response(status_code=204)

    @app.get("/api/planning/model-bom/{model_id}")
    def get_model_bom(model_id: int, request: Request, tab: str = "Model BOM Level 1", offset: int = 0, limit: int = 250, search: str = ""):
        current_user(request)
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        aliases = {"OEM BOM": "Model BOM Level 1"}
        tab = aliases.get(tab, tab)
        if tab != "Model BOM Level 1":
            raise HTTPException(status_code=422, detail="Select a parent part to view Level 2 or Level 3 BOM")
        with open_db() as db:
            model = db.get(ProductModel, model_id)
            if not model:
                raise HTTPException(status_code=404, detail="Model not found")
            rows = model_bom_rows(db, [model_id], _check_price_access(request))[model_id]
            if search.strip():
                needle = search.strip().lower()
                rows = [r for r in rows if needle in r["code"].lower() or needle in r["name"].lower()]
            return {"model_id": model_id, "model_code": model.code, "model_name": model.name, "tab": tab, "total": len(rows), "offset": offset, "limit": limit, "items": rows[offset:offset + limit]}

    @app.get("/api/planning/model-boms")
    def get_model_boms(request: Request, model_id: list[int] = Query(default=[])):
        """Level 1 lines of many models in one request (the Level 1 tab shows 50 models a page)."""
        current_user(request)
        if len(model_id) > 200:
            raise HTTPException(status_code=422, detail="Ask for at most 200 models at a time")
        with open_db() as db:
            models = {model.id: model for model in db.scalars(select(ProductModel).where(ProductModel.id.in_(model_id))).all()} if model_id else {}
            rows = model_bom_rows(db, list(models), _check_price_access(request))
            return {"models": [{"model_id": mid, "model_code": models[mid].code, "model_name": models[mid].name, "items": rows[mid]}
                               for mid in model_id if mid in models]}

    def model_bom_rows(db: Session, model_ids: list[int], can_price: bool) -> dict[int, list[dict]]:
        result: dict[int, list[dict]] = {mid: [] for mid in model_ids}
        if not model_ids:
            return result
        roots = db.scalars(select(ModelBOM).where(ModelBOM.model_id.in_(model_ids)).order_by(ModelBOM.id)).all()
        material_ids = {root.material_id for root in roots}
        items = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_(material_ids))).all()} if material_ids else {}
        parents_with_bom = set(db.scalars(
            select(BOMComponent.parent_item_id).where(BOMComponent.parent_item_id.in_(material_ids)).distinct()
        ).all()) if material_ids else set()
        price = bom_price_lookup(db, material_ids) if can_price else None
        for root in roots:
            item = items.get(root.material_id)
            if not item:
                continue
            row = {
                "level": 1, "material_id": item.id, "code": item.code, "name": item.name,
                "quantity": root.quantity, "uom": item.uom, "root_bom_id": root.id,
                "component_id": None, "parent_item_id": None, "edit_quantity": root.quantity,
                "has_bom": item.id in parents_with_bom,
            }
            if price:
                row.update(price(item))
            result[root.model_id].append(row)
        return result

    @app.post("/api/planning/model-bom/{model_id}", status_code=201)
    def add_model_bom(model_id: int, payload: ModelBOMInput, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            if not db.get(ProductModel, model_id) or not db.get(Item, payload.material_id):
                raise HTTPException(status_code=404, detail="Model or material not found")
            material = db.get(Item, payload.material_id)
            if payload.uom.strip() != material.uom:
                raise HTTPException(status_code=422, detail="BOM unit must match the material master unit")
            entry = ModelBOM(model_id=model_id, material_id=payload.material_id, quantity=payload.quantity, uom=payload.uom)
            db.add(entry)
            try:
                db.commit()
            except Exception as exc:
                db.rollback()
                raise HTTPException(status_code=409, detail="This material is already linked to the model BOM") from exc
            return {"id": entry.id, "model_id": model_id, "material_id": entry.material_id, "quantity": entry.quantity, "uom": entry.uom}

    @app.patch("/api/planning/model-bom/{model_id}/{bom_id}")
    def update_model_bom(model_id: int, bom_id: int, payload: ModelBOMInput, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            entry = db.get(ModelBOM, bom_id)
            if not entry or entry.model_id != model_id or not db.get(Item, payload.material_id):
                raise HTTPException(status_code=404, detail="BOM line or material not found")
            material = db.get(Item, payload.material_id)
            if payload.uom.strip() != material.uom:
                raise HTTPException(status_code=422, detail="BOM unit must match the material master unit")
            duplicate = db.scalar(select(ModelBOM.id).where(ModelBOM.model_id == model_id, ModelBOM.material_id == payload.material_id, ModelBOM.id != bom_id))
            if duplicate:
                raise HTTPException(status_code=409, detail="This material is already linked to the model BOM")
            entry.material_id, entry.quantity, entry.uom = payload.material_id, payload.quantity, payload.uom
            db.commit()
            return {"id": entry.id, "model_id": entry.model_id, "material_id": entry.material_id, "quantity": entry.quantity, "uom": entry.uom}

    @app.delete("/api/planning/model-bom/{model_id}/{bom_id}", status_code=204)
    def delete_model_bom(model_id: int, bom_id: int, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            entry = db.get(ModelBOM, bom_id)
            if not entry or entry.model_id != model_id:
                raise HTTPException(status_code=404, detail="BOM line not found")
            db.delete(entry)
            db.commit()
            return Response(status_code=204)

    def code_owner(db: Session, code: str) -> str | None:
        """Which master already holds this code (case-insensitive): 'T1', 'T2', 'T3' or None."""
        if db.scalar(select(Item.id).where(func.lower(Item.code) == code.lower()).limit(1)):
            return "T1"
        tier = db.scalar(select(TierMaterial.tier).where(func.lower(TierMaterial.code) == code.lower()).limit(1))
        return tier

    def sync_linked_supplier_materials(db: Session, *, item: Item | None = None, tier_material: TierMaterial | None = None) -> None:
        """Supplier material entries picked from a master follow its code / name / unit."""
        source = item or tier_material
        column = SupplierMaterial.item_id if item else SupplierMaterial.tier_material_id
        for entry in db.scalars(select(SupplierMaterial).where(column == source.id)).all():
            entry.code, entry.name, entry.uom = source.code, source.name, source.uom

    def not_purchased_reason(db: Session, material_id: int) -> str | None:
        """Parts you do not buy: made in-house (own BOM in Product Master) or generated in
        your factory (used with a negative quantity in any BOM). Supplier BOMs do not count."""
        if db.scalar(select(BOMComponent.id).where(BOMComponent.parent_item_id == material_id).limit(1)):
            return "is made in-house (it has its own BOM)"
        if db.scalar(select(BOMComponent.id).where(BOMComponent.child_item_id == material_id, BOMComponent.quantity < 0).limit(1)):
            return "is generated in your factory (negative BOM quantity)"
        return None

    def purchased_filter(statement):
        made = select(BOMComponent.parent_item_id)
        generated = select(BOMComponent.child_item_id).where(BOMComponent.quantity < 0)
        return statement.where(Item.id.not_in(made), Item.id.not_in(generated))

    def bom_price_lookup(db: Session, material_ids: Iterable[int] | None = None):
        """BOM part price = SOB-weighted price of the code's Direct suppliers:
        sum(price x SOB) / sum(SOB). Falls back to the Material Master price when no Direct
        supplier has a price, or when the suppliers quote in different currencies."""
        links: dict[int, list[SupplierPart]] = {}
        statement = select(SupplierPart).where(
            SupplierPart.tier == "Direct", SupplierPart.parent_supplier_id.is_(None),
            SupplierPart.sob_percent > 0, SupplierPart.unit_price > 0,
        )
        if material_ids is not None:
            statement = statement.where(SupplierPart.material_id.in_(list(material_ids)))
        for link in db.scalars(statement).all():
            links.setdefault(link.material_id, []).append(link)

        def price(item: Item) -> dict:
            quoted = links.get(item.id, [])
            if quoted and len({link.currency for link in quoted}) == 1:
                share = sum(link.sob_percent for link in quoted)
                return {"unit_price": sum(link.unit_price * link.sob_percent for link in quoted) / share,
                        "currency": quoted[0].currency, "price_source": f"SOB-weighted ({len(quoted)} supplier{'s' if len(quoted) > 1 else ''})"}
            source = "Material Master (suppliers in different currencies)" if quoted else "Material Master"
            return {"unit_price": item.unit_cost, "currency": item.currency, "price_source": source}
        return price

    def bom_parent_material_ids(db: Session, tab: str) -> set[int]:
        level1_ids = set(db.scalars(select(ModelBOM.material_id).distinct()).all())
        if tab == "Level 2 Part BOM":
            return level1_ids
        if tab == "Level 3 Part BOM":
            return set(db.scalars(select(BOMComponent.child_item_id).where(BOMComponent.parent_item_id.in_(level1_ids)).distinct()).all()) if level1_ids else set()
        raise HTTPException(status_code=422, detail="Unknown part BOM tab")

    @app.get("/api/planning/bom-part-parents")
    def list_bom_part_parents(request: Request, tab: str, model_id: list[int] = Query(default=[])):
        current_user(request)
        with open_db() as db:
            ids = bom_parent_material_ids(db, tab)
            model_ids_by_parent: dict[int, set[int]] = {}
            roots = db.scalars(select(ModelBOM)).all()
            roots_by_material: dict[int, set[int]] = {}
            for root in roots:
                roots_by_material.setdefault(root.material_id, set()).add(root.model_id)
                if tab == "Level 2 Part BOM":
                    model_ids_by_parent.setdefault(root.material_id, set()).add(root.model_id)
            if tab == "Level 3 Part BOM" and roots_by_material:
                for link in db.scalars(select(BOMComponent).where(BOMComponent.parent_item_id.in_(roots_by_material))).all():
                    model_ids_by_parent.setdefault(link.child_item_id, set()).update(roots_by_material[link.parent_item_id])
            selected_models = set(model_id)
            if selected_models:
                ids = {item_id for item_id in ids if model_ids_by_parent.get(item_id, set()) & selected_models}
            items = db.scalars(select(Item).where(Item.id.in_(ids)).order_by(Item.code)).all() if ids else []
            # Parents that already have their own BOM lines (shown in the parent dropdown);
            # the rest can be picked when starting a new BOM.
            with_bom = set(db.scalars(
                select(BOMComponent.parent_item_id).where(BOMComponent.parent_item_id.in_(ids)).distinct()
            ).all()) if ids else set()
            return {"items": [
                {"id": item.id, "code": item.code, "name": item.name, "uom": item.uom,
                 "model_ids": sorted(model_ids_by_parent.get(item.id, set())),
                 "has_bom": item.id in with_bom}
                for item in items
            ]}

    @app.get("/api/planning/part-bom/{parent_material_id}")
    def get_part_bom(
        parent_material_id: int, request: Request, tab: str,
        offset: int = 0, limit: int = 500, search: str = "",
        search_code: str = "", search_name: str = "",
    ):
        current_user(request)
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        with open_db() as db:
            if parent_material_id not in bom_parent_material_ids(db, tab):
                raise HTTPException(status_code=404, detail="Parent part does not exist at this BOM level")
            parent = db.get(Item, parent_material_id)
            if not parent:
                raise HTTPException(status_code=404, detail="Parent part not found")
            links = db.scalars(select(BOMComponent).where(BOMComponent.parent_item_id == parent_material_id).order_by(BOMComponent.id)).all()
            child_ids = [link.child_item_id for link in links]
            items_by_id = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_(child_ids))).all()} if child_ids else {}
            parents_with_bom = set(db.scalars(
                select(BOMComponent.parent_item_id)
                .where(BOMComponent.parent_item_id.in_(child_ids))
                .distinct()
            ).all()) if child_ids else set()
            rows = [{
                "component_id": link.id, "material_id": link.child_item_id,
                "code": items_by_id[link.child_item_id].code, "name": items_by_id[link.child_item_id].name,
                "quantity": link.quantity, "uom": items_by_id[link.child_item_id].uom,
                "has_bom": link.child_item_id in parents_with_bom,
            } for link in links if link.child_item_id in items_by_id]
            if search.strip():
                needle = search.strip().casefold()
                rows = [row for row in rows if needle in row["code"].casefold() or needle in row["name"].casefold()]
            if search_code.strip():
                code_needle = search_code.strip().casefold()
                rows = [row for row in rows if code_needle in row["code"].casefold()]
            if search_name.strip():
                name_needle = search_name.strip().casefold()
                rows = [row for row in rows if name_needle in row["name"].casefold()]
            can_price = _check_price_access(request)
            if can_price:
                price = bom_price_lookup(db, child_ids)
                for row in rows:
                    row.update(price(items_by_id[row["material_id"]]))
            return {
                "tab": tab, "parent": {"id": parent.id, "code": parent.code, "name": parent.name},
                "total": len(rows), "offset": offset, "limit": limit, "items": rows[offset:offset + limit],
            }

    @app.get("/api/planning/part-boms")
    def get_part_boms(
        request: Request, tab: str, parent_material_id: list[int] = Query(default=[]),
        offset: int = 0, limit: int = 500,
    ):
        current_user(request)
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        with open_db() as db:
            available_ids = bom_parent_material_ids(db, tab)
            selected_ids = set(parent_material_id) if parent_material_id else available_ids
            if not selected_ids.issubset(available_ids):
                raise HTTPException(status_code=404, detail="One or more parent parts do not exist at this BOM level")
            parents = db.scalars(select(Item).where(Item.id.in_(selected_ids))).all() if selected_ids else []
            parents_by_id = {parent.id: parent for parent in parents}
            links = db.execute(
                select(BOMComponent, Item.code, Item.name)
                .join(Item, Item.id == BOMComponent.parent_item_id)
                .where(BOMComponent.parent_item_id.in_(selected_ids))
                .order_by(Item.code, BOMComponent.id)
            ).all() if selected_ids else []
            child_ids = {link.child_item_id for link, _, _ in links}
            children = db.scalars(select(Item).where(Item.id.in_(child_ids))).all() if child_ids else []
            children_by_id = {child.id: child for child in children}
            parents_with_bom = set(db.scalars(
                select(BOMComponent.parent_item_id)
                .where(BOMComponent.parent_item_id.in_(child_ids))
                .distinct()
            ).all()) if child_ids else set()
            can_price = _check_price_access(request)
            price = bom_price_lookup(db, child_ids)
            rows = []
            for link, parent_code, parent_name in links:
                child = children_by_id.get(link.child_item_id)
                parent = parents_by_id.get(link.parent_item_id)
                if not child or not parent:
                    continue
                row = {
                    "component_id": link.id, "parent": {"id": parent.id, "code": parent_code, "name": parent_name},
                    "material_id": child.id, "code": child.code, "name": child.name,
                    "quantity": link.quantity, "uom": child.uom,
                    "has_bom": child.id in parents_with_bom,
                }
                if can_price:
                    row.update(price(child))
                rows.append(row)
            return {
                "tab": tab, "total": len(rows), "offset": offset, "limit": limit,
                "items": rows[offset:offset + limit],
            }

    @app.post("/api/planning/part-bom/{parent_material_id}/upload", status_code=201)
    async def upload_part_bom(parent_material_id: int, request: Request, tab: str):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        if tab not in {"Level 2 Part BOM", "Level 3 Part BOM"}:
            raise HTTPException(status_code=422, detail="Unknown part BOM tab")
        rows = _read_xlsx(await request.body())
        _required_columns(rows, {"parent_code", "material_code", "quantity", "uom"})
        with open_db() as db:
            if parent_material_id not in bom_parent_material_ids(db, tab):
                raise HTTPException(status_code=404, detail="Parent part does not exist at this BOM level")
            parent = db.get(Item, parent_material_id)
            items = {item.code: item for item in db.scalars(select(Item)).all()}
            graph: dict[int, set[int]] = {}
            existing_links = db.scalars(select(BOMComponent)).all()
            links_by_pair: dict[tuple[int, int], BOMComponent] = {}
            for link in existing_links:
                graph.setdefault(link.parent_item_id, set()).add(link.child_item_id)
                links_by_pair[(link.parent_item_id, link.child_item_id)] = link
            seen: set[int] = set()
            parsed: list[tuple[int, float]] = []

            def has_path(start_id: int, target_id: int) -> bool:
                pending = [start_id]
                visited: set[int] = set()
                while pending:
                    current = pending.pop()
                    if current == target_id:
                        return True
                    if current in visited:
                        continue
                    visited.add(current)
                    pending.extend(graph.get(current, ()))
                return False

            for row_number, row in enumerate(rows, start=2):
                parent_code = str(row.get("parent_code") or "").strip()
                if parent_code.casefold() != parent.code.casefold():
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: parent_code must be {parent.code}")
                material_code = str(row.get("material_code") or "").strip()
                child = items.get(material_code)
                if not child:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: unknown material_code {material_code!r}")
                unit = str(row.get("uom") or "").strip()
                if unit.casefold() != child.uom.casefold():
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: unit {unit!r} does not match {child.code} master unit {child.uom!r}")
                try:
                    quantity = float(row.get("quantity"))
                except (TypeError, ValueError) as exc:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: quantity must be a non-zero number") from exc
                if not math.isfinite(quantity) or quantity == 0 or abs(quantity) > 1_000_000_000:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: quantity must be non-zero and within -1,000,000,000 to 1,000,000,000")
                if child.id in seen:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: duplicate parent/material relationship")
                if child.id == parent_material_id or has_path(child.id, parent_material_id):
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: this BOM line would create a circular BOM")
                seen.add(child.id)
                parsed.append((child.id, quantity))
                graph.setdefault(parent_material_id, set()).add(child.id)

            keep_child_ids = {child_id for child_id, _ in parsed}
            for link in db.scalars(select(BOMComponent).where(BOMComponent.parent_item_id == parent_material_id)).all():
                if link.child_item_id not in keep_child_ids:
                    db.delete(link)
            for child_id, quantity in parsed:
                existing = links_by_pair.get((parent_material_id, child_id))
                if existing:
                    existing.quantity = quantity
                else:
                    db.add(BOMComponent(parent_item_id=parent_material_id, child_item_id=child_id, quantity=quantity))
            db.commit()
            return {"parent_id": parent_material_id, "parent_code": parent.code, "tab": tab, "rows": len(parsed)}

    @app.post("/api/planning/part-bom/{parent_material_id}/new", status_code=201)
    def create_part_bom(parent_material_id: int, payload: NewPartBOMInput, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            if parent_material_id not in bom_parent_material_ids(db, payload.tab):
                raise HTTPException(status_code=422, detail=f"This parent part is not linked at {payload.tab}")
            parent = db.get(Item, parent_material_id)
            if db.scalar(select(BOMComponent.id).where(BOMComponent.parent_item_id == parent_material_id).limit(1)):
                raise HTTPException(status_code=409, detail=f"{parent.code} already has a BOM; add lines to it instead")
            graph: dict[int, set[int]] = {}
            for link in db.scalars(select(BOMComponent)).all():
                graph.setdefault(link.parent_item_id, set()).add(link.child_item_id)

            def reaches(start_id: int, target_id: int) -> bool:
                pending, visited = [start_id], set()
                while pending:
                    current = pending.pop()
                    if current == target_id:
                        return True
                    if current not in visited:
                        visited.add(current)
                        pending.extend(graph.get(current, ()))
                return False

            seen: set[int] = set()
            for number, line in enumerate(payload.lines, start=1):
                child = db.get(Item, line.child_item_id)
                if not child:
                    raise HTTPException(status_code=422, detail=f"Line {number}: part is not in Material Master")
                if child.id in seen:
                    raise HTTPException(status_code=422, detail=f"Line {number}: {child.code} is entered twice")
                if child.id == parent_material_id or reaches(child.id, parent_material_id):
                    raise HTTPException(status_code=422, detail=f"Line {number}: {child.code} would create a circular BOM")
                seen.add(child.id)
                db.add(BOMComponent(parent_item_id=parent_material_id, child_item_id=child.id, quantity=line.quantity))
            db.commit()
            return {"parent_id": parent_material_id, "parent_code": parent.code, "tab": payload.tab, "rows": len(payload.lines)}

    @app.post("/api/planning/model-bom/{model_id}/upload", status_code=201)
    async def upload_model_bom_excel(model_id: int, request: Request, tab: str):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        allowed_tabs = {"OEM BOM", "Model BOM Level 1"}
        if tab not in allowed_tabs:
            raise HTTPException(status_code=422, detail="Level 2 and Level 3 BOM uploads must use a parent part")
        rows = _read_xlsx(await request.body())
        _required_columns(rows, {"model_code", "material_code", "quantity", "uom"})
        with open_db() as db:
            model = db.get(ProductModel, model_id)
            if not model:
                raise HTTPException(status_code=404, detail="Model not found")
            items = {item.code: item for item in db.scalars(select(Item)).all()}
            roots = db.scalars(select(ModelBOM).where(ModelBOM.model_id == model_id)).all()
            roots_by_material = {entry.material_id: entry for entry in roots}
            parsed: list[tuple[int, float, str]] = []
            seen: set[int] = set()
            for row_number, row in enumerate(rows, start=2):
                row_model_code = str(row.get("model_code") or "").strip()
                if row_model_code != model.code:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: model_code must be {model.code}")
                material_code = str(row.get("material_code") or "").strip()
                material = items.get(material_code)
                if not material:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: unknown material_code {material_code!r}")
                unit = str(row.get("uom") or "").strip()
                if not unit:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: uom is required")
                try:
                    quantity = float(row.get("quantity"))
                except (TypeError, ValueError) as exc:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: quantity must be a positive number") from exc
                if not math.isfinite(quantity) or quantity <= 0 or quantity > 1_000_000_000:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: quantity must be between zero and one billion")
                if unit.casefold() != material.uom.casefold():
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: unit {unit!r} does not match {material.code} master unit {material.uom!r}")
                if material.id in seen:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: duplicate Model BOM Level 1 material")
                seen.add(material.id)
                parsed.append((material.id, quantity, unit))

            for material_id, quantity, unit in parsed:
                existing = roots_by_material.get(material_id)
                if existing:
                    existing.quantity = quantity
                    existing.uom = unit
                else:
                    db.add(ModelBOM(model_id=model_id, material_id=material_id, quantity=quantity, uom=unit))
            db.commit()
            return {"model_id": model_id, "model_code": model.code, "tab": tab, "rows": len(parsed)}

    @app.get("/api/planning/supplier-parts")
    def list_supplier_parts(request: Request, material_id: int | None = None, offset: int = 0, limit: int = 500, search: str = ""):
        current_user(request)
        _check_supply_role(request)
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        with open_db() as db:
            statement = select(SupplierPart).join(Item, SupplierPart.material_id == Item.id).join(
                Supplier, SupplierPart.supplier_id == Supplier.id
            ).order_by(SupplierPart.material_id, SupplierPart.tier, SupplierPart.id)
            if material_id:
                statement = statement.where(SupplierPart.material_id == material_id)
            if search.strip():
                like = f"%{search.strip()}%"
                statement = statement.where(or_(
                    Item.code.ilike(like), Item.name.ilike(like), Supplier.code.ilike(like),
                    Supplier.name.ilike(like), SupplierPart.tier.ilike(like),
                ))
            total = db.scalar(select(func.count()).select_from(statement.subquery())) or 0
            links = db.scalars(statement.offset(offset).limit(limit)).all()
            can_price = _check_price_access(request)
            rows = []
            for link in links:
                item = db.get(Item, link.material_id)
                supplier = db.get(Supplier, link.supplier_id)
                parent = db.get(Supplier, link.parent_supplier_id) if link.parent_supplier_id else None
                rows.append({
                    "id": link.id, "material_id": link.material_id, "material_code": item.code if item else "",
                    "material_name": item.name if item else "", "supplier_id": link.supplier_id,
                    "supplier_code": supplier.code if supplier else "", "supplier_name": supplier.name if supplier else "",
                    "tier": link.tier, "parent_supplier_id": link.parent_supplier_id,
                    "parent_supplier_name": parent.name if parent else None, "sob_percent": link.sob_percent,
                    "lead_time_days": link.lead_time_days, "moq": link.moq, "buffer_days": link.buffer_days,
                    "price_mode": link.price_mode, "price_interval_months": link.price_interval_months,
                    "price_updated_month": link.price_updated_month,
                    "price_status": _price_status(link, _month_text(date.today())),
                    **({"unit_price": link.unit_price, "currency": link.currency} if can_price else {}),
                })
            return {"total": total, "offset": offset, "limit": limit, "items": rows}

    @app.post("/api/planning/supplier-parts", status_code=201)
    def add_supplier_part(payload: SupplierPartInput, request: Request):
        current_user(request)
        _check_supply_role(request)
        if payload.unit_price and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to set a part price")
        _month(_month_text(date.today()))
        if payload.tier != "Direct" and payload.parent_supplier_id is None:
            raise HTTPException(status_code=422, detail="T1/T2 supplier links must identify the customer supplier they serve")
        with open_db() as db:
            material = db.get(Item, payload.material_id)
            if not material or not db.get(Supplier, payload.supplier_id):
                raise HTTPException(status_code=404, detail="Material or supplier not found")
            reason = not_purchased_reason(db, material.id)
            if reason:
                raise HTTPException(status_code=422, detail=f"{material.code} {reason}, so it cannot have a supplier")
            share_total = db.scalar(select(func.coalesce(func.sum(SupplierPart.sob_percent), 0)).where(
                SupplierPart.material_id == payload.material_id,
                SupplierPart.tier == payload.tier,
                SupplierPart.parent_supplier_id == payload.parent_supplier_id,
            )) or 0
            if share_total + payload.sob_percent > 100:
                raise HTTPException(status_code=422, detail="Supplier SOB cannot exceed 100% within the same material tier and parent supplier")
            if payload.parent_supplier_id and not db.get(Supplier, payload.parent_supplier_id):
                raise HTTPException(status_code=404, detail="Parent supplier not found")
            existing = db.scalar(select(SupplierPart).where(
                SupplierPart.material_id == payload.material_id,
                SupplierPart.supplier_id == payload.supplier_id,
                SupplierPart.tier == payload.tier,
                SupplierPart.parent_supplier_id == payload.parent_supplier_id,
            ))
            if existing:
                raise HTTPException(status_code=409, detail="This supplier/material/tier relationship already exists")
            link = SupplierPart(**payload.model_dump())
            db.add(link)
            db.commit()
            db.refresh(link)
            return {"id": link.id, "material_id": link.material_id, "supplier_id": link.supplier_id, "tier": link.tier, "sob_percent": link.sob_percent}

    @app.post("/api/planning/supplier-parts/upload", status_code=201)
    async def upload_supplier_parts(request: Request):
        """Supplier/material relationships from Excel, all rows or none. One row per
        material + supplier (+ tier / parent supplier). Unknown supplier names are added to
        the Supplier register (Approved); an existing relationship is updated."""
        current_user(request)
        _check_supply_role(request)
        rows = _read_supplier_part_sheet(await request.body())
        can_price = _check_price_access(request)
        with open_db() as db:
            items = {item.code.casefold(): item for item in db.scalars(select(Item)).all()}
            suppliers = {supplier.name.strip().casefold(): supplier for supplier in db.scalars(select(Supplier)).all()}
            suppliers.update({supplier.code.casefold(): supplier for supplier in suppliers.values()})
            existing = {(link.material_id, link.supplier_id, link.tier, link.parent_supplier_id): link
                        for link in db.scalars(select(SupplierPart)).all()}
            next_number = int(db.scalar(select(func.count(Supplier.id))) or 0) + 1001
            used_codes = set(db.scalars(select(Supplier.code)).all())
            reasons: dict[int, str | None] = {}
            seen: set[tuple] = set()
            created = updated = added = 0

            def text(row: dict, key: str) -> str:
                value = row.get(key)
                return "" if value is None else str(value).strip()

            def number(row: dict, key: str, default: float, row_number: int) -> float:
                value = row.get(key)
                if value is None or str(value).strip() == "":
                    return default
                try:
                    result = float(value)
                except (TypeError, ValueError) as exc:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {key} must be a number") from exc
                if not math.isfinite(result) or result < 0:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {key} must be zero or more")
                return result

            def supplier_for(name: str, category: str) -> Supplier:
                nonlocal next_number, created
                found = suppliers.get(name.casefold())
                if found:
                    return found
                while f"SUP-{next_number:04d}" in used_codes:
                    next_number += 1
                supplier = Supplier(code=f"SUP-{next_number:04d}", name=name, category=category or "General", country="-",
                                    status="Approved", risk="Medium")
                used_codes.add(supplier.code)
                db.add(supplier)
                db.flush()
                suppliers[name.casefold()] = supplier
                created += 1
                return supplier

            for row in rows:
                row_number = row["_excel_row"]
                code, name = text(row, "material_code"), text(row, "supplier_name")
                if not code or not name:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: material_code and supplier_name are required")
                material = items.get(code.casefold())
                if not material:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: material_code {code!r} is not in Material Master")
                if material.id not in reasons:
                    reasons[material.id] = not_purchased_reason(db, material.id)
                if reasons[material.id]:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {material.code} {reasons[material.id]}, so it cannot have a supplier")
                tier = text(row, "tier") or "Direct"
                parent_name = text(row, "parent_supplier_name")
                if tier != "Direct" and not parent_name:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: tier {tier} needs parent_supplier_name (the supplier it supplies)")
                sob = number(row, "sob_percent", 100, row_number)
                if sob > 100:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: sob_percent cannot exceed 100")
                price = number(row, "unit_price", 0, row_number)
                if price and not can_price:
                    raise HTTPException(status_code=403, detail="Price access is required to upload part prices")
                currency = (text(row, "currency") or "INR").upper()
                if not re.fullmatch(r"[A-Z]{3}", currency):
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: currency must be a 3-letter code such as INR")
                price_mode = text(row, "price_mode") or "spot"
                if price_mode not in {"spot", "fixed_monthly", "every_n_months"}:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: price_mode must be spot, fixed_monthly or every_n_months")
                supplier = supplier_for(name, material.category)
                parent = supplier_for(parent_name, material.category) if parent_name else None
                key = (material.id, supplier.id, tier, parent.id if parent else None)
                if key in seen:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {material.code} / {supplier.name} / {tier} is repeated in the file")
                seen.add(key)
                values = {
                    "sob_percent": sob, "unit_price": price, "currency": currency,
                    "lead_time_days": int(number(row, "lead_time_days", 0, row_number)), "moq": number(row, "moq", 0, row_number),
                    "buffer_days": int(number(row, "buffer_days", 0, row_number)), "price_mode": price_mode,
                    "price_interval_months": max(1, int(number(row, "price_interval_months", 1, row_number))),
                }
                link = existing.get(key)
                if link:
                    for field_name, value in values.items():
                        setattr(link, field_name, value)
                    updated += 1
                else:
                    link = SupplierPart(material_id=material.id, supplier_id=supplier.id, tier=tier, parent_supplier_id=key[3], **values)
                    db.add(link)
                    existing[key] = link
                    added += 1
            # SOB within a material + tier + parent-supplier route cannot exceed 100 %
            db.flush()
            totals: dict[tuple, float] = {}
            for link in existing.values():
                route = (link.material_id, link.tier, link.parent_supplier_id)
                totals[route] = totals.get(route, 0) + link.sob_percent
            over = [route for route, total in totals.items() if total > 100.0001]
            if over:
                material = db.get(Item, over[0][0])
                db.rollback()
                raise HTTPException(status_code=422, detail=f"SOB for {material.code} ({over[0][1]}) adds up to more than 100% after this upload")
            db.commit()
            return {"rows": len(rows), "added": added, "updated": updated, "suppliers_created": created}

    # ------------------------------------------------------------------ supplier materials
    # ------------------------------------------------------------------ T2 / T3 Material Master
    def tier_material_json(material: TierMaterial, can_price: bool) -> dict:
        data = {"id": material.id, "tier": material.tier, "code": material.code, "name": material.name, "category": material.category, "uom": material.uom}
        if can_price:
            data.update(unit_price=material.unit_price, currency=material.currency)
        return data

    @app.get("/api/planning/tier-materials")
    def list_tier_materials(request: Request, tier: str, search: str = "", offset: int = 0, limit: int = 100):
        current_user(request)
        if tier not in {"T2", "T3"}:
            raise HTTPException(status_code=422, detail="tier must be T2 or T3")
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        with open_db() as db:
            statement = select(TierMaterial).where(TierMaterial.tier == tier).order_by(TierMaterial.code)
            if search.strip():
                like = f"%{search.strip()}%"
                statement = statement.where(TierMaterial.code.ilike(like) | TierMaterial.name.ilike(like))
            total = db.scalar(select(func.count()).select_from(statement.subquery())) or 0
            can_price = _check_price_access(request)
            return {"total": total, "offset": offset, "limit": limit,
                    "items": [tier_material_json(m, can_price) for m in db.scalars(statement.offset(offset).limit(limit)).all()]}

    def check_tier_material(db: Session, payload: TierMaterialInput, request: Request, current: TierMaterial | None = None) -> str:
        _check_supply_role(request)
        if payload.unit_price and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to set a price")
        code = payload.code.strip()
        if not code:
            raise HTTPException(status_code=422, detail="Code is required")
        if not (current and current.code.lower() == code.lower()):
            owner = code_owner(db, code)
            if owner:
                raise HTTPException(status_code=409, detail=f"{code} already exists in {owner} Material Master")
        return code

    @app.post("/api/planning/tier-materials", status_code=201)
    def add_tier_material(payload: TierMaterialInput, request: Request):
        current_user(request)
        with open_db() as db:
            code = check_tier_material(db, payload, request)
            material = TierMaterial(tier=payload.tier, code=code, name=payload.name.strip(), category=payload.category.strip(), uom=payload.uom.strip(),
                                    unit_price=payload.unit_price, currency=payload.currency)
            db.add(material)
            db.commit()
            return tier_material_json(material, True)

    @app.patch("/api/planning/tier-materials/{material_id}")
    def update_tier_material(material_id: int, payload: TierMaterialInput, request: Request):
        current_user(request)
        with open_db() as db:
            material = db.get(TierMaterial, material_id)
            if not material or material.tier != payload.tier:
                raise HTTPException(status_code=404, detail=f"{payload.tier} material not found")
            material.code = check_tier_material(db, payload, request, material)
            material.name, material.category, material.uom = payload.name.strip(), payload.category.strip(), payload.uom.strip()
            if _check_price_access(request):
                material.unit_price, material.currency = payload.unit_price, payload.currency
            sync_linked_supplier_materials(db, tier_material=material)
            db.commit()
            return tier_material_json(material, True)

    @app.delete("/api/planning/tier-materials/{material_id}", status_code=204)
    def delete_tier_material(material_id: int, request: Request):
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            material = db.get(TierMaterial, material_id)
            if not material:
                raise HTTPException(status_code=404, detail="Material not found")
            if db.scalar(select(SupplierMaterial.id).where(SupplierMaterial.tier_material_id == material_id).limit(1)):
                raise HTTPException(status_code=409, detail=f"{material.code} is used in a supplier's material master; remove it there first")
            db.delete(material)
            db.commit()
            return Response(status_code=204)

    @app.post("/api/planning/tier-materials/upload", status_code=201)
    async def upload_tier_materials(request: Request, tier: str):
        """T2 / T3 Material Master from Excel (heading row found automatically): Code, Name,
        Category, Unit, Unit price / Unit cost, Currency. Existing codes in this master are updated;
        a code already in another master is refused. All rows or none."""
        current_user(request)
        _check_supply_role(request)
        if tier not in {"T2", "T3"}:
            raise HTTPException(status_code=422, detail="tier must be T2 or T3")
        rows = _read_flexible_sheet(await request.body(), TIER_MATERIAL_COLUMN_ALIASES, {"code", "name"})
        can_price = _check_price_access(request)
        with open_db() as db:
            mine = {m.code.lower(): m for m in db.scalars(select(TierMaterial).where(TierMaterial.tier == tier)).all()}
            seen: set[str] = set()
            added = updated = 0
            for row in rows:
                row_number = row["_excel_row"]
                code, name = str(row.get("code") or "").strip(), str(row.get("name") or "").strip()
                if not code or not name:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: code and name are required")
                if code.lower() in seen:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {code} is repeated in the file")
                seen.add(code.lower())
                price, currency = _price_cells(row, row_number, can_price)
                uom = str(row.get("uom") or "ea").strip() or "ea"
                category = str(row.get("category") or "Raw material").strip() or "Raw material"
                material = mine.get(code.lower())
                if material:
                    material.name, material.category, material.uom = name, category, uom
                    if can_price:
                        material.unit_price, material.currency = price, currency
                    sync_linked_supplier_materials(db, tier_material=material)
                    updated += 1
                else:
                    owner = code_owner(db, code)
                    if owner:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: {code} already exists in {owner} Material Master")
                    db.add(TierMaterial(tier=tier, code=code, name=name, category=category, uom=uom, unit_price=price, currency=currency))
                    db.flush()
                    added += 1
            db.commit()
            return {"rows": len(rows), "added": added, "updated": updated}

    @app.get("/api/planning/tier-materials/{tier}.{file_format}")
    def export_tier_materials(tier: str, file_format: str, request: Request):
        current_user(request)
        if tier not in {"T2", "T3"} or file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export not found")
        can_price = _check_price_access(request)
        with open_db() as db:
            rows = db.scalars(select(TierMaterial).where(TierMaterial.tier == tier).order_by(TierMaterial.code)).all()
            headers = ["Code", "Name", "Category", "Unit"] + (["Reference price", "Currency"] if can_price else [])
            values = [[m.code, m.name, m.category, m.uom] + ([m.unit_price, m.currency] if can_price else []) for m in rows]
        title = f"{tier} Material Master"
        output = _generic_workbook(title, headers, values) if file_format == "xlsx" else _generic_pdf(title, headers, values)
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_format == "xlsx" else "application/pdf"
        return StreamingResponse(output, media_type=media, headers={"Content-Disposition": f'attachment; filename="{tier.lower()}-material-master.{file_format}"'})

    # ------------------------------------------------------------------ supplier materials
    def supplier_material_json(material: SupplierMaterial, can_price: bool) -> dict:
        source = "T1" if material.item_id else ("T2/T3" if material.tier_material_id else "Own")  # T2 vs T3 filled in by supplier_material_rows
        data = {"id": material.id, "supplier_id": material.supplier_id, "code": material.code, "name": material.name, "uom": material.uom,
                "item_id": material.item_id, "tier_material_id": material.tier_material_id, "source": source,
                "lead_time_days": material.lead_time_days or 0, "buffer_days": material.buffer_days or 0}
        if can_price:
            data.update(unit_price=material.unit_price, currency=material.currency)
        return data

    def supplier_material_rows(db: Session, materials: list[SupplierMaterial], can_price: bool) -> list[dict]:
        tiers = dict(db.execute(select(TierMaterial.id, TierMaterial.tier).where(
            TierMaterial.id.in_([m.tier_material_id for m in materials if m.tier_material_id])
        )).all()) if any(m.tier_material_id for m in materials) else {}
        rows = []
        for material in materials:
            row = supplier_material_json(material, can_price)
            if material.tier_material_id:
                row["source"] = tiers.get(material.tier_material_id, "T2/T3")
            rows.append(row)
        return rows

    @app.get("/api/planning/supplier-materials")
    def list_supplier_materials(request: Request, supplier_id: int, search: str = "", offset: int = 0, limit: int = 100):
        current_user(request)
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        with open_db() as db:
            statement = select(SupplierMaterial).where(SupplierMaterial.supplier_id == supplier_id).order_by(SupplierMaterial.code)
            if search.strip():
                like = f"%{search.strip()}%"
                statement = statement.where(SupplierMaterial.code.ilike(like) | SupplierMaterial.name.ilike(like))
            total = db.scalar(select(func.count()).select_from(statement.subquery())) or 0
            return {"total": total, "offset": offset, "limit": limit,
                    "items": supplier_material_rows(db, db.scalars(statement.offset(offset).limit(limit)).all(), _check_price_access(request))}

    @app.get("/api/planning/master-codes")
    def search_master_codes(request: Request, tier: str, search: str = "", limit: int = 100):
        """Codes a supplier can pick into its material master: T1 (your Material Master), T2 or T3."""
        current_user(request)
        if tier not in {"T1", "T2", "T3"} or limit < 1 or limit > 500:
            raise HTTPException(status_code=422, detail="tier must be T1, T2 or T3")
        with open_db() as db:
            model = Item if tier == "T1" else TierMaterial
            statement = select(model).order_by(model.code)
            if tier != "T1":
                statement = statement.where(TierMaterial.tier == tier)
            if search.strip():
                like = f"%{search.strip()}%"
                statement = statement.where(model.code.ilike(like) | model.name.ilike(like))
            return {"items": [{"id": m.id, "tier": tier, "code": m.code, "name": m.name, "uom": m.uom} for m in db.scalars(statement.limit(limit)).all()]}

    def resolve_supplier_material(db: Session, payload: SupplierMaterialInput, request: Request, current_id: int | None = None) -> tuple[str, str, str]:
        """(code, name, unit) for a supplier material entry: from the picked T1/T2/T3 code,
        or the supplier's own code (which must not be one of your master codes)."""
        if not db.get(Supplier, payload.supplier_id):
            raise HTTPException(status_code=404, detail="Supplier not found")
        if payload.unit_price and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to set a price")
        if payload.item_id and payload.tier_material_id:
            raise HTTPException(status_code=422, detail="Pick one code: from T1, or from T2/T3")
        if payload.item_id:
            master = db.get(Item, payload.item_id)
            if not master:
                raise HTTPException(status_code=422, detail="Code is not in T1 Material Master")
            code, name, uom = master.code, master.name, master.uom
        elif payload.tier_material_id:
            master = db.get(TierMaterial, payload.tier_material_id)
            if not master:
                raise HTTPException(status_code=422, detail="Code is not in T2/T3 Material Master")
            code, name, uom = master.code, master.name, master.uom
        else:
            code, name, uom = payload.code.strip(), payload.name.strip(), payload.uom.strip() or "ea"
            if not code or not name:
                raise HTTPException(status_code=422, detail="The supplier's own code needs a code and a name")
            owner = code_owner(db, code)
            if owner:
                raise HTTPException(status_code=422, detail=f"{code} is in your {owner} Material Master; pick it from {owner} instead of typing it as the supplier's own code")
        clash = db.scalar(select(SupplierMaterial.id).where(
            SupplierMaterial.supplier_id == payload.supplier_id, func.lower(SupplierMaterial.code) == code.lower()))
        if clash and clash != current_id:
            raise HTTPException(status_code=409, detail=f"{code} already exists in this supplier's material master")
        return code, name, uom

    @app.post("/api/planning/supplier-materials", status_code=201)
    def add_supplier_material(payload: SupplierMaterialInput, request: Request):
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            code, name, uom = resolve_supplier_material(db, payload, request)
            material = SupplierMaterial(supplier_id=payload.supplier_id, code=code, name=name, uom=uom, unit_price=payload.unit_price,
                                        currency=payload.currency, item_id=payload.item_id, tier_material_id=payload.tier_material_id,
                                        lead_time_days=payload.lead_time_days, buffer_days=payload.buffer_days)
            db.add(material)
            db.commit()
            return supplier_material_rows(db, [material], True)[0]

    @app.patch("/api/planning/supplier-materials/{material_id}")
    def update_supplier_material(material_id: int, payload: SupplierMaterialInput, request: Request):
        """Picked entries: only price / currency change (code and name follow your master).
        The supplier's own code: code, name, unit, price, currency."""
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            material = db.get(SupplierMaterial, material_id)
            if not material or material.supplier_id != payload.supplier_id:
                raise HTTPException(status_code=404, detail="Supplier material not found")
            if not (material.item_id or material.tier_material_id):
                payload.item_id = payload.tier_material_id = None
                material.code, material.name, material.uom = resolve_supplier_material(db, payload, request, material.id)
            elif payload.unit_price and not _check_price_access(request):
                raise HTTPException(status_code=403, detail="Price access is required to set a price")
            material.lead_time_days, material.buffer_days = payload.lead_time_days, payload.buffer_days
            if _check_price_access(request):
                material.unit_price, material.currency = payload.unit_price, payload.currency
            db.commit()
            return supplier_material_rows(db, [material], True)[0]

    @app.delete("/api/planning/supplier-materials/{material_id}", status_code=204)
    def delete_supplier_material(material_id: int, request: Request):
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            material = db.get(SupplierMaterial, material_id)
            if not material:
                raise HTTPException(status_code=404, detail="Supplier material not found")
            used = db.scalar(select(SupplierBOMLine.id).where(SupplierBOMLine.supplier_material_id == material_id).limit(1)) or \
                db.scalar(select(SupplierBOMLineSource.id).where(SupplierBOMLineSource.supplier_material_id == material_id).limit(1))
            if used:
                raise HTTPException(status_code=409, detail=f"{material.code} is used in a supplier BOM; remove it there first")
            db.delete(material)
            db.commit()
            return Response(status_code=204)

    @app.post("/api/planning/supplier-materials/upload", status_code=201)
    async def upload_supplier_materials(request: Request, supplier_id: int):
        """A supplier's material master from Excel (heading row found automatically): Code, Name,
        Unit, Unit price / Unit cost, Currency. A code found in your T1/T2/T3 Material Master is
        linked to it (its name and unit come from there); any other code is the supplier's own
        code and needs a name. Existing entries are updated. All rows or none."""
        current_user(request)
        _check_supply_role(request)
        rows = _read_flexible_sheet(await request.body(), SUPPLIER_MATERIAL_COLUMN_ALIASES, {"code"})
        can_price = _check_price_access(request)
        with open_db() as db:
            if not db.get(Supplier, supplier_id):
                raise HTTPException(status_code=404, detail="Supplier not found")
            existing = {m.code.lower(): m for m in db.scalars(select(SupplierMaterial).where(SupplierMaterial.supplier_id == supplier_id)).all()}
            t1 = {i.code.lower(): i for i in db.scalars(select(Item)).all()}
            t23 = {m.code.lower(): m for m in db.scalars(select(TierMaterial)).all()}
            seen: set[str] = set()
            added = updated = linked = 0
            for row in rows:
                row_number = row["_excel_row"]
                code = str(row.get("code") or "").strip()
                if not code:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: code is required")
                if code.lower() in seen:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {code} is repeated in the file")
                seen.add(code.lower())
                price, currency = _price_cells(row, row_number, can_price)
                item, tier_material = t1.get(code.lower()), t23.get(code.lower())
                if item or tier_material:
                    master = item or tier_material
                    code, name, uom = master.code, master.name, master.uom
                    linked += 1
                else:
                    name = str(row.get("name") or "").strip()
                    if not name:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: {code} is not in your T1/T2/T3 Material Master, so as the supplier's own code it needs a name")
                    uom = str(row.get("uom") or "ea").strip() or "ea"
                terms = {}
                for key in ("lead_time_days", "buffer_days"):
                    value = row.get(key)
                    if value is not None and str(value).strip() != "":
                        try:
                            terms[key] = int(float(value))
                        except (TypeError, ValueError) as exc:
                            raise HTTPException(status_code=422, detail=f"Row {row_number}: {key.replace('_', ' ')} must be a number") from exc
                        if terms[key] < 0:
                            raise HTTPException(status_code=422, detail=f"Row {row_number}: {key.replace('_', ' ')} must be zero or more")
                material = existing.get(code.lower())
                if material:
                    material.code, material.name, material.uom = code, name, uom
                    material.item_id, material.tier_material_id = (item.id if item else None), (tier_material.id if tier_material else None)
                    for key, value in terms.items():
                        setattr(material, key, value)
                    if can_price:
                        material.unit_price, material.currency = price, currency
                    updated += 1
                else:
                    material = SupplierMaterial(supplier_id=supplier_id, code=code, name=name, uom=uom, unit_price=price, currency=currency,
                                                item_id=item.id if item else None, tier_material_id=tier_material.id if tier_material else None,
                                                lead_time_days=terms.get("lead_time_days", 0), buffer_days=terms.get("buffer_days", 0))
                    db.add(material)
                    existing[code.lower()] = material
                    added += 1
            db.commit()
            return {"rows": len(rows), "added": added, "updated": updated, "linked": linked}

    # ------------------------------------------------------------------ supplier BOM (cost breakdown)
    def supplier_bom_view(db: Session, supplier_id: int, material_id: int, can_price: bool) -> dict:
        supplier, product = db.get(Supplier, supplier_id), db.get(Item, material_id)
        if not supplier or not product:
            raise HTTPException(status_code=404, detail="Supplier or part not found")
        quote = db.scalar(select(SupplierPart).where(
            SupplierPart.material_id == material_id, SupplierPart.supplier_id == supplier_id,
            SupplierPart.tier == "Direct", SupplierPart.parent_supplier_id.is_(None)))
        lines = db.scalars(select(SupplierBOMLine).where(
            SupplierBOMLine.supplier_id == supplier_id, SupplierBOMLine.product_material_id == material_id).order_by(SupplierBOMLine.id)).all()
        line_ids = [line.id for line in lines]
        sources: dict[int, list[SupplierBOMLineSource]] = {}
        if line_ids:
            for source in db.scalars(select(SupplierBOMLineSource).where(SupplierBOMLineSource.line_id.in_(line_ids)).order_by(SupplierBOMLineSource.id)).all():
                sources.setdefault(source.line_id, []).append(source)
        item_ids = {line.item_id for line in lines if line.item_id} | {src.item_id for group in sources.values() for src in group if src.item_id}
        own_ids = {line.supplier_material_id for line in lines if line.supplier_material_id} | {src.supplier_material_id for group in sources.values() for src in group if src.supplier_material_id}
        sub_ids = {src.sub_supplier_id for group in sources.values() for src in group if src.sub_supplier_id}
        items = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_(item_ids))).all()} if item_ids else {}
        owns = {m.id: m for m in db.scalars(select(SupplierMaterial).where(SupplierMaterial.id.in_(own_ids))).all()} if own_ids else {}
        tier_of = dict(db.execute(select(TierMaterial.id, TierMaterial.tier).where(
            TierMaterial.id.in_([m.tier_material_id for m in owns.values() if m.tier_material_id])
        )).all()) if any(m.tier_material_id for m in owns.values()) else {}
        subs = {sub.id: sub for sub in db.scalars(select(Supplier).where(Supplier.id.in_(sub_ids))).all()} if sub_ids else {}
        children: dict[int | None, list[SupplierBOMLine]] = {}
        for line in lines:
            children.setdefault(line.parent_line_id, []).append(line)
        base_currency = quote.currency if quote else (product.currency or "INR")

        def describe(item_id: int | None, own_id: int | None) -> dict:
            if item_id and item_id in items:
                item = items[item_id]
                return {"kind": "our", "origin": "T1", "id": item.id, "code": item.code, "name": item.name, "uom": item.uom, "price": item.unit_cost, "currency": item.currency}
            own = owns.get(own_id)
            origin = ("T1" if own.item_id else tier_of.get(own.tier_material_id, "T2/T3") if own.tier_material_id else "Own") if own else ""
            return {"kind": "supplier", "origin": origin, "id": own.id if own else None, "code": own.code if own else "?", "name": own.name if own else "?",
                    "uom": own.uom if own else "", "price": own.unit_price if own else 0, "currency": own.currency if own else base_currency}

        rows: list[dict] = []

        def unit_cost(line: SupplierBOMLine, depth: int) -> tuple[float, str, str, bool]:
            """(unit cost, currency, price source, currency mismatch) of one unit of this line."""
            component = describe(line.item_id, line.supplier_material_id)
            row = {"id": line.id, "parent_line_id": line.parent_line_id, "depth": depth, "level": f"L{depth}",
                   "component": {key: component[key] for key in ("kind", "origin", "id", "code", "name", "uom")}, "quantity": line.quantity,
                   "entered_price": line.unit_price, "entered_currency": line.currency,
                   "sources": [{"id": src.id, "sub_supplier_id": src.sub_supplier_id, "sub_supplier": subs[src.sub_supplier_id].name if src.sub_supplier_id in subs else "",
                                "alternate": (lambda c: None if not (src.item_id or src.supplier_material_id) else {key: c[key] for key in ("kind", "id", "code", "name")})(describe(src.item_id, src.supplier_material_id)),
                                "unit_price": src.unit_price, "currency": src.currency, "sob_percent": src.sob_percent} for src in sources.get(line.id, [])]}
            rows.append(row)
            kids = children.get(line.id, [])
            if kids:
                total, skipped = 0.0, []
                for kid in kids:
                    cost, currency, _, _ = unit_cost(kid, depth + 1)
                    if currency == base_currency:
                        total += kid.quantity * cost
                    else:
                        skipped.append(kid.id)
                price, currency, source = total, base_currency, "Sum of lines below"
            elif sources.get(line.id) and len({src.currency for src in sources[line.id]}) == 1:
                group = sources[line.id]
                share = sum(src.sob_percent for src in group)
                price = sum(src.unit_price * src.sob_percent for src in group) / share
                currency, source = group[0].currency, f"SOB-weighted ({len(group)} source{'s' if len(group) > 1 else ''})"
            elif line.unit_price is not None:
                price, currency, source = line.unit_price, line.currency or component["currency"], "Entered on line"
            else:
                price, currency = component["price"], component["currency"]
                source = "Our Material Master" if component["kind"] == "our" else "Supplier material list"
                if sources.get(line.id):
                    source += " (sources in different currencies)"
            row.update(has_children=bool(kids), unit_price=price, currency=currency, price_source=source,
                       line_cost=line.quantity * price, currency_mismatch=currency != base_currency)
            return price, currency, source, currency != base_currency

        should_cost = 0.0
        for line in children.get(None, []):
            cost, currency, _, _ = unit_cost(line, 1)
            if currency == base_currency:
                should_cost += line.quantity * cost
        result = {
            "supplier": {"id": supplier.id, "code": supplier.code, "name": supplier.name},
            "product": {"id": product.id, "code": product.code, "name": product.name, "uom": product.uom},
            "lines": rows, "currency": base_currency,
            "excluded_lines": [row["id"] for row in rows if row["currency_mismatch"]],
        }
        if can_price:
            result.update(should_cost=should_cost,
                          quoted_price=quote.unit_price if quote else None,
                          gap=(quote.unit_price - should_cost) if quote else None)
        else:
            for row in rows:
                for key in ("unit_price", "line_cost", "entered_price"):
                    row.pop(key, None)
                for src in row["sources"]:
                    src.pop("unit_price", None)
        return result

    @app.get("/api/planning/supplier-bom/products")
    def supplier_bom_products(request: Request, supplier_id: int):
        """Parts you buy from this supplier (Direct relationships), with BOM flag."""
        current_user(request)
        with open_db() as db:
            links = db.scalars(select(SupplierPart).where(SupplierPart.supplier_id == supplier_id, SupplierPart.tier == "Direct",
                                                          SupplierPart.parent_supplier_id.is_(None))).all()
            ids = sorted({link.material_id for link in links})
            items = db.scalars(select(Item).where(Item.id.in_(ids)).order_by(Item.code)).all() if ids else []
            with_bom = set(db.scalars(select(SupplierBOMLine.product_material_id).where(
                SupplierBOMLine.supplier_id == supplier_id).distinct()).all())
            return {"items": [{"id": item.id, "code": item.code, "name": item.name, "uom": item.uom, "has_bom": item.id in with_bom} for item in items]}

    @app.get("/api/planning/supplier-bom")
    def get_supplier_bom(request: Request, supplier_id: int, material_id: int):
        current_user(request)
        with open_db() as db:
            return supplier_bom_view(db, supplier_id, material_id, _check_price_access(request))

    def check_component(db: Session, supplier_id: int, item_id: int | None, own_id: int | None, label: str) -> None:
        # Supplier BOM codes come only from the supplier's material master (which holds the codes
        # picked from your T1/T2/T3 masters plus the supplier's own codes).
        if item_id:
            raise HTTPException(status_code=422, detail=f"{label}: add the code to this supplier's material master first, then pick it from there")
        if not own_id:
            raise HTTPException(status_code=422, detail=f"{label}: choose a part from the supplier's material master")
        if own_id:
            own = db.get(SupplierMaterial, own_id)
            if not own or own.supplier_id != supplier_id:
                raise HTTPException(status_code=422, detail=f"{label}: part is not in this supplier's material list")

    def line_depth(db: Session, line: SupplierBOMLine | None) -> int:
        depth = 0
        while line is not None:
            depth += 1
            line = db.get(SupplierBOMLine, line.parent_line_id) if line.parent_line_id else None
        return depth

    @app.post("/api/planning/supplier-bom/lines", status_code=201)
    def add_supplier_bom_line(payload: SupplierBOMLineInput, request: Request):
        current_user(request)
        _check_supply_role(request)
        if payload.quantity == 0:
            raise HTTPException(status_code=422, detail="Quantity cannot be zero")
        if payload.unit_price is not None and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to set a price")
        with open_db() as db:
            if not db.get(Supplier, payload.supplier_id) or not db.get(Item, payload.material_id):
                raise HTTPException(status_code=404, detail="Supplier or part not found")
            reason = not_purchased_reason(db, payload.material_id)
            if reason:
                raise HTTPException(status_code=422, detail=f"{db.get(Item, payload.material_id).code} {reason}; supplier BOMs are only for parts you buy")
            parent = None
            if payload.parent_line_id is not None:
                parent = db.get(SupplierBOMLine, payload.parent_line_id)
                if not parent or parent.supplier_id != payload.supplier_id or parent.product_material_id != payload.material_id:
                    raise HTTPException(status_code=404, detail="Parent line not found in this supplier BOM")
                if line_depth(db, parent) >= 3:
                    raise HTTPException(status_code=422, detail="Supplier BOMs go down to L3; an L3 line cannot have lines below it")
            check_component(db, payload.supplier_id, payload.item_id, payload.supplier_material_id, "Line")
            line = SupplierBOMLine(supplier_id=payload.supplier_id, product_material_id=payload.material_id, parent_line_id=payload.parent_line_id,
                                   item_id=payload.item_id, supplier_material_id=payload.supplier_material_id, quantity=payload.quantity,
                                   unit_price=payload.unit_price, currency=payload.currency)
            db.add(line)
            db.commit()
            return {"id": line.id}

    @app.patch("/api/planning/supplier-bom/lines/{line_id}")
    def update_supplier_bom_line(line_id: int, payload: SupplierBOMLineUpdate, request: Request):
        current_user(request)
        _check_supply_role(request)
        if payload.quantity == 0:
            raise HTTPException(status_code=422, detail="Quantity cannot be zero")
        with open_db() as db:
            line = db.get(SupplierBOMLine, line_id)
            if not line:
                raise HTTPException(status_code=404, detail="Supplier BOM line not found")
            line.quantity = payload.quantity
            if _check_price_access(request):
                line.unit_price, line.currency = payload.unit_price, payload.currency
            db.commit()
            return {"id": line.id}

    @app.delete("/api/planning/supplier-bom/lines/{line_id}", status_code=204)
    def delete_supplier_bom_line(line_id: int, request: Request):
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            line = db.get(SupplierBOMLine, line_id)
            if not line:
                raise HTTPException(status_code=404, detail="Supplier BOM line not found")
            pending, doomed = [line.id], []
            while pending:  # the line and every line below it
                current = pending.pop()
                doomed.append(current)
                pending.extend(db.scalars(select(SupplierBOMLine.id).where(SupplierBOMLine.parent_line_id == current)).all())
            db.execute(delete(SupplierBOMLineSource).where(SupplierBOMLineSource.line_id.in_(doomed)))
            for current in reversed(doomed):
                db.execute(delete(SupplierBOMLine).where(SupplierBOMLine.id == current))
            db.commit()
            return Response(status_code=204)

    @app.put("/api/planning/supplier-bom/lines/{line_id}/sources")
    def set_supplier_bom_sources(line_id: int, payload: list[SupplierBOMSourceInput], request: Request):
        """Replace the sub-suppliers / alternate codes of a line. SOB in total at most 100%."""
        current_user(request)
        _check_supply_role(request)
        if any(src.unit_price for src in payload) and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to set a price")
        if sum(src.sob_percent for src in payload) > 100.0001:
            raise HTTPException(status_code=422, detail="SOB of the sources adds up to more than 100%")
        with open_db() as db:
            line = db.get(SupplierBOMLine, line_id)
            if not line:
                raise HTTPException(status_code=404, detail="Supplier BOM line not found")
            if db.scalar(select(SupplierBOMLine.id).where(SupplierBOMLine.parent_line_id == line_id).limit(1)):
                raise HTTPException(status_code=422, detail="This line has lines below it; its price is their sum, so it takes no sources")
            for number, src in enumerate(payload, start=1):
                if src.sub_supplier_id is not None and not db.get(Supplier, src.sub_supplier_id):
                    raise HTTPException(status_code=422, detail=f"Source {number}: sub-supplier not found in the Supplier register")
                if src.item_id or src.supplier_material_id:
                    check_component(db, line.supplier_id, src.item_id, src.supplier_material_id, f"Source {number} alternate")
            db.execute(delete(SupplierBOMLineSource).where(SupplierBOMLineSource.line_id == line_id))
            for src in payload:
                db.add(SupplierBOMLineSource(line_id=line_id, **src.model_dump()))
            db.commit()
            return {"line_id": line_id, "sources": len(payload)}

    @app.post("/api/planning/supplier-bom/upload", status_code=201)
    async def upload_supplier_bom(request: Request, supplier_id: int, material_id: int):
        """Replace a supplier BOM from Excel: Parent code (blank = L1), Code, Quantity,
        Unit price (optional). Every code must be in this supplier's material master."""
        current_user(request)
        _check_supply_role(request)
        rows = _read_flexible_sheet(await request.body(), SUPPLIER_BOM_COLUMN_ALIASES, {"code", "quantity"})
        can_price = _check_price_access(request)
        with open_db() as db:
            if not db.get(Supplier, supplier_id) or not db.get(Item, material_id):
                raise HTTPException(status_code=404, detail="Supplier or part not found")
            reason = not_purchased_reason(db, material_id)
            if reason:
                raise HTTPException(status_code=422, detail=f"{db.get(Item, material_id).code} {reason}; supplier BOMs are only for parts you buy")
            own = {m.code.lower(): m for m in db.scalars(select(SupplierMaterial).where(SupplierMaterial.supplier_id == supplier_id)).all()}
            parsed: list[dict] = []
            by_code: dict[str, list[int]] = {}
            for row in rows:
                number_text = row["_excel_row"]
                code = str(row.get("code") or "").strip()
                parent_code = str(row.get("parent_code") or "").strip()
                if not code:
                    raise HTTPException(status_code=422, detail=f"Row {number_text}: code is required")
                target = ("supplier", own.get(code.lower()))
                if target[1] is None:
                    raise HTTPException(status_code=422, detail=f"Row {number_text}: {code} is not in this supplier's material master; add it there first (pick it from T1/T2/T3 or as the supplier's own code)")
                try:
                    quantity = float(row.get("quantity"))
                except (TypeError, ValueError) as exc:
                    raise HTTPException(status_code=422, detail=f"Row {number_text}: quantity must be a non-zero number") from exc
                if quantity == 0 or not math.isfinite(quantity):
                    raise HTTPException(status_code=422, detail=f"Row {number_text}: quantity must be a non-zero number")
                price = row.get("unit_price")
                if price is not None and str(price).strip() != "":
                    try:
                        price = float(price)
                    except (TypeError, ValueError) as exc:
                        raise HTTPException(status_code=422, detail=f"Row {number_text}: unit price must be a number") from exc
                    if price < 0:
                        raise HTTPException(status_code=422, detail=f"Row {number_text}: unit price must be zero or more")
                    if not can_price:
                        raise HTTPException(status_code=403, detail="Price access is required to upload prices")
                else:
                    price = None
                parent_index, depth = None, 1
                if parent_code:
                    matches = by_code.get(parent_code.lower(), [])
                    if not matches:
                        raise HTTPException(status_code=422, detail=f"Row {number_text}: parent code {parent_code} must appear on an earlier row")
                    if len(matches) > 1:
                        raise HTTPException(status_code=422, detail=f"Row {number_text}: parent code {parent_code} appears more than once above; it is ambiguous")
                    parent_index = matches[0]
                    depth = parsed[parent_index]["depth"] + 1
                    if depth > 3:
                        raise HTTPException(status_code=422, detail=f"Row {number_text}: supplier BOMs go down to L3")
                by_code.setdefault(code.lower(), []).append(len(parsed))
                parsed.append({"kind": target[0], "target": target[1], "quantity": quantity, "price": price, "parent": parent_index, "depth": depth})
            old = db.scalars(select(SupplierBOMLine.id).where(SupplierBOMLine.supplier_id == supplier_id, SupplierBOMLine.product_material_id == material_id)).all()
            if old:
                db.execute(delete(SupplierBOMLineSource).where(SupplierBOMLineSource.line_id.in_(old)))
                db.execute(delete(SupplierBOMLine).where(SupplierBOMLine.id.in_(old)))
            created: list[SupplierBOMLine] = []
            for entry in parsed:
                target = entry["target"]
                line = SupplierBOMLine(supplier_id=supplier_id, product_material_id=material_id,
                                       parent_line_id=created[entry["parent"]].id if entry["parent"] is not None else None,
                                       item_id=target.id if entry["kind"] == "our" else None,
                                       supplier_material_id=target.id if entry["kind"] == "supplier" else None,
                                       quantity=entry["quantity"], unit_price=entry["price"],
                                       currency=(target.currency if entry["price"] is not None else None))
                db.add(line)
                db.flush()
                created.append(line)
            db.commit()
            return {"rows": len(parsed)}

    @app.get("/api/planning/supplier-bom/{file_format}")
    def export_supplier_bom(file_format: str, request: Request, supplier_id: int, material_id: int):
        current_user(request)
        if file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export format not supported")
        can_price = _check_price_access(request)
        with open_db() as db:
            view = supplier_bom_view(db, supplier_id, material_id, can_price)
        headers = ["Level", "L1 code", "L1 name", "L2 code", "L2 name", "L3 code", "L3 name", "Source", "Quantity", "Unit"]
        if can_price:
            headers += ["Unit price", "Currency", "Line cost", "Price from"]
        path: dict[int, list[str]] = {}
        rows = []
        for line in view["lines"]:
            parent_path = path.get(line["parent_line_id"], [])
            mine = parent_path + [line["component"]["code"], line["component"]["name"]]
            path[line["id"]] = mine
            values = [line["level"], *(mine + [""] * 6)[:6], {"Own": "Supplier's own code"}.get(line["component"]["origin"], f"{line['component']['origin']} Material Master"),
                      line["quantity"], line["component"]["uom"]]
            if can_price:
                values += [line["unit_price"], line["currency"], line["line_cost"], line["price_source"]]
            rows.append({"depth": line["depth"] - 1, "has_children": line["has_children"], "values": values})
        title = f"Supplier BOM {view['supplier']['name']} - {view['product']['code']}"
        if can_price:
            gap = "" if view["quoted_price"] is None else f" | quoted {view['quoted_price']:,.2f} | gap {view['gap']:,.2f}"
            title += f" (should-cost {view['should_cost']:,.2f} {view['currency']}{gap})"
        if file_format == "xlsx":
            output, media_type = _bom_tree_workbook(title, headers, rows), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            output, media_type = _bom_tree_pdf(title, headers, rows), "application/pdf"
        filename = f"supplier-bom-{view['supplier']['code']}-{view['product']['code']}.{file_format}".replace(" ", "-")
        return StreamingResponse(output, media_type=media_type, headers={"Content-Disposition": f'attachment; filename="{filename}"'})

    @app.patch("/api/planning/supplier-parts/{link_id}")
    def update_supplier_part(link_id: int, payload: dict, request: Request):
        current_user(request)
        _check_supply_role(request)
        allowed = {"sob_percent", "lead_time_days", "moq", "buffer_days", "unit_price", "currency", "price_mode", "price_interval_months"}
        if set(payload) - allowed:
            raise HTTPException(status_code=422, detail="Unsupported supplier/material update field")
        if "unit_price" in payload and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to update a part price")
        with open_db() as db:
            link = db.get(SupplierPart, link_id)
            if not link:
                raise HTTPException(status_code=404, detail="Supplier/material link not found")
            if "sob_percent" in payload:
                share_total = db.scalar(select(func.coalesce(func.sum(SupplierPart.sob_percent), 0)).where(
                    SupplierPart.material_id == link.material_id,
                    SupplierPart.tier == link.tier,
                    SupplierPart.parent_supplier_id == link.parent_supplier_id,
                    SupplierPart.id != link.id,
                )) or 0
                if share_total + float(payload["sob_percent"]) > 100:
                    raise HTTPException(status_code=422, detail="Supplier SOB cannot exceed 100% within the same material tier and parent supplier")
            price_fields = {"unit_price", "price_mode", "price_interval_months"}
            price_changed = bool(price_fields.intersection(payload))
            for key, value in payload.items():
                setattr(link, key, value)
            if price_changed:
                link.price_updated_month = _month_text(date.today())
            db.commit()
            return {"id": link.id, "price_updated_month": link.price_updated_month}

    @app.post("/api/planning/plans", status_code=201)
    def upload_plan(payload: PlanUpload, request: Request):
        user = current_user(request)
        _check_workspace = _validate_workspace(payload.workspace)
        issue = _month(payload.issue_month)
        issue_month = _month_text(issue)
        with open_db() as db:
            models_by_code = {m.code: m for m in db.scalars(select(ProductModel)).all()}
            latest_revision = db.scalar(select(func.max(PlanningCycle.revision)).where(PlanningCycle.workspace == payload.workspace, PlanningCycle.issue_month == issue_month))
            revision = 0 if latest_revision is None else latest_revision + 1
            normalized = []
            seen = set()
            for row in payload.rows:
                code = str(row.get("model_code", "")).strip()
                target_date = _month(row.get("target_month"))
                target = _month_text(target_date)
                if target_date < issue or target_date > _plus_months(issue, 23):
                    raise HTTPException(status_code=422, detail="Plan target months must fall within the 24-month horizon starting from the issue month")
                quantity = float(row.get("quantity", 0))
                if quantity < 0 or quantity > 1_000_000_000:
                    raise HTTPException(status_code=422, detail="Plan quantity must be between zero and one billion")
                model = models_by_code.get(code)
                if not model:
                    raise HTTPException(status_code=422, detail=f"Unknown model code: {code}")
                key = (target, model.id)
                if key in seen:
                    raise HTTPException(status_code=422, detail=f"Duplicate plan row for model {code} in {target}")
                seen.add(key)
                normalized.append((target, model.id, quantity))
            cycle = PlanningCycle(workspace=payload.workspace, issue_month=issue_month, revision=revision, created_by=user["id"])
            db.add(cycle)
            db.flush()
            db.add_all(PlanLine(cycle_id=cycle.id, target_month=target, model_id=model_id, quantity=quantity) for target, model_id, quantity in normalized)
            db.commit()
            return {"id": cycle.id, "workspace": cycle.workspace, "issue_month": cycle.issue_month, "revision": f"Rev.{revision:02d}", "revision_number": revision, "rows": len(normalized)}

    @app.post("/api/planning/plans/{issue_month}/upload", status_code=201)
    async def upload_plan_excel(issue_month: str, request: Request, workspace: str = "trial"):
        raw = await request.body()
        rows = _read_xlsx(raw)
        payload = PlanUpload(issue_month=_month_text(_month(issue_month)), workspace=workspace, rows=_plan_rows_from_excel(rows))
        return upload_plan(payload, request)

    @app.get("/api/planning/plans")
    def list_plans(request: Request, workspace: str = "trial", issue_month: str | None = None):
        current_user(request)
        _validate_workspace(workspace)
        if issue_month:
            issue_month = _month_text(_month(issue_month))
        with open_db() as db:
            statement = select(PlanningCycle).where(PlanningCycle.workspace == workspace).order_by(PlanningCycle.issue_month.desc(), PlanningCycle.revision.desc())
            if issue_month:
                statement = statement.where(PlanningCycle.issue_month == issue_month)
            result = []
            for cycle in db.scalars(statement.limit(500)).all():
                result.append({
                    "id": cycle.id, "workspace": cycle.workspace, "issue_month": cycle.issue_month,
                    "revision": f"Rev.{cycle.revision:02d}", "revision_number": cycle.revision,
                    "created_at": cycle.created_at.isoformat(),
                    "row_count": db.scalar(select(func.count(PlanLine.id)).where(PlanLine.cycle_id == cycle.id)) or 0,
                })
            return result

    @app.get("/api/planning/plans/{cycle_id}/rows")
    def plan_revision_rows(
        cycle_id: int, request: Request, offset: int = 0, limit: int = 25, search: str = "",
    ):
        current_user(request)
        if offset < 0 or limit < 1 or limit > 100:
            raise HTTPException(status_code=422, detail="Use limit 1–100 and a non-negative offset")
        with open_db() as db:
            cycle = db.get(PlanningCycle, cycle_id)
            if not cycle:
                raise HTTPException(status_code=404, detail="Plan revision not found")
            model_query = select(ProductModel).order_by(ProductModel.code)
            if search.strip():
                like = f"%{search.strip()}%"
                model_query = model_query.where(ProductModel.code.ilike(like) | ProductModel.name.ilike(like))
            total = db.scalar(select(func.count()).select_from(model_query.subquery())) or 0
            models_page = db.scalars(model_query.offset(offset).limit(limit)).all()
            model_ids = [model.id for model in models_page]
            plan_lines = db.scalars(select(PlanLine).where(
                PlanLine.cycle_id == cycle_id, PlanLine.model_id.in_(model_ids),
            )).all() if model_ids else []
            quantities: dict[int, dict[str, float]] = {}
            for line in plan_lines:
                quantities.setdefault(line.model_id, {})[line.target_month] = line.quantity
            issue = _month(cycle.issue_month)
            months = [_month_text(_plus_months(issue, index)) for index in range(24)]
            latest_revision = db.scalar(select(func.max(PlanningCycle.revision)).where(
                PlanningCycle.workspace == cycle.workspace, PlanningCycle.issue_month == cycle.issue_month,
            ))
            return {
                "id": cycle.id, "workspace": cycle.workspace, "issue_month": cycle.issue_month,
                "revision": f"Rev.{cycle.revision:02d}", "revision_number": cycle.revision,
                "created_at": cycle.created_at.isoformat(), "latest": cycle.revision == latest_revision,
                "offset": offset, "limit": limit, "total_models": total, "months": months,
                "items": [{
                    "model_code": model.code, "model_name": model.name,
                    "quantities": quantities.get(model.id, {}),
                } for model in models_page],
            }

    @app.post("/api/planning/plans/{cycle_id}/revisions", status_code=201)
    def revise_plan(cycle_id: int, payload: PlanRevisionUpdate, request: Request):
        user = current_user(request)
        with open_db() as db:
            source = db.get(PlanningCycle, cycle_id)
            if not source:
                raise HTTPException(status_code=404, detail="Plan revision not found")
            latest_revision = db.scalar(select(func.max(PlanningCycle.revision)).where(
                PlanningCycle.workspace == source.workspace, PlanningCycle.issue_month == source.issue_month,
            ))
            if source.revision != latest_revision:
                raise HTTPException(status_code=409, detail="This revision is no longer the latest. Open the latest revision before saving changes.")
            issue = _month(source.issue_month)
            models_by_code = {model.code: model for model in db.scalars(select(ProductModel)).all()}
            changes: dict[tuple[int, str], float] = {}
            for row in payload.rows:
                model = models_by_code.get(row.model_code.strip())
                if not model:
                    raise HTTPException(status_code=422, detail=f"Unknown model code: {row.model_code}")
                target_date = _month(row.target_month)
                target_month = _month_text(target_date)
                if target_date < issue or target_date > _plus_months(issue, 23):
                    raise HTTPException(status_code=422, detail="Edited plan months must fall within the revision's 24-month horizon")
                if not math.isfinite(row.quantity):
                    raise HTTPException(status_code=422, detail="Plan quantity must be a finite number")
                key = (model.id, target_month)
                if key in changes:
                    raise HTTPException(status_code=422, detail=f"Duplicate edit for model {model.code} in {target_month}")
                changes[key] = row.quantity
            original = {
                (line.model_id, line.target_month): line.quantity
                for line in db.scalars(select(PlanLine).where(PlanLine.cycle_id == source.id)).all()
            }
            original.update(changes)
            new_cycle = PlanningCycle(
                workspace=source.workspace, issue_month=source.issue_month,
                revision=source.revision + 1, created_by=user["id"],
            )
            db.add(new_cycle)
            db.flush()
            db.add_all(
                PlanLine(cycle_id=new_cycle.id, model_id=model_id, target_month=month, quantity=quantity)
                for (model_id, month), quantity in original.items()
            )
            db.commit()
            db.refresh(new_cycle)
            return {
                "id": new_cycle.id, "workspace": new_cycle.workspace,
                "issue_month": new_cycle.issue_month, "revision": f"Rev.{new_cycle.revision:02d}",
                "revision_number": new_cycle.revision, "rows": len(original),
                "changed_rows": len(changes), "created_at": new_cycle.created_at.isoformat(),
            }

    @app.post("/api/planning/actuals", status_code=201)
    def upload_actuals(payload: ActualUpload, request: Request):
        user = current_user(request)
        _validate_workspace(payload.workspace)
        _month(payload.period)
        with open_db() as db:
            models_by_code = {m.code: m for m in db.scalars(select(ProductModel)).all()}
            saved = 0
            for row in payload.rows:
                code = str(row.get("model_code", "")).strip()
                model = models_by_code.get(code)
                if not model:
                    raise HTTPException(status_code=422, detail=f"Unknown model code: {code}")
                quantity = float(row.get("actual_quantity", row.get("quantity", 0)))
                oem_stock = float(row.get("oem_stock", 0))
                if min(quantity, oem_stock) < 0:
                    raise HTTPException(status_code=422, detail="Actual production and OEM stock cannot be negative")
                existing = db.scalar(select(ActualProduction).where(ActualProduction.workspace == payload.workspace, ActualProduction.period == payload.period, ActualProduction.model_id == model.id))
                if existing:
                    existing.actual_quantity = quantity
                    existing.oem_stock = oem_stock
                    existing.updated_by = user["id"]
                else:
                    db.add(ActualProduction(workspace=payload.workspace, period=payload.period, model_id=model.id, actual_quantity=quantity, oem_stock=oem_stock, updated_by=user["id"]))
                saved += 1
            db.commit()
            return {"workspace": payload.workspace, "period": payload.period, "rows": saved}

    @app.get("/api/planning/actuals")
    def list_actuals(request: Request, period: str, workspace: str = "trial"):
        current_user(request)
        _validate_workspace(workspace)
        _month(period)
        with open_db() as db:
            models_by_id = {m.id: m for m in db.scalars(select(ProductModel)).all()}
            return [
                {"model_code": models_by_id[row.model_id].code if row.model_id in models_by_id else "", "model_name": models_by_id[row.model_id].name if row.model_id in models_by_id else "", "actual_quantity": row.actual_quantity, "oem_stock": row.oem_stock}
                for row in db.scalars(select(ActualProduction).where(ActualProduction.workspace == workspace, ActualProduction.period == period).order_by(ActualProduction.model_id)).all()
            ]

    @app.post("/api/planning/actuals/{period}/upload", status_code=201)
    async def upload_actuals_excel(period: str, request: Request, workspace: str = "trial"):
        rows = _read_xlsx(await request.body())
        _required_columns(rows, {"model_code"})
        if "actual_quantity" not in rows[0] and "quantity" not in rows[0]:
            raise HTTPException(status_code=422, detail="Excel sheet needs an actual_quantity column")
        return upload_actuals(ActualUpload(period=period, workspace=workspace, rows=rows), request)

    @app.post("/api/planning/own-stock", status_code=201)
    def upload_own_stock(payload: StockUpload, request: Request):
        current_user(request)
        _validate_workspace(payload.workspace)
        _month(payload.period)
        with open_db() as db:
            items = {item.code: item for item in db.scalars(select(Item)).all()}
            saved = 0
            for row in payload.rows:
                code = str(row.get("material_code", "")).strip()
                item = items.get(code)
                if not item:
                    raise HTTPException(status_code=422, detail=f"Unknown material code: {code}")
                quantity = float(row.get("quantity", 0))
                if quantity < 0:
                    raise HTTPException(status_code=422, detail="Stock quantity cannot be negative")
                existing = db.scalar(select(OwnStockSnapshot).where(OwnStockSnapshot.workspace == payload.workspace, OwnStockSnapshot.period == payload.period, OwnStockSnapshot.material_id == item.id))
                if existing:
                    existing.quantity = quantity
                else:
                    db.add(OwnStockSnapshot(workspace=payload.workspace, period=payload.period, material_id=item.id, quantity=quantity))
                saved += 1
            db.commit()
            return {"workspace": payload.workspace, "period": payload.period, "rows": saved}

    @app.post("/api/planning/supplier-stock/{report_month}/upload", status_code=201)
    async def upload_supplier_stock_excel(report_month: str, request: Request, workspace: str = "trial"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        _month(report_month)
        rows = _read_xlsx(await request.body())
        _required_columns(rows, {"supplier_code", "material_code", "quantity"})
        with open_db() as db:
            links = db.scalars(select(SupplierPart)).all()
            suppliers = {s.code: s for s in db.scalars(select(Supplier)).all()}
            materials_by_code = {i.code: i for i in db.scalars(select(Item)).all()}
            by_key: dict[tuple[int, int], list[SupplierPart]] = {}
            for link in links:
                by_key.setdefault((link.supplier_id, link.material_id), []).append(link)
            suppliers_by_id = {supplier.id: supplier for supplier in suppliers.values()}
            saved = 0
            for row in rows:
                supplier = suppliers.get(str(row.get("supplier_code", "")).strip())
                item = materials_by_code.get(str(row.get("material_code", "")).strip())
                if not supplier or not item:
                    raise HTTPException(status_code=422, detail="Every supplier_code and material_code must exist in the master data")
                candidates = by_key.get((supplier.id, item.id), [])
                tier = str(row.get("tier") or "").strip()
                parent_code = str(row.get("parent_supplier_code") or "").strip()
                if tier:
                    candidates = [candidate for candidate in candidates if candidate.tier == tier]
                if parent_code:
                    candidates = [
                        candidate for candidate in candidates
                        if candidate.parent_supplier_id
                        and suppliers_by_id.get(candidate.parent_supplier_id)
                        and suppliers_by_id[candidate.parent_supplier_id].code == parent_code
                    ]
                if not candidates:
                    raise HTTPException(status_code=422, detail=f"No supplier/material relationship configured for {supplier.code} and {item.code}")
                if len(candidates) > 1:
                    raise HTTPException(status_code=422, detail=f"Supplier/material relationship is ambiguous for {supplier.code} and {item.code}; include tier and parent_supplier_code")
                link = candidates[0]
                quantity = float(row.get("quantity") or 0)
                available = float(row.get("available_quantity") if row.get("available_quantity") is not None else quantity)
                if min(quantity, available) < 0 or available > quantity:
                    raise HTTPException(status_code=422, detail="Available quantity must be between zero and reported supplier stock")
                arrival = str(row.get("arrival_month") or "").strip() or None
                if arrival:
                    _month(arrival)
                existing = db.scalar(select(SupplierStockSnapshot).where(SupplierStockSnapshot.workspace == workspace, SupplierStockSnapshot.report_month == report_month, SupplierStockSnapshot.supplier_part_id == link.id))
                if existing:
                    existing.quantity, existing.available_quantity, existing.arrival_month = quantity, available, arrival
                else:
                    db.add(SupplierStockSnapshot(workspace=workspace, report_month=report_month, supplier_part_id=link.id, quantity=quantity, available_quantity=available, arrival_month=arrival))
                saved += 1
            db.commit()
            return {"workspace": workspace, "report_month": report_month, "rows": saved}

    @app.get("/api/planning/report")
    def analysis_report(request: Request, start_month: str, end_month: str, workspace: str = "trial", offset: int = 0, limit: int = 1000):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        if limit < 1 or limit > 5000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–5000 and a non-negative offset")
        with open_db() as db:
            rows = _planning_report(db, workspace, start_month, end_month)
            return {"workspace": workspace, "start_month": start_month, "end_month": end_month, "total": len(rows), "offset": offset, "limit": limit, "items": rows[offset:offset + limit]}

    @app.get("/api/planning/report.xlsx")
    def analysis_excel(request: Request, start_month: str, end_month: str, workspace: str = "trial"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        with open_db() as db:
            rows = _planning_report(db, workspace, start_month, end_month)
        output = _report_workbook(_without_prices(rows, request), f"Shortage and excess · {start_month} to {end_month}")
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="dtc-analysis-{start_month}-{end_month}.xlsx"'})

    @app.get("/api/planning/report.pdf")
    def analysis_pdf(request: Request, start_month: str, end_month: str, workspace: str = "trial"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        with open_db() as db:
            rows = _planning_report(db, workspace, start_month, end_month)
        output = _report_pdf(rows, f"Shortage and excess · {start_month} to {end_month}")
        return StreamingResponse(output, media_type="application/pdf", headers={"Content-Disposition": f'attachment; filename="dtc-analysis-{start_month}-{end_month}.pdf"'})

    @app.get("/api/planning/templates/{kind}.xlsx")
    def planning_template(kind: str, request: Request, issue_month: str | None = None):
        current_user(request)
        templates = {
            "actuals": ["model_code", "actual_quantity", "oem_stock"],
            "own-stock": ["material_code", "quantity"],
            "supplier-stock": ["supplier_code", "material_code", "tier", "parent_supplier_code", "quantity", "available_quantity", "arrival_month"],
            "supplier-parts": ["material_code", "supplier_name", "tier", "parent_supplier_name", "sob_percent", "unit_price", "currency",
                               "lead_time_days", "moq", "buffer_days", "price_mode", "price_interval_months"],
        }
        bom_templates = {
            "bom-oem": ["model_code", "material_code", "quantity", "uom"],
            "bom-level-2": ["parent_code", "material_code", "quantity", "uom"],
            "bom-level-3": ["parent_code", "material_code", "quantity", "uom"],
        }
        if kind not in templates and kind not in bom_templates and kind != "plan":
            raise HTTPException(status_code=404, detail="Template type not found")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = kind[:31]
        if kind == "plan":
            first_month = _month(issue_month or _month_text(date.today()))
            headers: list[Any] = ["model_code", "model_MONTH"]
            headers.extend(_plus_months(first_month, offset) for offset in range(24))
            sheet.append(headers)
            for column in range(3, 27):
                sheet.cell(1, column).number_format = "mmm-yy"
            sheet.freeze_panes = "C2"
            sheet.column_dimensions["A"].width = 18
            sheet.column_dimensions["B"].width = 26
            for column in range(3, 27):
                sheet.column_dimensions[__import__("openpyxl").utils.get_column_letter(column)].width = 14
            for column in range(3, 27):
                for row_index in range(2, 1002):
                    sheet.cell(row_index, column).number_format = "#,##0.###"
        else:
            headers = bom_templates.get(kind, templates.get(kind, []))
            sheet.append(headers)
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="12263A")
            cell.alignment = Alignment(horizontal="center")
        if kind != "plan":
            for index, name in enumerate(headers, start=1):
                sheet.column_dimensions[__import__("openpyxl").utils.get_column_letter(index)].width = max(18, len(name) + 4)
        if kind in bom_templates:
            with open_db() as db:
                units = sorted({item.uom for item in db.scalars(select(Item)).all() if item.uom})
            units = list(dict.fromkeys([*units, "ea", "PCS", "Nos", "kg", "g", "grm", "m", "mtr", "mm", "cm", "L", "ml", "set", "pair", "roll", "box", "pack", "lot"]))
            unit_sheet = workbook.create_sheet("Lists")
            unit_sheet.sheet_state = "hidden"
            for row_index, unit in enumerate(units, start=1):
                unit_sheet.cell(row_index, 1, unit)
            from openpyxl.workbook.defined_name import DefinedName
            workbook.defined_names.add(DefinedName("UnitOptions", attr_text=f"'Lists'!$A$1:$A${len(units)}"))
            unit_column = __import__("openpyxl").utils.get_column_letter(headers.index("uom") + 1)
            unit_validation = DataValidation(type="list", formula1="=UnitOptions", allow_blank=False)
            sheet.add_data_validation(unit_validation)
            unit_validation.add(f"{unit_column}2:{unit_column}10000")
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="dtc-{kind}-template.xlsx"'})

    @app.post("/api/planning/purchase-orders", status_code=201)
    def create_planning_po(payload: POInput, request: Request):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(payload.workspace)
        _month(payload.expected_month)
        with open_db() as db:
            link = db.get(SupplierPart, payload.supplier_part_id)
            if not link:
                raise HTTPException(status_code=404, detail="Supplier/material relationship not found")
            if link.unit_price <= 0:
                raise HTTPException(status_code=422, detail="Enter a valid current price before issuing this PO")
            number = f"SC-{int((db.scalar(select(func.count(PlanningPO.id))) or 0) + 1):06d}"
            if payload.quantity_ordered < link.moq:
                raise HTTPException(status_code=422, detail=f"Quantity is below supplier MOQ of {link.moq:g}")
            order = PlanningPO(
                po_number=number, workspace=payload.workspace, supplier_part_id=link.id,
                quantity_ordered=payload.quantity_ordered, unit_price=link.unit_price,
                currency=link.currency, ordered_month=_month_text(date.today()),
                expected_month=payload.expected_month, ncnr=payload.ncnr, status="Issued",
            )
            db.add(order)
            db.commit()
            return {"id": order.id, "po_number": number, "quantity_ordered": order.quantity_ordered, "unit_price": order.unit_price, "currency": order.currency, "expected_month": order.expected_month, "ncnr": order.ncnr, "status": order.status}

    @app.post("/api/planning/purchase-orders/{po_id}/receipts", status_code=201)
    def receive_planning_po(po_id: int, payload: POReceipt, request: Request):
        current_user(request)
        _check_supply_role(request)
        _month(payload.received_month)
        with open_db() as db:
            order = db.get(PlanningPO, po_id)
            if not order:
                raise HTTPException(status_code=404, detail="Planning PO not found")
            remaining = order.quantity_ordered - order.quantity_received
            if payload.quantity_received > remaining:
                raise HTTPException(status_code=422, detail=f"Receipt exceeds the remaining {remaining:g} units")
            link = db.get(SupplierPart, order.supplier_part_id)
            if not link:
                raise HTTPException(status_code=404, detail="PO supplier/material relationship not found")
            own_stock = db.scalar(select(OwnStockSnapshot).where(
                OwnStockSnapshot.workspace == order.workspace,
                OwnStockSnapshot.period == payload.received_month,
                OwnStockSnapshot.material_id == link.material_id,
            ))
            if own_stock:
                own_stock.quantity += payload.quantity_received
            else:
                db.add(OwnStockSnapshot(
                    workspace=order.workspace, period=payload.received_month,
                    material_id=link.material_id, quantity=payload.quantity_received,
                ))
            order.quantity_received += payload.quantity_received
            order.status = "Received" if order.quantity_received == order.quantity_ordered else "Partially received"
            db.commit()
            return {"id": order.id, "po_number": order.po_number, "quantity_received": order.quantity_received, "remaining": order.quantity_ordered - order.quantity_received, "status": order.status}

    @app.patch("/api/planning/purchase-orders/{po_id}/schedule")
    def change_po_schedule(po_id: int, payload: dict, request: Request):
        current_user(request)
        _check_supply_role(request)
        expected = str(payload.get("expected_month", "")).strip()
        _month(expected)
        if not payload.get("change_note"):
            raise HTTPException(status_code=422, detail="Explain the requested delivery schedule change")
        with open_db() as db:
            order = db.get(PlanningPO, po_id)
            if not order:
                raise HTTPException(status_code=404, detail="Planning PO not found")
            order.expected_month = expected
            order.change_note = str(payload["change_note"])[:2000]
            db.commit()
            return {"id": order.id, "po_number": order.po_number, "expected_month": order.expected_month, "change_note": order.change_note}

    @app.get("/api/planning/purchase-orders")
    def planning_pos(request: Request, start_month: str, end_month: str, workspace: str = "trial"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        _month(start_month)
        _month(end_month)
        with open_db() as db:
            orders = db.scalars(select(PlanningPO).where(PlanningPO.workspace == workspace, PlanningPO.ordered_month >= start_month, PlanningPO.ordered_month <= end_month).order_by(PlanningPO.ordered_month.desc(), PlanningPO.id.desc())).all()
            result = []
            for order in orders:
                link = db.get(SupplierPart, order.supplier_part_id)
                item = db.get(Item, link.material_id) if link else None
                supplier = db.get(Supplier, link.supplier_id) if link else None
                row = {"id": order.id, "po_number": order.po_number, "workspace": workspace, "material_code": item.code if item else "", "material_name": item.name if item else "", "supplier_name": supplier.name if supplier else "", "tier": link.tier if link else "", "quantity_ordered": order.quantity_ordered, "quantity_received": order.quantity_received, "expected_month": order.expected_month, "ordered_month": order.ordered_month, "status": order.status, "ncnr": order.ncnr}
                if _check_price_access(request):
                    row.update(unit_price=order.unit_price, currency=order.currency, total=order.quantity_ordered * order.unit_price)
                result.append(row)
            return result

    @app.post("/api/planning/transfers", status_code=201)
    def create_transfer(payload: TransferInput, request: Request):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(payload.workspace)
        _month(payload.transfer_month)
        if payload.from_supplier_id == payload.to_supplier_id:
            raise HTTPException(status_code=422, detail="Choose two different suppliers")
        with open_db() as db:
            if not db.get(Item, payload.material_id) or not db.get(Supplier, payload.from_supplier_id) or not db.get(Supplier, payload.to_supplier_id):
                raise HTTPException(status_code=404, detail="Material or supplier not found")
            transfer = MaterialTransfer(**payload.model_dump())
            db.add(transfer)
            db.commit()
            db.refresh(transfer)
            return {"id": transfer.id, "material_id": transfer.material_id, "from_supplier_id": transfer.from_supplier_id, "to_supplier_id": transfer.to_supplier_id, "quantity": transfer.quantity, "transfer_month": transfer.transfer_month, "status": transfer.status}

    @app.get("/api/planning/transfers")
    def list_transfers(request: Request, workspace: str = "trial"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        with open_db() as db:
            suppliers = {s.id: s for s in db.scalars(select(Supplier)).all()}
            items = {i.id: i for i in db.scalars(select(Item)).all()}
            return [{
                "id": transfer.id, "material_code": items[transfer.material_id].code if transfer.material_id in items else "",
                "material_name": items[transfer.material_id].name if transfer.material_id in items else "",
                "from_supplier": suppliers[transfer.from_supplier_id].name if transfer.from_supplier_id in suppliers else "",
                "to_supplier": suppliers[transfer.to_supplier_id].name if transfer.to_supplier_id in suppliers else "",
                "quantity": transfer.quantity, "transfer_month": transfer.transfer_month, "status": transfer.status, "note": transfer.note,
            } for transfer in db.scalars(select(MaterialTransfer).where(MaterialTransfer.workspace == workspace).order_by(MaterialTransfer.transfer_month.desc(), MaterialTransfer.id.desc())).all()]

    def supplier_business_rows(db: Session, start_month: str, end_month: str, workspace: str) -> list[dict]:
        orders = db.scalars(select(PlanningPO).where(
            PlanningPO.workspace == workspace,
            PlanningPO.ordered_month >= start_month,
            PlanningPO.ordered_month <= end_month,
        )).all()
        links = {link.id: link for link in db.scalars(select(SupplierPart)).all()}
        suppliers = {supplier.id: supplier for supplier in db.scalars(select(Supplier)).all()}
        grouped: dict[tuple[int, str, str], dict] = {}
        for order in orders:
            link = links.get(order.supplier_part_id)
            if not link:
                continue
            supplier = suppliers.get(link.supplier_id)
            key = (link.supplier_id, link.tier, order.currency)
            row = grouped.setdefault(key, {
                "supplier_id": link.supplier_id,
                "supplier_code": supplier.code if supplier else "",
                "supplier_name": supplier.name if supplier else "",
                "tier": link.tier,
                "currency": order.currency,
                "po_count": 0,
                "po_value": 0,
            })
            row["po_count"] += 1
            row["po_value"] += order.quantity_ordered * order.unit_price
        return sorted(grouped.values(), key=lambda r: (r["tier"], r["supplier_name"], r["currency"]))

    @app.get("/api/planning/supplier-business")
    def supplier_business(request: Request, start_month: str, end_month: str, workspace: str = "production"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        _month(start_month)
        _month(end_month)
        with open_db() as db:
            result = supplier_business_rows(db, start_month, end_month, workspace)
            if not _check_price_access(request):
                for row in result:
                    row.pop("po_value", None)
            return result

    @app.get("/api/planning/supplier-business.xlsx")
    def supplier_business_excel(request: Request, start_month: str, end_month: str, workspace: str = "production"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        _month(start_month)
        _month(end_month)
        with open_db() as db:
            report = supplier_business_rows(db, start_month, end_month, workspace)
        headers = ["Supplier code", "Supplier", "Tier", "Currency", "PO count", "Issued PO value"]
        rows = [[r["supplier_code"], r["supplier_name"], r["tier"], r["currency"], r["po_count"], r.get("po_value") if _check_price_access(request) else "Restricted"] for r in report]
        output = _generic_workbook(f"Supplier business · {start_month} to {end_month}", headers, rows)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="dtc-supplier-business-{start_month}-{end_month}.xlsx"'})

    @app.get("/api/planning/supplier-business.pdf")
    def supplier_business_pdf(request: Request, start_month: str, end_month: str, workspace: str = "production"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        _month(start_month)
        _month(end_month)
        with open_db() as db:
            report = supplier_business_rows(db, start_month, end_month, workspace)
        headers = ["Supplier code", "Supplier", "Tier", "Currency", "PO count", "Issued PO value"]
        rows = [[r["supplier_code"], r["supplier_name"], r["tier"], r["currency"], r["po_count"], r.get("po_value") if _check_price_access(request) else "Restricted"] for r in report]
        output = _generic_pdf(f"Supplier business · {start_month} to {end_month}", headers, rows)
        return StreamingResponse(output, media_type="application/pdf", headers={"Content-Disposition": f'attachment; filename="dtc-supplier-business-{start_month}-{end_month}.pdf"'})

    BOM_TREE_LEVELS = {
        "Product Master": ["product", "model", "l1", "l2", "l3"],
        "Model BOM Level 1": ["model", "l1", "l2", "l3"],
        "Level 2 Part BOM": ["l1", "l2", "l3"],
        "Level 3 Part BOM": ["l2", "l3"],
    }
    BOM_TREE_LABELS = {
        "product": ("Product code", "Product name"), "model": ("Model code", "Model name"),
        "l1": ("L1 part code", "L1 part name"), "l2": ("L2 part code", "L2 part name"), "l3": ("L3 part code", "L3 part name"),
    }
    BOM_TREE_LEVEL_NAMES = {"product": "Product", "model": "Model", "l1": "L1", "l2": "L2", "l3": "L3"}

    @app.get("/api/planning/bom-tree/{file_format}")
    def export_bom_tree(
        file_format: str, request: Request, tab: str = "Product Master",
        model_id: list[int] = Query(default=[]), parent_item_id: list[int] = Query(default=[]),
    ):
        """Full BOM tree breakdown of a BOM tab (same columns as the screen), every row
        repeating its parents' codes and names."""
        current_user(request)
        if file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export format not supported")
        if tab not in BOM_TREE_LEVELS:
            raise HTTPException(status_code=422, detail="Unknown BOM tab")
        levels = BOM_TREE_LEVELS[tab]
        can_price = _check_price_access(request)
        tree: list[dict] = []
        with open_db() as db:
            items = {item.id: item for item in db.scalars(select(Item)).all()}
            price = bom_price_lookup(db)
            prices = {item_id: price(item) for item_id, item in items.items()}
            children: dict[int, list[BOMComponent]] = {}
            for link in db.scalars(select(BOMComponent).order_by(BOMComponent.id)).all():
                children.setdefault(link.parent_item_id, []).append(link)

            def emit(depth: int, level: str, path: dict, has_children: bool, quantity=None, item: Item | None = None) -> None:
                tree.append({"depth": depth, "level": level, "path": path, "has_children": has_children, "quantity": quantity, "item": item})

            def part_lines(parent_id: int, level: str, depth: int, path: dict) -> None:
                for link in children.get(parent_id, []):
                    child = items.get(link.child_item_id)
                    if not child:
                        continue
                    child_path = {**path, level: (child.code, child.name)}
                    deeper = level == "l2" and bool(children.get(child.id))
                    emit(depth, level, child_path, deeper, link.quantity, child)
                    if deeper:
                        part_lines(child.id, "l3", depth + 1, child_path)

            roots_by_model: dict[int, list[ModelBOM]] = {}
            for root in db.scalars(select(ModelBOM).order_by(ModelBOM.id)).all():
                roots_by_model.setdefault(root.model_id, []).append(root)

            def model_lines(model: ProductModel, depth: int, path: dict) -> None:
                for root in roots_by_model.get(model.id, []):
                    item = items.get(root.material_id)
                    if not item:
                        continue
                    line_path = {**path, "l1": (item.code, item.name)}
                    deeper = bool(children.get(item.id))
                    emit(depth, "l1", line_path, deeper, root.quantity, item)
                    if deeper:
                        part_lines(item.id, "l2", depth + 1, line_path)

            if tab == "Product Master":
                models = {model.id: model for model in db.scalars(select(ProductModel)).all()}
                links: dict[int, list[ProductModel]] = {}
                for link in db.scalars(select(ProductModelProduct)).all():
                    if link.model_id in models:
                        links.setdefault(link.product_id, []).append(models[link.model_id])
                products = db.scalars(select(Product).where(Product.code != MODEL_CATALOG_PRODUCT_CODE).order_by(Product.code)).all()
                for product in products:
                    product_models = sorted(links.get(product.id, []), key=lambda model: model.code)
                    if not product_models:
                        emit(0, "product", {"product": (product.code, product.name)}, False)
                    for model in product_models:
                        path = {"product": (product.code, product.name), "model": (model.code, model.name)}
                        emit(0, "model", path, bool(roots_by_model.get(model.id)))
                        model_lines(model, 1, path)
            elif tab == "Model BOM Level 1":
                statement = select(ProductModel).order_by(ProductModel.code)
                if model_id:
                    statement = statement.where(ProductModel.id.in_(model_id))
                for model in db.scalars(statement).all():
                    model_lines(model, 0, {"model": (model.code, model.name)})
            else:
                parent_level = "l1" if tab == "Level 2 Part BOM" else "l2"
                line_level = "l2" if tab == "Level 2 Part BOM" else "l3"
                parent_ids = {item_id for item_id in bom_parent_material_ids(db, tab) if children.get(item_id)}
                if parent_item_id:
                    parent_ids &= set(parent_item_id)
                for parent in sorted((items[item_id] for item_id in parent_ids if item_id in items), key=lambda item: item.code):
                    path = {parent_level: (parent.code, parent.name)}
                    emit(0, parent_level, path, True)
                    part_lines(parent.id, line_level, 1, path)

        headers = ["Level", *[label for level in levels for label in BOM_TREE_LABELS[level]], "Quantity", "Unit"]
        if can_price:
            headers += ["Part price", "Currency"]
        rows = []
        for entry in tree:
            item = entry["item"]
            values = [BOM_TREE_LEVEL_NAMES[entry["level"]]]
            for level in levels:
                values.extend(entry["path"].get(level, ("", "")))
            values += [entry["quantity"] if entry["quantity"] is not None else "", item.uom if item else ""]
            if can_price:
                priced = prices.get(item.id) if item else None
                values += [priced["unit_price"] if priced else "", priced["currency"] if priced else ""]
            rows.append({"depth": entry["depth"], "has_children": entry["has_children"], "values": values})
        title = f"{tab if tab.endswith('BOM') else tab + ' BOM'} tree"
        if file_format == "xlsx":
            output = _bom_tree_workbook(title, headers, rows)
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            output = _bom_tree_pdf(title, headers, rows)
            media_type = "application/pdf"
        filename = f"{title.lower().replace(' ', '-')}.{file_format}"
        return StreamingResponse(output, media_type=media_type, headers={"Content-Disposition": f'attachment; filename="{filename}"'})

    @app.get("/api/planning/tables/{kind}/{file_format}")
    def export_table(
        kind: str, file_format: str, request: Request,
        workspace: str = "trial", period: str | None = None,
        start_month: str | None = None, end_month: str | None = None,
        model_id: int | None = None, parent_item_id: int | None = None, tab: str = "Model BOM Level 1",
    ):
        current_user(request)
        if file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export format not supported")
        _validate_workspace(workspace)
        if kind in {"supplier-parts", "suppliers", "purchase-orders", "transfers"}:
            _check_supply_role(request)
        headers: list[str]
        rows: list[list[Any]]
        title = kind.replace("-", " ").title()
        if kind == "materials":
            with open_db() as db:
                materials = db.scalars(select(Item).order_by(Item.code)).all()
                can_price = _check_price_access(request)
                headers = ["Material code", "Material name", "Category", "Unit", "On hand", "Buffer days"]
                if can_price:
                    headers += ["Unit cost", "Currency"]
                rows = [[item.code, item.name, item.category, item.uom, item.stock_on_hand, getattr(item, "buffer_days", 0), *([item.unit_cost, item.currency] if can_price else [])] for item in materials]
        elif kind == "products":
            with open_db() as db:
                rows = [[product.code, product.name] for product in db.scalars(select(Product).where(Product.code != MODEL_CATALOG_PRODUCT_CODE).order_by(Product.code)).all()]
            headers = ["Product code", "Product name"]
        elif kind == "models":
            with open_db() as db:
                product_by_id = {product.id: product for product in db.scalars(select(Product)).all()}
                assigned: dict[int, list[str]] = {}
                for link in db.scalars(select(ProductModelProduct)).all():
                    product = product_by_id.get(link.product_id)
                    if product and product.code != MODEL_CATALOG_PRODUCT_CODE:
                        assigned.setdefault(link.model_id, []).append(f"{product.code} · {product.name}")
                rows = [[", ".join(assigned.get(model.id, [])), model.code, model.name] for model in db.scalars(select(ProductModel).order_by(ProductModel.code)).all()]
            headers = ["Assigned products", "Model code", "Model name"]
        elif kind == "bom":
            if parent_item_id is not None:
                page = get_part_bom(parent_item_id, request, tab=tab, offset=0, limit=1000)
                bom_rows = list(page["items"])
                while len(bom_rows) < page["total"]:
                    page = get_part_bom(parent_item_id, request, tab=tab, offset=len(bom_rows), limit=1000)
                    bom_rows.extend(page["items"])
                headers = ["Parent part code", "Parent part name", "Part code", "Part name", "Quantity per parent", "Unit"]
                rows = [[page["parent"]["code"], page["parent"]["name"], row["code"], row["name"], row["quantity"], row["uom"]] for row in bom_rows]
                title = f"{page['parent']['code']} {tab}"
            else:
                if model_id is None:
                    raise HTTPException(status_code=422, detail="Select a model or parent part before exporting its BOM")
                page = get_model_bom(model_id, request, tab=tab, offset=0, limit=1000)
                bom_rows = list(page["items"])
                while len(bom_rows) < page["total"]:
                    page = get_model_bom(model_id, request, tab=tab, offset=len(bom_rows), limit=1000)
                    bom_rows.extend(page["items"])
                headers = ["Model code", "Model name", "Level", "Part code", "Part name", "Quantity in model", "Unit"]
                rows = [[page["model_code"], page["model_name"], row["level"], row["code"], row["name"], row["quantity"], row["uom"]] for row in bom_rows]
                title = f"{page['model_code']} {tab}"
            if _check_price_access(request):
                headers += ["Unit price", "Currency"]
                rows = [
                    row + [bom_row.get("unit_price"), bom_row.get("currency")]
                    for row, bom_row in zip(rows, bom_rows)
                ]
        elif kind == "supplier-parts":
            part_rows = []
            offset = 0
            while True:
                page = list_supplier_parts(request, offset=offset, limit=1000)
                part_rows.extend(page["items"])
                offset += len(page["items"])
                if offset >= page["total"]:
                    break
            headers = ["Material code", "Material name", "Supplier code", "Supplier", "Tier", "Parent supplier", "SOB %", "Lead time days", "MOQ", "Buffer days", "Price mode", "Price status"]
            if _check_price_access(request):
                headers += ["Unit price", "Currency"]
            rows = [[row["material_code"], row["material_name"], row["supplier_code"], row["supplier_name"], row["tier"], row.get("parent_supplier_name") or "", row["sob_percent"], row["lead_time_days"], row["moq"], row["buffer_days"], row["price_mode"], row["price_status"], *([row.get("unit_price"), row.get("currency")] if _check_price_access(request) else [])] for row in part_rows]
        elif kind == "suppliers":
            with open_db() as db:
                suppliers = db.scalars(select(Supplier).order_by(Supplier.code)).all()
                rows = [[supplier.code, supplier.name, supplier.category, supplier.status, supplier.country, supplier.contact_name, supplier.email, supplier.phone] for supplier in suppliers]
            headers = ["Supplier code", "Supplier", "Category", "Approval status", "Country", "Contact", "Email", "Phone"]
        elif kind == "plans":
            plans = list_plans(request, workspace=workspace)
            headers = ["Issue month", "Revision", "Rows", "Uploaded at", "Workspace"]
            rows = [[row["issue_month"], row["revision"], row["row_count"], row["created_at"], row["workspace"]] for row in plans]
        elif kind == "actuals":
            if not period:
                raise HTTPException(status_code=422, detail="Choose a production month before exporting actuals")
            _month(period)
            actual_rows = list_actuals(request, period=period, workspace=workspace)
            headers = ["Model code", "Model name", "Actual production", "OEM stock"]
            rows = [[row["model_code"], row["model_name"], row["actual_quantity"], row["oem_stock"]] for row in actual_rows]
            title = f"Actual production and OEM stock · {period}"
        elif kind == "purchase-orders":
            if not start_month or not end_month:
                raise HTTPException(status_code=422, detail="Choose a start and end month before exporting purchase orders")
            orders = planning_pos(request, start_month=start_month, end_month=end_month, workspace=workspace)
            headers = ["PO", "Material code", "Material name", "Supplier", "Tier", "Quantity ordered", "Received", "Expected month", "Status", "NCNR"]
            if _check_price_access(request):
                headers += ["Unit price", "Currency", "Total"]
            rows = [[row["po_number"], row["material_code"], row["material_name"], row["supplier_name"], row["tier"], row["quantity_ordered"], row["quantity_received"], row["expected_month"], row["status"], row["ncnr"], *([row.get("unit_price"), row.get("currency"), row.get("total")] if _check_price_access(request) else [])] for row in orders]
        elif kind == "transfers":
            transfers = list_transfers(request, workspace=workspace)
            headers = ["Material code", "Material name", "From supplier", "To supplier", "Quantity", "Transfer month", "Status", "Note"]
            rows = [[row["material_code"], row["material_name"], row["from_supplier"], row["to_supplier"], row["quantity"], row["transfer_month"], row["status"], row["note"]] for row in transfers]
        elif kind == "users":
            admin_only(request)
            with open_db() as db:
                users = db.scalars(select(AppUser).order_by(AppUser.username)).all()
                rows = [[user.username, user.full_name, user.role, user.can_view_prices, user.enabled] for user in users]
            headers = ["Username", "Full name", "Role", "Can view prices", "Enabled"]
        else:
            raise HTTPException(status_code=404, detail="Table export not found")
        output = _generic_workbook(title, headers, rows) if file_format == "xlsx" else _generic_pdf(title, headers, rows)
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_format == "xlsx" else "application/pdf"
        return StreamingResponse(output, media_type=media_type, headers={"Content-Disposition": f'attachment; filename="dtc-{kind}.{file_format}"'})

    @app.delete("/api/planning/trial-data", status_code=204)
    def clear_trial_data(request: Request):
        admin_only(request)
        with open_db() as db:
            db.query(PlanLine).filter(
                PlanLine.cycle_id.in_(select(PlanningCycle.id).where(PlanningCycle.workspace == "trial"))
            ).delete(synchronize_session=False)
            for model in (PlanningPO, MaterialTransfer, SupplierStockSnapshot, OwnStockSnapshot, ActualProduction, PlanningCycle):
                db.query(model).filter(model.workspace == "trial").delete(synchronize_session=False)
            db.commit()
        return Response(status_code=204)

    @app.get("/api/planning/trial-data/counts")
    def trial_data_counts(request: Request):
        admin_only(request)
        with open_db() as db:
            return {
                "plans": db.scalar(select(func.count(PlanningCycle.id)).where(PlanningCycle.workspace == "trial")) or 0,
                "plan_rows": db.scalar(select(func.count(PlanLine.id)).join(PlanningCycle).where(PlanningCycle.workspace == "trial")) or 0,
                "actuals": db.scalar(select(func.count(ActualProduction.id)).where(ActualProduction.workspace == "trial")) or 0,
                "supplier_stock_rows": db.scalar(select(func.count(SupplierStockSnapshot.id)).where(SupplierStockSnapshot.workspace == "trial")) or 0,
                "purchase_orders": db.scalar(select(func.count(PlanningPO.id)).where(PlanningPO.workspace == "trial")) or 0,
                "transfers": db.scalar(select(func.count(MaterialTransfer.id)).where(MaterialTransfer.workspace == "trial")) or 0,
            }
