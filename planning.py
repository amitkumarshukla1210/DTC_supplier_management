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
from sqlalchemy import Boolean, Date, DateTime, Float, ForeignKey, Index, Integer, String, Text, UniqueConstraint, delete, func, or_, select, update as sa_update
from sqlalchemy.orm import Mapped, Session, mapped_column, sessionmaker
from starlette.responses import JSONResponse
import xlsxwriter

from app import Base, BOMComponent, Category, CATEGORY_TABLES, Item, MaterialRequest, SUPPLIER_TIERS, Supplier, _ensure_category, _tier_list, _tier_text


TOP_TIER = "T1"  # supplies you directly (formerly called Direct)


def _normalize_tier(value: Any) -> str:
    """'T1', 't2', 'Direct' (= T1) ... -> 'T1', 'T2' ...; anything else is refused."""
    text = str(value or "").strip().upper().replace(" ", "").replace("-", "")
    if text in {"", "DIRECT"}:
        return TOP_TIER
    if re.fullmatch(r"TIER[1-9]", text):  # "Tier 2" -> T2
        text = "T" + text[4:]
    if not re.fullmatch(r"T[1-9]", text):
        raise HTTPException(status_code=422, detail=f"Tier must be T1 (supplies you), T2, T3 …; got {value!r}")
    return text


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
    "SOB change from", "Demand change from SOB change",
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
    """Who supplies a part: tier T1 = supplies you (a T1 Material Master code you buy); T2 = supplies
    a T1 supplier (parent), T3 = supplies a T2 supplier ... A T2/T3 row may use a code from the T2 / T3
    Material Master (tier_material_id) instead of a T1 code (material_id): exactly one is set."""
    __tablename__ = "supplier_parts"
    __table_args__ = (UniqueConstraint("material_id", "tier_material_id", "supplier_id", "tier", "parent_supplier_id", "parent_link_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    material_id: Mapped[int | None] = mapped_column(ForeignKey("items.id", ondelete="CASCADE"), nullable=True, index=True)
    tier_material_id: Mapped[int | None] = mapped_column(ForeignKey("tier_materials.id"), nullable=True, index=True)
    supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
    tier: Mapped[str] = mapped_column(String(24), default="T1", index=True)
    parent_supplier_id: Mapped[int | None] = mapped_column(ForeignKey("suppliers.id"), nullable=True, index=True)
    # T2/T3: the parent relationship = the parent supplier's part this material goes into
    # (e.g. JGI steel T2 -> Supangita's Base Front Beam T1). Empty on rows saved before this.
    parent_link_id: Mapped[int | None] = mapped_column(ForeignKey("supplier_parts.id"), nullable=True, index=True)
    sob_percent: Mapped[float] = mapped_column(Float, default=100)
    lead_time_days: Mapped[int] = mapped_column(Integer, default=0)
    moq: Mapped[float] = mapped_column(Float, default=0)
    buffer_days: Mapped[int] = mapped_column(Integer, default=0)
    unit_price: Mapped[float] = mapped_column(Float, default=0)
    currency: Mapped[str] = mapped_column(String(3), default="USD")
    price_mode: Mapped[str] = mapped_column(String(24), default="spot")
    price_interval_months: Mapped[int] = mapped_column(Integer, default=1)
    price_updated_month: Mapped[str] = mapped_column(String(7), default=lambda: _month_text(date.today()))


class AppSetting(Base):
    """One-time data updates already done (key = name of the update)."""
    __tablename__ = "app_settings"

    key: Mapped[str] = mapped_column(String(64), primary_key=True)
    value: Mapped[str] = mapped_column(String(200), default="")


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
    release_date: Mapped[str | None] = mapped_column(String(10), nullable=True)  # YYYY-MM-DD from the plan sheet
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
    actual_quantity: Mapped[float] = mapped_column(Float, default=0)  # total produced in the month till the as-on date
    oem_stock: Mapped[float] = mapped_column(Float, default=0)
    as_of: Mapped[str | None] = mapped_column(String(10), nullable=True)  # YYYY-MM-DD of the latest entry
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


class SupplierSourceStock(Base):
    """Stock at a T2 / T3 / T4 supplier for its place in a T1 part's BOM (Upload supplier stock tree): the
    T1 supplier and part, the material line's code and the supplier of that line."""
    __tablename__ = "supplier_source_stock"
    __table_args__ = (UniqueConstraint("workspace", "report_month", "t1_supplier_id", "product_material_id", "material_code", "sub_supplier_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    workspace: Mapped[str] = mapped_column(String(16), index=True)
    report_month: Mapped[str] = mapped_column(String(7), index=True)
    t1_supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
    product_material_id: Mapped[int] = mapped_column(ForeignKey("items.id"), index=True)
    material_code: Mapped[str] = mapped_column(String(48))
    sub_supplier_id: Mapped[int] = mapped_column(ForeignKey("suppliers.id"), index=True)
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
    # the T2 supplier's own terms for this material (Supplier/material coverage page)
    lead_time_days: Mapped[int] = mapped_column(Integer, default=0)
    moq: Mapped[float] = mapped_column(Float, default=0)
    buffer_days: Mapped[int] = mapped_column(Integer, default=0)
    price_mode: Mapped[str] = mapped_column(String(24), default="spot")
    price_interval_months: Mapped[int] = mapped_column(Integer, default=1)
    price_updated_month: Mapped[str] = mapped_column(String(7), default=lambda: _month_text(date.today()))


class AlternateCode(Base):
    """Your own alternate code for a part code (e.g. FL101STEEL0.8 for FL001STEEL0.8). The SOB
    is set on each BOM line that uses the main code (BOMLineAlternate), 0% until set there;
    sob_percent here is no longer used."""
    __tablename__ = "alternate_codes"
    __table_args__ = (UniqueConstraint("main_item_id", "alternate_item_id"),)

    id: Mapped[int] = mapped_column(primary_key=True)
    main_item_id: Mapped[int] = mapped_column(ForeignKey("items.id", ondelete="CASCADE"), index=True)
    alternate_item_id: Mapped[int] = mapped_column(ForeignKey("items.id", ondelete="CASCADE"), unique=True)
    sob_percent: Mapped[float] = mapped_column(Float)


class BOMLineAlternate(Base):
    """Alternate code SOB on one BOM line: a Model BOM Level 1 line (model_bom_id) or a
    Level 2 / Level 3 part BOM line (component_id). The main code keeps 100 - sum of these.
    main_item_id is the line's code when the SOB was saved; it applies only while it still matches.
    from_month: the SOB applies from this month ("YYYY-MM") until the next change; "" = from the
    start. Every change is kept, so older months are planned on the SOB valid then (history)."""
    __tablename__ = "bom_line_alternates"

    id: Mapped[int] = mapped_column(primary_key=True)
    model_bom_id: Mapped[int | None] = mapped_column(ForeignKey("model_bom.id", ondelete="CASCADE"), nullable=True, index=True)
    component_id: Mapped[int | None] = mapped_column(ForeignKey("bom_components.id", ondelete="CASCADE"), nullable=True, index=True)
    main_item_id: Mapped[int] = mapped_column(ForeignKey("items.id"))
    alternate_item_id: Mapped[int] = mapped_column(ForeignKey("items.id"))
    sob_percent: Mapped[float] = mapped_column(Float)
    from_month: Mapped[str] = mapped_column(String(7), default="")


class SOBChangeLog(Base):
    """Every alternate SOB change (set or removed), one row per BOM line: who, when, which
    code, where it is used, from which month and the SOB split."""
    __tablename__ = "sob_change_log"

    id: Mapped[int] = mapped_column(primary_key=True)
    changed_at: Mapped[datetime] = mapped_column(DateTime, index=True)
    changed_by: Mapped[str] = mapped_column(String(120))
    action: Mapped[str] = mapped_column(String(12))  # "Set" or "Removed"
    main_item_id: Mapped[int] = mapped_column(Integer, index=True)
    main_code: Mapped[str] = mapped_column(String(48))
    level: Mapped[str] = mapped_column(String(4))
    used_in: Mapped[str] = mapped_column(String(300))
    from_month: Mapped[str] = mapped_column(String(7), default="")
    sob_text: Mapped[str] = mapped_column(String(600), default="")


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


class AlternateCodeInput(BaseModel):
    main_item_id: int
    alternate_item_id: int


class LineAlternateSOB(BaseModel):
    alternate_item_id: int
    sob_percent: float = Field(ge=0, le=100)


class SOBLine(BaseModel):
    model_bom_id: int | None = None
    component_id: int | None = None


class AlternateSOBInput(BaseModel):
    """SOB for a main code's alternates from a month, on the ticked uses (BOM lines)."""
    main_item_id: int
    from_month: str = ""
    alternates: list[LineAlternateSOB] = []
    lines: list[SOBLine]
    action: str = "set"  # "set" or "remove" (remove the change of from_month)


class LineAlternatesInput(BaseModel):
    model_bom_id: int | None = None
    component_id: int | None = None
    alternates: list[LineAlternateSOB] = []
    from_month: str = ""  # "" = from the start (all months)
    apply_everywhere: bool = False  # same change on every BOM line using this main code


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
    lead_time_days: int = Field(default=0, ge=0, le=5000)
    moq: float = Field(default=0, ge=0, allow_inf_nan=False)
    buffer_days: int = Field(default=0, ge=0, le=5000)
    price_mode: str = Field(default="spot", pattern=r"^(spot|fixed_monthly|every_n_months)$")
    price_interval_months: int = Field(default=1, ge=1, le=120)


class CoverageLineInput(BaseModel):
    """A material line of the part BOM on the coverage page: any T1 / T2 / T3 Material Master code
    or a code of the supplier's own material master; it joins the supplier's material master."""
    link_id: int
    parent_line_id: int | None = None  # an L2 line goes under an L1 line (e.g. iron ore under steel)
    item_id: int | None = None
    tier_material_id: int | None = None
    supplier_material_id: int | None = None
    quantity: float = Field(gt=0, le=1_000_000_000, allow_inf_nan=False)
    sources: list[SupplierBOMSourceInput] = []  # its suppliers, typed in the same window (SOB up to 100% in total)


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
    material_id: int | None = None  # T1 Material Master code
    tier_material_id: int | None = None  # T2 / T3 Material Master code (T2/T3 rows only)
    supplier_id: int
    tier: str = Field(default="T1", min_length=2, max_length=24)
    parent_supplier_id: int | None = None
    parent_link_id: int | None = None  # T2/T3: the parent supplier's relationship (its part)
    sob_percent: float = Field(default=100, ge=0, le=100)
    lead_time_days: int = Field(default=0, ge=0, le=5000)
    moq: float = Field(default=0, ge=0)
    buffer_days: int = Field(default=0, ge=0, le=5000)
    unit_price: float = Field(default=0, ge=0)
    currency: str = Field(default="USD", pattern=r"^[A-Z]{3}$")
    price_mode: str = Field(default="spot", pattern=r"^(spot|fixed_monthly|every_n_months)$")
    price_interval_months: int = Field(default=1, ge=1, le=120)


class SupplierPartRows(BaseModel):
    rows: list[SupplierPartInput]


class SupplierStockRow(BaseModel):
    supplier_part_id: int
    quantity: float = Field(ge=0)
    available_quantity: float | None = Field(default=None, ge=0)
    arrival_month: str | None = None


class SupplierStockRows(BaseModel):
    rows: list[SupplierStockRow]


class PlanUpload(BaseModel):
    issue_month: str
    workspace: str = "trial"
    revision: int | None = Field(default=None, ge=0, le=999)  # from the sheet's Plan revision (Rev.01); none = next
    release_date: str | None = None  # YYYY-MM-DD
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
    as_of: str | None = None  # YYYY-MM-DD: actual = total produced in the month till this date
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
    "material_code": "material_code", "part_code": "material_code", "item_code": "material_code", "t1_part_code": "material_code",
    "supplier_name": "supplier_name", "supplier": "supplier_name", "vendor": "supplier_name", "vendor_name": "supplier_name",
    "supplier_code": "supplier_code", "vendor_code": "supplier_code",
    "tier": "tier", "parent_supplier_name": "parent_supplier_name", "parent_supplier": "parent_supplier_name",
    "sob_percent": "sob_percent", "sob": "sob_percent", "sob_%": "sob_percent", "sob_(%)": "sob_percent",
    "unit_price": "unit_price", "unit_cost": "unit_price", "price": "unit_price",
    "currency": "currency",
    "lead_time_days": "lead_time_days", "lead_time": "lead_time_days", "lead_time_(days)": "lead_time_days",
    "moq": "moq", "buffer_days": "buffer_days", "buffer": "buffer_days",
    "price_mode": "price_mode", "price_interval_months": "price_interval_months",
}
# the relationship page's columns, in its order (template and download)
RELATIONSHIP_SHEET_COLUMNS = ["Supplier Code", "Supplier Name", "T1 Part Code", "T1 Part Name", "SOB", "Lead time", "MOQ", "Buffer days", "Unit price", "Currency"]
# supplier-wise coverage workbook: every T1 part of a supplier, its BOM lines (L1-L3) and the suppliers of each line (T2-T4)
COVERAGE_SHEET_COLUMNS = ["T1 Part Code", "T1 Part Name", "Level", "Material code", "Material name", "BOM qty", "Unit", "Supplier Code", "Supplier Name",
                          "SOB %", "Lead time days", "MOQ", "Buffer days", "Price mode", "Price interval months", "Unit price", "Currency"]
COVERAGE_COLUMN_ALIASES = {
    "t1_part_code": "part_code", "part_code": "part_code", "level": "level", "material_code": "material_code", "code": "material_code",
    "bom_qty": "quantity", "quantity": "quantity", "qty": "quantity", "supplier_code": "supplier_code", "supplier_name": "supplier_name",
    "sob_%": "sob_percent", "sob": "sob_percent", "sob_percent": "sob_percent", "lead_time_days": "lead_time_days", "lead_time": "lead_time_days",
    "moq": "moq", "buffer_days": "buffer_days", "price_mode": "price_mode", "price_interval_months": "price_interval_months",
    "unit_price": "unit_price", "price": "unit_price", "currency": "currency",
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

SUPPLIER_COLUMN_ALIASES = {
    "supplier_code": "code", "code": "code", "vendor_code": "code", "supplier_id": "code", "vendor_id": "code",
    "supplier_name": "name", "supplier": "name", "name": "name", "vendor_name": "name", "vendor": "name",
    "category": "category", "supplier_category": "category", "commodity": "category",
    "country": "country", "contact": "contact_name", "contact_person": "contact_name", "contact_name": "contact_name",
    "email": "email", "e-mail": "email", "email_id": "email", "mail": "email",
    "phone": "phone", "mobile": "phone", "phone_number": "phone", "contact_number": "phone", "mobile_number": "phone",
    "approval": "status", "status": "status", "approval_status": "status",
    "tier": "tiers", "tiers": "tiers", "supplier_tier": "tiers", "supplier_tiers": "tiers",
}

ALTERNATE_CODE_COLUMN_ALIASES = {
    "main_code": "main_code", "main_part_code": "main_code", "part_code": "main_code", "material_code": "main_code",
    "main_material_code": "main_code", "original_code": "main_code", "code": "main_code",
    "alternate_code": "alternate_code", "alternate_part_code": "alternate_code", "alt_code": "alternate_code",
    "alternate": "alternate_code", "alternative_code": "alternate_code", "new_code": "alternate_code",
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


def _sheet_date(value: Any) -> date | None:
    """A date cell or text such as 01-01-2026 (DD-MM-YYYY), 2026-01-01 or 01/01/2026."""
    if value is None or str(value).strip() == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for date_format in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(str(value).strip(), date_format).date()
        except ValueError:
            continue
    raise HTTPException(status_code=422, detail=f"{value} is not a date; use DD-MM-YYYY")


def _is_month(value: Any) -> bool:
    try:
        _month(value)
        return True
    except HTTPException:
        return False


def _plan_blocks(raw: bytes) -> list[dict[str, Any]] | None:
    """Plan sheet as in the template: blocks of Model Code | Model Name | Release date | Plan revision | one
    column per month (01-01-2026 ...). A sheet may hold several blocks (one per revision). None: not this layout."""
    try:
        workbook = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        lines = list(workbook.active.iter_rows(values_only=True))
        workbook.close()
    except Exception as exc:
        raise HTTPException(status_code=422, detail="Could not read the Excel workbook; use a valid .xlsx file") from exc

    def key(value: Any) -> str:
        return re.sub(r"\s+", "_", str(value or "").strip().lower())

    blocks: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for number, line in enumerate(lines, start=1):
        keys = [key(value) for value in line]
        month_like = sum(1 for value in line if value is not None and str(value).strip() and _is_month(value))
        if "model_code" in keys and (month_like or "release_date" in keys or "plan_revision" in keys):
            columns, months = {}, []
            for index, value in enumerate(line):
                if keys[index] in ("model_code", "model_name", "release_date", "plan_revision"):
                    columns[keys[index]] = index
                elif value is not None and str(value).strip():
                    try:
                        months.append((index, _month_text(_month(value))))
                    except HTTPException:
                        continue
            if not months:
                raise HTTPException(status_code=422, detail=f"Row {number}: the heading row has no month columns (01-01-2026 ...)")
            if len({month for _, month in months}) != len(months):
                raise HTTPException(status_code=422, detail=f"Row {number}: the same month is twice in the heading row")
            current = {"row": number, "columns": columns, "months": months, "rows": []}
            blocks.append(current)
            continue
        if current is None:
            continue
        if all(value is None or str(value).strip() == "" for value in line):
            current = None  # a blank row ends the block
            continue
        cell = lambda name: line[current["columns"][name]] if name in current["columns"] and current["columns"][name] < len(line) else None
        code = str(cell("model_code") or "").strip()
        if not code:
            continue
        try:
            released = _sheet_date(cell("release_date"))
        except HTTPException as exc:
            raise HTTPException(status_code=422, detail=f"Row {number}: Release date {exc.detail}") from exc
        revision_text = str(cell("plan_revision") or "").strip()
        found = re.search(r"\d+", revision_text)
        quantities = {}
        for index, month in current["months"]:
            value = line[index] if index < len(line) else None
            if value is None or str(value).strip() == "":
                continue
            try:
                quantities[month] = float(value)
            except (TypeError, ValueError) as exc:
                raise HTTPException(status_code=422, detail=f"Row {number}: {value} is not a quantity") from exc
        current["rows"].append({"row": number, "code": code, "name": str(cell("model_name") or "").strip(), "released": released,
                                "revision": int(found.group()) if found else None, "quantities": quantities})
    blocks = [block for block in blocks if block["rows"]]
    if not blocks:
        return None
    for block in blocks:
        dates = {row["released"] for row in block["rows"] if row["released"]}
        revisions = {row["revision"] for row in block["rows"] if row["revision"] is not None}
        if len(dates) > 1 or len(revisions) > 1:
            raise HTTPException(status_code=422, detail=f"The block starting on row {block['row']} has more than one release date or revision; one block = one revision")
        block["released"] = next(iter(dates), None)
        block["revision"] = next(iter(revisions), None)
    return blocks


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


def _item_map(db: Session) -> tuple[dict[int, Item], dict[int, list[tuple[int, float, int]]]]:
    items = {item.id: item for item in db.scalars(select(Item)).all()}
    children: dict[int, list[tuple[int, float, int]]] = {}
    for row in db.execute(select(BOMComponent.parent_item_id, BOMComponent.child_item_id, BOMComponent.quantity, BOMComponent.id)):
        children.setdefault(row.parent_item_id, []).append((row.child_item_id, row.quantity, row.id))
    return items, children


def _sob_at(entries: Iterable[tuple[str, int, float]], month: str) -> dict[int, float]:
    """Alternate SOB valid in a month from a line's history [(from_month, alternate id, SOB)]:
    for each alternate the latest change on or before the month ("" = from the start)."""
    latest: dict[int, tuple[str, float]] = {}
    for from_month, alternate_id, sob in entries:
        if from_month <= month and (alternate_id not in latest or from_month >= latest[alternate_id][0]):
            latest[alternate_id] = (from_month, sob)
    return {alternate_id: sob for alternate_id, (_, sob) in latest.items()}


def _line_alternate_history(db: Session) -> tuple[dict[int, list[tuple[str, int, float]]], dict[int, list[tuple[str, int, float]]], list[str]]:
    """Alternate SOB history per BOM line: (by Model BOM Level 1 line id, by part BOM line id)
    -> [(from_month, alternate item id, SOB %)], and the sorted months in which an SOB changes.
    Only current main / alternate pairs, and only while the line still uses the main code."""
    pairs = set(db.execute(select(AlternateCode.main_item_id, AlternateCode.alternate_item_id)).all())
    root_items = dict(db.execute(select(ModelBOM.id, ModelBOM.material_id)).all())
    component_items = dict(db.execute(select(BOMComponent.id, BOMComponent.child_item_id)).all())
    by_root: dict[int, list[tuple[str, int, float]]] = {}
    by_component: dict[int, list[tuple[str, int, float]]] = {}
    months: set[str] = set()
    for entry in db.scalars(select(BOMLineAlternate)).all():
        if (entry.main_item_id, entry.alternate_item_id) not in pairs:
            continue
        record = (entry.from_month or "", entry.alternate_item_id, entry.sob_percent)
        if entry.model_bom_id is not None and root_items.get(entry.model_bom_id) == entry.main_item_id:
            by_root.setdefault(entry.model_bom_id, []).append(record)
        elif entry.component_id is not None and component_items.get(entry.component_id) == entry.main_item_id:
            by_component.setdefault(entry.component_id, []).append(record)
        else:
            continue
        if record[0]:
            months.add(record[0])
    return by_root, by_component, sorted(months)


def _splits_at(history: dict[int, list[tuple[str, int, float]]], month: str) -> dict[int, list[tuple[int, float]]]:
    result: dict[int, list[tuple[int, float]]] = {}
    for line_id, entries in history.items():
        used = [(alternate_id, sob) for alternate_id, sob in _sob_at(entries, month).items() if sob > 0]
        if used:
            result[line_id] = used
    return result


def _rollup_model_bom(
    roots: list[ModelBOM], children: dict[int, list[tuple[int, float, int]]],
    root_splits: dict[int, list[tuple[int, float]]] | None = None,
    component_splits: dict[int, list[tuple[int, float]]] | None = None,
) -> dict[int, float]:
    """Material quantity per model. On a BOM line with alternate SOB, each alternate takes its
    SOB % of that line's quantity and the main code keeps the rest."""
    result: dict[int, float] = {}
    root_splits = root_splits or {}
    component_splits = component_splits or {}

    def walk(item_id: int, qty: float, ancestors: frozenset[int], split: list[tuple[int, float]] = ()) -> None:
        if item_id in ancestors:
            raise HTTPException(status_code=422, detail="Circular BOM detected")
        for alternate_id, sob in split:
            result[alternate_id] = result.get(alternate_id, 0) + qty * sob / 100
        qty = qty * (100 - sum(sob for _, sob in split)) / 100
        result[item_id] = result.get(item_id, 0) + qty
        for child_id, child_qty, component_id in children.get(item_id, ()):
            walk(child_id, qty * child_qty, ancestors | {item_id}, component_splits.get(component_id, ()))

    for root in roots:
        walk(root.material_id, root.quantity, frozenset(), root_splits.get(root.id, ()))
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
    # Alternate codes take their SOB share of each BOM line's demand. The SOB can change from a
    # month (history), so the usage per model is worked out for each SOB period.
    root_history, component_history, change_months = _line_alternate_history(db)
    periods = [""] + change_months  # period key = month its SOB starts ("" = from the start)
    usage_by_period = {
        key: {
            model_id: _rollup_model_bom(roots_by_model.get(model_id, []), children,
                                        _splits_at(root_history, key or "0000-00"), _splits_at(component_history, key or "0000-00"))
            for model_id in models
        }
        for key in periods
    }

    def period_of(month_text: str) -> str:
        return max((key for key in periods if key <= month_text), default="")

    # SOB changes that moved a material's usage: material -> [change months]
    material_changes: dict[int, list[str]] = {}
    for previous, key in zip(periods, periods[1:]):
        for model_id in models:
            before, after = usage_by_period[previous][model_id], usage_by_period[key][model_id]
            for material_id in set(before) | set(after):
                if abs(before.get(material_id, 0) - after.get(material_id, 0)) > 1e-12 and key not in material_changes.get(material_id, []):
                    material_changes.setdefault(material_id, []).append(key)
    material_usage_by_model: dict[int, dict[int, float]] = {}
    for usage in usage_by_period.values():
        for model_id, rolled in usage.items():
            merged = material_usage_by_model.setdefault(model_id, {})
            for material_id in rolled:
                merged.setdefault(material_id, 0)
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
        month_usage = usage_by_period[period_of(month)]
        # demand change caused by the latest SOB change (on or before this month) that moved each material
        sob_change_by_material: dict[int, tuple[str, float]] = {}
        for material_id, changes in material_changes.items():
            done = [key for key in changes if key <= month]
            if done:
                previous = periods[periods.index(done[-1]) - 1]
                delta = sum(
                    (plan_by_cycle_month.get((cycle.id, month), {}).get(model_id, 0) if cycle else 0)
                    * (month_usage[model_id].get(material_id, 0) - usage_by_period[previous][model_id].get(material_id, 0))
                    for model_id in models
                )
                sob_change_by_material[material_id] = (done[-1], delta)
        for model_id, model in models.items():
            planned_qty = plan_by_cycle_month.get((cycle.id, month), {}).get(model_id, 0) if cycle else 0
            actual = actual_by_key.get((month, model_id))
            rolled = month_usage.get(model_id, {})
            for material_id, unit_qty in rolled.items():
                plan_demand_by_material[material_id] = plan_demand_by_material.get(material_id, 0) + planned_qty * unit_qty
                if actual:
                    actual_demand_by_material[material_id] = actual_demand_by_material.get(material_id, 0) + actual.actual_quantity * unit_qty
                    actual_models_by_material.setdefault(material_id, set()).add(model_id)
        for material_id in set(plan_demand_by_material) | set(actual_demand_by_material):
            plan_demand = plan_demand_by_material.get(material_id, 0)
            sob_change_month, sob_change_delta = sob_change_by_material.get(material_id, ("", 0.0))
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
            direct_links = groups.get((TOP_TIER, None), [])
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
                        if tier == TOP_TIER and parent_id is None else 0
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
                            "sob_change_from": sob_change_month,
                            "sob_demand_change": sob_change_delta * link.sob_percent / 100,
                        })
                    projected[link.id] = max(0, supplier_balance)
                if shares < 100 and material:
                    unassigned_share = 100 - shares
                    unassigned = material_demand * unassigned_share / 100
                    own_stock_allocated = (
                        own_stock_opening.get(material_id, 0) * unassigned_share / 100
                        if (tier == TOP_TIER and parent_id is None) or not links else 0
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
                        "sob_change_from": sob_change_month,
                        "sob_demand_change": sob_change_delta * unassigned_share / 100,
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
        row.get("sob_change_from") or "", row.get("sob_demand_change") or 0,
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


def add_supplier_tier(supplier: Supplier | None, tier: str) -> None:
    """A supplier used at a tier (T1-T3) shows in that Supplier Master tab too."""
    if supplier is not None and tier in SUPPLIER_TIERS and tier not in _tier_list(supplier.tiers):
        supplier.tiers = ",".join(_tier_list(_tier_list(supplier.tiers) + [tier]))


def tiers_from_links(db: Session, link: "SupplierPart") -> None:
    add_supplier_tier(db.get(Supplier, link.supplier_id), link.tier)
    if link.parent_supplier_id and re.fullmatch(r"T[2-9]", link.tier or ""):
        add_supplier_tier(db.get(Supplier, link.parent_supplier_id), f"T{int(link.tier[1:]) - 1}")


def upgrade_supplier_parts(db: Session) -> None:
    """Start-up updates of supplier/material relationships (each runs once):
    1. the table gets tier_material_id and material_id may be empty (T2/T3 codes); SQLite cannot change a
       column, so the table is rebuilt with its rows (a copy of the database file is kept first);
    2. tiers renamed: the supplier that supplies you is T1 (was "Direct"); old T1 -> T2, T2 -> T3 ..."""
    from sqlalchemy import inspect as sa_inspect, text as sa_text
    bind = db.get_bind()
    columns = {column["name"]: column for column in sa_inspect(bind).get_columns("supplier_parts")}
    needs_rebuild = "tier_material_id" not in columns or "parent_link_id" not in columns or not columns["material_id"]["nullable"]
    done = {row.key for row in db.scalars(select(AppSetting)).all()}
    needs_rename = "tiers_t1_is_direct" not in done
    if (needs_rebuild or needs_rename) and db.scalar(select(func.count(SupplierPart.id)) if not needs_rebuild else sa_text("SELECT COUNT(*) FROM supplier_parts")):
        url = str(bind.url)
        if url.startswith("sqlite:///"):
            source = Path(url.removeprefix("sqlite:///"))
            backup = source.with_name(f"{source.name}.before-tier-update-{datetime.now():%Y%m%d-%H%M%S}.bak")
            if source.is_file():
                db.commit()
                backup.write_bytes(source.read_bytes())
    if needs_rebuild:
        old_columns = [name for name in columns if name in SupplierPart.__table__.columns]
        db.execute(sa_text("PRAGMA legacy_alter_table=ON"))
        db.execute(sa_text("ALTER TABLE supplier_parts RENAME TO supplier_parts_old"))
        for (index_name,) in db.execute(sa_text("SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='supplier_parts_old' AND sql IS NOT NULL")).all():
            db.execute(sa_text(f'DROP INDEX "{index_name}"'))
        db.commit()
        SupplierPart.__table__.create(bind)
        names = ", ".join(old_columns)
        db.execute(sa_text(f"INSERT INTO supplier_parts ({names}) SELECT {names} FROM supplier_parts_old"))
        db.execute(sa_text("DROP TABLE supplier_parts_old"))
        db.execute(sa_text("PRAGMA legacy_alter_table=OFF"))
        db.commit()
    if needs_rename:
        tiers = sorted({tier for tier in db.scalars(select(SupplierPart.tier).distinct()).all() if re.fullmatch(r"T[1-9]", tier or "")}, reverse=True)
        for tier in tiers:  # T2 -> T3 before T1 -> T2
            db.execute(sa_text("UPDATE supplier_parts SET tier = :new WHERE tier = :old"), {"new": f"T{int(tier[1:]) + 1}", "old": tier})
        db.execute(sa_text("UPDATE supplier_parts SET tier = 'T1' WHERE tier = 'Direct' OR tier = '' OR tier IS NULL"))
        db.add(AppSetting(key="tiers_t1_is_direct", value=datetime.now().isoformat(timespec="seconds")))
        db.commit()
    # suppliers' tiers (Supplier Master tabs), once: from their relationships; T1 when none
    if "supplier_tiers_set" not in done:
        for supplier in db.scalars(select(Supplier)).all():
            if not _tier_list(supplier.tiers):
                supplier.tiers = ""
        db.flush()
        for link in db.scalars(select(SupplierPart)).all():
            tiers_from_links(db, link)
        for supplier in db.scalars(select(Supplier)).all():
            if not _tier_list(supplier.tiers):
                supplier.tiers = "T1"
        db.add(AppSetting(key="supplier_tiers_set", value=datetime.now().isoformat(timespec="seconds")))
        db.commit()


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
        upgrade_supplier_parts(db)
        # supplier BOM sources (T2 suppliers of a material line) gained their own terms
        present = {column["name"] for column in sa_inspect(db.get_bind()).get_columns("supplier_bom_line_sources")}
        for column, definition in (("lead_time_days", "INTEGER NOT NULL DEFAULT 0"), ("moq", "FLOAT NOT NULL DEFAULT 0"), ("buffer_days", "INTEGER NOT NULL DEFAULT 0"),
                                   ("price_mode", "VARCHAR(24) NOT NULL DEFAULT 'spot'"), ("price_interval_months", "INTEGER NOT NULL DEFAULT 1"),
                                   ("price_updated_month", f"VARCHAR(7) NOT NULL DEFAULT '{_month_text(date.today())}'")):
            if column not in present:
                db.execute(sa_text(f"ALTER TABLE supplier_bom_line_sources ADD COLUMN {column} {definition}"))
        if "as_of" not in {column["name"] for column in sa_inspect(db.get_bind()).get_columns("actual_production")}:
            db.execute(sa_text("ALTER TABLE actual_production ADD COLUMN as_of VARCHAR(10)"))
        if "release_date" not in {column["name"] for column in sa_inspect(db.get_bind()).get_columns("planning_cycles")}:
            db.execute(sa_text("ALTER TABLE planning_cycles ADD COLUMN release_date VARCHAR(10)"))
        # bom_line_alternates gained from_month (SOB history); earlier SOB applies from the start
        if "from_month" not in {column["name"] for column in sa_inspect(db.get_bind()).get_columns("bom_line_alternates")}:
            db.execute(sa_text("ALTER TABLE bom_line_alternates ADD COLUMN from_month VARCHAR(7) NOT NULL DEFAULT ''"))
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
            db.execute(delete(BOMLineAlternate).where(BOMLineAlternate.model_bom_id.in_(select(ModelBOM.id).where(ModelBOM.model_id == model_id))))
            db.execute(delete(ModelBOM).where(ModelBOM.model_id == model_id))
            db.execute(delete(ProductModelProduct).where(ProductModelProduct.model_id == model_id))
            db.delete(model)
            db.commit()
            return Response(status_code=204)

    def bom_levels(db: Session) -> dict[int, list[str]]:
        """Where each T1 code is used: L1 = a line of Model BOM Level 1; L2 = a line of a Level 2
        Part BOM (under an L1 part); L3 = a line of a Level 3 Part BOM (under an L2 part).
        An alternate code with SOB above 0% on a line (in any month) is used at that line's level."""
        level1 = set(db.scalars(select(ModelBOM.material_id).distinct()).all())
        level2 = set(db.scalars(select(BOMComponent.child_item_id).where(BOMComponent.parent_item_id.in_(level1)).distinct()).all()) if level1 else set()
        level3 = set(db.scalars(select(BOMComponent.child_item_id).where(BOMComponent.parent_item_id.in_(level2)).distinct()).all()) if level2 else set()
        direct: dict[int, set[str]] = {}
        for name, ids in (("L1", level1), ("L2", level2), ("L3", level3)):
            for item_id in ids:
                direct.setdefault(item_id, set()).add(name)
        # alternates: the level of each BOM line where they take a share
        as_alternate: dict[int, set[str]] = {}
        root_history, component_history, _ = _line_alternate_history(db)
        for line_id, entries in root_history.items():
            for _, alternate_id, sob in entries:
                if sob > 0:
                    as_alternate.setdefault(alternate_id, set()).add("L1")
        if component_history:
            parents = dict(db.execute(select(BOMComponent.id, BOMComponent.parent_item_id).where(BOMComponent.id.in_(list(component_history)))).all())
            for line_id, entries in component_history.items():
                line_levels = {name for name, ids in (("L2", level1), ("L3", level2)) if parents.get(line_id) in ids}
                for _, alternate_id, sob in entries:
                    if sob > 0:
                        as_alternate.setdefault(alternate_id, set()).update(line_levels)
        levels: dict[int, list[str]] = {}
        for item_id in set(direct) | set(as_alternate):
            names = sorted(direct.get(item_id, set()) | as_alternate.get(item_id, set()))
            label = " & ".join(names)
            if as_alternate.get(item_id, set()) - direct.get(item_id, set()):
                label += " (alternate)"
            levels[item_id] = [label]
        return levels

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
            levels = bom_levels(db)
            page_ids = [item.id for item in rows]
            in_alternates = set(db.scalars(select(AlternateCode.main_item_id).where(AlternateCode.main_item_id.in_(page_ids))).all()) | \
                set(db.scalars(select(AlternateCode.alternate_item_id).where(AlternateCode.alternate_item_id.in_(page_ids))).all()) if page_ids else set()
            return {
                "total": total, "offset": offset, "limit": limit,
                "items": [{
                    "id": item.id, "code": item.code, "name": item.name, "category": item.category,
                    "used_in": " & ".join(levels.get(item.id, [])),  # "" = not used in any BOM
                    # Move to T2/T3 is offered only for codes in no BOM (L1/L2/L3) and not in Alternate codes
                    "movable": not levels.get(item.id) and item.id not in in_alternates,
                    "uom": item.uom, "stock_on_hand": item.stock_on_hand, "reorder_level": item.reorder_level,
                    "buffer_days": getattr(item, "buffer_days", 0),
                    **({"unit_cost": item.unit_cost, "currency": item.currency} if can_price else {}),
                } for item in rows],
            }

    @app.get("/api/planning/material-categories")
    def material_categories(request: Request):
        """The category list (names) for every category dropdown."""
        current_user(request)
        with open_db() as db:
            return {"items": sorted(db.scalars(select(Category.name)).all(), key=str.casefold)}

    def category_usage(db: Session) -> dict[str, dict[str, int]]:
        """Records per category (lower-case key): parts (T1), T2/T3 codes and suppliers."""
        from sqlalchemy import text as sa_text
        usage: dict[str, dict[str, int]] = {}
        labels = {"items": "parts", "suppliers": "suppliers", "tier_materials": "tier_codes"}
        for table in CATEGORY_TABLES:
            label = labels[table]
            for name, count in db.execute(sa_text(f"SELECT trim(category), COUNT(*) FROM {table} GROUP BY lower(trim(category))")).all():
                if name:
                    entry = usage.setdefault(name.lower(), {"name": name, "parts": 0, "tier_codes": 0, "suppliers": 0})
                    entry[label] += count
        return usage

    def recategorise(db: Session, old: str, new: str) -> int:
        from sqlalchemy import text as sa_text
        changed = 0
        for table in CATEGORY_TABLES:
            changed += db.execute(sa_text(f"UPDATE {table} SET category = :new WHERE lower(trim(category)) = :old"), {"new": new, "old": old.strip().lower()}).rowcount
        return changed

    @app.get("/api/planning/categories")
    def list_categories(request: Request):
        """The category list with how many parts, T2/T3 codes and suppliers use each, plus categories
        used by records but not in the list (to replace or add)."""
        current_user(request)
        with open_db() as db:
            usage = category_usage(db)
            items = []
            for category in sorted(db.scalars(select(Category)).all(), key=lambda c: c.name.casefold()):
                used = usage.pop(category.name.lower(), {})
                items.append({"id": category.id, "name": category.name, "parts": used.get("parts", 0),
                              "tier_codes": used.get("tier_codes", 0), "suppliers": used.get("suppliers", 0)})
            unlisted = sorted(usage.values(), key=lambda entry: entry["name"].casefold())
            return {"items": items, "unlisted": unlisted}

    @app.post("/api/planning/categories", status_code=201)
    def add_category(payload: dict, request: Request):
        current_user(request)
        _check_supply_role(request)
        name = " ".join(str(payload.get("name") or "").split())
        if len(name) < 2:
            raise HTTPException(status_code=422, detail="Category needs at least 2 characters")
        with open_db() as db:
            if db.scalar(select(Category.id).where(func.lower(Category.name) == name.lower())):
                raise HTTPException(status_code=409, detail=f"{name} is already in the category list")
            saved = _ensure_category(db, name)
            # records already using this name (any spelling) take the list's spelling
            recategorise(db, saved, saved)
            db.commit()
            return {"name": saved}

    @app.patch("/api/planning/categories/{category_id}")
    def rename_category(category_id: int, payload: dict, request: Request):
        """Rename a category; every part, T2/T3 code and supplier using it follows."""
        current_user(request)
        _check_supply_role(request)
        name = " ".join(str(payload.get("name") or "").split())
        if len(name) < 2 or len(name) > 80:
            raise HTTPException(status_code=422, detail="Category needs 2 to 80 characters")
        with open_db() as db:
            category = db.get(Category, category_id)
            if not category:
                raise HTTPException(status_code=404, detail="Category not found")
            other = db.scalar(select(Category).where(func.lower(Category.name) == name.lower(), Category.id != category_id))
            if other:
                raise HTTPException(status_code=409, detail=f"{other.name} is already in the list; use Replace to move {category.name} into it")
            changed = recategorise(db, category.name, name)
            category.name = name
            db.commit()
            return {"name": name, "records": changed}

    @app.delete("/api/planning/categories/{category_id}", status_code=204)
    def delete_category(category_id: int, request: Request):
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            category = db.get(Category, category_id)
            if not category:
                raise HTTPException(status_code=404, detail="Category not found")
            used = category_usage(db).get(category.name.lower())
            if used:
                parts = [f"{used[k]} {label}" for k, label in (("parts", "parts"), ("tier_codes", "T2/T3 codes"), ("suppliers", "suppliers")) if used[k]]
                raise HTTPException(status_code=409, detail=f"{category.name} is used by {', '.join(parts)}; use Replace to move them to another category first")
            db.delete(category)
            db.commit()
            return Response(status_code=204)

    @app.post("/api/planning/categories/replace")
    def replace_category(payload: dict, request: Request):
        """Move every record from one category (in the list or not) to a category in the list.
        An old category that was in the list is removed from it."""
        current_user(request)
        _check_supply_role(request)
        old = " ".join(str(payload.get("old") or "").split())
        new = " ".join(str(payload.get("new") or "").split())
        with open_db() as db:
            target = db.scalar(select(Category).where(func.lower(Category.name) == new.lower()))
            if not old or not target:
                raise HTTPException(status_code=422, detail="Choose the old category and a category from your list")
            if old.lower() == target.name.lower():
                raise HTTPException(status_code=422, detail="Choose a different category to replace with")
            changed = recategorise(db, old, target.name)
            listed = db.scalar(select(Category).where(func.lower(Category.name) == old.lower()))
            if listed:
                db.delete(listed)
            db.commit()
            return {"old": old, "new": target.name, "records": changed}

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
                code=code, name=payload.name.strip(), category=_ensure_category(db, payload.category),
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
            item.code, item.name, item.category, item.uom = code, payload.name.strip(), _ensure_category(db, payload.category), payload.uom.strip()
            item.stock_on_hand, item.buffer_days = payload.stock_on_hand, payload.buffer_days
            if _check_price_access(request):
                item.unit_cost, item.currency = payload.unit_cost, payload.currency
            sync_linked_supplier_materials(db, item=item)
            db.commit()
            response_data = {"id": item.id, "code": item.code, "name": item.name, "category": item.category, "uom": item.uom, "stock_on_hand": item.stock_on_hand, "buffer_days": item.buffer_days}
            if _check_price_access(request):
                response_data.update(unit_cost=item.unit_cost, currency=item.currency)
            return response_data

    def material_uses(db: Session, material_id: int) -> list[tuple[str, int, bool]]:
        """What uses a T1 code: (description, count, can follow the code to T2/T3)."""
        def count(model, *conditions) -> int:
            return db.scalar(select(func.count()).select_from(model).where(*conditions)) or 0
        from sqlalchemy import text as sa_text
        sources = db.execute(sa_text("SELECT COUNT(*) FROM supplier_bom_line_sources WHERE item_id = :id"), {"id": material_id}).scalar() or 0
        uses = [
            ("Model BOM Level 1 line", count(ModelBOM, ModelBOM.material_id == material_id), False),
            ("part BOM line", count(BOMComponent, BOMComponent.child_item_id == material_id), False),
            ("line in its own part BOM", count(BOMComponent, BOMComponent.parent_item_id == material_id), False),
            ("T1 supplier/material relationship", count(SupplierPart, SupplierPart.material_id == material_id, SupplierPart.tier == TOP_TIER), False),
            ("T2/T3 supplier/material relationship", count(SupplierPart, SupplierPart.material_id == material_id, SupplierPart.tier != TOP_TIER), True),
            ("supplier material master entry", count(SupplierMaterial, SupplierMaterial.item_id == material_id), True),
            ("supplier BOM (as the part you buy)", count(SupplierBOMLine, SupplierBOMLine.product_material_id == material_id), False),
            ("supplier BOM line", count(SupplierBOMLine, SupplierBOMLine.item_id == material_id), False),
            ("supplier BOM source", sources, False),
            ("alternate code entry", count(AlternateCode, or_(AlternateCode.main_item_id == material_id, AlternateCode.alternate_item_id == material_id)), False),
            ("own stock entry", count(OwnStockSnapshot, OwnStockSnapshot.material_id == material_id), False),
            ("supplier-to-supplier transfer", count(MaterialTransfer, MaterialTransfer.material_id == material_id), False),
        ]
        return [(label, number, follows) for label, number, follows in uses if number]

    def describe_uses(uses: list[tuple[str, int, bool]]) -> str:
        return ", ".join(f"{number} {label}{'s' if number != 1 and not label.endswith('BOM)') else ''}" for label, number, _ in uses)

    def remove_retired_records(db: Session, material_id: int) -> None:
        """Material requests of the retired stock screens (install sample) are not shown anywhere."""
        from sqlalchemy import text as sa_text
        for statement in (
            "DELETE FROM stock_deliveries WHERE po_id IN (SELECT id FROM stock_purchase_orders WHERE request_id IN (SELECT id FROM material_requests WHERE item_id = :id))",
            "DELETE FROM stock_purchase_orders WHERE request_id IN (SELECT id FROM material_requests WHERE item_id = :id)",
            "DELETE FROM material_requests WHERE item_id = :id",
        ):
            db.execute(sa_text(statement), {"id": material_id})

    @app.delete("/api/planning/materials/{material_id}", status_code=204)
    def delete_material(material_id: int, request: Request):
        current_user(request)
        if request.state.user["role"] not in SUPPLY_ROLES:
            raise HTTPException(status_code=403, detail="Master-data editing access is required")
        with open_db() as db:
            item = db.get(Item, material_id)
            if not item:
                raise HTTPException(status_code=404, detail="Material not found")
            uses = material_uses(db, material_id)
            if uses:
                raise HTTPException(status_code=409, detail=f"{item.code} cannot be deleted: it is used in {describe_uses(uses)}. Remove those first, or use Move to T2/T3 if it belongs there")
            remove_retired_records(db, material_id)
            db.delete(item)
            db.commit()
            return Response(status_code=204)

    @app.post("/api/planning/materials/{material_id}/move")
    def move_material_to_tier(material_id: int, request: Request, tier: str):
        """A code put in T1 Material Master by mistake moves to T2 or T3 Material Master (code, name,
        category, unit, price). Supplier material-master entries and T2/T3 relationships follow it;
        anything that needs a T1 code (your BOMs, T1 relationships, alternates, stock ...) must be
        removed first and is listed."""
        current_user(request)
        _check_supply_role(request)
        if tier not in {"T2", "T3"}:
            raise HTTPException(status_code=422, detail="Move to T2 or T3 Material Master")
        with open_db() as db:
            item = db.get(Item, material_id)
            if not item:
                raise HTTPException(status_code=404, detail="Material not found")
            uses = material_uses(db, material_id)
            blocking = [use for use in uses if not use[2]]
            if blocking:
                raise HTTPException(status_code=409, detail=f"{item.code} cannot move to {tier} Material Master: it is used in {describe_uses(blocking)}, which need a T1 code. Remove those first")
            moved = TierMaterial(tier=tier, code=item.code, name=item.name, category=item.category, uom=item.uom,
                                 unit_price=item.unit_cost, currency=item.currency)
            db.add(moved)
            db.flush()
            relationships = db.execute(sa_update(SupplierPart).where(SupplierPart.material_id == material_id)
                                       .values(material_id=None, tier_material_id=moved.id)).rowcount
            entries = db.execute(sa_update(SupplierMaterial).where(SupplierMaterial.item_id == material_id)
                                 .values(item_id=None, tier_material_id=moved.id)).rowcount
            remove_retired_records(db, material_id)
            db.delete(item)
            db.commit()
            return {"code": moved.code, "tier": tier, "relationships": relationships, "supplier_materials": entries}

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
        add_alternates(db, [row for rows in result.values() for row in rows], can_price)
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
            db.execute(delete(BOMLineAlternate).where(BOMLineAlternate.model_bom_id == bom_id))
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

    def alternate_groups(db: Session, main_ids: Iterable[int] | None = None) -> dict[int, list[AlternateCode]]:
        statement = select(AlternateCode).order_by(AlternateCode.id)
        if main_ids is not None:
            main_ids = list(main_ids)
            if not main_ids:
                return {}
            statement = statement.where(AlternateCode.main_item_id.in_(main_ids))
        groups: dict[int, list[AlternateCode]] = {}
        for alternate in db.scalars(statement).all():
            groups.setdefault(alternate.main_item_id, []).append(alternate)
        return groups

    def bom_price_lookup(db: Session, material_ids: Iterable[int] | None = None):
        """BOM part price. Each code's own price = SOB-weighted price of its Direct suppliers:
        sum(price x SOB) / sum(SOB); Material Master price when no Direct supplier has a price,
        or when the suppliers quote in different currencies. On a BOM line with alternate SOB
        the line price is blended by price.blend (main share x main price + SOB x alternate price).
        price.alternates(item_id) lists the code's alternates with their own price."""
        if material_ids is not None:
            material_ids = set(material_ids)
        groups = alternate_groups(db, material_ids)
        alternate_ids = {alternate.alternate_item_id for group in groups.values() for alternate in group}
        alternate_items = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_(alternate_ids))).all()} if alternate_ids else {}
        links: dict[int, list[SupplierPart]] = {}
        statement = select(SupplierPart).where(
            SupplierPart.tier == TOP_TIER, SupplierPart.parent_supplier_id.is_(None),
            SupplierPart.sob_percent > 0, SupplierPart.unit_price > 0,
        )
        if material_ids is not None:
            statement = statement.where(SupplierPart.material_id.in_(list(material_ids | alternate_ids)))
        for link in db.scalars(statement).all():
            links.setdefault(link.material_id, []).append(link)

        def own_price(item: Item) -> dict:
            quoted = links.get(item.id, [])
            if quoted and len({link.currency for link in quoted}) == 1:
                share = sum(link.sob_percent for link in quoted)
                return {"unit_price": sum(link.unit_price * link.sob_percent for link in quoted) / share,
                        "currency": quoted[0].currency, "price_source": f"SOB-weighted ({len(quoted)} supplier{'s' if len(quoted) > 1 else ''})"}
            source = "Material Master (suppliers in different currencies)" if quoted else "Material Master"
            return {"unit_price": item.unit_cost, "currency": item.currency, "price_source": source}

        def alternates(item_id: int) -> list[dict]:
            rows = []
            for alternate in groups.get(item_id, []):
                item = alternate_items.get(alternate.alternate_item_id)
                if item:
                    rows.append({"id": alternate.id, "material_id": item.id, "code": item.code, "name": item.name,
                                 "uom": item.uom, **own_price(item)})
            return rows

        def blend(item: Item, main: dict, listed: list[dict]) -> dict:
            """Line price with alternate SOB: listed = alternates with 'sob_percent' for this line."""
            used = [a for a in listed if a["sob_percent"] > 0]
            if not used:
                return main
            main_share = 100 - sum(a["sob_percent"] for a in used)
            parts = ([(main_share, item.code, main)] if main_share > 0 else []) + [(a["sob_percent"], a["code"], a) for a in used]
            if len({part["currency"] for _, _, part in parts}) != 1:
                return {**main, "price_source": f"{main['price_source']}; main code only (alternate codes priced in another currency)"}
            detail = " + ".join(f"{code} {share:g}% x {part['unit_price']:,.2f}" for share, code, part in parts)
            return {"unit_price": sum(share * part["unit_price"] for share, _, part in parts) / 100,
                    "currency": parts[0][2]["currency"], "price_source": f"SOB-weighted with alternate codes: {detail}"}

        price = own_price
        price.alternates = alternates
        price.own = own_price
        price.blend = blend
        return price

    def line_sobs(db: Session, model_bom_ids: Iterable[int] = (), component_ids: Iterable[int] = (), everything: bool = False) -> dict[tuple[str, int], list[tuple[str, int, int, float]]]:
        """Saved alternate SOB history per BOM line: ('root', id) / ('component', id) ->
        [(from_month, main id, alternate id, SOB)]."""
        model_bom_ids, component_ids = list(model_bom_ids), list(component_ids)
        conditions = [BOMLineAlternate.id.is_not(None)] if everything else []
        if model_bom_ids:
            conditions.append(BOMLineAlternate.model_bom_id.in_(model_bom_ids))
        if component_ids:
            conditions.append(BOMLineAlternate.component_id.in_(component_ids))
        result: dict[tuple[str, int], list[tuple[str, int, int, float]]] = {}
        if not conditions:
            return result
        for entry in db.scalars(select(BOMLineAlternate).where(or_(*conditions))).all():
            key = ("root", entry.model_bom_id) if entry.model_bom_id is not None else ("component", entry.component_id)
            result.setdefault(key, []).append((entry.from_month or "", entry.main_item_id, entry.alternate_item_id, entry.sob_percent))
        return result

    def line_alternates(price, item: Item, history: list[tuple[str, int, int, float]], quantity: float, month: str) -> tuple[list[dict], float, dict | None]:
        """The code's alternates with this line's SOB valid in the month (0% until set), their share
        of the quantity, and the next SOB change after the month (if any)."""
        listed = price.alternates(item.id)
        allowed = {alternate["material_id"] for alternate in listed}
        entries = [(from_month, alternate_id, sob) for from_month, main_id, alternate_id, sob in history
                   if main_id == item.id and alternate_id in allowed]
        now = _sob_at(entries, month)
        later = sorted({from_month for from_month, _, _ in entries if from_month > month})
        upcoming = _sob_at(entries, later[0]) if later else None
        for alternate in listed:
            alternate["sob_percent"] = now.get(alternate["material_id"], 0.0)
            alternate["quantity"] = quantity * alternate["sob_percent"] / 100
            if upcoming is not None:
                alternate["next_sob_percent"] = upcoming.get(alternate["material_id"], 0.0)
        next_change = ({"from_month": later[0], "main_sob_percent": 100 - sum(upcoming.get(a["material_id"], 0.0) for a in listed)}
                       if upcoming is not None else None)
        return listed, 100 - sum(alternate["sob_percent"] for alternate in listed), next_change

    def add_alternates(db: Session, rows: list[dict], can_price: bool, month: str | None = None) -> None:
        """Each BOM row whose code has alternate codes gets 'alternates' (code, name, this line's
        SOB % this month, its share of the line quantity, own price, next SOB), 'main_sob_percent',
        'main_quantity', 'sob_next_change' and, with price access, the line price blended by SOB."""
        groups = alternate_groups(db, {row["material_id"] for row in rows})
        if not groups:
            return
        month = month or _month_text(date.today())
        price = bom_price_lookup(db, groups)
        main_rows = [row for row in rows if row["material_id"] in groups]
        saved = line_sobs(db, [row["root_bom_id"] for row in main_rows if row.get("root_bom_id")],
                          [row["component_id"] for row in main_rows if row.get("component_id")])
        items = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_(list(groups)))).all()}
        for row in main_rows:
            item = items.get(row["material_id"])
            if not item:
                continue
            key = ("root", row["root_bom_id"]) if row.get("root_bom_id") else ("component", row.get("component_id"))
            listed, main_share, next_change = line_alternates(price, item, saved.get(key, []), row["quantity"], month)
            row["main_sob_percent"] = main_share
            row["sob_month"] = month
            row["sob_next_change"] = next_change
            # The main code's own share of the BOM quantity (main + alternates = BOM quantity)
            row["main_quantity"] = row["quantity"] * main_share / 100
            if can_price:
                row.update(price.blend(item, price.own(item), listed))
            else:
                for alternate in listed:
                    for k in ("unit_price", "currency", "price_source"):
                        alternate.pop(k, None)
            row["alternates"] = listed

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
            add_alternates(db, rows, can_price)
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
            add_alternates(db, rows, can_price)
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
                    db.execute(delete(BOMLineAlternate).where(BOMLineAlternate.component_id == link.id))
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
    def list_supplier_parts(request: Request, material_id: int | None = None, offset: int = 0, limit: int = 500, search: str = "", supplier_id: int | None = None, tier: str = ""):
        current_user(request)
        _check_supply_role(request)
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        with open_db() as db:
            statement = select(SupplierPart).outerjoin(Item, SupplierPart.material_id == Item.id).outerjoin(
                TierMaterial, SupplierPart.tier_material_id == TierMaterial.id
            ).join(Supplier, SupplierPart.supplier_id == Supplier.id).order_by(SupplierPart.material_id, SupplierPart.tier, SupplierPart.id)
            if material_id:
                statement = statement.where(SupplierPart.material_id == material_id)
            if supplier_id:
                statement = statement.where(SupplierPart.supplier_id == supplier_id)
            if tier:
                statement = statement.where(SupplierPart.tier == _normalize_tier(tier))
            if search.strip():
                like = f"%{search.strip()}%"
                statement = statement.where(or_(
                    Item.code.ilike(like), Item.name.ilike(like), TierMaterial.code.ilike(like), TierMaterial.name.ilike(like),
                    Supplier.code.ilike(like), Supplier.name.ilike(like), SupplierPart.tier.ilike(like),
                ))
            total = db.scalar(select(func.count()).select_from(statement.subquery())) or 0
            links = db.scalars(statement.offset(offset).limit(limit)).all()
            can_price = _check_price_access(request)
            rows = []
            for link in links:
                item = db.get(Item, link.material_id) if link.material_id else db.get(TierMaterial, link.tier_material_id)
                supplier = db.get(Supplier, link.supplier_id)
                parent = db.get(Supplier, link.parent_supplier_id) if link.parent_supplier_id else None
                rows.append({
                    "id": link.id, "material_id": link.material_id, "tier_material_id": link.tier_material_id,
                    "code_master": "T1" if link.material_id else (item.tier if item else "T2"),
                    "material_code": item.code if item else "",
                    "material_name": item.name if item else "", "supplier_id": link.supplier_id,
                    "supplier_code": supplier.code if supplier else "", "supplier_name": supplier.name if supplier else "",
                    "tier": link.tier, "parent_supplier_id": link.parent_supplier_id,
                    "parent_supplier_name": parent.name if parent else None, "parent_link_id": link.parent_link_id, "sob_percent": link.sob_percent,
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
        with open_db() as db:
            link = new_supplier_part(db, payload, request)
            db.commit()
            db.refresh(link)
            return {"id": link.id, "material_id": link.material_id, "supplier_id": link.supplier_id, "tier": link.tier, "sob_percent": link.sob_percent}

    def chain_link_json(db: Session, link: SupplierPart, can_price: bool, cache: dict) -> dict:
        key = ("item", link.material_id) if link.material_id else ("tier", link.tier_material_id)
        if key not in cache:
            cache[key] = db.get(Item, link.material_id) if link.material_id else db.get(TierMaterial, link.tier_material_id)
        part = cache[key]
        if ("supplier", link.supplier_id) not in cache:
            cache[("supplier", link.supplier_id)] = db.get(Supplier, link.supplier_id)
        supplier = cache[("supplier", link.supplier_id)]
        data = {"id": link.id, "tier": link.tier, "supplier_id": link.supplier_id, "supplier_code": supplier.code if supplier else "",
                "supplier_name": supplier.name if supplier else "", "material_id": link.material_id, "tier_material_id": link.tier_material_id,
                "material_code": part.code if part else "", "material_name": part.name if part else "",
                "code_master": "T1" if link.material_id else (part.tier if part else "T2"),
                "parent_supplier_id": link.parent_supplier_id, "parent_link_id": link.parent_link_id, "sob_percent": link.sob_percent,
                "lead_time_days": link.lead_time_days, "moq": link.moq, "buffer_days": link.buffer_days, "price_mode": link.price_mode,
                "price_interval_months": link.price_interval_months, "price_status": _price_status(link, _month_text(date.today()))}
        if can_price:
            data.update(unit_price=link.unit_price, currency=link.currency)
        return data

    def supply_chain_tree(rows: list[dict]) -> list[dict]:
        """Chain rows as a tree, as on the coverage screen: the part (depth 0, with its T1 supplier count and
        total SOB), each T1 supplier (1), its T2 (2) and T3 (3). Each node carries its parents and its own link."""
        by_part: dict[str, dict[int, dict]] = {}
        for row in rows:
            if row["t1"]:
                by_part.setdefault(row["t1"]["material_code"], {})[row["t1"]["id"]] = row["t1"]
        seen: set = set()
        nodes: list[dict] = []
        for row in rows:
            t1, t2, t3 = row["t1"], row["t2"], row["t3"]
            if t1 and ("p", t1["material_code"]) not in seen:
                seen.add(("p", t1["material_code"]))
                suppliers = list(by_part[t1["material_code"]].values())
                nodes.append({"depth": 0, "t1": t1, "t2": None, "t3": None, "link": None, "suppliers": len(suppliers),
                              "total_sob": round(sum(link["sob_percent"] or 0 for link in suppliers), 4)})
            for depth, key, link in ((1, (1, t1 and t1["id"]), t1), (2, (2, t1 and t1["id"], t2 and t2["id"]), t2),
                                     (3, (3, t1 and t1["id"], t2 and t2["id"], t3 and t3["id"]), t3)):
                if link and key not in seen:
                    seen.add(key)
                    nodes.append({"depth": depth, "t1": t1, "t2": t2 if depth >= 2 else None, "t3": t3 if depth == 3 else None, "link": link})
        return nodes

    def supply_chain_rows(db: Session, can_price: bool, search: str = "") -> list[dict]:
        """Coverage as chains: one row per path T1 -> T2 -> T3 (as in the user's sheet). A T2/T3
        relationship sits under its parent relationship; rows saved without one sit under every
        relationship of their parent supplier at the tier above; rows with no parent there are
        listed on their own."""
        links = db.scalars(select(SupplierPart)).all()
        cache: dict = {}
        by_parent_link: dict[int, list[SupplierPart]] = {}
        unlinked: dict[tuple[int, str], list[SupplierPart]] = {}
        for link in links:
            if link.parent_link_id:
                by_parent_link.setdefault(link.parent_link_id, []).append(link)
            elif link.tier != TOP_TIER and link.parent_supplier_id:
                unlinked.setdefault((link.parent_supplier_id, link.tier), []).append(link)
        placed: set[int] = set()

        def children(link: SupplierPart) -> list[SupplierPart]:
            next_tier = f"T{int(link.tier[1:]) + 1}" if re.fullmatch(r"T[1-9]", link.tier or "") else ""
            found = by_parent_link.get(link.id, []) + unlinked.get((link.supplier_id, next_tier), [])
            return sorted(found, key=lambda child: (chain_link_json(db, child, False, cache)["material_code"], chain_link_json(db, child, False, cache)["supplier_name"]))

        def paths(link: SupplierPart) -> list[list[SupplierPart]]:
            placed.add(link.id)
            below = children(link) if int(link.tier[1:]) < 3 else []
            if not below:
                return [[link]]
            return [[link] + path for child in below for path in paths(child)]

        groups: list[list[list[SupplierPart]]] = []
        top = sorted([link for link in links if link.tier == TOP_TIER], key=lambda link: (chain_link_json(db, link, False, cache)["material_code"], chain_link_json(db, link, False, cache)["supplier_name"]))
        for link in top:
            groups.append(paths(link))
        for link in sorted(links, key=lambda item: (item.tier, item.id)):  # T2/T3 rows with no parent relationship shown
            if link.id not in placed and link.tier in ("T2", "T3"):
                groups.append([[None] * (int(link.tier[1:]) - 1) + path for path in paths(link)])
        needle = search.strip().casefold()
        rows = []
        for group_index, group in enumerate(groups):
            group_rows = []
            for path in group:
                blocks = {}
                for position, link in enumerate(path):
                    blocks[f"t{position + 1}"] = chain_link_json(db, link, can_price, cache) if link else None
                last = next(block for block in reversed([blocks.get(f"t{i}") for i in (1, 2, 3)]) if block)
                group_rows.append({"group": group_index, "t1": blocks.get("t1"), "t2": blocks.get("t2"), "t3": blocks.get("t3"), "last": last})
            if needle and not any(needle in " ".join(f"{b['supplier_code']} {b['supplier_name']} {b['material_code']} {b['material_name']}".casefold()
                                                     for b in (row["t1"], row["t2"], row["t3"]) if b) for row in group_rows):
                continue
            rows.extend(group_rows)
        return rows

    @app.get("/api/planning/supplier-chain")
    def supplier_chain(request: Request, search: str = ""):
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            rows = supply_chain_rows(db, _check_price_access(request), search)
            return {"total": len(rows), "items": rows}

    @app.get("/api/planning/supplier-chain.{file_format}")
    def export_supplier_chain(file_format: str, request: Request, search: str = ""):
        """Coverage in the chain layout: T1 / T2 / T3 Supplier blocks, then the last supplier's terms."""
        current_user(request)
        _check_supply_role(request)
        if file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export not found")
        can_price = _check_price_access(request)
        with open_db() as db:
            rows = supply_chain_rows(db, can_price, search)
        # T1 block starts with your part (Material code, name), then its supplier; T2/T3 start with the supplier
        # tree as on the screen: Level, T1 block (your part first), T2 / T3 blocks (supplier first), then the
        # terms of the supplier on that row (each level has its own SOB, lead time, MOQ, buffer, price)
        t1_heads = ["Material code", "Material name", "Supplier code", "Supplier name"]
        block_heads = ["Supplier code", "Supplier name", "Material code", "Material name"]
        terms = ["SOB %", "Lead time days", "MOQ", "Buffer days", "Price mode", "Price status"] + (["Unit price", "Currency"] if can_price else [])
        mode_label = {"spot": "Spot", "fixed_monthly": "Fixed monthly", "every_n_months": "Every N months"}

        def ident(link: dict | None, first: bool = False) -> list:
            if not link:
                return [""] * 4
            supplier, part = [link["supplier_code"], link["supplier_name"]], [link["material_code"], link["material_name"]]
            return part + supplier if first else supplier + part

        values, depths = [], []
        for node in supply_chain_tree(rows):
            link = node["link"]
            if node["depth"] == 0:
                t1 = node["t1"]
                line = ["Part", t1["material_code"], t1["material_name"], "", f'{node["suppliers"]} T1 supplier{"s" if node["suppliers"] != 1 else ""}'] + [""] * 8 + [node["total_sob"]] + [""] * (len(terms) - 1)
            else:
                line = [link["tier"]] + ident(node["t1"], True) + ident(node["t2"]) + ident(node["t3"]) + [
                    link["sob_percent"], link["lead_time_days"], link["moq"], link["buffer_days"], mode_label.get(link["price_mode"], link["price_mode"]), link["price_status"]] + (
                    [link.get("unit_price"), link.get("currency")] if can_price else [])
            values.append(line)
            depths.append(node["depth"])
        if file_format == "pdf":
            headers = ["Level"] + [f"T1 {h}" for h in t1_heads] + [f"{tier} {h}" for tier in ("T2", "T3") for h in block_heads] + terms
            output = _generic_pdf("Supplier/material coverage", headers, values)
            return StreamingResponse(output, media_type="application/pdf", headers={"Content-Disposition": 'attachment; filename="supplier-material-coverage.pdf"'})
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Coverage"
        dark = PatternFill("solid", fgColor="12263A")
        white = Font(bold=True, color="FFFFFF")
        sheet.append(["Design to Cost · Supplier/material coverage"])
        sheet.append([])
        sheet.append(["Level", "T1 Supplier", None, None, None, "T2 Supplier", None, None, None, "T3 Supplier", None, None, None, "Terms of the supplier on the row"] + [None] * (len(terms) - 1))
        sheet.append([None] + t1_heads + block_heads * 2 + terms)
        sheet.merge_cells(start_row=3, start_column=1, end_row=4, end_column=1)
        for index in range(3):
            sheet.merge_cells(start_row=3, start_column=2 + index * 4, end_row=3, end_column=5 + index * 4)
        sheet.merge_cells(start_row=3, start_column=14, end_row=3, end_column=13 + len(terms))
        for row_cells in sheet.iter_rows(min_row=3, max_row=4):
            for cell in row_cells:
                cell.fill, cell.font = dark, white
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        part_font = Font(bold=True)
        for line, depth in zip(values, depths):
            sheet.append(line)
            number_row = sheet.max_row
            if depth == 0:
                for cell in sheet[number_row]:
                    cell.font = part_font
            else:  # + / - in Excel, like the screen
                sheet.row_dimensions[number_row].outline_level = depth
        sheet.sheet_properties.outlinePr.summaryBelow = False
        for column in range(1, 14 + len(terms)):
            sheet.column_dimensions[__import__("openpyxl").utils.get_column_letter(column)].width = 8 if column == 1 else 16
        sheet.freeze_panes = "B5"
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 headers={"Content-Disposition": 'attachment; filename="supplier-material-coverage.xlsx"'})

    @app.post("/api/planning/supplier-chain/rows", status_code=201)
    def save_supplier_chain_rows(payload: dict, request: Request):
        """Chain rows typed on the relationship page, all or none. Each row: t1 / t2 / t3 blocks
        {supplier_id, material_id | tier_material_id, sob_percent, unit_price} (a block may be left out:
        the previous row's T1 / T2 is used, as in the coverage sheet) and the terms of the last
        supplier in the row. A relationship already there is updated."""
        current_user(request)
        _check_supply_role(request)
        rows = payload.get("rows") or []
        if not rows:
            raise HTTPException(status_code=422, detail="Add at least one row")
        term_keys = ("lead_time_days", "moq", "buffer_days", "currency", "price_mode", "price_interval_months")
        with open_db() as db:
            previous: dict[str, SupplierPart | None] = {"t1": None, "t2": None}
            touched: set[int] = set()
            saved = 0
            for index, row in enumerate(rows, start=1):
                try:
                    chain: list[SupplierPart] = []
                    present = [key for key in ("t1", "t2", "t3") if row.get(key)]
                    if not present:
                        continue
                    deepest = present[-1]
                    for position, key in enumerate(("t1", "t2", "t3")):
                        block = row.get(key)
                        tier = f"T{position + 1}"
                        if not block:
                            if position < ["t1", "t2", "t3"].index(deepest):
                                inherited = previous.get(key)
                                if inherited is None:
                                    raise HTTPException(status_code=422, detail=f"fill the {tier} supplier (no {tier} in the row above to continue)")
                                chain.append(inherited)
                            continue
                        parent = chain[-1] if chain else None
                        if position > 0 and parent is None:
                            raise HTTPException(status_code=422, detail=f"a {tier} supplier needs its T{position} supplier on the left")
                        values = {"supplier_id": block.get("supplier_id"), "material_id": block.get("material_id") or None,
                                  "tier_material_id": block.get("tier_material_id") or None, "tier": tier,
                                  "parent_link_id": parent.id if parent else None, "parent_supplier_id": parent.supplier_id if parent else None,
                                  "sob_percent": float(block.get("sob_percent") if block.get("sob_percent") not in (None, "") else 100),
                                  "unit_price": float(block.get("unit_price") or 0)}
                        if key == deepest:
                            values.update({term: row[term] for term in term_keys if row.get(term) not in (None, "")})
                        elif row.get("currency"):
                            values["currency"] = row["currency"]  # each block's unit price is in the row's currency
                        # each level can have its own terms (typed in its block)
                        values.update({term: block[term] for term in term_keys if block.get(term) not in (None, "")})
                        try:
                            clean = SupplierPartInput(**values).model_dump()
                        except Exception as exc:
                            raise HTTPException(status_code=422, detail=f"check the {tier} values (SOB 0-100, numbers 0 or more)") from exc
                        part_column = SupplierPart.material_id if clean["material_id"] else SupplierPart.tier_material_id
                        existing = db.scalar(select(SupplierPart).where(
                            SupplierPart.supplier_id == clean["supplier_id"], SupplierPart.tier == tier,
                            part_column == (clean["material_id"] or clean["tier_material_id"]),
                            (SupplierPart.parent_link_id == clean["parent_link_id"]) if clean["parent_link_id"] else SupplierPart.parent_link_id.is_(None),
                            (SupplierPart.parent_supplier_id == clean["parent_supplier_id"]) if clean["parent_supplier_id"] else SupplierPart.parent_supplier_id.is_(None)))
                        if existing and existing.id in touched:
                            link = existing  # same relationship repeated on a later row: keep the first values
                        elif existing:
                            update = {k: v for k, v in clean.items() if k in values or k in ("material_id", "tier_material_id")}
                            checked = check_supplier_part(db, {**{k: getattr(existing, k) for k in clean}, **update}, request, exclude_id=existing.id)
                            for field_name, value in checked.items():
                                setattr(existing, field_name, value)
                            db.flush()
                            link = existing
                        else:
                            link = SupplierPart(**check_supplier_part(db, clean, request))
                            db.add(link)
                            db.flush()
                            tiers_from_links(db, link)
                        touched.add(link.id)
                        chain.append(link)
                        saved += 1
                    previous["t1"] = chain[0] if chain else previous["t1"]
                    previous["t2"] = chain[1] if len(chain) > 1 else (None if row.get("t1") else previous["t2"])
                except HTTPException as exc:
                    db.rollback()
                    raise HTTPException(status_code=exc.status_code, detail=f"Row {index}: {exc.detail}") from exc
            db.commit()
            return {"rows": len(rows), "relationships": len(touched)}

    @app.post("/api/planning/supplier-parts/rows", status_code=201)
    def add_supplier_part_rows(payload: SupplierPartRows, request: Request):
        """Several relationships typed in the entry table: all rows are saved, or none."""
        current_user(request)
        _check_supply_role(request)
        if not payload.rows:
            raise HTTPException(status_code=422, detail="Add at least one row")
        with open_db() as db:
            for index, row in enumerate(payload.rows, start=1):
                try:
                    new_supplier_part(db, row, request)
                except HTTPException as exc:
                    db.rollback()
                    raise HTTPException(status_code=exc.status_code, detail=f"Row {index}: {exc.detail}") from exc
            db.commit()
            return {"rows": len(payload.rows)}

    @app.delete("/api/planning/supplier-parts/{link_id}", status_code=204)
    def delete_supplier_part(link_id: int, request: Request):
        """Delete a supplier/material relationship; refused while it has POs or supplier stock."""
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            link = db.get(SupplierPart, link_id)
            if not link:
                raise HTTPException(status_code=404, detail="Supplier/material link not found")
            pos = db.scalar(select(func.count()).select_from(PlanningPO).where(PlanningPO.supplier_part_id == link_id)) or 0
            stock = db.scalar(select(func.count()).select_from(SupplierStockSnapshot).where(SupplierStockSnapshot.supplier_part_id == link_id)) or 0
            children = db.scalar(select(func.count()).select_from(SupplierPart).where(SupplierPart.parent_link_id == link_id)) or 0
            if children:
                raise HTTPException(status_code=409, detail=f"This relationship has {children} {'supplier' if children == 1 else 'suppliers'} under it (next tier); delete or move those first")
            if pos or stock:
                used = ", ".join(text for text in (f"{pos} PO{'s' if pos != 1 else ''}" if pos else "", f"{stock} supplier stock row{'s' if stock != 1 else ''}" if stock else "") if text)
                raise HTTPException(status_code=409, detail=f"This relationship has {used}; it cannot be deleted. Set its SOB to 0 instead")
            db.delete(link)
            db.commit()
            return Response(status_code=204)

    @app.post("/api/suppliers/upload", status_code=201)
    async def upload_suppliers(request: Request, tier: str = "T1"):
        """Supplier register from Excel, all rows or none (heading row found automatically; your own
        headings are fine). Needed: Supplier name. Optional: Supplier code (blank = next free SUP-
        code), Category (new ones are added to the list), Country, Contact person, Email, Phone,
        Approval (Approved / Pending review / On hold; new suppliers default to Approved). A row
        whose code (or, without a code, name) is already in the register updates that supplier."""
        current_user(request)
        _check_supply_role(request)
        rows = _read_flexible_sheet(await request.body(), SUPPLIER_COLUMN_ALIASES, {"name"})
        statuses = {"approved": "Approved", "pending review": "Pending review", "pending": "Pending review", "on hold": "On hold", "hold": "On hold"}
        cell = lambda row, key: " ".join(str(row.get(key) if row.get(key) is not None else "").split())
        with open_db() as db:
            suppliers = db.scalars(select(Supplier)).all()
            by_code = {supplier.code.lower(): supplier for supplier in suppliers}
            by_name = {supplier.name.lower(): supplier for supplier in suppliers}
            used_codes = {supplier.code.lower() for supplier in suppliers}
            next_number = len(suppliers) + 1001
            seen_codes: set[str] = set()
            seen_names: set[str] = set()
            added = updated = 0
            for row in rows:
                number = row["_excel_row"]
                name, code = cell(row, "name"), cell(row, "code")
                if len(name) < 2:
                    raise HTTPException(status_code=422, detail=f"Row {number}: supplier name is required (at least 2 characters)")
                if code.lower() in seen_codes and code:
                    raise HTTPException(status_code=422, detail=f"Row {number}: supplier code {code} is repeated in the file")
                if name.lower() in seen_names:
                    raise HTTPException(status_code=422, detail=f"Row {number}: supplier {name} is repeated in the file")
                seen_names.add(name.lower())
                if code:
                    seen_codes.add(code.lower())
                status_text = cell(row, "status").lower()
                if status_text and status_text not in statuses:
                    raise HTTPException(status_code=422, detail=f"Row {number}: approval must be Approved, Pending review or On hold")
                supplier = by_code.get(code.lower()) if code else by_name.get(name.lower())
                if supplier is None and code and name.lower() in by_name:
                    raise HTTPException(status_code=422, detail=f"Row {number}: {name} is already in the register as {by_name[name.lower()].code}; use that code to update it")
                category_text = cell(row, "category")
                if supplier:
                    other = by_name.get(name.lower())
                    if other and other.id != supplier.id:
                        raise HTTPException(status_code=422, detail=f"Row {number}: the name {name} is already used by {other.code}")
                    by_name.pop(supplier.name.lower(), None)
                    supplier.name = name
                    by_name[name.lower()] = supplier
                    if category_text:
                        supplier.category = _ensure_category(db, category_text)
                    for key in ("country", "contact_name", "email", "phone"):
                        if cell(row, key):
                            setattr(supplier, key, cell(row, key))
                    if status_text:
                        supplier.status = statuses[status_text]
                    for value in (_tier_list(cell(row, "tiers")) or [tier]):
                        add_supplier_tier(supplier, value)
                    updated += 1
                else:
                    if not code:
                        while f"sup-{next_number:04d}" in used_codes:
                            next_number += 1
                        code = f"SUP-{next_number:04d}"
                    if cell(row, "tiers") and not _tier_list(cell(row, "tiers")):
                        raise HTTPException(status_code=422, detail=f"Row {number}: tier must be T1, T2 and/or T3")
                    supplier = Supplier(code=code, name=name, tiers=",".join(_tier_list(cell(row, "tiers")) or _tier_list(tier) or ["T1"]),
                                        category=_ensure_category(db, category_text or "Raw material"),
                                        country=cell(row, "country") or "-", contact_name=cell(row, "contact_name"), email=cell(row, "email"),
                                        phone=cell(row, "phone"), status=statuses.get(status_text, "Approved"), risk="Medium")
                    db.add(supplier)
                    db.flush()
                    used_codes.add(code.lower())
                    by_code[code.lower()] = supplier
                    by_name[name.lower()] = supplier
                    added += 1
            db.commit()
            return {"rows": len(rows), "added": added, "updated": updated}

    @app.delete("/api/suppliers/{supplier_id}", status_code=204)
    def delete_supplier(supplier_id: int, request: Request):
        """Delete a supplier from the register; refused (with where it is used) while anything uses it."""
        current_user(request)
        _check_supply_role(request)
        from sqlalchemy import text as sa_text
        with open_db() as db:
            supplier = db.get(Supplier, supplier_id)
            if not supplier:
                raise HTTPException(status_code=404, detail="Supplier not found")
            checks = [
                ("supplier/material relationship", "SELECT COUNT(*) FROM supplier_parts WHERE supplier_id = :id OR parent_supplier_id = :id"),
                ("supplier material", "SELECT COUNT(*) FROM supplier_materials WHERE supplier_id = :id"),
                ("supplier BOM line", "SELECT COUNT(*) FROM supplier_bom_lines WHERE supplier_id = :id"),
                ("supplier BOM source", "SELECT COUNT(*) FROM supplier_bom_line_sources WHERE sub_supplier_id = :id"),
                ("material transfer", "SELECT COUNT(*) FROM material_transfers WHERE from_supplier_id = :id OR to_supplier_id = :id"),
            ]
            used = []
            for label, query in checks:
                count = db.execute(sa_text(query), {"id": supplier_id}).scalar() or 0
                if count:
                    used.append(f"{count} {label}{'s' if count != 1 and not label.endswith('RFQ') else ''}")
            if used:
                raise HTTPException(status_code=409, detail=f"{supplier.name} is used: {', '.join(used)}. Remove those first, or set its status to On hold")
            # Old records of the retired RFQ / purchase-order / material-request screens (the sample
            # data the app was installed with) are not shown anywhere: remove them with the supplier.
            for statement in (
                "DELETE FROM deliveries WHERE po_id IN (SELECT id FROM purchase_orders WHERE supplier_id = :id)",
                "DELETE FROM purchase_orders WHERE supplier_id = :id",
                "DELETE FROM stock_deliveries WHERE po_id IN (SELECT id FROM stock_purchase_orders WHERE supplier_id = :id)",
                "DELETE FROM stock_purchase_orders WHERE supplier_id = :id",
                "DELETE FROM quotes WHERE supplier_id = :id",
                "DELETE FROM rfq_invitations WHERE supplier_id = :id",
                "UPDATE rfqs SET awarded_supplier_id = NULL WHERE awarded_supplier_id = :id",
                "UPDATE items SET preferred_supplier_id = NULL WHERE preferred_supplier_id = :id",
            ):
                db.execute(sa_text(statement), {"id": supplier_id})
            db.delete(supplier)
            db.commit()
            return Response(status_code=204)

    def check_supplier_part(db: Session, values: dict, request: Request, exclude_id: int | None = None) -> dict:
        """Rules for a supplier/material relationship (new or edited); returns the clean values."""
        values = dict(values)
        values["tier"] = _normalize_tier(values.get("tier"))
        if values.get("unit_price") and not _check_price_access(request) and exclude_id is None:
            raise HTTPException(status_code=403, detail="Price access is required to set a part price")
        material_id, tier_material_id = values.get("material_id"), values.get("tier_material_id")
        if bool(material_id) == bool(tier_material_id):
            raise HTTPException(status_code=422, detail="Choose one part code (from T1, or for T2/T3 rows from T2/T3 Material Master)")
        if values["tier"] != TOP_TIER and not values.get("parent_supplier_id"):
            raise HTTPException(status_code=422, detail=f"A {values['tier']} supplier must have its parent supplier (the supplier it supplies)")
        if values["tier"] == TOP_TIER:
            values["parent_supplier_id"] = None
            values["parent_link_id"] = None
        if values.get("parent_link_id"):
            parent_link = db.get(SupplierPart, values["parent_link_id"])
            if not parent_link:
                raise HTTPException(status_code=404, detail="Parent relationship not found")
            if parent_link.id == exclude_id:
                raise HTTPException(status_code=422, detail="A relationship cannot be its own parent")
            expected = f"T{int(parent_link.tier[1:]) + 1}" if re.fullmatch(r"T[1-9]", parent_link.tier or "") else ""
            if values["tier"] != expected:
                raise HTTPException(status_code=422, detail=f"Under a {parent_link.tier} relationship the tier must be {expected}")
            values["parent_supplier_id"] = parent_link.supplier_id
        if not db.get(Supplier, values["supplier_id"]):
            raise HTTPException(status_code=404, detail="Supplier not found")
        if values.get("parent_supplier_id") and not db.get(Supplier, values["parent_supplier_id"]):
            raise HTTPException(status_code=404, detail="Parent supplier not found")
        if values.get("parent_supplier_id") == values["supplier_id"]:
            raise HTTPException(status_code=422, detail="A supplier cannot be its own parent supplier")
        if material_id:
            material = db.get(Item, material_id)
            if not material:
                raise HTTPException(status_code=404, detail="Material not found")
            reason = not_purchased_reason(db, material.id)
            if reason:
                raise HTTPException(status_code=422, detail=f"{material.code} {reason}, so it cannot have a supplier")
            column = SupplierPart.material_id
        else:
            material = db.get(TierMaterial, tier_material_id)
            if not material:
                raise HTTPException(status_code=404, detail="T2/T3 code not found")
            if values["tier"] == TOP_TIER:
                raise HTTPException(status_code=422, detail=f"{material.code} is a {material.tier} Material Master code: a T1 supplier needs a T1 code; use tier T2 or T3 for your suppliers' suppliers")
            column = SupplierPart.tier_material_id
        # SOB group: same part and tier under the same parent relationship (or parent supplier, for rows without one)
        same_route = [column == (material_id or tier_material_id), SupplierPart.tier == values["tier"],
                      (SupplierPart.parent_supplier_id == values["parent_supplier_id"]) if values.get("parent_supplier_id") else SupplierPart.parent_supplier_id.is_(None),
                      (SupplierPart.parent_link_id == values["parent_link_id"]) if values.get("parent_link_id") else SupplierPart.parent_link_id.is_(None)]
        if exclude_id:
            same_route.append(SupplierPart.id != exclude_id)
        share_total = db.scalar(select(func.coalesce(func.sum(SupplierPart.sob_percent), 0)).where(*same_route)) or 0
        if share_total + float(values.get("sob_percent", 100)) > 100 + 1e-9:
            left = max(0.0, 100 - share_total)
            raise HTTPException(status_code=422, detail=f"Supplier SOB cannot exceed 100% within the same material tier and parent supplier: other suppliers already have {share_total:g}%, so at most {left:g}% is left. Lower another supplier's SOB first (Edit)")
        if db.scalar(select(SupplierPart.id).where(*same_route, SupplierPart.supplier_id == values["supplier_id"]).limit(1)):
            raise HTTPException(status_code=409, detail=f"{material.code} with this supplier, tier and parent already exists")
        return values

    def new_supplier_part(db: Session, payload: SupplierPartInput, request: Request) -> SupplierPart:
        """Check and add one supplier/material relationship (flushed, not committed)."""
        values = check_supplier_part(db, payload.model_dump(), request)
        link = SupplierPart(**values)
        db.add(link)
        db.flush()  # later rows of the same batch see it (SOB total, duplicates)
        tiers_from_links(db, link)
        return link

    @app.post("/api/planning/supplier-parts/upload", status_code=201)
    async def upload_supplier_parts(request: Request):
        """Supplier/material relationships from Excel, all rows or none. One row per
        material + supplier (+ tier / parent supplier). Unknown supplier names are added to
        the Supplier register (Approved); an existing relationship is updated."""
        current_user(request)
        _check_supply_role(request)
        rows = _read_supplier_part_sheet(await request.body(), required={"material_code"})
        can_price = _check_price_access(request)
        with open_db() as db:
            items = {item.code.casefold(): item for item in db.scalars(select(Item)).all()}
            tier_codes = {code.code.casefold(): code for code in db.scalars(select(TierMaterial)).all()}
            suppliers = {supplier.name.strip().casefold(): supplier for supplier in db.scalars(select(Supplier)).all()}
            suppliers.update({supplier.code.casefold(): supplier for supplier in suppliers.values()})
            existing = {(("t1", link.material_id) if link.material_id else ("tm", link.tier_material_id), link.supplier_id, link.tier, link.parent_supplier_id): link
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
                supplier = Supplier(code=f"SUP-{next_number:04d}", name=name, tiers="", category=_ensure_category(db, category or "Raw material"), country="-",
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
                supplier_code = text(row, "supplier_code")
                if supplier_code:  # Supplier Code column: the register's supplier (its name in the sheet is only a label)
                    known = suppliers.get(supplier_code.casefold())
                    if not known and not name:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: supplier code {supplier_code} is not in the Supplier register")
                    if known:
                        name = known.name
                if not code or not name:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: part code and supplier (code or name) are required")
                try:
                    tier = _normalize_tier(text(row, "tier"))
                except HTTPException as exc:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {exc.detail}") from exc
                material = items.get(code.casefold())
                tier_code = None if material else tier_codes.get(code.casefold())
                if not material and not tier_code:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: material_code {code!r} is not in T1, T2 or T3 Material Master")
                if tier_code and tier == TOP_TIER:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {code} is a {tier_code.tier} Material Master code: a T1 supplier needs a T1 code; use tier T2 or T3")
                if material:
                    if material.id not in reasons:
                        reasons[material.id] = not_purchased_reason(db, material.id)
                    if reasons[material.id]:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: {material.code} {reasons[material.id]}, so it cannot have a supplier")
                parent_name = text(row, "parent_supplier_name")
                if tier != TOP_TIER and not parent_name:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: tier {tier} needs parent_supplier_name (the supplier it supplies)")
                if tier == TOP_TIER:
                    parent_name = ""
                part = material or tier_code
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
                supplier = supplier_for(name, part.category)
                parent = supplier_for(parent_name, part.category) if parent_name else None
                key = (("t1", material.id) if material else ("tm", tier_code.id), supplier.id, tier, parent.id if parent else None)
                if key in seen:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {part.code} / {supplier.name} / {tier} is repeated in the file")
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
                    link = SupplierPart(material_id=material.id if material else None, tier_material_id=tier_code.id if tier_code else None,
                                        supplier_id=supplier.id, tier=tier, parent_supplier_id=key[3], **values)
                    tiers_from_links(db, link)
                    db.add(link)
                    existing[key] = link
                    added += 1
            # SOB within a material + tier + parent-supplier route cannot exceed 100 %
            db.flush()
            totals: dict[tuple, float] = {}
            for link in existing.values():
                route = (link.material_id, link.tier_material_id, link.tier, link.parent_supplier_id)
                totals[route] = totals.get(route, 0) + link.sob_percent
            over = [route for route, total in totals.items() if total > 100.0001]
            if over:
                material = db.get(Item, over[0][0]) if over[0][0] else db.get(TierMaterial, over[0][1])
                db.rollback()
                raise HTTPException(status_code=422, detail=f"SOB for {material.code} ({over[0][2]}) adds up to more than 100% after this upload")
            db.commit()
            return {"rows": len(rows), "added": added, "updated": updated, "suppliers_created": created}

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
            material = TierMaterial(tier=payload.tier, code=code, name=payload.name.strip(), category=_ensure_category(db, payload.category), uom=payload.uom.strip(),
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
            material.name, material.category, material.uom = payload.name.strip(), _ensure_category(db, payload.category), payload.uom.strip()
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
            if db.scalar(select(SupplierPart.id).where(SupplierPart.tier_material_id == material_id).limit(1)):
                raise HTTPException(status_code=409, detail=f"{material.code} is used in a supplier/material relationship; remove it there first")
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
                category = _ensure_category(db, str(row.get("category") or "").strip() or "Raw material")
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

    # ------------------------------------------------------------------ alternate codes
    def check_alternate(db: Session, main: Item | None, alternate: Item | None, existing_id: int | None = None) -> None:
        """Rules for an alternate code: both in T1 Material Master, both bought parts, same unit,
        one level only (an alternate has no alternates of its own) and one main code per alternate."""
        if not main or not alternate:
            raise HTTPException(status_code=404, detail="Main or alternate code not found in T1 Material Master")
        if main.id == alternate.id:
            raise HTTPException(status_code=422, detail=f"{main.code} cannot be its own alternate")
        if main.uom != alternate.uom:
            raise HTTPException(status_code=422, detail=f"{alternate.code} is in {alternate.uom} but {main.code} is in {main.uom}; an alternate must use the same unit")
        for item in (main, alternate):
            reason = not_purchased_reason(db, item.id)
            if reason:
                raise HTTPException(status_code=422, detail=f"{item.code} {reason}; alternate codes are only for parts you buy")
        others = select(AlternateCode).where(AlternateCode.id != existing_id) if existing_id else select(AlternateCode)
        taken = db.scalar(others.where(AlternateCode.alternate_item_id == alternate.id).limit(1))
        if taken:
            owner = db.get(Item, taken.main_item_id)
            raise HTTPException(status_code=422, detail=f"{alternate.code} is already an alternate of {owner.code if owner else 'another code'}")
        if db.scalar(others.where(AlternateCode.main_item_id == alternate.id).limit(1)):
            raise HTTPException(status_code=422, detail=f"{alternate.code} has alternate codes of its own, so it cannot be an alternate")
        if db.scalar(others.where(AlternateCode.alternate_item_id == main.id).limit(1)):
            raise HTTPException(status_code=422, detail=f"{main.code} is itself an alternate code, so it cannot have alternates")

    def alternate_rows(db: Session, entries: list[AlternateCode], can_price: bool) -> list[dict]:
        ids = {entry.main_item_id for entry in entries} | {entry.alternate_item_id for entry in entries}
        items = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_(ids))).all()} if ids else {}
        price = bom_price_lookup(db, {entry.main_item_id for entry in entries}) if can_price else None
        mains = list({entry.main_item_id for entry in entries})
        use_counts: dict[int, dict[str, int]] = {}
        if mains:
            level1 = set(db.scalars(select(ModelBOM.material_id).distinct()).all())
            for material_id in db.scalars(select(ModelBOM.material_id).where(ModelBOM.material_id.in_(mains))).all():
                use_counts.setdefault(material_id, {}).setdefault("L1", 0)
                use_counts[material_id]["L1"] += 1
            for child_id, parent_id in db.execute(select(BOMComponent.child_item_id, BOMComponent.parent_item_id).where(BOMComponent.child_item_id.in_(mains))).all():
                level = "L2" if parent_id in level1 else "L3"
                use_counts.setdefault(child_id, {}).setdefault(level, 0)
                use_counts[child_id][level] += 1
        rows = []
        for entry in entries:
            main, alternate = items.get(entry.main_item_id), items.get(entry.alternate_item_id)
            if not main or not alternate:
                continue
            row = {"id": entry.id, "main_item_id": main.id, "main_code": main.code, "main_name": main.name,
                   "alternate_item_id": alternate.id, "alternate_code": alternate.code,
                   "alternate_name": alternate.name, "uom": main.uom, "uses": use_counts.get(main.id, {})}
            if price:
                main_price, alternate_price = price.own(main), price.own(alternate)
                row.update(main_unit_price=main_price["unit_price"], main_currency=main_price["currency"],
                           unit_price=alternate_price["unit_price"], currency=alternate_price["currency"])
            rows.append(row)
        return rows

    def bom_line(db: Session, model_bom_id: int | None, component_id: int | None):
        if (model_bom_id is None) == (component_id is None):
            raise HTTPException(status_code=422, detail="Give either model_bom_id or component_id")
        line = db.get(ModelBOM, model_bom_id) if model_bom_id is not None else db.get(BOMComponent, component_id)
        if not line:
            raise HTTPException(status_code=404, detail="BOM line not found")
        main = db.get(Item, line.material_id if model_bom_id is not None else line.child_item_id)
        return line, main

    def line_history_rows(db: Session, model_bom_id: int | None, component_id: int | None, main: Item) -> list[dict]:
        """SOB history of one line, one entry per change month (oldest first)."""
        allowed = [a.alternate_item_id for a in alternate_groups(db, [main.id]).get(main.id, [])]
        column = BOMLineAlternate.model_bom_id if model_bom_id is not None else BOMLineAlternate.component_id
        entries = [(e.from_month or "", e.alternate_item_id, e.sob_percent) for e in db.scalars(
            select(BOMLineAlternate).where(column == (model_bom_id if model_bom_id is not None else component_id),
                                           BOMLineAlternate.main_item_id == main.id)).all() if e.alternate_item_id in allowed]
        rows = []
        for from_month in sorted({entry[0] for entry in entries}):
            state = _sob_at(entries, from_month)
            values = {alternate_id: state.get(alternate_id, 0.0) for alternate_id in allowed}
            rows.append({"from_month": from_month, "alternates": [{"alternate_item_id": k, "sob_percent": v} for k, v in values.items()],
                         "main_sob_percent": 100 - sum(values.values())})
        return rows

    @app.get("/api/planning/bom-line-alternates")
    def get_line_alternates(request: Request, model_bom_id: int | None = None, component_id: int | None = None):
        """Alternate SOB history of one BOM line: what changed and from which month."""
        current_user(request)
        with open_db() as db:
            _, main = bom_line(db, model_bom_id, component_id)
            return {"main_code": main.code, "history": line_history_rows(db, model_bom_id, component_id, main)}

    def code_uses(db: Session, main: Item) -> list[dict]:
        """Every BOM line using a code: L1 direct lines (per model) and Level 2 / Level 3 part BOM lines."""
        uses = []
        for root, model in db.execute(select(ModelBOM, ProductModel).join(ProductModel, ProductModel.id == ModelBOM.model_id)
                                      .where(ModelBOM.material_id == main.id).order_by(ProductModel.code)).all():
            uses.append({"key": f"root-{root.id}", "model_bom_id": root.id, "component_id": None, "level": "L1",
                         "used_in": f"Model {model.code} · {model.name}", "quantity": root.quantity})
        level1 = set(db.scalars(select(ModelBOM.material_id).distinct()).all())
        for link, parent in db.execute(select(BOMComponent, Item).join(Item, Item.id == BOMComponent.parent_item_id)
                                       .where(BOMComponent.child_item_id == main.id).order_by(Item.code)).all():
            level = "L2" if parent.id in level1 else "L3"
            uses.append({"key": f"component-{link.id}", "model_bom_id": None, "component_id": link.id, "level": level,
                         "used_in": f"{parent.code} · {parent.name} BOM", "quantity": link.quantity})
        uses.sort(key=lambda use: (use["level"], use["used_in"]))
        return uses

    def sob_text(values: dict[int, float], codes: dict[int, str]) -> str:
        parts = [f"{codes.get(alternate_id, alternate_id)} {sob:g}%" for alternate_id, sob in values.items()]
        return ", ".join(parts + [f"main {100 - sum(values.values()):g}%"])

    def write_sob_change(db: Session, request: Request, main: Item, line_keys: list[tuple[str, int]], values: dict[int, float],
                         from_month: str, action: str) -> int:
        """Set (or remove) the SOB change of from_month on BOM lines of the main code, and log it."""
        uses = {(("root", use["model_bom_id"]) if use["model_bom_id"] else ("component", use["component_id"])): use for use in code_uses(db, main)}
        for key in line_keys:
            if key not in uses:
                raise HTTPException(status_code=422, detail=f"A selected BOM line does not use {main.code}; refresh and try again")
        codes = dict(db.execute(select(Item.id, Item.code).where(Item.id.in_(list(values) or [-1]))).all())
        user = request.state.user
        changed = 0
        for kind, line_id in line_keys:
            column = BOMLineAlternate.model_bom_id if kind == "root" else BOMLineAlternate.component_id
            removed = db.execute(delete(BOMLineAlternate).where(column == line_id, BOMLineAlternate.from_month == from_month)).rowcount
            if action == "set":
                for alternate_id, sob in values.items():
                    db.add(BOMLineAlternate(model_bom_id=line_id if kind == "root" else None, component_id=line_id if kind == "component" else None,
                                            main_item_id=main.id, alternate_item_id=alternate_id, sob_percent=sob, from_month=from_month))
            elif not removed:
                continue
            use = uses[(kind, line_id)]
            db.add(SOBChangeLog(changed_at=datetime.now(), changed_by=user.get("full_name") or user["username"],
                                action="Set" if action == "set" else "Removed", main_item_id=main.id, main_code=main.code,
                                level=use["level"], used_in=use["used_in"], from_month=from_month,
                                sob_text=sob_text(values, codes) if action == "set" else ""))
            changed += 1
        return changed

    def checked_values(db: Session, main: Item, alternates: list[LineAlternateSOB]) -> dict[int, float]:
        allowed = [a.alternate_item_id for a in alternate_groups(db, [main.id]).get(main.id, [])]
        seen: set[int] = set()
        for entry in alternates:
            if entry.alternate_item_id not in allowed:
                other = db.get(Item, entry.alternate_item_id)
                raise HTTPException(status_code=422, detail=f"{other.code if other else 'This code'} is not an alternate of {main.code}; add it in Product Master → Alternate codes first")
            if entry.alternate_item_id in seen:
                raise HTTPException(status_code=422, detail="An alternate code is listed twice")
            seen.add(entry.alternate_item_id)
        total = sum(entry.sob_percent for entry in alternates)
        if total > 100 + 1e-9:
            raise HTTPException(status_code=422, detail=f"Alternate SOB adds up to {total:g}%; the total cannot be more than 100%")
        values = {alternate_id: 0.0 for alternate_id in allowed}
        values.update({entry.alternate_item_id: entry.sob_percent for entry in alternates})
        return values

    def clean_month(value: str) -> str:
        value = (value or "").strip()
        return _month_text(_month(value)) if value else ""

    @app.get("/api/planning/alternate-codes/uses")
    def alternate_code_uses(request: Request, main_item_id: int, month: str | None = None):
        """Where a main code is used (L1 / L2 / L3 BOM lines), with each line's SOB this month,
        the next change and the full history: the Manage SOB view of Alternate codes."""
        current_user(request)
        month = clean_month(month) if month else _month_text(date.today())
        with open_db() as db:
            main = db.get(Item, main_item_id)
            if not main:
                raise HTTPException(status_code=404, detail="Main code not found")
            price = bom_price_lookup(db, [main.id])
            alternates = price.alternates(main.id)
            if not alternates:
                raise HTTPException(status_code=404, detail=f"{main.code} has no alternate codes")
            uses = code_uses(db, main)
            saved = line_sobs(db, [u["model_bom_id"] for u in uses if u["model_bom_id"]], [u["component_id"] for u in uses if u["component_id"]])
            for use in uses:
                key = ("root", use["model_bom_id"]) if use["model_bom_id"] else ("component", use["component_id"])
                listed, main_share, next_change = line_alternates(price, main, saved.get(key, []), use["quantity"], month)
                use["sob"] = {a["material_id"]: a["sob_percent"] for a in listed}
                use["main_sob_percent"] = main_share
                use["next_change"] = next_change
                use["history"] = line_history_rows(db, use["model_bom_id"], use["component_id"], main)
            return {"main": {"id": main.id, "code": main.code, "name": main.name, "uom": main.uom}, "month": month,
                    "alternates": [{"id": a["material_id"], "code": a["code"], "name": a["name"]} for a in alternates], "uses": uses}

    @app.post("/api/planning/alternate-sob")
    def save_alternate_sob(payload: AlternateSOBInput, request: Request):
        """Set the alternates' SOB from a month on the ticked uses of the main code (or remove
        that month's change from them). Logged in the SOB change log."""
        current_user(request)
        _check_supply_role(request)
        if payload.action not in {"set", "remove"}:
            raise HTTPException(status_code=422, detail="action must be set or remove")
        if not payload.lines:
            raise HTTPException(status_code=422, detail="Tick at least one use (BOM line)")
        from_month = clean_month(payload.from_month)
        with open_db() as db:
            main = db.get(Item, payload.main_item_id)
            if not main:
                raise HTTPException(status_code=404, detail="Main code not found")
            values = checked_values(db, main, payload.alternates) if payload.action == "set" else {}
            keys = []
            for line in payload.lines:
                if (line.model_bom_id is None) == (line.component_id is None):
                    raise HTTPException(status_code=422, detail="Each line needs either model_bom_id or component_id")
                keys.append(("root", line.model_bom_id) if line.model_bom_id is not None else ("component", line.component_id))
            changed = write_sob_change(db, request, main, keys, values, from_month, payload.action)
            if payload.action == "remove" and not changed:
                raise HTTPException(status_code=404, detail=f"None of the ticked uses has an SOB change from {from_month or 'the start'}")
            db.commit()
            return {"main_code": main.code, "from_month": from_month, "lines": changed,
                    "main_sob_percent": 100 - sum(values.values()) if values else None}

    @app.put("/api/planning/bom-line-alternates")
    def save_line_alternates(payload: LineAlternatesInput, request: Request):
        """One BOM line's SOB from a month (or every line using the code); same as Manage SOB."""
        current_user(request)
        _check_supply_role(request)
        from_month = clean_month(payload.from_month)
        with open_db() as db:
            _, main = bom_line(db, payload.model_bom_id, payload.component_id)
            values = checked_values(db, main, payload.alternates)
            if payload.apply_everywhere:
                keys = [("root", u["model_bom_id"]) if u["model_bom_id"] else ("component", u["component_id"]) for u in code_uses(db, main)]
            else:
                keys = [("root", payload.model_bom_id) if payload.model_bom_id is not None else ("component", payload.component_id)]
            changed = write_sob_change(db, request, main, keys, values, from_month, "set")
            db.commit()
            return {"main_code": main.code, "main_sob_percent": 100 - sum(values.values()), "from_month": from_month, "lines": changed,
                    "history": line_history_rows(db, payload.model_bom_id, payload.component_id, main)}

    @app.delete("/api/planning/bom-line-alternates", status_code=204)
    def delete_line_alternate_change(request: Request, from_month: str, model_bom_id: int | None = None, component_id: int | None = None):
        """Remove one SOB change (one from-month) from a BOM line's history."""
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            _, main = bom_line(db, model_bom_id, component_id)
            key = ("root", model_bom_id) if model_bom_id is not None else ("component", component_id)
            if not write_sob_change(db, request, main, [key], {}, from_month, "remove"):
                raise HTTPException(status_code=404, detail="No SOB change from this month on this line")
            db.commit()
            return Response(status_code=204)

    def sob_log_rows(db: Session, search: str = "") -> Any:
        statement = select(SOBChangeLog)
        if search.strip():
            needle = f"%{search.strip().lower()}%"
            statement = statement.where(or_(func.lower(SOBChangeLog.main_code).like(needle), func.lower(SOBChangeLog.used_in).like(needle),
                                            func.lower(SOBChangeLog.sob_text).like(needle), func.lower(SOBChangeLog.changed_by).like(needle)))
        return statement

    def sob_log_json(entry: SOBChangeLog) -> dict:
        return {"id": entry.id, "changed_at": entry.changed_at.strftime("%Y-%m-%d %H:%M"), "changed_by": entry.changed_by,
                "action": entry.action, "main_code": entry.main_code, "level": entry.level, "used_in": entry.used_in,
                "from_month": entry.from_month, "sob_text": entry.sob_text}

    @app.get("/api/planning/sob-change-log")
    def sob_change_log(request: Request, search: str = "", offset: int = 0, limit: int = 500):
        current_user(request)
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        with open_db() as db:
            statement = sob_log_rows(db, search)
            total = db.scalar(select(func.count()).select_from(statement.subquery()))
            entries = db.scalars(statement.order_by(SOBChangeLog.id.desc()).offset(offset).limit(limit)).all()
            return {"total": total, "offset": offset, "limit": limit, "items": [sob_log_json(e) for e in entries]}

    @app.get("/api/planning/sob-change-log.{file_format}")
    def export_sob_change_log(file_format: str, request: Request, search: str = ""):
        current_user(request)
        if file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export not found")
        with open_db() as db:
            rows = [sob_log_json(e) for e in db.scalars(sob_log_rows(db, search).order_by(SOBChangeLog.id.desc())).all()]
        headers = ["Changed on", "Changed by", "Action", "Main code", "Level", "Used in", "SOB from month", "SOB split"]
        values = [[r["changed_at"], r["changed_by"], r["action"], r["main_code"], r["level"], r["used_in"],
                   r["from_month"] or "Start (all months)", r["sob_text"]] for r in rows]
        output = _generic_workbook("SOB change log", headers, values) if file_format == "xlsx" else _generic_pdf("SOB change log", headers, values)
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_format == "xlsx" else "application/pdf"
        return StreamingResponse(output, media_type=media, headers={"Content-Disposition": f'attachment; filename="sob-change-log.{file_format}"'})

    @app.get("/api/planning/alternate-codes")
    def list_alternate_codes(request: Request, search: str = "", offset: int = 0, limit: int = 500):
        current_user(request)
        if limit < 1 or limit > 1000 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–1000 and a non-negative offset")
        with open_db() as db:
            main_item, alternate_item = Item.__table__.alias("main_item"), Item.__table__.alias("alternate_item")
            statement = (select(AlternateCode)
                         .join(main_item, main_item.c.id == AlternateCode.main_item_id)
                         .join(alternate_item, alternate_item.c.id == AlternateCode.alternate_item_id))
            if search.strip():
                needle = f"%{search.strip().lower()}%"
                statement = statement.where(or_(func.lower(main_item.c.code).like(needle), func.lower(main_item.c.name).like(needle),
                                                func.lower(alternate_item.c.code).like(needle), func.lower(alternate_item.c.name).like(needle)))
            total = db.scalar(select(func.count()).select_from(statement.subquery()))
            entries = db.scalars(statement.order_by(main_item.c.code, AlternateCode.id).offset(offset).limit(limit)).all()
            return {"total": total, "offset": offset, "limit": limit,
                    "items": alternate_rows(db, list(entries), _check_price_access(request))}

    @app.post("/api/planning/alternate-codes", status_code=201)
    def create_alternate_code(payload: AlternateCodeInput, request: Request):
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            check_alternate(db, db.get(Item, payload.main_item_id), db.get(Item, payload.alternate_item_id))
            # SOB is set on each BOM line (0% until then)
            entry = AlternateCode(main_item_id=payload.main_item_id, alternate_item_id=payload.alternate_item_id, sob_percent=0)
            db.add(entry)
            db.commit()
            return alternate_rows(db, [entry], _check_price_access(request))[0]

    @app.delete("/api/planning/alternate-codes/{alternate_id}", status_code=204)
    def delete_alternate_code(alternate_id: int, request: Request):
        current_user(request)
        _check_supply_role(request)
        with open_db() as db:
            entry = db.get(AlternateCode, alternate_id)
            if not entry:
                raise HTTPException(status_code=404, detail="Alternate code not found")
            # its SOB on BOM lines goes back to the main code
            db.execute(delete(BOMLineAlternate).where(BOMLineAlternate.main_item_id == entry.main_item_id,
                                                      BOMLineAlternate.alternate_item_id == entry.alternate_item_id))
            db.delete(entry)
            db.commit()
            return Response(status_code=204)

    @app.post("/api/planning/alternate-codes/upload", status_code=201)
    async def upload_alternate_codes(request: Request):
        """Alternate codes from Excel (heading row found automatically): Main code, Alternate
        code. Pairs already there are left as they are; SOB is set on each BOM line. All rows or none."""
        current_user(request)
        _check_supply_role(request)
        rows = _read_flexible_sheet(await request.body(), ALTERNATE_CODE_COLUMN_ALIASES, {"main_code", "alternate_code"})
        with open_db() as db:
            seen: set[tuple[str, str]] = set()
            added = existing_pairs = 0
            for row in rows:
                row_number = row["_excel_row"]
                main_code, alternate_code = str(row.get("main_code") or "").strip(), str(row.get("alternate_code") or "").strip()
                if not main_code or not alternate_code:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: main code and alternate code are required")
                key = (main_code.lower(), alternate_code.lower())
                if key in seen:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {main_code} / {alternate_code} is repeated in the file")
                seen.add(key)
                main = db.scalar(select(Item).where(func.lower(Item.code) == key[0]))
                alternate = db.scalar(select(Item).where(func.lower(Item.code) == key[1]))
                if not main or not alternate:
                    missing = main_code if not main else alternate_code
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {missing} is not in T1 Material Master")
                existing = db.scalar(select(AlternateCode).where(AlternateCode.main_item_id == main.id, AlternateCode.alternate_item_id == alternate.id))
                try:
                    check_alternate(db, main, alternate, existing.id if existing else None)
                except HTTPException as exc:
                    raise HTTPException(status_code=exc.status_code, detail=f"Row {row_number}: {exc.detail}") from exc
                if existing:
                    existing_pairs += 1
                else:
                    db.add(AlternateCode(main_item_id=main.id, alternate_item_id=alternate.id, sob_percent=0))
                    added += 1
                db.flush()
            db.commit()
            return {"rows": len(rows), "added": added, "already_there": existing_pairs}

    @app.get("/api/planning/alternate-codes.{file_format}")
    def export_alternate_codes(file_format: str, request: Request):
        current_user(request)
        if file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export not found")
        can_price = _check_price_access(request)
        with open_db() as db:
            rows = alternate_rows(db, list(db.scalars(select(AlternateCode)).all()), can_price)
        rows.sort(key=lambda row: (row["main_code"], row["alternate_code"]))
        headers = ["Main code", "Main name", "Alternate code", "Alternate name", "Unit"]
        if can_price:
            headers += ["Main price", "Main currency", "Alternate price", "Alternate currency"]
        values = [[r["main_code"], r["main_name"], r["alternate_code"], r["alternate_name"], r["uom"]]
                  + ([r["main_unit_price"], r["main_currency"], r["unit_price"], r["currency"]] if can_price else []) for r in rows]
        output = _generic_workbook("Alternate codes", headers, values) if file_format == "xlsx" else _generic_pdf("Alternate codes", headers, values)
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if file_format == "xlsx" else "application/pdf"
        return StreamingResponse(output, media_type=media, headers={"Content-Disposition": f'attachment; filename="alternate-codes.{file_format}"'})

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
            SupplierPart.tier == TOP_TIER, SupplierPart.parent_supplier_id.is_(None)))
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
                                "unit_price": src.unit_price, "currency": src.currency, "sob_percent": src.sob_percent,
                                "sub_supplier_code": subs[src.sub_supplier_id].code if src.sub_supplier_id in subs else "",
                                "lead_time_days": src.lead_time_days, "moq": src.moq, "buffer_days": src.buffer_days, "price_mode": src.price_mode,
                                "price_interval_months": src.price_interval_months, "price_status": _price_status(src, _month_text(date.today()))} for src in sources.get(line.id, [])]}
            rows.append(row)
            kids = children.get(line.id, [])
            if sources.get(line.id) and len({src.currency for src in sources[line.id]}) == 1:
                for kid in kids:  # its BOM (e.g. steel from ore + coal) is listed; the price is what its suppliers charge
                    unit_cost(kid, depth + 1)
                group = sources[line.id]
                share = sum(src.sob_percent for src in group)
                price = sum(src.unit_price * src.sob_percent for src in group) / share
                currency, source = group[0].currency, f"SOB-weighted ({len(group)} source{'s' if len(group) > 1 else ''})"
            elif kids:
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
        """Parts you buy from this supplier (T1 relationships), with BOM flag."""
        current_user(request)
        with open_db() as db:
            links = db.scalars(select(SupplierPart).where(SupplierPart.supplier_id == supplier_id, SupplierPart.tier == TOP_TIER,
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
            for number, src in enumerate(payload, start=1):
                if src.sub_supplier_id is not None and not db.get(Supplier, src.sub_supplier_id):
                    raise HTTPException(status_code=422, detail=f"Source {number}: sub-supplier not found in the Supplier register")
                if src.item_id or src.supplier_material_id:
                    check_component(db, line.supplier_id, src.item_id, src.supplier_material_id, f"Source {number} alternate")
            before = {(old.sub_supplier_id, old.unit_price, old.price_mode): old.price_updated_month
                      for old in db.scalars(select(SupplierBOMLineSource).where(SupplierBOMLineSource.line_id == line_id)).all()}
            db.execute(delete(SupplierBOMLineSource).where(SupplierBOMLineSource.line_id == line_id))
            for src in payload:
                kept = before.get((src.sub_supplier_id, src.unit_price, src.price_mode))  # price unchanged: keep its month
                db.add(SupplierBOMLineSource(line_id=line_id, **src.model_dump(), **({"price_updated_month": kept} if kept else {})))
            db.commit()
            return {"line_id": line_id, "sources": len(payload)}

    def styled_sheet(sheet, title: str, note: str, headers: list[str]) -> None:
        sheet.append([title])
        sheet.append([note])
        sheet.append(headers)
        sheet["A1"].font = Font(bold=True, size=13)
        sheet["A2"].font = Font(italic=True, color="5B6B7B")
        for cell in sheet[3]:
            cell.fill, cell.font = PatternFill("solid", fgColor="12263A"), Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for index in range(1, len(headers) + 1):
            sheet.column_dimensions[__import__("openpyxl").utils.get_column_letter(index)].width = 16
        sheet.freeze_panes = "A4"

    def xlsx_response(workbook, filename: str) -> StreamingResponse:
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 headers={"Content-Disposition": f'attachment; filename="{filename}"'})

    @app.get("/api/planning/supplier-relationships/template.xlsx")
    def relationship_template(request: Request, supplier_id: int | None = None):
        """The relationship page's columns; with supplier_id, that supplier's current T1 rows (edit and upload back)."""
        current_user(request)
        _check_supply_role(request)
        can_price = _check_price_access(request)
        with open_db() as db:
            supplier = db.get(Supplier, supplier_id) if supplier_id else None
            statement = select(SupplierPart).where(SupplierPart.tier == TOP_TIER, SupplierPart.material_id.is_not(None))
            if supplier_id:
                statement = statement.where(SupplierPart.supplier_id == supplier_id)
            links = db.scalars(statement.order_by(SupplierPart.supplier_id, SupplierPart.material_id)).all() if supplier_id else []
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Relationships"
            styled_sheet(sheet, f"Supplier/material relationship{f' · {supplier.code} {supplier.name}' if supplier else ''}",
                         "One row per part and supplier. Needed: Supplier Code (or Supplier Name) and T1 Part Code. SOB blank = 100. A row already there is updated.",
                         RELATIONSHIP_SHEET_COLUMNS)
            for link in links:
                item, who = db.get(Item, link.material_id), db.get(Supplier, link.supplier_id)
                sheet.append([who.code, who.name, item.code, item.name, link.sob_percent, link.lead_time_days, link.moq, link.buffer_days,
                              link.unit_price if can_price else None, link.currency])
            if not links and supplier:
                sheet.append([supplier.code, supplier.name])
        return xlsx_response(workbook, f"supplier-relationships{f'-{supplier.code}' if supplier else ''}.xlsx")

    @app.get("/api/planning/supplier-relationships.{file_format}")
    def export_relationships(file_format: str, request: Request, supplier_id: int | None = None, search: str = ""):
        """The relationship page's list (one supplier or all, with its search) in the upload's columns, so the
        Excel can be edited and uploaded back; or PDF."""
        if file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export not found")
        current_user(request)
        _check_supply_role(request)
        can_price = _check_price_access(request)
        needle = search.strip().casefold()
        with open_db() as db:
            supplier = db.get(Supplier, supplier_id) if supplier_id else None
            statement = select(SupplierPart).where(SupplierPart.tier == TOP_TIER, SupplierPart.material_id.is_not(None))
            if supplier_id:
                statement = statement.where(SupplierPart.supplier_id == supplier_id)
            suppliers = {who.id: who for who in db.scalars(select(Supplier)).all()}
            links = db.scalars(statement).all()
            items = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_({link.material_id for link in links})))} if links else {}
            values = []
            for link in sorted(links, key=lambda l: (suppliers[l.supplier_id].code.casefold(), items[l.material_id].code.casefold())):
                who, item = suppliers[link.supplier_id], items[link.material_id]
                if needle and needle not in f"{who.code} {who.name} {item.code} {item.name}".casefold():
                    continue
                values.append([who.code, who.name, item.code, item.name, link.sob_percent, link.lead_time_days, link.moq, link.buffer_days,
                               link.unit_price if can_price else None, link.currency])
        title = f"Supplier/material relationship{f' · {supplier.code} {supplier.name}' if supplier else ' · all suppliers'}"
        filename = f"supplier-relationships{f'-{supplier.code}' if supplier else ''}"
        headers = RELATIONSHIP_SHEET_COLUMNS if can_price else [h for h in RELATIONSHIP_SHEET_COLUMNS if h != "Unit price"]
        if not can_price:
            values = [row[:8] + row[9:] for row in values]
        if file_format == "pdf":
            output = _generic_pdf(title, headers, values)
            return StreamingResponse(output, media_type="application/pdf", headers={"Content-Disposition": f'attachment; filename="{filename}.pdf"'})
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Relationships"
        styled_sheet(sheet, title, "Same columns as the upload: edit and upload back (a row already there is updated).", headers)
        for row in values:
            sheet.append(row)
        return xlsx_response(workbook, f"{filename}.xlsx")

    def own_material_for(db: Session, supplier_id: int, item_id: int | None, tier_material_id: int | None) -> SupplierMaterial:
        """The supplier's material master entry for a T1 / T2 / T3 code (added when missing)."""
        column = SupplierMaterial.item_id if item_id else SupplierMaterial.tier_material_id
        own = db.scalar(select(SupplierMaterial).where(SupplierMaterial.supplier_id == supplier_id, column == (item_id or tier_material_id)))
        if own:
            return own
        master = db.get(Item, item_id) if item_id else db.get(TierMaterial, tier_material_id)
        if not master:
            raise HTTPException(status_code=404, detail="Material code not found")
        if db.scalar(select(SupplierMaterial.id).where(SupplierMaterial.supplier_id == supplier_id, func.lower(SupplierMaterial.code) == master.code.lower())):
            raise HTTPException(status_code=409, detail=f"{master.code} is already a code of the supplier's own material master")
        own = SupplierMaterial(supplier_id=supplier_id, code=master.code, name=master.name, uom=master.uom or "ea",
                               unit_price=getattr(master, "unit_cost", 0) or 0, currency=getattr(master, "currency", None) or "INR",
                               item_id=item_id, tier_material_id=tier_material_id)
        db.add(own)
        db.flush()
        return own

    @app.get("/api/planning/coverage-workbook/template.xlsx")
    def coverage_template(request: Request, supplier_id: int):
        """Supplier-wise coverage workbook: each T1 part the supplier supplies, its BOM lines (L1-L3, quantity per
        unit of the line above) and the suppliers of each line (T2-T4) - filled with what is saved now."""
        current_user(request)
        _check_supply_role(request)
        can_price = _check_price_access(request)
        with open_db() as db:
            supplier = db.get(Supplier, supplier_id)
            if not supplier:
                raise HTTPException(status_code=404, detail="Supplier not found")
            links = db.scalars(select(SupplierPart).where(SupplierPart.supplier_id == supplier_id, SupplierPart.tier == TOP_TIER,
                                                          SupplierPart.material_id.is_not(None))).all()
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Coverage"
            styled_sheet(sheet, f"Supplier/material coverage · {supplier.code} {supplier.name}",
                         "Level L1/L2/L3 = a material of the BOM (BOM qty per 1 unit of the line above; L1 per 1 part). Level T2/T3/T4 = a supplier of the line above it "
                         "(T2 under L1, T3 under L2, T4 under L3) with SOB. T1 rows are for reading only. Uploading replaces the BOM of each part in the file.",
                         COVERAGE_SHEET_COLUMNS)
            mode_label = {"spot": "spot", "fixed_monthly": "fixed_monthly", "every_n_months": "every_n_months"}
            for link in sorted(links, key=lambda l: db.get(Item, l.material_id).code):
                item = db.get(Item, link.material_id)
                sheet.append([item.code, item.name, "T1", item.code, item.name, 1, item.uom, supplier.code, supplier.name, link.sob_percent, link.lead_time_days,
                              link.moq, link.buffer_days, link.price_mode, link.price_interval_months, link.unit_price if can_price else None, link.currency])
                view = supplier_bom_view(db, supplier_id, link.material_id, can_price)
                for line in view["lines"]:
                    component = line["component"]
                    sheet.append([item.code, item.name, f"L{line['depth']}", component["code"], component["name"], line["quantity"], component["uom"]])
                    for src in line["sources"]:
                        sheet.append([item.code, item.name, f"T{line['depth'] + 1}", component["code"], component["name"], None, component["uom"], src["sub_supplier_code"],
                                      src["sub_supplier"], src["sob_percent"], src["lead_time_days"], src["moq"], src["buffer_days"],
                                      mode_label.get(src["price_mode"], "spot"), src["price_interval_months"], src.get("unit_price") if can_price else None, src["currency"]])
        return xlsx_response(workbook, f"supplier-coverage-{supplier.code}.xlsx")

    @app.post("/api/planning/coverage-workbook/upload", status_code=201)
    async def upload_coverage(request: Request, supplier_id: int):
        """Replace the BOM (lines and their suppliers) of every T1 part in the file, for this supplier. All or nothing."""
        current_user(request)
        _check_supply_role(request)
        can_price = _check_price_access(request)
        rows = _read_supplier_part_sheet(await request.body(), COVERAGE_COLUMN_ALIASES, {"part_code", "level", "material_code"})
        with open_db() as db:
            supplier = db.get(Supplier, supplier_id)
            if not supplier:
                raise HTTPException(status_code=404, detail="Supplier not found")
            links = {db.get(Item, link.material_id).code.casefold(): link for link in db.scalars(select(SupplierPart).where(
                SupplierPart.supplier_id == supplier_id, SupplierPart.tier == TOP_TIER, SupplierPart.material_id.is_not(None))).all()}
            items = {item.code.casefold(): item for item in db.scalars(select(Item)).all()}
            tier_codes = {code.code.casefold(): code for code in db.scalars(select(TierMaterial)).all()}
            owns = {m.code.casefold(): m for m in db.scalars(select(SupplierMaterial).where(SupplierMaterial.supplier_id == supplier_id)).all()}
            register = {}
            for who in db.scalars(select(Supplier)).all():
                register[who.code.casefold()] = who
                register.setdefault(who.name.strip().casefold(), who)

            def text(row: dict, key: str) -> str:
                value = row.get(key)
                if isinstance(value, float) and value.is_integer():
                    value = int(value)
                return "" if value is None else str(value).strip()

            def number(row: dict, key: str, default: float, row_number: int) -> float:
                raw = text(row, key)
                if raw == "":
                    return default
                try:
                    value = float(raw.rstrip("%"))
                except ValueError as exc:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {key.replace('_', ' ')} must be a number") from exc
                if not math.isfinite(value) or value < 0:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {key.replace('_', ' ')} must be zero or more")
                return value

            plan: dict[int, list[dict]] = {}  # link id -> lines (each with sources)
            part_code = ""
            last_line: dict[tuple[int, int], dict] = {}  # (link id, depth) -> last line at that depth
            for row in rows:
                row_number = row["_excel_row"]
                part_code = text(row, "part_code") or part_code
                level = text(row, "level").upper().replace(" ", "").replace("BOM", "")
                if not part_code:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: T1 Part Code is missing")
                link = links.get(part_code.casefold())
                if not link:
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: {part_code} is not a part {supplier.name} supplies; add it in Supplier/material relationship first")
                lines = plan.setdefault(link.id, [])
                if level == "T1":
                    continue  # for reading only
                match = re.fullmatch(r"([LT])([1-4])", level)
                if not match or (match.group(1) == "L" and match.group(2) == "4") or (match.group(1) == "T" and match.group(2) == "1"):
                    raise HTTPException(status_code=422, detail=f"Row {row_number}: Level must be L1, L2, L3 (a material) or T2, T3, T4 (its supplier)")
                depth = int(match.group(2))
                code = text(row, "material_code")
                if match.group(1) == "L":
                    if not code:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: Material code is missing")
                    quantity = number(row, "quantity", 0, row_number)
                    if quantity <= 0:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: BOM qty must be more than 0")
                    parent = None
                    if depth > 1:
                        parent = last_line.get((link.id, depth - 1))
                        if parent is None:
                            raise HTTPException(status_code=422, detail=f"Row {row_number}: an L{depth} line must come after the L{depth - 1} line it belongs to")
                    found = items.get(code.casefold()) or tier_codes.get(code.casefold()) or owns.get(code.casefold())
                    if not found:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: {code} is not in the T1, T2 or T3 Material Master")
                    siblings = parent["children"] if parent else lines
                    if any(other["code"].casefold() == code.casefold() for other in siblings):
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: {code} is repeated at the same place in the BOM of {part_code}")
                    line = {"code": code, "found": found, "quantity": quantity, "sources": [], "children": [], "row": row_number}
                    siblings.append(line)
                    last_line[(link.id, depth)] = line
                    for deeper in range(depth + 1, 4):
                        last_line.pop((link.id, deeper), None)
                else:
                    line = last_line.get((link.id, depth - 1))
                    if line is None:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: a T{depth} supplier must come after the L{depth - 1} material it supplies")
                    if code and code.casefold() != line["code"].casefold():
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: a T{depth} row supplies the line above ({line['code']}); its material code must be {line['code']} or blank")
                    who_key = text(row, "supplier_code") or text(row, "supplier_name")
                    who = register.get(who_key.casefold()) if who_key else None
                    if who is None:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: supplier {who_key or '(blank)'} is not in the Supplier register" if who_key else f"Row {row_number}: Supplier Code is missing")
                    if any(src["supplier"].id == who.id for src in line["sources"]):
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: {who.name} is repeated as a supplier of {line['code']}")
                    sob = number(row, "sob_percent", 100, row_number)
                    if sob <= 1 and text(row, "sob_percent") and "%" not in text(row, "sob_percent") and sob > 0 and isinstance(row.get("sob_percent"), float):
                        sob *= 100  # Excel percent cell (0.3 = 30%)
                    if not 0 < sob <= 100:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: SOB must be more than 0 and at most 100")
                    if sum(src["sob_percent"] for src in line["sources"]) + sob > 100.0001:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: SOB of the suppliers of {line['code']} adds up to more than 100%")
                    price = number(row, "unit_price", 0, row_number)
                    if price and not can_price:
                        raise HTTPException(status_code=403, detail="Price access is required to upload prices")
                    mode = text(row, "price_mode") or "spot"
                    mode = {"spot": "spot", "fixed monthly": "fixed_monthly", "fixed_monthly": "fixed_monthly", "every n months": "every_n_months", "every_n_months": "every_n_months"}.get(mode.lower())
                    if not mode:
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: Price mode must be spot, fixed_monthly or every_n_months")
                    currency = (text(row, "currency") or "INR").upper()
                    if not re.fullmatch(r"[A-Z]{3}", currency):
                        raise HTTPException(status_code=422, detail=f"Row {row_number}: Currency must be a 3-letter code such as INR")
                    line["sources"].append({"supplier": who, "depth": depth, "sob_percent": sob, "unit_price": price, "currency": currency, "price_mode": mode,
                                            "lead_time_days": int(number(row, "lead_time_days", 0, row_number)), "moq": number(row, "moq", 0, row_number),
                                            "buffer_days": int(number(row, "buffer_days", 0, row_number)),
                                            "price_interval_months": max(1, int(number(row, "price_interval_months", 1, row_number)))})
            if not plan:
                raise HTTPException(status_code=422, detail="The sheet has no rows for this supplier's parts")
            counts = {"parts": 0, "lines": 0, "suppliers": 0}
            for link_id, lines in plan.items():
                link = db.get(SupplierPart, link_id)
                old = db.scalars(select(SupplierBOMLine.id).where(SupplierBOMLine.supplier_id == supplier_id, SupplierBOMLine.product_material_id == link.material_id)).all()
                if old:
                    db.execute(delete(SupplierBOMLineSource).where(SupplierBOMLineSource.line_id.in_(old)))
                    db.execute(sa_update(SupplierBOMLine).where(SupplierBOMLine.id.in_(old)).values(parent_line_id=None))
                    db.execute(delete(SupplierBOMLine).where(SupplierBOMLine.id.in_(old)))
                counts["parts"] += 1

                def save(entries: list[dict], parent_id: int | None) -> None:
                    for entry in entries:
                        found = entry["found"]
                        if isinstance(found, SupplierMaterial):
                            own = found
                        else:
                            own = own_material_for(db, supplier_id, found.id if isinstance(found, Item) else None, found.id if isinstance(found, TierMaterial) else None)
                        line = SupplierBOMLine(supplier_id=supplier_id, product_material_id=link.material_id, parent_line_id=parent_id,
                                               supplier_material_id=own.id, quantity=entry["quantity"])
                        db.add(line)
                        db.flush()
                        counts["lines"] += 1
                        for src in entry["sources"]:
                            who = src.pop("supplier")
                            depth = src.pop("depth")
                            db.add(SupplierBOMLineSource(line_id=line.id, sub_supplier_id=who.id, **src))
                            add_supplier_tier(who, f"T{depth}")
                            counts["suppliers"] += 1
                        save(entry["children"], line.id)
                save(lines, None)
            db.commit()
            return counts

    def coverage_link(db: Session, link_id: int) -> SupplierPart:
        link = db.get(SupplierPart, link_id)
        if not link or link.tier != TOP_TIER or not link.material_id:
            raise HTTPException(status_code=404, detail="T1 supplier/material relationship not found")
        return link

    @app.get("/api/planning/supplier-coverage/{link_id}")
    def supplier_coverage(link_id: int, request: Request):
        """Coverage page of one T1 relationship (e.g. Savera -> Base Front Beam): the T1 row and the
        supplier's part BOM (L1 material lines, quantity per piece) with the T2 suppliers of each line."""
        current_user(request)
        _check_supply_role(request)
        can_price = _check_price_access(request)
        with open_db() as db:
            link = coverage_link(db, link_id)
            t1 = chain_link_json(db, link, can_price, {})
            item = db.get(Item, link.material_id)
            t1["uom"] = item.uom if item else ""
            bom = supplier_bom_view(db, link.supplier_id, link.material_id, can_price)
            lines = bom["lines"]  # L1 materials, their L2 BOM and L3 BOM, each with its suppliers
            month = _month_text(_month(request.query_params.get("month") or _month_text(date.today())))
            cycle, demand = month_material_demand(db, request.query_params.get("workspace") or "trial", month)
            plan = {"month": month, "revision": f"Rev.{cycle.revision:02d}", "part_requirement": demand.get(link.material_id, 0.0)} if cycle else None
            return {"t1": t1, "lines": lines, "plan": plan}

    @app.post("/api/planning/supplier-coverage/lines", status_code=201)
    def add_coverage_line(payload: CoverageLineInput, request: Request):
        """Add a material line (e.g. FL001STEEL0.7, 2 kg per beam) with the suppliers typed for it. A T1/T2/T3
        Material Master code joins the supplier's material master."""
        current_user(request)
        _check_supply_role(request)
        if sum(1 for value in (payload.item_id, payload.tier_material_id, payload.supplier_material_id) if value) != 1:
            raise HTTPException(status_code=422, detail="Pick one material code")
        with open_db() as db:
            link = coverage_link(db, payload.link_id)
            supplier_id = link.supplier_id
            if payload.supplier_material_id:
                own = db.get(SupplierMaterial, payload.supplier_material_id)
                if not own or own.supplier_id != supplier_id:
                    raise HTTPException(status_code=422, detail="Material is not in this supplier's material master")
            else:
                own = own_material_for(db, supplier_id, payload.item_id, payload.tier_material_id)
            if payload.parent_line_id is not None:
                parent = db.get(SupplierBOMLine, payload.parent_line_id)
                if not parent or parent.supplier_id != supplier_id or parent.product_material_id != link.material_id:
                    raise HTTPException(status_code=404, detail="Parent line not found in this part BOM")
                if line_depth(db, parent) >= 3:
                    raise HTTPException(status_code=422, detail="Lines go down to L3 (L1 material, its L2 BOM, their L3 BOM)")
            same_place = SupplierBOMLine.parent_line_id == payload.parent_line_id if payload.parent_line_id else SupplierBOMLine.parent_line_id.is_(None)
            if db.scalar(select(SupplierBOMLine.id).where(SupplierBOMLine.supplier_id == supplier_id, SupplierBOMLine.product_material_id == link.material_id,
                                                          same_place, SupplierBOMLine.supplier_material_id == own.id)):
                raise HTTPException(status_code=409, detail=f"{own.code} is already a line {'under that line' if payload.parent_line_id else 'of this part BOM'}; change its quantity with Edit")
            line = SupplierBOMLine(supplier_id=supplier_id, product_material_id=link.material_id, supplier_material_id=own.id, quantity=payload.quantity,
                                   parent_line_id=payload.parent_line_id)
            db.add(line)
            db.flush()
            # its suppliers, as typed (nothing is added that was not typed)
            if sum(src.sob_percent for src in payload.sources) > 100.0001:
                raise HTTPException(status_code=422, detail=f"SOB of the suppliers of {own.code} adds up to more than 100%")
            if any(src.unit_price for src in payload.sources) and not _check_price_access(request):
                raise HTTPException(status_code=403, detail="Price access is required to set a price")
            seen_suppliers: set[int] = set()
            depth = line_depth(db, line)
            for number, src in enumerate(payload.sources, start=1):
                who = db.get(Supplier, src.sub_supplier_id) if src.sub_supplier_id else None
                if not who:
                    raise HTTPException(status_code=422, detail=f"Supplier {number} of {own.code} is not in the Supplier register")
                if who.id in seen_suppliers:
                    raise HTTPException(status_code=422, detail=f"{who.name} is entered twice as a supplier of {own.code}")
                seen_suppliers.add(who.id)
                if src.item_id or src.supplier_material_id:
                    check_component(db, supplier_id, src.item_id, src.supplier_material_id, f"Supplier {number} alternate")
                db.add(SupplierBOMLineSource(line_id=line.id, **src.model_dump()))
                add_supplier_tier(who, f"T{depth + 1}")
            db.commit()
            return {"id": line.id, "code": own.code, "sources": len(payload.sources)}

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
        """Edit any field of a relationship: supplier, part (T1, or T2/T3 for T2/T3 rows), tier,
        parent supplier, SOB, lead time, MOQ, buffer, price, currency, price agreement."""
        current_user(request)
        _check_supply_role(request)
        allowed = {"material_id", "tier_material_id", "supplier_id", "tier", "parent_supplier_id", "parent_link_id", "sob_percent", "lead_time_days",
                   "moq", "buffer_days", "unit_price", "currency", "price_mode", "price_interval_months"}
        if set(payload) - allowed:
            raise HTTPException(status_code=422, detail="Unsupported supplier/material update field")
        if "unit_price" in payload and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to update a part price")
        with open_db() as db:
            link = db.get(SupplierPart, link_id)
            if not link:
                raise HTTPException(status_code=404, detail="Supplier/material link not found")
            current = {key: getattr(link, key) for key in allowed}
            merged = {**current, **payload}
            if payload.get("material_id"):
                merged["tier_material_id"] = payload.get("tier_material_id")
            if payload.get("tier_material_id"):
                merged["material_id"] = payload.get("material_id")
            # a different parent supplier (or tier) without a new parent relationship drops the old one
            if ("parent_supplier_id" in payload or "tier" in payload) and "parent_link_id" not in payload and link.parent_link_id:
                parent_link = db.get(SupplierPart, link.parent_link_id)
                if not parent_link or parent_link.supplier_id != merged.get("parent_supplier_id") or _normalize_tier(merged.get("tier")) != f"T{int(parent_link.tier[1:]) + 1}":
                    merged["parent_link_id"] = None
            try:
                values = SupplierPartInput(**{k: v for k, v in merged.items() if k in SupplierPartInput.model_fields}).model_dump()
            except Exception as exc:
                raise HTTPException(status_code=422, detail="Check the values: SOB 0-100, lead time / MOQ / buffer 0 or more, currency like INR") from exc
            values = check_supplier_part(db, values, request, exclude_id=link_id)
            price_changed = any(values[key] != current[key] for key in ("unit_price", "price_mode", "price_interval_months"))
            for key, value in values.items():
                setattr(link, key, value)
            tiers_from_links(db, link)
            if price_changed:
                link.price_updated_month = _month_text(date.today())
            db.commit()
            return {"id": link.id, "tier": link.tier, "price_updated_month": link.price_updated_month}

    @app.post("/api/planning/plans", status_code=201)
    def upload_plan(payload: PlanUpload, request: Request):
        user = current_user(request)
        _check_workspace = _validate_workspace(payload.workspace)
        issue = _month(payload.issue_month)
        issue_month = _month_text(issue)
        with open_db() as db:
            models_by_code: dict[str, list[ProductModel]] = {}
            for model in db.scalars(select(ProductModel)).all():
                models_by_code.setdefault(model.code, []).append(model)
            latest_revision = db.scalar(select(func.max(PlanningCycle.revision)).where(PlanningCycle.workspace == payload.workspace, PlanningCycle.issue_month == issue_month))
            revision = 0 if latest_revision is None else latest_revision + 1
            if payload.revision is not None:
                taken = db.scalar(select(PlanningCycle).where(PlanningCycle.workspace == payload.workspace, PlanningCycle.issue_month == issue_month,
                                                              PlanningCycle.revision == payload.revision))
                if taken:
                    raise HTTPException(status_code=409, detail=f"Rev.{payload.revision:02d} of {issue_month} is already loaded"
                                        + (f" (released {date.fromisoformat(taken.release_date).strftime('%d-%m-%Y')})" if taken.release_date else "")
                                        + "; change the Plan revision in the sheet")
                revision = payload.revision
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
                where = f"Row {row['_row']}: " if row.get("_row") else ""
                matches = models_by_code.get(code, [])
                name = str(row.get("model_name") or "").strip()
                if len(matches) > 1 and name:  # a model code used by more than one model (e.g. FL12000 Eva / Serena): the name decides
                    matches = [model for model in matches if model.name.strip().casefold() == name.casefold()]
                if not matches:
                    raise HTTPException(status_code=422, detail=f"{where}Unknown model code: {code}{f' ({name})' if name else ''}")
                if len(matches) > 1:
                    raise HTTPException(status_code=422, detail=f"{where}Model code {code} is used by {len(matches)} models; give the Model Name")
                model = matches[0]
                if name and " ".join(model.name.split()).casefold() != " ".join(name.split()).casefold():
                    raise HTTPException(status_code=422, detail=f"{where}{code} is {model.name} in Product Master, not {name}; check the Model Code")
                key = (target, model.id)
                if key in seen:
                    raise HTTPException(status_code=422, detail=f"Duplicate plan row for model {code} in {target}")
                seen.add(key)
                normalized.append((target, model.id, quantity))
            cycle = PlanningCycle(workspace=payload.workspace, issue_month=issue_month, revision=revision, created_by=user["id"],
                                  release_date=_sheet_date(payload.release_date).isoformat() if payload.release_date else None)
            db.add(cycle)
            db.flush()
            db.add_all(PlanLine(cycle_id=cycle.id, target_month=target, model_id=model_id, quantity=quantity) for target, model_id, quantity in normalized)
            db.commit()
            return {"id": cycle.id, "workspace": cycle.workspace, "issue_month": cycle.issue_month, "revision": f"Rev.{revision:02d}", "revision_number": revision,
                    "rows": len(normalized), "release_date": cycle.release_date}

    @app.post("/api/planning/plans/{issue_month}/upload", status_code=201)
    async def upload_plan_excel(issue_month: str, request: Request, workspace: str = "trial", release_date: str = ""):
        raw = await request.body()
        blocks = _plan_blocks(raw)
        if blocks is not None:  # the plan template: one plan per upload, released on the date picked in the app (next revision)
            picked_date = _sheet_date(release_date) if release_date else None
            shown = lambda day: day.strftime("%d-%m-%Y") if day else "(blank)"
            dated = [block for block in blocks if block["released"]]
            if len(blocks) == 1:
                chosen = blocks
            elif picked_date and dated:  # an older sheet with several dated blocks: the one released on the picked date
                chosen = [block for block in blocks if block["released"] == picked_date]
                if not chosen:
                    listed = ", ".join(shown(block["released"]) for block in blocks)
                    raise HTTPException(status_code=422, detail=f"No plan in the file is released on {shown(picked_date)}. Plans in the file: {listed}.")
            else:
                raise HTTPException(status_code=422, detail=f"The file has {len(blocks)} plans; keep one plan (one heading row) per upload")
            if len(chosen) > 1:
                raise HTTPException(status_code=422, detail=f"The file has {len(chosen)} plans released on {shown(picked_date)}; keep one")
            block = chosen[0]
            rows = [{"model_code": row["code"], "model_name": row["name"], "target_month": month, "quantity": quantity, "_row": row["row"]}
                    for row in block["rows"] for month, quantity in row["quantities"].items()]
            if not rows:
                raise HTTPException(status_code=422, detail="The plan has no quantities")
            released = picked_date or block["released"]
            plan_month = _month_text(_month(released)) if released else _month_text(_month(issue_month))  # no date: the month in the address (older uploads)
            payload = PlanUpload(issue_month=plan_month, workspace=workspace, rows=rows, release_date=released.isoformat() if released else None)
            return upload_plan(payload, request)
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
                    "created_at": cycle.created_at.isoformat(), "release_date": cycle.release_date,
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
        payload.period = _month_text(_month(payload.period))
        as_of = _sheet_date(payload.as_of) if payload.as_of else None
        if as_of and _month_text(_month(as_of)) != payload.period:
            raise HTTPException(status_code=422, detail=f"The as-on date {as_of.strftime('%d-%m-%Y')} is not in {payload.period}")
        with open_db() as db:
            models_by_code = {m.code: m for m in db.scalars(select(ProductModel)).all()}
            saved = 0
            for row in payload.rows:
                where = f"Row {row['_excel_row']}: " if row.get("_excel_row") else ""
                code = str(row.get("model_code", "")).strip()
                if not code:
                    continue
                model = models_by_code.get(code)  # actuals: one model per code
                if not model:
                    raise HTTPException(status_code=422, detail=f"{where}Unknown model code: {code}")
                name = str(row.get("model_name") or "").strip()
                if name and " ".join(model.name.split()).casefold() != " ".join(name.split()).casefold():
                    raise HTTPException(status_code=422, detail=f"{where}{code} is {model.name} in Product Master, not {name}; check the Model Code")
                raw_quantity = row.get("actual_quantity", row.get("quantity"))
                if raw_quantity is None or str(raw_quantity).strip() == "":
                    continue  # nothing typed for this model
                try:
                    quantity = float(raw_quantity)
                    oem_stock = float(row.get("oem_stock") or 0)
                except (TypeError, ValueError) as exc:
                    raise HTTPException(status_code=422, detail=f"{where}actual production and OEM stock must be numbers") from exc
                if min(quantity, oem_stock) < 0:
                    raise HTTPException(status_code=422, detail=f"{where}Actual production and OEM stock cannot be negative")
                existing = db.scalar(select(ActualProduction).where(ActualProduction.workspace == payload.workspace, ActualProduction.period == payload.period, ActualProduction.model_id == model.id))
                if existing:  # the total till the new date replaces the earlier total
                    existing.actual_quantity = quantity
                    existing.oem_stock = oem_stock
                    existing.updated_by = user["id"]
                    existing.as_of = as_of.isoformat() if as_of else existing.as_of
                else:
                    db.add(ActualProduction(workspace=payload.workspace, period=payload.period, model_id=model.id, actual_quantity=quantity, oem_stock=oem_stock,
                                            as_of=as_of.isoformat() if as_of else None, updated_by=user["id"]))
                saved += 1
            if not saved:
                raise HTTPException(status_code=422, detail="No actual production typed")
            db.commit()
            return {"workspace": payload.workspace, "period": payload.period, "rows": saved, "as_of": as_of.isoformat() if as_of else None}

    def month_plan(db: Session, workspace: str, period: str) -> tuple[PlanningCycle | None, dict[int, float]]:
        """The plan in force for a production month: the latest revision of the latest plan issued up to that month."""
        cycle = db.scalar(select(PlanningCycle).where(PlanningCycle.workspace == workspace, PlanningCycle.issue_month <= period)
                          .order_by(PlanningCycle.issue_month.desc(), PlanningCycle.revision.desc()).limit(1))
        if not cycle:
            return None, {}
        return cycle, {line.model_id: line.quantity for line in db.scalars(select(PlanLine).where(PlanLine.cycle_id == cycle.id, PlanLine.target_month == period)).all()}

    def actual_rows(db: Session, workspace: str, period: str) -> tuple[PlanningCycle | None, list[dict]]:
        """Every model planned for the month or with an actual: plan, actual till date, balance, % achieved."""
        cycle, planned = month_plan(db, workspace, period)
        actuals = {row.model_id: row for row in db.scalars(select(ActualProduction).where(ActualProduction.workspace == workspace, ActualProduction.period == period)).all()}
        models = {m.id: m for m in db.scalars(select(ProductModel)).all()}
        ids = [model_id for model_id in models if planned.get(model_id) or model_id in actuals]
        rows = []
        for model_id in sorted(ids, key=lambda i: (models[i].code.casefold(), models[i].name.casefold())):
            plan, actual = planned.get(model_id, 0.0), actuals.get(model_id)
            done = actual.actual_quantity if actual else None
            rows.append({"model_code": models[model_id].code, "model_name": models[model_id].name, "plan_quantity": plan,
                         "actual_quantity": done, "oem_stock": actual.oem_stock if actual else None, "as_of": actual.as_of if actual else None,
                         "balance": plan - (done or 0), "achieved_percent": round((done or 0) * 100 / plan, 1) if plan else None})
        return cycle, rows

    @app.get("/api/planning/actuals")
    def list_actuals(request: Request, period: str, workspace: str = "trial", with_plan: bool = False):
        current_user(request)
        _validate_workspace(workspace)
        period = _month_text(_month(period))
        with open_db() as db:
            cycle, rows = actual_rows(db, workspace, period)
            if with_plan:
                return {"period": period, "plan": {"revision": f"Rev.{cycle.revision:02d}", "issue_month": cycle.issue_month, "release_date": cycle.release_date} if cycle else None,
                        "items": rows}
            return [row for row in rows if row["actual_quantity"] is not None]

    def month_material_demand(db: Session, workspace: str, month: str) -> tuple[PlanningCycle | None, dict[int, float]]:
        """Plan demand of every material in a month: the plan in force x each model's BOM (L1-L3, alternates' SOB)."""
        cycle, planned = month_plan(db, workspace, month)
        if not planned:
            return cycle, {}
        _, children = _item_map(db)
        roots_by_model: dict[int, list[ModelBOM]] = {}
        for root in db.scalars(select(ModelBOM).where(ModelBOM.model_id.in_(list(planned)))).all():
            roots_by_model.setdefault(root.model_id, []).append(root)
        root_history, component_history, _ = _line_alternate_history(db)
        root_splits, component_splits = _splits_at(root_history, month), _splits_at(component_history, month)
        demand: dict[int, float] = {}
        for model_id, quantity in planned.items():
            if not quantity:
                continue
            for material_id, per_model in _rollup_model_bom(roots_by_model.get(model_id, []), children, root_splits, component_splits).items():
                demand[material_id] = demand.get(material_id, 0) + quantity * per_model
        return cycle, demand

    @app.get("/api/planning/supplier-overview/{supplier_id}")
    def supplier_overview(supplier_id: int, request: Request, month: str = "", workspace: str = "trial"):
        """One supplier at a glance (Supplier register click): the parts it supplies you with SOB, rate, FG stock and
        requirement; each part's BOM with its sub-suppliers (SOB, rate, requirement, stock); the sub-supplier list."""
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        can_price = _check_price_access(request)
        month = _month_text(_month(month)) if month else _month_text(date.today())
        with open_db() as db:
            supplier = db.get(Supplier, supplier_id)
            if not supplier:
                raise HTTPException(status_code=404, detail="Supplier not found")
            cycle, demand = month_material_demand(db, workspace, month)
            links = db.scalars(select(SupplierPart).where(SupplierPart.supplier_id == supplier_id, SupplierPart.tier == TOP_TIER, SupplierPart.material_id.is_not(None))).all()
            items = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_({link.material_id for link in links})))} if links else {}
            stock = {row.supplier_part_id: row for row in db.scalars(select(SupplierStockSnapshot).where(
                SupplierStockSnapshot.workspace == workspace, SupplierStockSnapshot.report_month == month,
                SupplierStockSnapshot.supplier_part_id.in_([link.id for link in links]))).all()} if links else {}
            source_stock = {(row.product_material_id, row.material_code.casefold(), row.sub_supplier_id): row for row in db.scalars(select(SupplierSourceStock).where(
                SupplierSourceStock.workspace == workspace, SupplierSourceStock.report_month == month, SupplierSourceStock.t1_supplier_id == supplier_id)).all()}
            parts, sub_rows = [], []
            for link in sorted(links, key=lambda l: items[l.material_id].code.casefold()):
                item = items[link.material_id]
                held = stock.get(link.id)
                part_need = demand.get(item.id, 0.0)
                share = part_need * (link.sob_percent or 0) / 100  # this supplier's part of the requirement
                part = {"link_id": link.id, "code": item.code, "name": item.name, "uom": item.uom, "sob_percent": link.sob_percent,
                        "lead_time_days": link.lead_time_days, "price_status": _price_status(link, _month_text(date.today())),
                        "stock": held.quantity if held else None, "available": held.available_quantity if held else None,
                        "part_requirement": part_need, "supplier_requirement": share, "lines": []}
                net = max(0.0, share - (held.quantity if held else 0.0))  # what is left to make after the FG stock
                part["net_requirement"] = net
                if can_price:
                    part.update(unit_price=link.unit_price, currency=link.currency)
                view = supplier_bom_view(db, supplier_id, item.id, can_price)
                per_unit: dict[int, float] = {}
                for line in view["lines"]:
                    per_piece = line["quantity"] * (per_unit.get(line["parent_line_id"], 1.0) if line["parent_line_id"] else 1.0)
                    per_unit[line["id"]] = per_piece
                    component = line["component"]
                    need = net * per_piece  # the BOM is needed for what is still to be made
                    entry = {"id": line["id"], "depth": line["depth"], "parent_line_id": line["parent_line_id"], "code": component["code"], "name": component["name"],
                             "origin": component["origin"], "uom": component["uom"], "bom_qty": line["quantity"], "per_part": per_piece, "requirement": need, "sources": []}
                    for src in line["sources"]:
                        kept = source_stock.get((item.id, component["code"].casefold(), src["sub_supplier_id"]))
                        source = {"sub_supplier_id": src["sub_supplier_id"], "code": src["sub_supplier_code"], "name": src["sub_supplier"], "tier": f"T{line['depth'] + 1}",
                                  "sob_percent": src["sob_percent"], "lead_time_days": src["lead_time_days"], "requirement": need * src["sob_percent"] / 100,
                                  "stock": kept.quantity if kept else None, "available": kept.available_quantity if kept else None}
                        if can_price:
                            source.update(unit_price=src.get("unit_price"), currency=src["currency"])
                        entry["sources"].append(source)
                        sub_rows.append({**source, "part_code": item.code, "part_name": item.name, "material_code": component["code"], "material_name": component["name"], "uom": component["uom"]})
                    part["lines"].append(entry)
                parts.append(part)
            sub_rows.sort(key=lambda row: (row["code"].casefold(), row["part_code"].casefold(), row["material_code"].casefold()))
            # one line per sub-supplier and material: requirement summed over the parts; SOB = its share of the
            # material's total requirement (plain average when there is no requirement); rate weighted by requirement
            material_total: dict[str, float] = {}
            for row in sub_rows:
                material_total[row["material_code"].casefold()] = material_total.get(row["material_code"].casefold(), 0) + row["requirement"]
            blended: dict[tuple[int, str], dict] = {}
            for row in sub_rows:
                key = (row["sub_supplier_id"], row["material_code"].casefold())
                line = blended.setdefault(key, {"sub_supplier_id": row["sub_supplier_id"], "code": row["code"], "name": row["name"], "tiers": [], "material_code": row["material_code"],
                                                "material_name": row["material_name"], "uom": row["uom"], "requirement": 0.0, "stock": None, "parts": [], "_sob": [], "_rate": []})
                if row["tier"] not in line["tiers"]:
                    line["tiers"].append(row["tier"])
                line["requirement"] += row["requirement"]
                if row["stock"] is not None:
                    line["stock"] = (line["stock"] or 0) + row["stock"]
                line["parts"].append({"part_code": row["part_code"], "sob_percent": row["sob_percent"], "requirement": row["requirement"], "unit_price": row.get("unit_price")})
                line["_sob"].append(row["sob_percent"])
                if can_price:
                    line["_rate"].append((row.get("unit_price") or 0, row["requirement"], row.get("currency")))
            summary = []
            for (sub_id, material_key), line in blended.items():
                total = material_total.get(material_key, 0)
                line["sob_percent"] = round(line["requirement"] * 100 / total, 4) if total else round(sum(line["_sob"]) / len(line["_sob"]), 4)
                line["tier"] = ", ".join(line.pop("tiers"))
                rates = line.pop("_rate")
                line.pop("_sob")
                if can_price and rates:
                    weight = sum(need for _, need, _ in rates)
                    line["unit_price"] = round(sum(price * need for price, need, _ in rates) / weight, 4) if weight else round(sum(price for price, _, _ in rates) / len(rates), 4)
                    line["currency"] = rates[0][2] if len({currency for _, _, currency in rates}) == 1 else "mixed"
                line["net_requirement"] = max(0.0, line["requirement"] - (line["stock"] or 0.0))
                summary.append(line)
            return {"supplier": {"id": supplier.id, "code": supplier.code, "name": supplier.name, "category": supplier.category, "tiers": _tier_list(supplier.tiers)},
                    "month": month, "workspace": workspace,
                    "plan": {"revision": f"Rev.{cycle.revision:02d}", "issue_month": cycle.issue_month, "release_date": cycle.release_date} if cycle else None,
                    "parts": parts, "sub_suppliers": sub_rows, "sub_supplier_summary": summary}

    @app.post("/api/planning/supplier-overview/{supplier_id}/sub-supplier")
    def edit_sub_supplier(supplier_id: int, payload: dict, request: Request):
        """From the overview's sub-supplier list: set this sub-supplier's SOB and / or rate for one material in every
        part of the supplier where it supplies that material. All parts or none (SOB per part at most 100%)."""
        current_user(request)
        _check_supply_role(request)
        sub_id, code = payload.get("sub_supplier_id"), str(payload.get("material_code") or "").strip()
        sob, price = payload.get("sob_percent"), payload.get("unit_price")
        stock = payload.get("stock")
        if sob is None and price is None and stock in (None, ""):
            raise HTTPException(status_code=422, detail="Give the SOB, the rate or the stock")
        try:
            sob = None if sob in (None, "") else float(sob)
            price = None if price in (None, "") else float(price)
        except (TypeError, ValueError) as exc:
            raise HTTPException(status_code=422, detail="SOB and rate must be numbers") from exc
        if sob is not None and not 0 < sob <= 100:
            raise HTTPException(status_code=422, detail="SOB must be more than 0 and at most 100")
        try:
            stock = None if stock in (None, "") else float(stock)
        except (TypeError, ValueError) as exc:
            raise HTTPException(status_code=422, detail="Stock must be a number") from exc
        if stock is not None and (stock < 0 or not math.isfinite(stock)):
            raise HTTPException(status_code=422, detail="Stock must be zero or more")
        workspace = str(payload.get("workspace") or "trial")
        _validate_workspace(workspace)
        stock_month = _month_text(_month(payload.get("month") or _month_text(date.today())))
        if price is not None and (price < 0 or not math.isfinite(price)):
            raise HTTPException(status_code=422, detail="Rate must be zero or more")
        if price is not None and not _check_price_access(request):
            raise HTTPException(status_code=403, detail="Price access is required to set a rate")
        with open_db() as db:
            rows = db.execute(select(SupplierBOMLineSource, SupplierBOMLine).join(SupplierBOMLine, SupplierBOMLine.id == SupplierBOMLineSource.line_id)
                              .join(SupplierMaterial, SupplierMaterial.id == SupplierBOMLine.supplier_material_id)
                              .where(SupplierBOMLine.supplier_id == supplier_id, SupplierBOMLineSource.sub_supplier_id == sub_id,
                                     func.lower(SupplierMaterial.code) == code.casefold())).all()
            if not rows:
                raise HTTPException(status_code=404, detail=f"That sub-supplier does not supply {code} in this supplier's BOMs")
            month = _month_text(date.today())
            for source, line in rows:
                if sob is not None:
                    others = db.scalar(select(func.coalesce(func.sum(SupplierBOMLineSource.sob_percent), 0)).where(
                        SupplierBOMLineSource.line_id == line.id, SupplierBOMLineSource.id != source.id)) or 0
                    if others + sob > 100.0001:
                        part = db.get(Item, line.product_material_id)
                        raise HTTPException(status_code=422, detail=f"{part.code if part else 'A part'}: the other suppliers of {code} have {others:g}%, so at most {100 - others:g}% is left; lower them first")
                    source.sob_percent = sob
                if price is not None and price != source.unit_price:
                    source.unit_price = price
                    source.price_updated_month = month
            if stock is not None:
                # one stock per sub-supplier and material for this supplier: kept on the first part (by code) that uses
                # it; other parts' entries of the month are cleared so the total is what was typed
                products = sorted({line.product_material_id for _, line in rows}, key=lambda pid: (db.get(Item, pid).code.casefold() if db.get(Item, pid) else ""))
                same = (SupplierSourceStock.workspace == workspace, SupplierSourceStock.report_month == stock_month, SupplierSourceStock.t1_supplier_id == supplier_id,
                        SupplierSourceStock.sub_supplier_id == sub_id, func.lower(SupplierSourceStock.material_code) == code.casefold())
                db.execute(delete(SupplierSourceStock).where(*same, SupplierSourceStock.product_material_id != products[0]))
                held = db.scalar(select(SupplierSourceStock).where(*same, SupplierSourceStock.product_material_id == products[0]))
                if not held:
                    held = SupplierSourceStock(workspace=workspace, report_month=stock_month, t1_supplier_id=supplier_id, product_material_id=products[0],
                                               material_code=code, sub_supplier_id=sub_id)
                    db.add(held)
                held.quantity = held.available_quantity = stock
            db.commit()
            return {"parts": len(rows)}

    @app.get("/api/planning/actuals/template.xlsx")
    def actual_template(request: Request, as_of: str = "", workspace: str = "trial"):
        """Actual production template like the plan template: logo and title, the month and as-on date, then the
        models of the month's plan with Plan qty, Actual production (total till date) and OEM stock."""
        current_user(request)
        _validate_workspace(workspace)
        day = _sheet_date(as_of) if as_of else date.today()
        period = _month_text(_month(day))
        with open_db() as db:
            cycle, rows = actual_rows(db, workspace, period)
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Actual production"
        try:  # the app logo (needs Pillow; the template works without it)
            from openpyxl.drawing.image import Image as SheetImage
            logo = SheetImage(str(Path(__file__).resolve().parent / "static" / "logo.png"))
            logo.width = logo.height = 48
            sheet.add_image(logo, "A1")
        except Exception:
            pass
        sheet.row_dimensions[1].height = 22
        sheet.row_dimensions[2].height = 22
        sheet["B1"] = "Design to Cost · Actual production"
        sheet["B1"].font = Font(bold=True, size=14, color="12263A")
        if cycle:
            plan_date = date.fromisoformat(cycle.release_date).strftime("%d-%m-%Y") if cycle.release_date else ""
            plan_text = f"plan Rev.{cycle.revision:02d}" + (f" of {plan_date}" if plan_date else "")
        else:
            plan_text = "no plan uploaded for this month"
        sheet["B2"] = f"Month: {_month(day).strftime('%b-%Y')} · as on {day.strftime('%d-%m-%Y')} · {plan_text}. Type the total produced in the month till the as-on date; a blank row is skipped."
        sheet["B2"].font = Font(italic=True, color="5B6B7B")
        headers = ["Model Code", "Model Name", f"Plan qty ({f'Rev.{cycle.revision:02d}' if cycle else 'no plan'})", "Actual production (till date)", "OEM stock"]
        for index, value in enumerate(headers, start=1):
            cell = sheet.cell(4, index, value)
            cell.fill, cell.font = PatternFill("solid", fgColor="12263A"), Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row_index, row in enumerate(rows, start=5):
            sheet.cell(row_index, 1, row["model_code"])
            sheet.cell(row_index, 2, row["model_name"])
            sheet.cell(row_index, 3, row["plan_quantity"])
            sheet.cell(row_index, 4, row["actual_quantity"])
            sheet.cell(row_index, 5, row["oem_stock"])
            for column in (3, 4, 5):
                sheet.cell(row_index, column).number_format = "#,##0.###"
        for column, width in zip("ABCDE", (16, 26, 16, 22, 14)):
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "C5"
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 headers={"Content-Disposition": f'attachment; filename="actual-production-{day.strftime("%d-%m-%Y")}.xlsx"'})

    @app.post("/api/planning/actuals/{period}/upload", status_code=201)
    async def upload_actuals_excel(period: str, request: Request, workspace: str = "trial", as_of: str = ""):
        aliases = {"model_code": "model_code", "model_name": "model_name", "actual_production_(till_date)": "actual_quantity", "actual_production": "actual_quantity",
                   "actual_quantity": "actual_quantity", "actual": "actual_quantity", "quantity": "actual_quantity", "oem_stock": "oem_stock"}
        rows = _read_supplier_part_sheet(await request.body(), aliases, {"model_code", "actual_quantity"})
        return upload_actuals(ActualUpload(period=period, workspace=workspace, rows=rows, as_of=as_of or None), request)

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

    def save_stock(db: Session, workspace: str, report_month: str, link: SupplierPart, quantity: float, available: float | None, arrival: str | None) -> None:
        available = quantity if available is None else available
        if min(quantity, available) < 0 or available > quantity:
            raise HTTPException(status_code=422, detail="Available quantity must be between zero and reported supplier stock")
        arrival = (arrival or "").strip() or None
        if arrival:
            arrival = _month_text(_month(arrival))
        existing = db.scalar(select(SupplierStockSnapshot).where(SupplierStockSnapshot.workspace == workspace, SupplierStockSnapshot.report_month == report_month, SupplierStockSnapshot.supplier_part_id == link.id))
        if existing:
            existing.quantity, existing.available_quantity, existing.arrival_month = quantity, available, arrival
        else:
            db.add(SupplierStockSnapshot(workspace=workspace, report_month=report_month, supplier_part_id=link.id, quantity=quantity, available_quantity=available, arrival_month=arrival))
            db.flush()

    @app.post("/api/planning/supplier-stock/{report_month}", status_code=201)
    def save_supplier_stock_rows(report_month: str, payload: SupplierStockRows, request: Request, workspace: str = "trial"):
        """Supplier stock typed in the entry table (no Excel): one row per relationship. All rows or none;
        a relationship already reported for the month is updated."""
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        report_month = _month_text(_month(report_month))
        if not payload.rows:
            raise HTTPException(status_code=422, detail="Add at least one row")
        with open_db() as db:
            seen: set[int] = set()
            for index, row in enumerate(payload.rows, start=1):
                link = db.get(SupplierPart, row.supplier_part_id)
                if not link:
                    raise HTTPException(status_code=422, detail=f"Row {index}: choose the supplier's material")
                if row.supplier_part_id in seen:
                    raise HTTPException(status_code=422, detail=f"Row {index}: this supplier and material is entered twice")
                seen.add(row.supplier_part_id)
                try:
                    save_stock(db, workspace, report_month, link, row.quantity, row.available_quantity, row.arrival_month)
                except HTTPException as exc:
                    raise HTTPException(status_code=exc.status_code, detail=f"Row {index}: {exc.detail}") from exc
            db.commit()
            return {"workspace": workspace, "report_month": report_month, "rows": len(payload.rows)}

    @app.get("/api/planning/supply-tree")
    def supply_tree(request: Request, report_month: str = "", workspace: str = "trial"):
        """Every T1 supplier -> the parts it supplies you -> each part's BOM (L1-L3) and the suppliers of each
        line (T2-T4), depth first, with the month's stock on the T1 parts (Upload supplier stock page)."""
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        month = _month_text(_month(report_month)) if report_month else _month_text(date.today())
        with open_db() as db:
            links = db.scalars(select(SupplierPart).where(SupplierPart.tier == TOP_TIER, SupplierPart.material_id.is_not(None))).all()
            suppliers = {who.id: who for who in db.scalars(select(Supplier)).all()}
            items = {item.id: item for item in db.scalars(select(Item).where(Item.id.in_({link.material_id for link in links})))} if links else {}
            stock = {row.supplier_part_id: row for row in db.scalars(select(SupplierStockSnapshot).where(
                SupplierStockSnapshot.workspace == workspace, SupplierStockSnapshot.report_month == month)).all()}
            with_bom = set(db.execute(select(SupplierBOMLine.supplier_id, SupplierBOMLine.product_material_id).distinct()).all())
            source_stock = {(row.t1_supplier_id, row.product_material_id, row.material_code.casefold(), row.sub_supplier_id): row
                            for row in db.scalars(select(SupplierSourceStock).where(SupplierSourceStock.workspace == workspace, SupplierSourceStock.report_month == month)).all()}
            nodes: list[dict] = []
            by_supplier: dict[int, list[SupplierPart]] = {}
            for link in links:
                by_supplier.setdefault(link.supplier_id, []).append(link)
            for supplier_id in sorted(by_supplier, key=lambda sid: (suppliers[sid].code or "").casefold()):
                who = suppliers[supplier_id]
                parts = sorted(by_supplier[supplier_id], key=lambda link: items[link.material_id].code.casefold())
                nodes.append({"depth": 0, "level": "T1 supplier", "supplier_code": who.code, "supplier_name": who.name, "parts": len(parts)})
                for link in parts:
                    item = items[link.material_id]
                    snapshot = stock.get(link.id)
                    nodes.append({"depth": 1, "level": "T1 part", "link_id": link.id, "supplier_code": who.code, "supplier_name": who.name,
                                  "material_code": item.code, "material_name": item.name, "sob_percent": link.sob_percent, "uom": item.uom,
                                  "stock": snapshot.quantity if snapshot else None, "available": snapshot.available_quantity if snapshot else None,
                                  "arrival_month": snapshot.arrival_month if snapshot else None})
                    if (supplier_id, link.material_id) not in with_bom:
                        continue
                    for line in supplier_bom_view(db, supplier_id, link.material_id, False)["lines"]:
                        component = line["component"]
                        nodes.append({"depth": line["depth"] + 1, "level": f"L{line['depth']} material", "material_code": component["code"],
                                      "material_name": component["name"], "bom_qty": line["quantity"], "uom": component["uom"],
                                      "sob_percent": sum(src["sob_percent"] for src in line["sources"]) if line["sources"] else None})
                        for src in line["sources"]:
                            held = source_stock.get((supplier_id, link.material_id, component["code"].casefold(), src["sub_supplier_id"]))
                            nodes.append({"depth": line["depth"] + 2, "level": f"T{line['depth'] + 1} supplier", "supplier_code": src["sub_supplier_code"],
                                          "supplier_name": src["sub_supplier"], "material_code": component["code"], "material_name": component["name"],
                                          "sob_percent": src["sob_percent"], "uom": component["uom"], "lead_time_days": src["lead_time_days"],
                                          "stock": held.quantity if held else None, "available": held.available_quantity if held else None,
                                          "arrival_month": held.arrival_month if held else None})
            return {"report_month": month, "workspace": workspace, "items": nodes}

    @app.get("/api/planning/supply-tree.{file_format}")
    def export_supply_tree(file_format: str, request: Request, report_month: str = "", workspace: str = "trial", template: bool = False):
        """The supply chain tree of the Upload supplier stock page as Excel (grouped levels: + / - in Excel) or PDF."""
        if file_format not in {"xlsx", "pdf"}:
            raise HTTPException(status_code=404, detail="Export not found")
        tree = supply_tree(request, report_month, workspace)
        headers = ["Level", "Supplier Code", "Supplier Name", "Material code", "Material name", "SOB %", "BOM qty", "Unit", "Stock quantity", "Available quantity", "Arrival month"]
        indent = "   "
        values = [[indent * node["depth"] + node["level"], node.get("supplier_code") or "", node.get("supplier_name") or "", node.get("material_code") or "",
                   (f'{node["parts"]} part{"s" if node["parts"] != 1 else ""}' if node["depth"] == 0 else node.get("material_name") or ""),
                   node.get("sob_percent"), node.get("bom_qty"), node.get("uom") or "", node.get("stock"), node.get("available"), node.get("arrival_month") or ""]
                  for node in tree["items"]]
        title = f"Supply chain tree · stock {tree['report_month']} ({tree['workspace']})"
        filename = f"supply-chain-tree-{tree['report_month']}"
        if file_format == "pdf":
            output = _generic_pdf(title, headers, values)
            return StreamingResponse(output, media_type="application/pdf", headers={"Content-Disposition": f'attachment; filename="{filename}.pdf"'})
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Supply chain tree"
        styled_sheet(sheet, title, ("Stock template: type Stock quantity, Available quantity and Arrival month (YYYY-MM, if not yet there) on the T1 part rows, "
                                    "then Upload Excel on the Upload supplier stock page (other rows are for reading; a blank Stock quantity is skipped). ") if template else ""
                     + "T1 supplier -> parts it supplies you -> BOM materials (L1-L3) -> their T2 / T3 / T4 suppliers. Use + / - at the left to open and close levels.", headers)
        fills = {0: PatternFill("solid", fgColor="FFFFFF"), 1: PatternFill("solid", fgColor="F7FAFD")}
        supplier_fill = PatternFill("solid", fgColor="FFFAF0")
        for node, line in zip(tree["items"], values):
            sheet.append(line)
            row = sheet.max_row
            if node["depth"]:
                sheet.row_dimensions[row].outline_level = min(node["depth"], 7)
            is_supplier = node["depth"] > 0 and node["level"].endswith("supplier")
            for cell in sheet[row]:
                if node["depth"] <= 1:
                    cell.font = Font(bold=True)
                if is_supplier:
                    cell.fill, cell.font = supplier_fill, Font(italic=True, color="7A5A12")
                elif node["depth"] in fills:
                    cell.fill = fills[node["depth"]]
        sheet.sheet_properties.outlinePr.summaryBelow = False
        sheet.column_dimensions["A"].width = 28
        return xlsx_response(workbook, f"{'supplier-stock-template' if template else filename}.xlsx")

    @app.get("/api/planning/supplier-stock/{report_month}")
    def list_supplier_stock(report_month: str, request: Request, workspace: str = "trial"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        report_month = _month_text(_month(report_month))
        with open_db() as db:
            rows = db.execute(
                select(SupplierStockSnapshot, SupplierPart, Supplier)
                .join(SupplierPart, SupplierPart.id == SupplierStockSnapshot.supplier_part_id)
                .join(Supplier, Supplier.id == SupplierPart.supplier_id)
                .where(SupplierStockSnapshot.workspace == workspace, SupplierStockSnapshot.report_month == report_month)
                .order_by(Supplier.name, SupplierPart.id)
            ).all()
            rows = [(stock, link, db.get(Item, link.material_id) if link.material_id else db.get(TierMaterial, link.tier_material_id), supplier)
                    for stock, link, supplier in rows]
            items = [
                {"id": stock.id, "supplier_part_id": link.id, "supplier_code": supplier.code, "supplier_name": supplier.name,
                 "material_code": item.code, "material_name": item.name, "uom": item.uom, "tier": link.tier,
                 "quantity": stock.quantity, "available_quantity": stock.available_quantity, "arrival_month": stock.arrival_month}
                for stock, link, item, supplier in rows]
            for held in db.scalars(select(SupplierSourceStock).where(SupplierSourceStock.workspace == workspace, SupplierSourceStock.report_month == report_month)).all():
                who, t1, part = db.get(Supplier, held.sub_supplier_id), db.get(Supplier, held.t1_supplier_id), db.get(Item, held.product_material_id)
                own = db.scalar(select(SupplierMaterial).where(SupplierMaterial.supplier_id == held.t1_supplier_id, func.lower(SupplierMaterial.code) == held.material_code.casefold()))
                items.append({"id": None, "supplier_part_id": None, "supplier_code": who.code, "supplier_name": who.name, "material_code": held.material_code,
                              "material_name": own.name if own else "", "uom": own.uom if own else "", "tier": f"T2+ for {t1.name} · {part.code}",
                              "quantity": held.quantity, "available_quantity": held.available_quantity, "arrival_month": held.arrival_month})
            return {"report_month": report_month, "items": items}

    def stock_rows_from_tree(raw: bytes) -> list[dict] | None:
        """Rows of the supply chain tree workbook with stock typed on T1 part rows (None: not that workbook)."""
        aliases = {"level": "level", "supplier_code": "supplier_code", "material_code": "material_code", "stock_quantity": "quantity",
                   "available_quantity": "available_quantity", "arrival_month": "arrival_month"}
        try:
            rows = _read_supplier_part_sheet(raw, aliases, {"level", "supplier_code", "material_code", "quantity"})
        except HTTPException:
            return None
        out = []
        part: tuple[str, str] | None = None  # the T1 supplier and part above (for T2 / T3 supplier rows)
        for row in rows:
            level = str(row.get("level") or "").strip().casefold()
            if level == "t1 supplier":
                part = None
            elif level == "t1 part":
                part = (str(row.get("supplier_code") or "").strip(), str(row.get("material_code") or "").strip())
            source = re.fullmatch(r"t([2-9]) supplier", level)
            if (level != "t1 part" and not source) or row.get("quantity") in (None, ""):
                continue
            if source and part is None:
                raise HTTPException(status_code=422, detail=f"Row {row['_excel_row']}: a {level.upper().replace(' SUPPLIER', '')} supplier row must be under its T1 part row")
            arrival = row.get("arrival_month")
            if hasattr(arrival, "strftime"):
                arrival = arrival.strftime("%Y-%m")
            try:
                quantity = float(row["quantity"])
                available = float(row["available_quantity"]) if row.get("available_quantity") not in (None, "") else None
            except (TypeError, ValueError) as exc:
                raise HTTPException(status_code=422, detail=f"Row {row['_excel_row']}: stock quantities must be numbers") from exc
            entry = {"supplier_code": str(row["supplier_code"]).strip(), "material_code": str(row["material_code"]).strip(), "tier": "T1",
                     "quantity": quantity, "available_quantity": available, "arrival_month": arrival or "", "_excel_row": row["_excel_row"]}
            if source:
                entry.update(tier=f"T{source.group(1)}", t1_supplier_code=part[0], part_code=part[1])
            out.append(entry)
        return out

    def save_source_stock(db: Session, workspace: str, report_month: str, rows: list[dict]) -> int:
        """Stock typed on T2 / T3 / T4 supplier rows of the tree: checked against the T1 part's BOM."""
        suppliers = {who.code.casefold(): who for who in db.scalars(select(Supplier)).all()}
        items = {item.code.casefold(): item for item in db.scalars(select(Item)).all()}
        saved = 0
        for row in rows:
            where = f"Row {row['_excel_row']}: "
            t1, part, who = suppliers.get(row["t1_supplier_code"].casefold()), items.get(row["part_code"].casefold()), suppliers.get(row["supplier_code"].casefold())
            if not t1 or not part:
                raise HTTPException(status_code=422, detail=f"{where}T1 part {row['part_code']} of {row['t1_supplier_code']} is not found")
            if not who:
                raise HTTPException(status_code=422, detail=f"{where}supplier {row['supplier_code']} is not in the Supplier register")
            supplies = db.scalar(select(SupplierBOMLineSource.id).join(SupplierBOMLine, SupplierBOMLine.id == SupplierBOMLineSource.line_id)
                                 .join(SupplierMaterial, SupplierMaterial.id == SupplierBOMLine.supplier_material_id)
                                 .where(SupplierBOMLine.supplier_id == t1.id, SupplierBOMLine.product_material_id == part.id,
                                        func.lower(SupplierMaterial.code) == row["material_code"].casefold(), SupplierBOMLineSource.sub_supplier_id == who.id).limit(1))
            if not supplies:
                raise HTTPException(status_code=422, detail=f"{where}{who.name} is not a supplier of {row['material_code']} in the BOM of {part.code} ({t1.name})")
            if row["quantity"] < 0 or (row["available_quantity"] is not None and row["available_quantity"] < 0):
                raise HTTPException(status_code=422, detail=f"{where}stock cannot be negative")
            arrival = _month_text(_month(row["arrival_month"])) if row["arrival_month"] else None
            stock = db.scalar(select(SupplierSourceStock).where(
                SupplierSourceStock.workspace == workspace, SupplierSourceStock.report_month == report_month, SupplierSourceStock.t1_supplier_id == t1.id,
                SupplierSourceStock.product_material_id == part.id, func.lower(SupplierSourceStock.material_code) == row["material_code"].casefold(),
                SupplierSourceStock.sub_supplier_id == who.id))
            if not stock:
                stock = SupplierSourceStock(workspace=workspace, report_month=report_month, t1_supplier_id=t1.id, product_material_id=part.id,
                                            material_code=row["material_code"], sub_supplier_id=who.id)
                db.add(stock)
            stock.quantity = row["quantity"]
            stock.available_quantity = row["quantity"] if row["available_quantity"] is None else row["available_quantity"]
            stock.arrival_month = arrival
            saved += 1
        return saved

    @app.post("/api/planning/supplier-source-stock/{report_month}", status_code=201)
    def save_source_stock_rows(report_month: str, payload: dict, request: Request, workspace: str = "trial"):
        """Stock at a T2 / T3 supplier typed on the supplier overview: rows of t1_supplier_code, part_code,
        material_code, supplier_code, quantity (available = quantity unless given). All rows or none."""
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        month = _month_text(_month(report_month))
        rows = []
        for number, row in enumerate(payload.get("rows") or [], start=1):
            try:
                quantity = float(row.get("quantity"))
                available = None if row.get("available_quantity") in (None, "") else float(row["available_quantity"])
            except (TypeError, ValueError) as exc:
                raise HTTPException(status_code=422, detail=f"Row {number}: stock must be a number") from exc
            rows.append({"t1_supplier_code": str(row.get("t1_supplier_code") or ""), "part_code": str(row.get("part_code") or ""), "material_code": str(row.get("material_code") or ""),
                         "supplier_code": str(row.get("supplier_code") or ""), "quantity": quantity, "available_quantity": available, "arrival_month": "", "_excel_row": number})
        if not rows:
            raise HTTPException(status_code=422, detail="Add at least one row")
        with open_db() as db:
            saved = save_source_stock(db, workspace, month, rows)
            db.commit()
        return {"report_month": month, "rows": saved}

    @app.post("/api/planning/supplier-stock/{report_month}/upload", status_code=201)
    async def upload_supplier_stock_excel(report_month: str, request: Request, workspace: str = "trial"):
        current_user(request)
        _check_supply_role(request)
        _validate_workspace(workspace)
        _month(report_month)
        raw = await request.body()
        source_rows: list[dict] = []
        try:
            rows = _read_xlsx(raw)
            plain_error = None
        except HTTPException as exc:  # e.g. the tree workbook's title rows
            rows, plain_error = [], exc
        if not rows or not {"supplier_code", "material_code", "quantity"} <= set(rows[0]):
            tree_rows = stock_rows_from_tree(raw)  # the tree workbook (template / download), stock typed on T1 part rows
            if tree_rows is not None:
                rows = [row for row in tree_rows if row["tier"] == "T1"]
                source_rows = [row for row in tree_rows if row["tier"] != "T1"]
                if not tree_rows:
                    raise HTTPException(status_code=422, detail="No stock typed: fill Stock quantity on the T1 part or T2 / T3 supplier rows")
                if not rows:
                    with open_db() as db:
                        saved = save_source_stock(db, workspace, _month_text(_month(report_month)), source_rows)
                        db.commit()
                    return {"workspace": workspace, "report_month": _month_text(_month(report_month)), "rows": saved}
            elif plain_error is not None:
                raise plain_error
        _required_columns(rows, {"supplier_code", "material_code", "quantity"})
        with open_db() as db:
            links = db.scalars(select(SupplierPart)).all()
            suppliers = {s.code: s for s in db.scalars(select(Supplier)).all()}
            materials_by_code = {i.code: ("t1", i.id) for i in db.scalars(select(Item)).all()}
            materials_by_code.update({t.code: ("tm", t.id) for t in db.scalars(select(TierMaterial)).all()})
            by_key: dict[tuple, list[SupplierPart]] = {}
            for link in links:
                by_key.setdefault((link.supplier_id, ("t1", link.material_id) if link.material_id else ("tm", link.tier_material_id)), []).append(link)
            suppliers_by_id = {supplier.id: supplier for supplier in suppliers.values()}
            saved = 0
            for row in rows:
                supplier = suppliers.get(str(row.get("supplier_code", "")).strip())
                item = materials_by_code.get(str(row.get("material_code", "")).strip())
                if not supplier or not item:
                    raise HTTPException(status_code=422, detail="Every supplier_code and material_code must exist in the master data")
                candidates = by_key.get((supplier.id, item), [])
                tier = _normalize_tier(row.get("tier")) if str(row.get("tier") or "").strip() else ""
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
                    raise HTTPException(status_code=422, detail=f"No supplier/material relationship configured for {supplier.code} and {row.get('material_code')}")
                if len(candidates) > 1:
                    raise HTTPException(status_code=422, detail=f"Supplier/material relationship is ambiguous for {supplier.code} and {row.get('material_code')}; include tier and parent_supplier_code")
                link = candidates[0]
                quantity = float(row.get("quantity") or 0)
                available = float(row.get("available_quantity")) if row.get("available_quantity") is not None else None
                save_stock(db, workspace, report_month, link, quantity, available, str(row.get("arrival_month") or ""))
                saved += 1
            saved += save_source_stock(db, workspace, _month_text(_month(report_month)), source_rows)
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

    def plan_template_workbook(issue_month: str | None, release_date: str | None) -> StreamingResponse:
        """Plan template: logo and title, then Model Code | Model Name | 24 months (DD-MM-YYYY) from the month of the
        plan date, one row per model. The plan date and revision are set by the app on upload."""
        released = _sheet_date(release_date) if release_date else None
        first_month = _month(released) if released else _month(issue_month or _month_text(date.today()))
        with open_db() as db:
            rows = [(model.code, model.name) for model in db.scalars(select(ProductModel).order_by(ProductModel.code, ProductModel.name)).all()]
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Plan"
        dark, white = PatternFill("solid", fgColor="12263A"), Font(bold=True, color="FFFFFF")
        try:  # the app logo (needs Pillow; the template works without it)
            from openpyxl.drawing.image import Image as SheetImage
            logo = SheetImage(str(Path(__file__).resolve().parent / "static" / "logo.png"))
            logo.width = logo.height = 48
            sheet.add_image(logo, "A1")
        except Exception:
            pass
        sheet.row_dimensions[1].height = 22
        sheet.row_dimensions[2].height = 22
        sheet["B1"] = "Design to Cost · Production plan"
        sheet["B1"].font = Font(bold=True, size=14, color="12263A")
        sheet["B2"] = (f"Plan from {first_month.strftime('%d-%m-%Y')} for 24 months · dates DD-MM-YYYY. On upload pick the plan date in the app: "
                       "each upload is saved as the next revision of that month (Rev.00, Rev.01 ...).")
        sheet["B2"].font = Font(italic=True, color="5B6B7B")
        months = [_plus_months(first_month, offset) for offset in range(24)]
        headers = ["Model Code", "Model Name"] + months
        for index, value in enumerate(headers, start=1):
            cell = sheet.cell(4, index, value)
            cell.fill, cell.font = dark, white
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if index > 2:
                cell.number_format = "DD-MM-YYYY"
        for row_index, (code, name) in enumerate(rows or [("", "")], start=5):
            sheet.cell(row_index, 1, code)
            sheet.cell(row_index, 2, name)
            for column in range(3, 27):
                sheet.cell(row_index, column).number_format = "#,##0.###"
        sheet.column_dimensions["A"].width = 16
        sheet.column_dimensions["B"].width = 26
        for column in range(3, 27):
            sheet.column_dimensions[__import__("openpyxl").utils.get_column_letter(column)].width = 12
        sheet.freeze_panes = "C5"
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                 headers={"Content-Disposition": f'attachment; filename="plan-template-{first_month.strftime("%m-%Y")}.xlsx"'})

    @app.get("/api/planning/templates/{kind}.xlsx")
    def planning_template(kind: str, request: Request, issue_month: str | None = None, release_date: str | None = None):
        current_user(request)
        templates = {
            "actuals": ["model_code", "actual_quantity", "oem_stock"],
            "own-stock": ["material_code", "quantity"],
            "supplier-stock": ["supplier_code", "material_code", "tier", "parent_supplier_code", "quantity", "available_quantity", "arrival_month"],
            "alternate-codes": ["main_code", "alternate_code"],
            "suppliers": ["supplier_code", "supplier_name", "supplier_tier", "category", "country", "contact_person", "email", "phone", "approval"],
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
            return plan_template_workbook(issue_month, release_date)
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
            if not link.material_id:
                raise HTTPException(status_code=422, detail="POs are issued for T1 Material Master codes; this relationship uses a T2/T3 code")
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
            line_prices: dict[int, dict] = {}  # tree index -> price blended with that line's alternate SOB
            children: dict[int, list[BOMComponent]] = {}
            for link in db.scalars(select(BOMComponent).order_by(BOMComponent.id)).all():
                children.setdefault(link.parent_item_id, []).append(link)

            saved_sobs = line_sobs(db, everything=True)
            this_month = _month_text(date.today())

            def emit(depth: int, level: str, path: dict, has_children: bool, quantity=None, item: Item | None = None, line: tuple | None = None) -> None:
                alternates, main_share, next_change = (line_alternates(price, item, saved_sobs.get(line, []), quantity, this_month)
                                                       if item is not None and quantity is not None and price.alternates(item.id) else ([], 100, None))
                later = f", {next_change['main_sob_percent']:g}% from {next_change['from_month']}" if next_change else ""
                if alternates:
                    line_prices[len(tree)] = price.blend(item, prices[item.id], alternates)
                    # Main code row: its own SOB share of the BOM quantity
                    path = {**path, level: (path[level][0], f"{path[level][1]} (main {main_share:g}% of {quantity:g}{later})")}
                    tree.append({"depth": depth, "level": level, "path": path, "has_children": has_children, "quantity": quantity * main_share / 100, "item": item})
                else:
                    tree.append({"depth": depth, "level": level, "path": path, "has_children": has_children, "quantity": quantity, "item": item})
                # Alternate codes follow their main code's line, with their SOB share of its quantity.
                if alternates:
                    for alternate in alternates:
                        tree.append({"depth": depth, "level": level, "has_children": False, "item": items.get(alternate["material_id"]),
                                     "path": {**path, level: (alternate["code"], f"{alternate['name']} (alternate of {item.code}, SOB {alternate['sob_percent']:g}%"
                                                                     + (f", {alternate['next_sob_percent']:g}% from {next_change['from_month']}" if next_change else "") + ")")},
                                     "quantity": quantity * alternate["sob_percent"] / 100, "alternate": alternate})

            def part_lines(parent_id: int, level: str, depth: int, path: dict) -> None:
                for link in children.get(parent_id, []):
                    child = items.get(link.child_item_id)
                    if not child:
                        continue
                    child_path = {**path, level: (child.code, child.name)}
                    deeper = level == "l2" and bool(children.get(child.id))
                    emit(depth, level, child_path, deeper, link.quantity, child, ("component", link.id))
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
                    emit(depth, "l1", line_path, deeper, root.quantity, item, ("root", root.id))
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
        for index, entry in enumerate(tree):
            item = entry["item"]
            values = [BOM_TREE_LEVEL_NAMES[entry["level"]] + (" alternate" if entry.get("alternate") else "")]
            for level in levels:
                values.extend(entry["path"].get(level, ("", "")))
            values += [entry["quantity"] if entry["quantity"] is not None else "", item.uom if item else ""]
            if can_price:
                priced = (entry["alternate"] if entry.get("alternate") else line_prices.get(index) or prices.get(item.id)) if item else None
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
                headers = ["Material code", "Material name", "Used in", "Category", "Unit", "On hand", "Buffer days"]
                if can_price:
                    headers += ["Unit cost", "Currency"]
                levels = bom_levels(db)
                rows = [[item.code, item.name, " & ".join(levels.get(item.id, [])) or "Not used", item.category, item.uom, item.stock_on_hand, getattr(item, "buffer_days", 0), *([item.unit_cost, item.currency] if can_price else [])] for item in materials]
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
            headers = ["Supplier code", "Supplier name", "Material code", "Material name", "Tier", "Parent supplier", "SOB %", "Lead time days", "MOQ", "Buffer days", "Price mode", "Price status"]
            if _check_price_access(request):
                headers += ["Unit price", "Currency"]
            rows = [[row["supplier_code"], row["supplier_name"], row["material_code"], row["material_name"], row["tier"], row.get("parent_supplier_name") or "", row["sob_percent"], row["lead_time_days"], row["moq"], row["buffer_days"], row["price_mode"], row["price_status"], *([row.get("unit_price"), row.get("currency")] if _check_price_access(request) else [])] for row in part_rows]
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
