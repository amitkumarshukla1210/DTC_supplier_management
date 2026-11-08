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
from pydantic import BaseModel, Field
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
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
    def materials(request: Request, offset: int = 0, limit: int = 100, search: str = ""):
        current_user(request)
        if limit < 1 or limit > 500 or offset < 0:
            raise HTTPException(status_code=422, detail="Use limit 1–500 and a non-negative offset")
        with open_db() as db:
            statement = select(Item).order_by(Item.code)
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
            item.code, item.name, item.category, item.uom = code, payload.name.strip(), payload.category.strip(), payload.uom.strip()
            item.stock_on_hand, item.buffer_days = payload.stock_on_hand, payload.buffer_days
            if _check_price_access(request):
                item.unit_cost, item.currency = payload.unit_cost, payload.currency
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
            )
            if referenced:
                raise HTTPException(status_code=409, detail="This material is used by a BOM, supplier, stock snapshot or transfer; remove those references first")
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
            items, _ = _item_map(db)
            roots = db.scalars(select(ModelBOM).where(ModelBOM.model_id == model_id).order_by(ModelBOM.id)).all()
            root_material_ids = [root.material_id for root in roots]
            parents_with_bom = set(db.scalars(
                select(BOMComponent.parent_item_id)
                .where(BOMComponent.parent_item_id.in_(root_material_ids))
                .distinct()
            ).all()) if root_material_ids else set()
            rows = []
            for root in roots:
                item = items.get(root.material_id)
                if not item:
                    continue
                rows.append({
                    "level": 1,
                    "material_id": item.id,
                    "code": item.code,
                    "name": item.name,
                    "quantity": root.quantity,
                    "uom": item.uom,
                    "root_bom_id": root.id,
                    "component_id": None,
                    "parent_item_id": None,
                    "edit_quantity": root.quantity,
                    "has_bom": item.id in parents_with_bom,
                })

            if search.strip():
                needle = search.strip().lower()
                rows = [r for r in rows if needle in r["code"].lower() or needle in r["name"].lower()]
            can_price = _check_price_access(request)
            if can_price:
                for row in rows:
                    row["unit_price"] = items[row["material_id"]].unit_cost
                    row["currency"] = items[row["material_id"]].currency
            return {"model_id": model_id, "model_code": model.code, "model_name": model.name, "tab": tab, "total": len(rows), "offset": offset, "limit": limit, "items": rows[offset:offset + limit]}

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
            return {"items": [
                {"id": item.id, "code": item.code, "name": item.name, "uom": item.uom,
                 "model_ids": sorted(model_ids_by_parent.get(item.id, set()))}
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
            items_by_id = {item.id: item for item in db.scalars(select(Item)).all()}
            child_ids = [link.child_item_id for link in links]
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
                for row in rows:
                    material = items_by_id[row["material_id"]]
                    row.update(unit_price=material.unit_cost, currency=material.currency)
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
                    row.update(unit_price=child.unit_cost, currency=child.currency)
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
            if not db.get(Item, payload.material_id) or not db.get(Supplier, payload.supplier_id):
                raise HTTPException(status_code=404, detail="Material or supplier not found")
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
