"""Generate a large test database: 1,000 suppliers, 10,000 materials, 1,000 products/models,
each model with 20 L1 lines; 300 L1 parts with L2 BOMs (5 lines), 300 L2 parts with L3 BOMs
(4 lines); ~12,000 supplier/material links. Usage: python generate.py <sqlite-path>"""
import random
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from sqlalchemy import create_engine, select  # noqa: E402
from sqlalchemy.orm import Session  # noqa: E402

from app import Base, BOMComponent, Item, Supplier, create_app  # noqa: E402
from planning import ModelBOM, Product, ProductModel, ProductModelProduct, SupplierPart  # noqa: E402

path = Path(sys.argv[1])
if path.exists():
    path.unlink()
create_app(f"sqlite:///{path}")  # creates tables + seed
engine = create_engine(f"sqlite:///{path}")
random.seed(7)
with Session(engine) as db:
    suppliers = [Supplier(code=f"LS-{i:05d}", name=f"Load Supplier {i:04d}", category="Raw material", country="India", status="Approved") for i in range(1000)]
    db.add_all(suppliers)
    db.flush()
    finished = [Item(code=f"LF-{i:05d}", name=f"Assembly {i}", category="Subassembly", uom="PCS", unit_cost=0) for i in range(600)]  # make parts (have BOMs)
    bought = [Item(code=f"LM-{i:05d}", name=f"Bought part {i}", category="Raw material", uom="kg", unit_cost=round(random.uniform(1, 500), 2), currency="INR") for i in range(9400)]
    db.add_all(finished + bought)
    db.flush()
    l1_make, l2_make = finished[:300], finished[300:]
    # L2 BOMs for L1 make parts: 4 bought + 1 L2 make part
    for index, parent in enumerate(l1_make):
        db.add(BOMComponent(parent_item_id=parent.id, child_item_id=l2_make[index].id, quantity=1))
        for child in random.sample(bought, 4):
            db.add(BOMComponent(parent_item_id=parent.id, child_item_id=child.id, quantity=round(random.uniform(0.1, 5), 3)))
    # L3 BOMs for L2 make parts
    for parent in l2_make:
        for child in random.sample(bought, 4):
            db.add(BOMComponent(parent_item_id=parent.id, child_item_id=child.id, quantity=round(random.uniform(0.1, 5), 3)))
    products, models = [], []
    for i in range(1000):
        product = Product(code=f"LP-{i:05d}", name=f"Product {i}")
        db.add(product)
        db.flush()
        model = ProductModel(code=f"LMOD-{i:05d}", name=f"Model {i}", product_id=product.id)
        db.add(model)
        db.flush()
        db.add(ProductModelProduct(product_id=product.id, model_id=model.id))
        lines = random.sample(l1_make, 5) + random.sample(bought, 15)
        for part in lines:
            db.add(ModelBOM(model_id=model.id, material_id=part.id, quantity=1))
    # supplier links: every bought part 1 supplier, 2,600 parts a second supplier
    for part in bought:
        db.add(SupplierPart(material_id=part.id, supplier_id=random.choice(suppliers).id, sob_percent=100 if random.random() > 0.28 else 60,
                            unit_price=part.unit_cost, currency="INR", lead_time_days=30))
    db.commit()
    counts = {name: db.scalar(select(__import__("sqlalchemy").func.count()).select_from(model)) for name, model in
              [("suppliers", Supplier), ("materials", Item), ("models", ProductModel), ("model BOM lines", ModelBOM), ("part BOM lines", BOMComponent), ("supplier links", SupplierPart)]}
    print(counts)
