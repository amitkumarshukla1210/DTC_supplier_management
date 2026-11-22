import io
import math
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app import create_app


class SupplierHubApiTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        db_path = Path(self.temp_dir.name) / "test.db"
        self.client = TestClient(create_app(f"sqlite:///{db_path}"))
        self.client.__enter__()
        setup = self.client.post("/api/auth/setup", json={
            "username": "admin",
            "full_name": "Test Admin",
            "password": "Secure-test-password-2026!",
        })
        self.assertEqual(setup.status_code, 201, setup.text)

    def tearDown(self):
        self.client.__exit__(None, None, None)
        self.temp_dir.cleanup()

    def test_seeded_workspace_loads(self):
        response = self.client.get("/api/dashboard")
        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertGreater(payload["metrics"]["suppliers"], 0)
        self.assertIn("spend_by_currency", payload["metrics"])
        self.assertTrue(payload["orders"])

    def test_homepage_and_brand_asset_are_served(self):
        page = self.client.get("/")
        logo = self.client.get("/static/logo.png")
        frontend = self.client.get("/static/app.js")
        self.assertEqual(page.status_code, 200)
        self.assertIn("Design to Cost", page.text)
        self.assertIn("SUPPLY CHAIN MANAGEMENT", page.text)
        self.assertEqual(logo.status_code, 200)
        self.assertTrue(logo.content.startswith(b"\x89PNG\r\n\x1a\n"))
        self.assertEqual(frontend.status_code, 200)
        self.assertIn("state.products.flatMap", frontend.text)
        self.assertIn('table(treeHeaders("Product Master"), productRows', frontend.text)
        self.assertIn('l1: ["L1 part code", "L1 part name"], l2: ["L2 part code", "L2 part name"], l3: ["L3 part code", "L3 part name"]', frontend.text)
        self.assertIn('row.insertAdjacentHTML("afterend", rows.join(""))', frontend.text)
        self.assertIn('"Model BOM Level 1"', frontend.text)
        self.assertIn('data-tree-expand="model"', frontend.text)
        self.assertIn('data-tree-expand="part"', frontend.text)
        self.assertIn('data-has-bom="${model.has_bom}"', frontend.text)
        self.assertIn('bom-controls-part', frontend.text)
        self.assertIn('attributes: "data-tree-group"', frontend.text)
        self.assertIn('data-action="new-part-bom"', frontend.text)
        self.assertIn('id="categorySelect"', frontend.text)
        self.assertIn('Level ${state.bomTab === "Model BOM Level 1" ? "1"', frontend.text)
        self.assertNotIn('data-model-apply>Execute', frontend.text)
        self.assertIn('aria-label="Select ${esc(state.bomTab)} Excel workbook"', frontend.text)
        self.assertIn('data-model-option', frontend.text)
        self.assertIn('data-execute-master-search', frontend.text)
        self.assertIn('data-parent-search', frontend.text)
        self.assertIn('data-parent-select-all', frontend.text)
        self.assertIn('data-parent-clear', frontend.text)
        self.assertIn('data-parent-option', frontend.text)
        self.assertIn('state.models.map((model) => model.id)', frontend.text)
        self.assertIn('state.bomParentOptions.map((item) => item.id)', frontend.text)
        self.assertNotIn('partBomCodeSearch', frontend.text)
        self.assertNotIn('data-execute-part-search', frontend.text)
        self.assertIn('data-column-apply', frontend.text)
        self.assertIn('data-column-select-all', frontend.text)
        self.assertIn('class="column-resizer"', frontend.text)
        self.assertIn('id="modalMaterialNameSelect"', frontend.text)
        self.assertIn('picker.open = false', frontend.text)
        self.assertIn('modelSelectionAction.closest(".model-picker") || $(".model-picker")', frontend.text)
        self.assertIn('productModelRow.dataset.hasBom === "true" && !event.target.closest("button, a, input, select")', frontend.text)
        self.assertIn('state.selectedModelIds = [Number(level1Row.dataset.modelId)]', frontend.text)
        self.assertIn('state.selectedBomParent = String(level1Row.dataset.materialId)', frontend.text)
        self.assertIn('state.selectedBomParent = String(level2Row.dataset.materialId)', frontend.text)
        self.assertIn('data-bom-level1-row', frontend.text)
        self.assertIn('data-bom-level2-row', frontend.text)
        self.assertIn('data-product-model-row="${model.id}"', frontend.text)
        self.assertIn('id="modalModelCode"', frontend.text)
        self.assertIn('id="modalMaterialName"', frontend.text)
        self.assertIn('id="newPartBomButton"', frontend.text)
        self.assertNotIn('\nif (type === "component")', frontend.text)
        styles = (Path(__file__).parents[1] / "static" / "styles.css").read_text(encoding="utf-8")
        self.assertIn(".part-bom-page th { padding: 10px 7px; font-size: 12px;", styles)

    def test_supplier_to_rfq_quote_award_and_delivery(self):
        response = self.client.post("/api/rfqs", json={})
        self.assertEqual(response.status_code, 410)
        self.assertIn("retired", response.json()["detail"])

    def test_rfq_rejects_supplier_not_approved(self):
        response = self.client.get("/api/rfqs")
        self.assertEqual(response.status_code, 410)

    def test_nested_bom_and_circular_bom_protection(self):
        items = self.client.get("/api/items").json()
        assembly = next(item for item in items if item["code"] == "ASSY-001")
        bom = self.client.get(f"/api/items/{assembly['id']}/bom")
        self.assertEqual(bom.status_code, 200)
        subassembly = next(row for row in bom.json()["children"] if row["item"]["code"] == "SUB-001")
        housing = next(row for row in subassembly["children"] if row["item"]["code"] == "MAT-101")
        billet = next(row for row in housing["children"] if row["item"]["code"] == "MAT-105")
        self.assertEqual(subassembly["level"], 1)
        self.assertEqual(housing["level"], 2)
        self.assertEqual(billet["level"], 3)
        self.assertEqual(billet["quantity_in_assembly"], 1.8)

        updated = self.client.patch(f"/api/items/{housing['item']['id']}/bom/{billet['component_id']}", json={
            "child_item_id": billet["item"]["id"],
            "quantity": 2.25,
        })
        self.assertEqual(updated.status_code, 200, updated.text)
        self.assertEqual(updated.json()["children"][0]["quantity_in_assembly"], 2.25)

        removed = self.client.delete(f"/api/items/{housing['item']['id']}/bom/{billet['component_id']}")
        self.assertEqual(removed.status_code, 204)
        self.assertEqual(housing["quantity_in_assembly"], 1)

        new_component = next(item for item in items if item["code"] == "MAT-102")
        added = self.client.post(f"/api/items/{housing['item']['id']}/bom", json={
            "child_item_id": new_component["id"],
            "quantity": 3,
        })
        self.assertEqual(added.status_code, 201, added.text)
        self.assertTrue(any(row["item"]["id"] == new_component["id"] for row in added.json()["children"]))
        added_component = next(row for row in added.json()["children"] if row["item"]["id"] == new_component["id"])
        negative_update = self.client.patch(f"/api/items/{housing['item']['id']}/bom/{added_component['component_id']}", json={
            "child_item_id": new_component["id"],
            "quantity": -0.25,
        })
        self.assertEqual(negative_update.status_code, 200, negative_update.text)
        negative_row = next(row for row in negative_update.json()["children"] if row["item"]["id"] == new_component["id"])
        self.assertEqual(negative_row["quantity_per_parent"], -0.25)
        self.assertEqual(negative_row["quantity_in_assembly"], -0.25)
        zero_quantity = self.client.post(f"/api/items/{housing['item']['id']}/bom", json={
            "child_item_id": new_component["id"],
            "quantity": 0,
        })
        self.assertEqual(zero_quantity.status_code, 422)
        self.assertEqual(self.client.delete(f"/api/items/{housing['item']['id']}/bom/{added_component['component_id']}").status_code, 204)

        rejected = self.client.post(f"/api/items/{housing['item']['id']}/bom", json={
            "child_item_id": assembly["id"],
            "quantity": 1,
        })
        self.assertEqual(rejected.status_code, 422)
        self.assertIn("circular BOM", rejected.json()["detail"])

    def test_material_master_models_products_and_bom_lines_are_editable(self):
        model = self.client.post("/api/planning/models", json={"code": "CATALOG-MODEL", "name": "Catalog model"})
        self.assertEqual(model.status_code, 201, model.text)
        model_id = model.json()["id"]
        self.assertIsNone(model.json()["product_id"])
        second_model = self.client.post("/api/planning/models", json={
            "code": "CATALOG-MODEL-2", "name": "Second catalog model",
        })
        self.assertEqual(second_model.status_code, 201, second_model.text)

        product = self.client.post("/api/planning/products", json={
            "code": "FRONT-LOAD", "name": "Front Loading", "model_ids": [model_id, second_model.json()["id"]],
        })
        self.assertEqual(product.status_code, 201, product.text)
        product_id = product.json()["id"]
        second_product = self.client.post("/api/planning/products", json={
            "code": "WASHER", "name": "Washer", "model_ids": [model_id],
        })
        self.assertEqual(second_product.status_code, 201, second_product.text)
        model_list = self.client.get("/api/planning/models").json()
        catalog_model = next(row for row in model_list if row["id"] == model_id)
        self.assertCountEqual(catalog_model["product_ids"], [product_id, second_product.json()["id"]])
        product_list = self.client.get("/api/planning/products").json()
        front_load = next(row for row in product_list if row["id"] == product_id)
        self.assertCountEqual(front_load["model_ids"], [model_id, second_model.json()["id"]])
        self.assertEqual(self.client.patch(f"/api/planning/products/{product_id}", json={
            "code": "FRONT-LOAD", "name": "Front Loading Updated", "model_ids": [],
        }).status_code, 200)

        duplicate_model = self.client.post("/api/planning/models", json={"code": "catalog-model", "name": "Duplicate"})
        self.assertEqual(duplicate_model.status_code, 409)
        material = self.client.post("/api/planning/materials", json={
            "code": "UNIQUE-MAT", "name": "Unique material", "category": "Raw material", "uom": "PCS",
        })
        self.assertEqual(material.status_code, 201, material.text)
        material_id = material.json()["id"]
        self.assertEqual(self.client.post("/api/planning/materials", json={
            "code": "unique-mat", "name": "Duplicate material", "category": "Raw material", "uom": "PCS",
        }).status_code, 409)
        update_material = self.client.patch(f"/api/planning/materials/{material_id}", json={
            "code": "UNIQUE-MAT", "name": "Updated material", "category": "Component", "uom": "PCS",
            "stock_on_hand": 2, "buffer_days": 4,
        })
        self.assertEqual(update_material.status_code, 200, update_material.text)
        self.assertEqual(update_material.json()["name"], "Updated material")
        self.assertEqual(self.client.post(f"/api/planning/model-bom/{model_id}", json={
            "material_id": material_id, "quantity": 1, "uom": "ea",
        }).status_code, 422)

        bom = self.client.post(f"/api/planning/model-bom/{model_id}", json={
            "material_id": material_id, "quantity": 2, "uom": "PCS",
        })
        self.assertEqual(bom.status_code, 201, bom.text)
        bom_id = bom.json()["id"]
        updated_bom = self.client.patch(f"/api/planning/model-bom/{model_id}/{bom_id}", json={
            "material_id": material_id, "quantity": 3, "uom": "PCS",
        })
        self.assertEqual(updated_bom.status_code, 200, updated_bom.text)
        self.assertEqual(updated_bom.json()["quantity"], 3)
        listed_bom = self.client.get(f"/api/planning/model-bom/{model_id}?tab=OEM%20BOM").json()
        self.assertEqual(listed_bom["items"][0]["root_bom_id"], bom_id)
        self.assertTrue(all(row["level"] == 1 for row in listed_bom["items"]))
        self.assertEqual(self.client.delete(f"/api/planning/materials/{material_id}").status_code, 409)
        self.assertEqual(self.client.delete(f"/api/planning/model-bom/{model_id}/{bom_id}").status_code, 204)
        self.assertEqual(self.client.delete(f"/api/planning/materials/{material_id}").status_code, 204)
        self.assertEqual(self.client.delete(f"/api/planning/products/{product_id}").status_code, 204)

    def test_material_request_purchase_order_updates_inventory(self):
        requests = self.client.get("/api/material-requests").json()
        request = next(row for row in requests if row["number"] == "MR-26012")
        initial_stock = self.client.get("/api/items").json()
        item = next(row for row in initial_stock if row["id"] == request["item_id"])

        created = self.client.post(f"/api/material-requests/{request['id']}/purchase-order", json={})
        self.assertEqual(created.status_code, 201, created.text)
        order = created.json()
        self.assertEqual(order["order_type"], "stock")
        self.assertEqual(order["quantity_ordered"], 28)
        duplicate_order = self.client.post(f"/api/material-requests/{request['id']}/purchase-order", json={})
        self.assertEqual(duplicate_order.status_code, 409)

        partial = self.client.post(f"/api/stock-purchase-orders/{order['id']}/deliveries", json={
            "quantity_received": 8,
            "delivered_at": date.today().isoformat(),
            "quality_accepted": True,
        })
        self.assertEqual(partial.status_code, 201, partial.text)
        self.assertEqual(partial.json()["status"], "Partially received")
        self.assertEqual(partial.json()["progress"], 29)

        updated_items = self.client.get("/api/items").json()
        updated_item = next(row for row in updated_items if row["id"] == item["id"])
        self.assertEqual(updated_item["stock_on_hand"], item["stock_on_hand"] + 8)

        too_many = self.client.post(f"/api/stock-purchase-orders/{order['id']}/deliveries", json={
            "quantity_received": 21,
        })
        self.assertEqual(too_many.status_code, 422)

        final = self.client.post(f"/api/stock-purchase-orders/{order['id']}/deliveries", json={
            "quantity_received": 20,
            "quality_accepted": True,
        })
        self.assertEqual(final.status_code, 201, final.text)
        self.assertEqual(final.json()["status"], "Received")
        request_after = next(row for row in self.client.get("/api/material-requests").json() if row["id"] == request["id"])
        self.assertEqual(request_after["status"], "Received")
        self.assertEqual(request_after["quantity_received"], 28)

    def test_reorder_requests_are_generated_once_per_item(self):
        first = self.client.post("/api/material-requests/reorder", json={})
        self.assertEqual(first.status_code, 201, first.text)
        self.assertTrue(first.json())
        second = self.client.post("/api/material-requests/reorder", json={})
        self.assertEqual(second.status_code, 201)
        self.assertEqual(second.json(), [])

    def test_item_supplier_and_lead_time_are_editable(self):
        item = next(item for item in self.client.get("/api/items").json() if item["code"] == "MAT-101")
        response = self.client.patch(f"/api/items/{item['id']}", json={"lead_time_days": 31})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["lead_time_days"], 31)
        self.assertEqual(response.json()["preferred_supplier_name"], "Apex Precision Works")

    def test_planning_revisions_roll_up_bom_and_report_actuals_and_supplier_stock(self):
        product = self.client.post("/api/planning/products", json={"code": "PRD-TEST", "name": "Planning product"})
        self.assertEqual(product.status_code, 201, product.text)
        model = self.client.post("/api/planning/models", json={
            "product_id": product.json()["id"], "code": "MODEL-TEST", "name": "Test model",
        })
        self.assertEqual(model.status_code, 201, model.text)
        materials = self.client.get("/api/planning/materials?limit=500").json()["items"]
        material = next(row for row in materials if row["code"] == "MAT-104")
        bom = self.client.post(f"/api/planning/model-bom/{model.json()['id']}", json={
            "material_id": material["id"], "quantity": 2, "uom": material["uom"],
        })
        self.assertEqual(bom.status_code, 201, bom.text)

        suppliers = self.client.get("/api/suppliers").json()
        supplier = next(row for row in suppliers if row["status"] == "Approved")
        supplier_part = self.client.post("/api/planning/supplier-parts", json={
            "material_id": material["id"], "supplier_id": supplier["id"], "tier": "Direct",
            "sob_percent": 100, "lead_time_days": 200, "buffer_days": 30,
            "unit_price": 5, "currency": "USD",
        })
        self.assertEqual(supplier_part.status_code, 201, supplier_part.text)

        base_plan = {
            "issue_month": "2026-10", "workspace": "trial",
            "rows": [{"model_code": "MODEL-TEST", "target_month": "2026-10", "quantity": 100}],
        }
        first = self.client.post("/api/planning/plans", json=base_plan)
        second = self.client.post("/api/planning/plans", json=base_plan)
        self.assertEqual(first.status_code, 201, first.text)
        self.assertEqual(first.json()["revision"], "Rev.00")
        self.assertEqual(second.json()["revision"], "Rev.01")

        actual = self.client.post("/api/planning/actuals", json={
            "period": "2026-10", "workspace": "trial",
            "rows": [{"model_code": "MODEL-TEST", "actual_quantity": 40, "oem_stock": 12}],
        })
        self.assertEqual(actual.status_code, 201, actual.text)
        own_stock = self.client.post("/api/planning/own-stock", json={
            "period": "2026-10", "workspace": "trial",
            "rows": [{"material_code": material["code"], "quantity": 10}],
        })
        self.assertEqual(own_stock.status_code, 201, own_stock.text)
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["supplier_code", "material_code", "quantity", "available_quantity", "arrival_month"])
        sheet.append([supplier["code"], material["code"], 25, 10, "2026-10"])
        upload = io.BytesIO()
        workbook.save(upload)
        stock = self.client.post(
            "/api/planning/supplier-stock/2026-10/upload?workspace=trial",
            content=upload.getvalue(),
            headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        )
        self.assertEqual(stock.status_code, 201, stock.text)

        report = self.client.get("/api/planning/report?start_month=2026-10&end_month=2026-10&workspace=trial")
        self.assertEqual(report.status_code, 200, report.text)
        row = next(r for r in report.json()["items"] if r["material_code"] == material["code"] and r["supplier_name"] == supplier["name"])
        self.assertEqual(row["plan_demand"], 200)
        self.assertEqual(row["actual_demand"], 80)
        self.assertEqual(row["supplier_stock"], 10)
        self.assertEqual(row["own_stock"], 10)
        self.assertEqual(row["own_stock_allocated"], 10)
        self.assertEqual(row["recommended_po_qty"], row["shortage_qty"])
        self.assertEqual(row["criticality"], "Very High")
        self.assertEqual(row["plan_revision"], "Rev.01")

        excel = self.client.get("/api/planning/report.xlsx?start_month=2026-10&end_month=2026-10&workspace=trial")
        pdf = self.client.get("/api/planning/report.pdf?start_month=2026-10&end_month=2026-10&workspace=trial")
        self.assertEqual(excel.status_code, 200)
        self.assertTrue(excel.content.startswith(b"PK"))
        report_workbook = load_workbook(io.BytesIO(excel.content), read_only=False)
        self.assertEqual(report_workbook.active["D1"].value, "Design to Cost")
        self.assertEqual(report_workbook.active["A4"].value, "Month")
        self.assertEqual(len(report_workbook.active._images), 1)
        report_workbook.close()
        self.assertEqual(pdf.status_code, 200)
        self.assertTrue(pdf.content.startswith(b"%PDF"))

    def test_user_price_permission_and_trial_reset_are_scoped(self):
        material = next(row for row in self.client.get("/api/planning/materials?limit=500").json()["items"] if row["code"] == "MAT-104")
        supplier = next(row for row in self.client.get("/api/suppliers").json() if row["status"] == "Approved")
        supplier_part = self.client.post("/api/planning/supplier-parts", json={
            "material_id": material["id"], "supplier_id": supplier["id"], "tier": "Direct",
            "sob_percent": 100, "lead_time_days": 20, "unit_price": 12, "currency": "USD",
        })
        self.assertEqual(supplier_part.status_code, 201, supplier_part.text)
        po = self.client.post("/api/planning/purchase-orders", json={
            "supplier_part_id": supplier_part.json()["id"], "quantity_ordered": 5,
            "expected_month": date.today().strftime("%Y-%m"), "workspace": "trial",
        })
        self.assertEqual(po.status_code, 201, po.text)
        report_user = self.client.post("/api/users", json={
            "username": "supplychain", "full_name": "Supply Chain",
            "password": "Secure-supplychain-password-2026!", "role": "Supply Chain",
            "can_view_prices": False,
        })
        self.assertEqual(report_user.status_code, 201, report_user.text)
        self.client.post("/api/auth/logout")
        self.client.post("/api/auth/login", json={
            "username": "supplychain", "password": "Secure-supplychain-password-2026!",
        })
        workbook_response = self.client.get(
            f"/api/planning/supplier-business.xlsx?start_month={date.today():%Y-%m}&end_month={date.today():%Y-%m}&workspace=trial"
        )
        self.assertEqual(workbook_response.status_code, 200, workbook_response.text)
        workbook = load_workbook(io.BytesIO(workbook_response.content), read_only=True)
        self.assertEqual(workbook.active["F5"].value, "Restricted")
        workbook.close()
        workbook = load_workbook(io.BytesIO(workbook_response.content), read_only=False)
        self.assertEqual(len(workbook.active._images), 1)
        workbook.close()
        self.client.post("/api/auth/logout")
        self.client.post("/api/auth/login", json={
            "username": "admin", "password": "Secure-test-password-2026!",
        })

        created = self.client.post("/api/users", json={
            "username": "planner", "full_name": "Planning User",
            "password": "Secure-planner-password-2026!", "role": "User",
            "can_view_prices": False,
        })
        self.assertEqual(created.status_code, 201, created.text)
        self.client.post("/api/auth/logout")
        login = self.client.post("/api/auth/login", json={
            "username": "planner", "password": "Secure-planner-password-2026!",
        })
        self.assertEqual(login.status_code, 200, login.text)
        materials = self.client.get("/api/planning/materials?limit=1").json()["items"]
        self.assertNotIn("unit_cost", materials[0])
        self.assertEqual(self.client.get("/api/planning/supplier-parts").status_code, 403)
        self.assertEqual(self.client.delete("/api/planning/trial-data").status_code, 403)
        self.client.post("/api/auth/logout")
        self.client.post("/api/auth/login", json={
            "username": "admin", "password": "Secure-test-password-2026!",
        })
        deleted = self.client.delete("/api/planning/trial-data")
        self.assertEqual(deleted.status_code, 204, deleted.text)
        self.assertEqual(self.client.get("/api/planning/trial-data/counts").json()["plans"], 0)

    def test_planning_po_receipts_schedule_and_supplier_transfer(self):
        material = next(row for row in self.client.get("/api/planning/materials?limit=500").json()["items"] if row["code"] == "MAT-104")
        approved = [row for row in self.client.get("/api/suppliers").json() if row["status"] == "Approved"]
        first, second = approved[:2]
        links = []
        for supplier in (first, second):
            response = self.client.post("/api/planning/supplier-parts", json={
                "material_id": material["id"], "supplier_id": supplier["id"], "tier": "Direct",
                "sob_percent": 50, "lead_time_days": 30, "moq": 5, "unit_price": 3.5,
                "currency": "USD",
            })
            self.assertEqual(response.status_code, 201, response.text)
            links.append(response.json())

        order = self.client.post("/api/planning/purchase-orders", json={
            "supplier_part_id": links[0]["id"], "quantity_ordered": 10,
            "expected_month": "2026-10", "workspace": "trial",
        })
        self.assertEqual(order.status_code, 201, order.text)
        schedule = self.client.patch(f"/api/planning/purchase-orders/{order.json()['id']}/schedule", json={
            "expected_month": "2026-11", "change_note": "Align delivery with revised production plan",
        })
        self.assertEqual(schedule.status_code, 200, schedule.text)
        self.assertEqual(schedule.json()["expected_month"], "2026-11")

        receipt = self.client.post(f"/api/planning/purchase-orders/{order.json()['id']}/receipts", json={
            "quantity_received": 4, "received_month": "2026-11",
        })
        self.assertEqual(receipt.status_code, 201, receipt.text)
        self.assertEqual(receipt.json()["status"], "Partially received")
        completed = self.client.post(f"/api/planning/purchase-orders/{order.json()['id']}/receipts", json={
            "quantity_received": 6, "received_month": "2026-11",
        })
        self.assertEqual(completed.status_code, 201, completed.text)
        self.assertEqual(completed.json()["status"], "Received")

        transfer = self.client.post("/api/planning/transfers", json={
            "material_id": material["id"], "from_supplier_id": first["id"], "to_supplier_id": second["id"],
            "quantity": 3, "transfer_month": "2026-11", "workspace": "trial", "note": "Rebalance stock",
        })
        self.assertEqual(transfer.status_code, 201, transfer.text)
        self.assertEqual(self.client.get("/api/planning/transfers?workspace=trial").json()[0]["quantity"], 3)

    def test_plan_uploads_get_next_revision_on_the_picked_date(self):
        # The user's way: in Nov-26 the 24-month plan is uploaded on 01-11, 10-11 and 15-11; each upload = next revision
        api = self.client
        product = api.post("/api/planning/products", json={"code": "RV-PRD", "name": "Fans"}).json()["id"]
        api.post("/api/planning/models", json={"product_id": product, "code": "RV12000", "name": "Eva 8014"})
        api.post("/api/planning/models", json={"product_id": product, "code": "RV12001", "name": "Serena 8014"})
        xlsx = {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        # template: no Release date / Plan revision columns; 24 months from the plan date's month
        template = load_workbook(io.BytesIO(api.get("/api/planning/templates/plan.xlsx", params={"release_date": "2027-02-10"}).content)).active
        heads = [template.cell(4, c).value for c in range(1, 28)]
        self.assertEqual(heads[:2], ["Model Code", "Model Name"])
        self.assertEqual((heads[2].date(), heads[25].date(), heads[26]), (date(2027, 2, 1), date(2029, 1, 1), None))
        self.assertEqual([(template.cell(r, 1).value, template.cell(r, 2).value) for r in (5, 6)], [("RV12000", "Eva 8014"), ("RV12001", "Serena 8014")])

        def sheet_with(eva, serena, extra=()):
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Design to Cost · Production plan"])
            sheet.append([])
            sheet.append([])
            sheet.append(["Model Code", "Model Name"] + [f"01-{(10 + m) % 12 + 1:02d}-{2026 + (10 + m) // 12}" for m in range(24)])
            sheet.append(["RV12000", "Eva 8014"] + [eva] * 24)
            sheet.append(["RV12001", "Serena 8014"] + [serena] * 24)
            for row in extra:
                sheet.append(row)
            content = io.BytesIO()
            workbook.save(content)
            return content.getvalue()

        upload = lambda day, body: api.post(f"/api/planning/plans/2026-11/upload?workspace=trial&release_date={day}", content=body, headers=xlsx)
        results = [upload(day, sheet_with(qty, qty * 2, [[None], ["Total", ""] + [qty * 3] * 24])) for day, qty in (("2026-11-01", 100), ("10-11-2026", 200), ("2026-11-15", 300))]
        self.assertEqual([(r.status_code, r.json()["revision"], r.json()["issue_month"], r.json()["release_date"], r.json()["rows"]) for r in results],
                         [(201, "Rev.00", "2026-11", "2026-11-01", 48), (201, "Rev.01", "2026-11", "2026-11-10", 48), (201, "Rev.02", "2026-11", "2026-11-15", 48)])
        listed = {p["revision"]: p["release_date"] for p in api.get("/api/planning/plans", params={"workspace": "trial", "issue_month": "2026-11"}).json()}
        self.assertEqual(listed, {"Rev.00": "2026-11-01", "Rev.01": "2026-11-10", "Rev.02": "2026-11-15"})
        by_name = {item["model_name"]: item["quantities"] for item in api.get(f"/api/planning/plans/{results[1].json()['id']}/rows", params={"limit": 100}).json()["items"]}
        self.assertEqual((by_name["Eva 8014"]["2026-11"], by_name["Eva 8014"]["2028-10"], by_name["Serena 8014"]["2027-05"]), (200, 200, 400))
        # one plan per upload; a model name that does not match its code is reported (nothing saved)
        two = sheet_with(1, 2, [[None], ["Model Code", "Model Name", "01-11-2026"], ["RV12000", "Eva 8014", 5]])
        self.assertIn("keep one plan", upload("2026-11-20", two).json()["detail"])
        wrong = sheet_with(1, 2, [["RV12000", "Serena  8014"] + [1] * 24])
        self.assertIn("Row 7: RV12000 is Eva 8014 in Product Master, not Serena  8014", upload("2026-11-20", wrong).json()["detail"])
        # a plan month before the plan date's month is outside the horizon
        self.assertIn("24-month horizon", upload("2026-12-01", sheet_with(1, 2)).json()["detail"])
        self.assertEqual(len(api.get("/api/planning/plans", params={"workspace": "trial", "issue_month": "2026-11"}).json()), 3)

    def test_actual_production_till_date_with_plan_template_and_upload(self):
        api = self.client
        xlsx = {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        product = api.post("/api/planning/products", json={"code": "AC-PRD", "name": "Fans"}).json()["id"]
        for code, name in (("AC12000", "Eva 8014"), ("AC12001", "Serena 8014"), ("AC12002", "Old model")):
            api.post("/api/planning/models", json={"product_id": product, "code": code, "name": name})
        # Nov-26 plan, two revisions: the latest (Rev.01, 10-11) is the plan in force; AC12002 not planned
        for day, eva, serena in (("2026-11-01", 20000, 15000), ("2026-11-10", 18000, 16000)):
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Model Code", "Model Name", "01-11-2026", "01-12-2026"])
            sheet.append(["AC12000", "Eva 8014", eva, 1])
            sheet.append(["AC12001", "Serena 8014", serena, 1])
            content = io.BytesIO()
            workbook.save(content)
            self.assertEqual(api.post(f"/api/planning/plans/2026-11/upload?workspace=trial&release_date={day}", content=content.getvalue(), headers=xlsx).status_code, 201)
        template_buffer = io.BytesIO(api.get("/api/planning/actuals/template.xlsx", params={"as_of": "2026-11-15"}).content)  # kept open: the logo image reads from it on save
        template = load_workbook(template_buffer).active
        self.assertEqual((template["B1"].value, len(template._images)), ("Design to Cost · Actual production", 1))
        self.assertIn("Month: Nov-2026 · as on 15-11-2026 · plan Rev.01 of 10-11-2026", template["B2"].value)
        self.assertEqual([template.cell(4, c).value for c in range(1, 6)], ["Model Code", "Model Name", "Plan qty (Rev.01)", "Actual production (till date)", "OEM stock"])
        self.assertEqual([[template.cell(r, c).value for c in range(1, 6)] for r in (5, 6, 7)], [["AC12000", "Eva 8014", 18000, None, None], ["AC12001", "Serena 8014", 16000, None, None], [None] * 5])
        # 15-11: the total produced from 01-11 till 15-11 (Serena left blank: skipped)
        template.cell(5, 4).value, template.cell(5, 5).value = 9000, 1200
        template._images = []  # openpyxl cannot re-save an image read from memory (Excel keeps it)
        content = io.BytesIO()
        template.parent.save(content)
        first = api.post("/api/planning/actuals/2026-11/upload?workspace=trial&as_of=2026-11-15", content=content.getvalue(), headers=xlsx)
        self.assertEqual((first.status_code, first.json()["rows"], first.json()["as_of"]), (201, 1, "2026-11-15"), first.text)
        view = api.get("/api/planning/actuals", params={"period": "2026-11", "with_plan": "true"}).json()
        self.assertEqual(view["plan"]["revision"], "Rev.01")
        got = [(r["model_code"], r["plan_quantity"], r["actual_quantity"], r["oem_stock"], r["balance"], r["achieved_percent"], r["as_of"]) for r in view["items"]]
        self.assertEqual(got, [("AC12000", 18000, 9000, 1200, 9000, 50.0, "2026-11-15"), ("AC12001", 16000, None, None, 16000, 0.0, None)])
        # 25-11: the new total replaces it (typed in the table: JSON rows)
        again = api.post("/api/planning/actuals", json={"period": "2026-11", "workspace": "trial", "as_of": "25-11-2026", "rows": [{"model_code": "AC12000", "actual_quantity": 16000, "oem_stock": 800}, {"model_code": "AC12001", "actual_quantity": 12000, "oem_stock": 0}]})
        self.assertEqual(again.status_code, 201, again.text)
        rows = {r["model_code"]: r for r in api.get("/api/planning/actuals", params={"period": "2026-11", "with_plan": "true"}).json()["items"]}
        self.assertEqual((rows["AC12000"]["actual_quantity"], rows["AC12000"]["achieved_percent"], rows["AC12000"]["as_of"], rows["AC12001"]["balance"]), (16000, 88.9, "2026-11-25", 4000))
        self.assertEqual(len(api.get("/api/planning/actuals", params={"period": "2026-11"}).json()), 2)  # plain list: models with an actual
        # errors: a date outside the month, a name that is not the code's, nothing typed
        self.assertIn("is not in 2026-11", api.post("/api/planning/actuals", json={"period": "2026-11", "as_of": "2026-12-01", "rows": [{"model_code": "AC12000", "actual_quantity": 1}]}).json()["detail"])
        template.cell(6, 2).value, template.cell(6, 4).value = "Eva 8014", 5
        content = io.BytesIO()
        template.parent.save(content)
        self.assertIn("Row 6: AC12001 is Serena 8014 in Product Master, not Eva 8014", api.post("/api/planning/actuals/2026-11/upload?as_of=2026-11-26", content=content.getvalue(), headers=xlsx).json()["detail"])
        self.assertIn("No actual production typed", api.post("/api/planning/actuals", json={"period": "2026-11", "rows": [{"model_code": "AC12000", "actual_quantity": ""}]}).json()["detail"])

    def test_demand_and_po_plan_t1_t2_t3_month_wise(self):
        # Plan: Eva 60,000 in Nov-26 and 30,000 in Dec-26, 1 beam each. Savera makes the beam at 60% (MOQ 5,000, lead 45 days,
        # FG stock 9,000 in Nov, open PO 5,000 due Dec). Beam: 2 kg steel from JGI 30% (MOQ 10,000, lead 20 days, stock 5,000)
        # and TATA 70%; per kg steel 1.7 kg ore from XYZ.
        api = self.client
        xlsx = {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        # your opening stock (01st month) of the beam 10,000 and buffer norm 15 days
        beam = api.post("/api/planning/materials", json={"code": "DP-BEAM", "name": "Base Front Beam", "uom": "Nos", "stock_on_hand": 10000, "buffer_days": 15}).json()["id"]
        steel = api.post("/api/planning/tier-materials", json={"tier": "T2", "code": "DP-STEEL", "name": "STEEL0.7", "uom": "kg"}).json()["id"]
        ore = api.post("/api/planning/tier-materials", json={"tier": "T3", "code": "DP-ORE", "name": "Iron ore", "uom": "kg"}).json()["id"]
        product = api.post("/api/planning/products", json={"code": "DP-PRD", "name": "Fans"}).json()["id"]
        eva = api.post("/api/planning/models", json={"product_id": product, "code": "DP12000", "name": "Eva 8014"}).json()["id"]
        api.post(f"/api/planning/model-bom/{eva}", json={"material_id": beam, "quantity": 1, "uom": "Nos"})
        sup = lambda code, name: api.post("/api/suppliers", json={"code": code, "name": name, "category": "Sheet Metal", "country": "India"}).json()["id"]
        savera, jgi, tata, xyz = sup("DP-1013", "DP Savera"), sup("DP-1004", "DP JGI"), sup("DP-1014", "DP TATA"), sup("DP-1030", "DP XYZ")
        link = api.post("/api/planning/supplier-parts", json={"material_id": beam, "supplier_id": savera, "sob_percent": 60, "moq": 5000, "lead_time_days": 45, "unit_price": 85, "currency": "INR"}).json()["id"]
        steel_line = api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "tier_material_id": steel, "quantity": 2, "sources": [
            {"sub_supplier_id": jgi, "sob_percent": 30, "moq": 10000, "lead_time_days": 20, "unit_price": 74, "currency": "INR"}, {"sub_supplier_id": tata, "sob_percent": 70, "unit_price": 72, "currency": "INR"}]}).json()["id"]
        api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "parent_line_id": steel_line, "tier_material_id": ore, "quantity": 1.7, "sources": [{"sub_supplier_id": xyz, "sob_percent": 100}]})
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Model Code", "Model Name", "01-11-2026", "01-12-2026"])
        sheet.append(["DP12000", "Eva 8014", 60000, 30000])
        content = io.BytesIO()
        workbook.save(content)
        self.assertEqual(api.post("/api/planning/plans/2026-11/upload?workspace=trial&release_date=2026-11-10", content=content.getvalue(), headers=xlsx).status_code, 201)
        api.post("/api/planning/supplier-stock/2026-11?workspace=trial", json={"rows": [{"supplier_part_id": link, "quantity": 9000}]})
        api.post("/api/planning/supplier-source-stock/2026-11?workspace=trial", json={"rows": [{"t1_supplier_code": "DP-1013", "part_code": "DP-BEAM", "material_code": "DP-STEEL", "supplier_code": "DP-1004", "quantity": 5000}]})
        self.assertEqual(api.post("/api/planning/purchase-orders", json={"supplier_part_id": link, "quantity_ordered": 5000, "expected_month": "2026-12", "workspace": "trial", "ncnr": False}).status_code, 201)
        plan = lambda tier: api.get("/api/planning/demand-plan", params={"tier": tier, "start_month": "2026-11", "months": 2, "search": "DP-"}).json()
        cells = lambda row: [(c["month"], c["requirement"], c["net"], c["po"], c["order_by"]) for c in row["months"]]
        # T1 (your PO to Savera), worked out per part. Savera's FG stock is NOT counted.
        # Nov: need 60,000, buffer 60,000 / 30.4375 x 15 (29,568.8), you have 10,000 -> order 79,568.8; Savera 60% = 47,741.3 -> 50,000 (MOQ 5,000).
        # What arrives: Savera's 50,000 + the 40% not linked here (assumed as planned) -> Nov closing 31,827.5.
        # Dec: 31,827.5 + open PO 5,000 - 30,000 = 6,827.5; buffer 14,784.4 -> order 7,956.9; 60% = 4,774.1 -> 5,000.
        b1 = 60000 / 30.4375 * 15
        o1 = 60000 + b1 - 10000
        c1 = 10000 - 60000 + o1 * 0.4 + 50000
        b2 = 30000 / 30.4375 * 15
        o2 = b2 - (c1 + 5000 - 30000)
        c2 = c1 + 5000 - 30000 + o2 * 0.4 + 5000
        r4 = lambda value: round(value, 4)
        t1 = plan("T1")
        self.assertEqual((t1["months"], t1["plans"]), (["2026-11", "2026-12"], {"2026-11": "Rev.00", "2026-12": "Rev.00"}))
        row = t1["rows"][0]
        self.assertEqual((row["key"], row["supplier_code"], row["code"], row["for"], row["stock"], row["buffer_days"], row["supplier_fg_stock"]), (f"T1:{link}", "DP-1013", "DP-BEAM", "OEM", 10000, 15, 9000))
        got = [(c["month"], c["requirement"], round(c["net"], 4), c["po"], c["actual"], c["saved"], c["order_by"], round(c["closing"], 4)) for c in row["months"]]
        self.assertEqual(got, [("2026-11", 36000, r4(o1 * 0.6), 50000, 50000, False, "2026-09", r4(c1)), ("2026-12", 18000, r4(o2 * 0.6), 5000, 5000, False, "2026-10", r4(c2))])
        # T2: Savera's balance demand = your actual PO - its FG stock 9,000: Nov 41,000, Dec 5,000; x 2 kg x SOB.
        # JGI's own stock (5,000) is its property: not deducted. JGI 30%: Nov 24,600 -> 30,000 (MOQ 10,000) by Oct; the extra 5,400
        # carries: Dec 3,000 is covered -> no PO. TATA 70%: 57,400 and 7,000 (no MOQ, no lead time).
        t2 = {r["supplier_code"]: r for r in plan("T2")["rows"]}
        self.assertEqual((t2["DP-1004"]["key"], t2["DP-1004"]["for"], t2["DP-1004"]["code"], t2["DP-1004"]["stock"], round(t2["DP-1004"]["sob_percent"], 4), t2["DP-1004"]["unit_price"]),
                         (f"T2:{jgi}:dp-steel:{savera}", "DP Savera", "DP-STEEL", 5000, 30, 74))
        self.assertEqual(cells(t2["DP-1004"]), [("2026-11", 24600, 24600, 30000, "2026-10"), ("2026-12", 3000, 0, 0, None)])
        self.assertEqual(cells(t2["DP-1014"]), [("2026-11", 57400, 57400, 57400, "2026-11"), ("2026-12", 7000, 7000, 7000, "2026-12")])
        # T3: ore 1.7 kg per kg steel = 3.4 kg per beam x Savera's balance demand
        t3 = plan("T3")["rows"]
        self.assertEqual([(r["supplier_code"], r["code"]) for r in t3], [("DP-1030", "DP-ORE")])
        self.assertEqual([c["requirement"] for c in t3[0]["months"]], [139400, 17000])
        # Excel: blocks Requirement | PO demand | Order by | Actual PO, each with the months
        book = load_workbook(io.BytesIO(api.get("/api/planning/demand-plan.xlsx", params={"tier": "T2", "start_month": "2026-11", "months": 2, "search": "DP-"}).content)).active
        self.assertEqual([book.cell(3, c).value for c in (11, 13, 15, 17)], ["Requirement", "PO demand", "Order by", "Actual PO"])
        self.assertEqual([book.cell(4, c).value for c in (11, 12, 17, 18)], ["2026-11", "2026-12", "2026-11", "2026-12"])
        self.assertEqual([book.cell(5, c).value for c in (1, 3, 5, 11, 12, 13, 14, 15, 17, 18)], ["DP-1004", "DP-STEEL", "DP Savera", 24600, 3000, 30000, 0, "2026-10", 30000, 0])

        # Actual PO: you issue 40,000 in Nov instead of 50,000 -> Nov closing 10,000 less (21,827.5); Dec: 21,827.5 + 5,000 - 30,000
        # = -3,172.5, order 14,784.4 + 3,172.5 = 17,956.9; 60% = 10,774.1 -> PO demand 15,000. Savera's balance Nov 40,000 - 9,000 = 31,000:
        # JGI 18,600 -> 20,000; Dec balance 15,000 -> JGI 9,000 - 1,400 carried = 7,600 -> 10,000.
        save = lambda key, month, quantity: api.post("/api/planning/demand-plan/actual-po", json={"workspace": "trial", "row_key": key, "month": month, "quantity": quantity})
        self.assertEqual(save(f"T1:{link}", "2026-11", 40000).status_code, 200)
        row = plan("T1")["rows"][0]
        self.assertEqual([(c["po"], c["actual"], c["saved"], round(c["net"], 4), round(c["closing"], 4)) for c in row["months"]],
                         [(50000, 40000, True, r4(o1 * 0.6), r4(c1 - 10000)), (15000, 15000, False, r4((b2 - (c1 - 10000 + 5000 - 30000)) * 0.6), r4(c1 - 10000 + 5000 - 30000 + (b2 - (c1 - 10000 + 5000 - 30000)) * 0.4 + 15000))])
        jgi_row = {r["supplier_code"]: r for r in plan("T2")["rows"]}["DP-1004"]
        self.assertEqual(cells(jgi_row), [("2026-11", 18600, 18600, 20000, "2026-10"), ("2026-12", 9000, 7600, 10000, "2026-11")])
        self.assertEqual(save(f"T1:{link}", "2026-11", -5).status_code, 422)
        self.assertEqual(save("T1:abc", "2026-11", 5).status_code, 422)
        self.assertEqual(save(f"T1:{link}", "2026-11", None).status_code, 200)  # blank = back to the PO demand
        self.assertEqual([c["actual"] for c in plan("T1")["rows"][0]["months"]], [50000, 5000])

        # Rev. Change Impact: Nov Rev.01 drops Dec to 15,000 (-50 %). Savera's lead time 45 days = 2 months: Nov and Dec POs
        # (order by Sep / Oct, on or before Nov) are fixed at 50,000 / 5,000. Nov: unchanged, closing 31,827.5 (OK).
        # Dec: 31,827.5 + open PO 5,000 + fixed 5,000 - 15,000 = 26,827.5; buffer 15,000 / 30.4375 x 15 = 7,392.2; above 15,000 + 7,392.2
        # -> Over stock 26,827.5 - 7,392.2 = 19,435.3.
        b3 = 15000 / 30.4375 * 15
        c3 = c1 + 5000 + 5000 - 15000
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Model Code", "Model Name", "01-11-2026", "01-12-2026"])
        sheet.append(["DP12000", "Eva 8014", 60000, 15000])
        content = io.BytesIO()
        workbook.save(content)
        self.assertEqual(api.post("/api/planning/plans/2026-11/upload?workspace=trial&release_date=2026-11-20", content=content.getvalue(), headers=xlsx).status_code, 201)
        revisions = {r["revision"]: r["id"] for r in api.get("/api/planning/plans", params={"issue_month": "2026-11"}).json()}
        impact = lambda tier: api.get("/api/planning/rev-impact", params={"old_cycle": revisions["Rev.00"], "new_cycle": revisions["Rev.01"], "tier": tier, "months": 2}).json()
        t1i = [r for r in impact("T1")["rows"] if r["code"] == "DP-BEAM"][0]
        self.assertEqual(t1i["suppliers"], "DP Savera")
        got = [(c["month"], c["old_requirement"], c["new_requirement"], c["change"], c["change_percent"], c["big_change"], c["locked"], c["fixed_po"], round(c["closing"], 4), round(c["buffer"], 4), c["status"], round(c["excess"], 4), c["shortage"]) for c in t1i["months"]]
        self.assertEqual(got, [("2026-11", 60000, 60000, 0, 0.0, False, True, 50000, r4(c1), r4(b1), "OK", 0, 0),
                               ("2026-12", 30000, 15000, -15000, -50.0, True, True, 5000, r4(c3), r4(b3), "Over stock", r4(c3 - b3), 0)])
        # T2: Savera's POs are fixed inside its lead time, so its balance demand and JGI's requirement do not change
        jgi_i = [r for r in impact("T2")["rows"] if r["suppliers"] == "DP JGI"][0]
        self.assertEqual([(c["old_requirement"], c["new_requirement"], c["status"], c["locked"], c["fixed_po"]) for c in jgi_i["months"]], [(24600, 24600, "OK", True, 30000), (3000, 3000, "OK", True, 0)])
        self.assertEqual(api.get("/api/planning/rev-impact", params={"old_cycle": revisions["Rev.01"], "new_cycle": revisions["Rev.01"]}).status_code, 422)
        self.assertEqual(api.get("/api/planning/rev-impact", params={"old_cycle": 99999, "new_cycle": revisions["Rev.01"]}).status_code, 404)
        book = load_workbook(io.BytesIO(api.get("/api/planning/rev-impact.xlsx", params={"old_cycle": revisions["Rev.00"], "new_cycle": revisions["Rev.01"], "tier": "T1", "months": 2, "issues_only": True}).content)).active
        beam_rows = [[c.value for c in r] for r in book.iter_rows(min_row=4) if r[2].value == "DP-BEAM"]
        self.assertEqual([(r[5], r[9], r[15], r[17]) for r in beam_rows], [("2026-12", -50.0, "Over stock", r4(c3 - b3))])
        # Rev.02 raises Dec to 90,000 (+200 %): Dec 31,827.5 + open PO 5,000 - 90,000 = -53,172.5; buffer 90,000 / 30.4375 x 15;
        # Savera's Dec PO is fixed at 5,000, the 40 % not linked arrives as re-planned -> closing below 0 = Shortage.
        # 12 months asked, but both plans stop at Dec -> only Nov and Dec.
        sheet = workbook.active
        sheet.delete_rows(2)
        sheet.append(["DP12000", "Eva 8014", 60000, 90000])
        content = io.BytesIO()
        workbook.save(content)
        self.assertEqual(api.post("/api/planning/plans/2026-11/upload?workspace=trial&release_date=2026-11-25", content=content.getvalue(), headers=xlsx).status_code, 201)
        revisions = {r["revision"]: r["id"] for r in api.get("/api/planning/plans", params={"issue_month": "2026-11"}).json()}
        result = api.get("/api/planning/rev-impact", params={"old_cycle": revisions["Rev.00"], "new_cycle": revisions["Rev.02"], "tier": "T1", "months": 12}).json()
        self.assertEqual(result["months"], ["2026-11", "2026-12"])
        b4 = 90000 / 30.4375 * 15
        p4 = c1 + 5000 - 90000
        c4 = p4 + (b4 - p4) * 0.4 + 5000
        dec = [r for r in result["rows"] if r["code"] == "DP-BEAM"][0]["months"][1]
        self.assertEqual((dec["change_percent"], dec["status"], r4(dec["closing"]), r4(dec["shortage"]), dec["excess"]), (200.0, "Shortage", r4(c4), r4(-c4), 0))
        self.assertEqual(api.get("/api/planning/demand-plan", params={"tier": "T9"}).status_code, 422)

    def test_supplier_overview_parts_bom_sub_suppliers_rates_and_requirement(self):
        # Savera makes Base Front Beam (60%) and Top Beam (100%). Plan Nov-26: Eva 60,000 (1 beam + 1 top beam each),
        # Serena 10,000 (2 beams each). Beam: 2 kg steel (JGI 30% @74, TATA 70% @72), per kg steel 1.7 kg ore (XYZ 100%).
        # Top beam: 1.5 kg steel (JGI 100% @76).
        api = self.client
        mat = lambda code, name, uom: api.post("/api/planning/materials", json={"code": code, "name": name, "uom": uom}).json()["id"]
        beam, top = mat("OV-BEAM", "Base Front Beam", "Nos"), mat("OV-TOPB", "Top Beam", "Nos")
        steel = api.post("/api/planning/tier-materials", json={"tier": "T2", "code": "OV-STEEL", "name": "STEEL0.7", "uom": "kg"}).json()["id"]
        ore = api.post("/api/planning/tier-materials", json={"tier": "T3", "code": "OV-ORE", "name": "Iron ore", "uom": "kg"}).json()["id"]
        product = api.post("/api/planning/products", json={"code": "OV-PRD", "name": "Fans"}).json()["id"]
        eva = api.post("/api/planning/models", json={"product_id": product, "code": "OV12000", "name": "Eva 8014"}).json()["id"]
        serena = api.post("/api/planning/models", json={"product_id": product, "code": "OV12001", "name": "Serena 8014"}).json()["id"]
        for model, material, qty in ((eva, beam, 1), (eva, top, 1), (serena, beam, 2)):
            self.assertEqual(api.post(f"/api/planning/model-bom/{model}", json={"material_id": material, "quantity": qty, "uom": "Nos"}).status_code, 201)
        sup = lambda code, name: api.post("/api/suppliers", json={"code": code, "name": name, "category": "Sheet Metal", "country": "India"}).json()["id"]
        savera, other, jgi, tata, xyz = sup("OV-1013", "OV Savera"), sup("OV-1007", "OV Supangita"), sup("OV-1004", "OV JGI"), sup("OV-1014", "OV TATA"), sup("OV-1030", "OV XYZ")
        link = lambda who, material, sob, price: api.post("/api/planning/supplier-parts", json={"material_id": material, "supplier_id": who, "sob_percent": sob, "unit_price": price, "currency": "INR", "lead_time_days": 15}).json()["id"]
        beam_link, top_link = link(savera, beam, 60, 85), link(savera, top, 100, 40)
        link(other, beam, 40, 90)
        src = lambda who, sob, price=0: {"sub_supplier_id": who, "sob_percent": sob, "unit_price": price, "currency": "INR"}
        steel_line = api.post("/api/planning/supplier-coverage/lines", json={"link_id": beam_link, "tier_material_id": steel, "quantity": 2, "sources": [src(jgi, 30, 74), src(tata, 70, 72)]}).json()["id"]
        api.post("/api/planning/supplier-coverage/lines", json={"link_id": beam_link, "parent_line_id": steel_line, "tier_material_id": ore, "quantity": 1.7, "sources": [src(xyz, 100, 9)]})
        api.post("/api/planning/supplier-coverage/lines", json={"link_id": top_link, "tier_material_id": steel, "quantity": 1.5, "sources": [src(jgi, 100, 76)]})
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Model Code", "Model Name", "01-11-2026"])
        sheet.append(["OV12000", "Eva 8014", 60000])
        sheet.append(["OV12001", "Serena 8014", 10000])
        content = io.BytesIO()
        workbook.save(content)
        api.post("/api/planning/plans/2026-11/upload?workspace=trial&release_date=2026-11-10", content=content.getvalue(), headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
        api.post("/api/planning/supplier-stock/2026-11?workspace=trial", json={"rows": [{"supplier_part_id": beam_link, "quantity": 9000, "available_quantity": 8000}]})
        view = api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-11"}).json()
        self.assertEqual((view["supplier"]["code"], view["plan"]["revision"], view["plan"]["release_date"]), ("OV-1013", "Rev.00", "2026-11-10"))
        # 1. parts: SOB, rate, FG stock, requirement (beam 60,000 x 1 + 10,000 x 2 = 80,000; Savera 60% = 48,000)
        parts = {p["code"]: p for p in view["parts"]}
        self.assertEqual([p["code"] for p in view["parts"]], ["OV-BEAM", "OV-TOPB"])
        self.assertEqual((parts["OV-BEAM"]["sob_percent"], parts["OV-BEAM"]["unit_price"], parts["OV-BEAM"]["stock"], parts["OV-BEAM"]["available"], parts["OV-BEAM"]["part_requirement"], parts["OV-BEAM"]["supplier_requirement"]),
                         (60, 85, 9000, 8000, 80000, 48000))
        self.assertEqual((parts["OV-TOPB"]["part_requirement"], parts["OV-TOPB"]["supplier_requirement"], parts["OV-TOPB"]["stock"]), (60000, 60000, None))
        # net requirement after the FG stock: beam 48,000 - 9,000 = 39,000 still to make; top beam 60,000
        self.assertEqual((parts["OV-BEAM"]["net_requirement"], parts["OV-TOPB"]["net_requirement"]), (39000, 60000))
        # the coverage page starts from the part's requirement in the month's plan (all suppliers: 80,000 beams)
        coverage_plan = api.get(f"/api/planning/supplier-coverage/{beam_link}", params={"month": "2026-11"}).json()["plan"]
        self.assertEqual((coverage_plan["revision"], coverage_plan["month"], coverage_plan["part_requirement"]), ("Rev.00", "2026-11", 80000))
        # same as the Shortage & excess report's plan demand for Savera's beam (worked out there separately)
        report = api.get("/api/planning/report", params={"start_month": "2026-11", "end_month": "2026-11"}).json()["items"]
        self.assertEqual(next(r["plan_demand"] for r in report if r["material_code"] == "OV-BEAM" and r["supplier_code"] == "OV-1013"), 48000)
        # 2-3. BOM from the net requirement: steel 39,000 x 2 = 78,000 kg (JGI 30% 23,400, TATA 70% 54,600); ore x 1.7 = 132,600
        lines = parts["OV-BEAM"]["lines"]
        self.assertEqual([(l["depth"], l["code"], l["bom_qty"], l["per_part"], l["requirement"]) for l in lines], [(1, "OV-STEEL", 2, 2, 78000), (2, "OV-ORE", 1.7, 3.4, 132600)])
        self.assertEqual([(s["code"], s["tier"], s["sob_percent"], s["unit_price"], s["requirement"]) for s in lines[0]["sources"]], [("OV-1004", "T2", 30, 74, 23400), ("OV-1014", "T2", 70, 72, 54600)])
        self.assertEqual((lines[1]["sources"][0]["code"], lines[1]["sources"][0]["tier"], lines[1]["sources"][0]["requirement"]), ("OV-1030", "T3", 132600))
        # 4-6. sub-supplier list: who, for which part and material, SOB, rate, requirement (top beam steel 60,000 x 1.5 = 90,000 kg)
        listed = [(r["code"], r["part_code"], r["material_code"], r["sob_percent"], r["unit_price"], r["requirement"]) for r in view["sub_suppliers"]]
        self.assertEqual(listed, [("OV-1004", "OV-BEAM", "OV-STEEL", 30, 74, 23400), ("OV-1004", "OV-TOPB", "OV-STEEL", 100, 76, 90000),
                                  ("OV-1014", "OV-BEAM", "OV-STEEL", 70, 72, 54600), ("OV-1030", "OV-BEAM", "OV-ORE", 100, 9, 132600)])
        # one line per sub-supplier and material, blended over Savera's parts: steel total 23,400 + 90,000 + 54,600 = 168,000 kg;
        # JGI 113,400 kg -> 67.5%, rate (74 x 23,400 + 76 x 90,000) / 113,400 = 75.5873; TATA 54,600 -> 32.5% @72; XYZ ore 100% @9
        blend = [(r["code"], r["material_code"], round(r["sob_percent"], 3), round(r["unit_price"], 4), r["requirement"], r["tier"], r["net_requirement"]) for r in view["sub_supplier_summary"]]
        self.assertEqual(blend, [("OV-1004", "OV-STEEL", 67.5, 75.5873, 113400, "T2", 113400), ("OV-1014", "OV-STEEL", 32.5, 72, 54600, "T2", 54600), ("OV-1030", "OV-ORE", 100, 9, 132600, "T3", 132600)])
        # edit from that line: TATA steel SOB 50 (beam only); JGI 60 would make the beam 110% -> refused, nothing changed
        edit = lambda body: api.post(f"/api/planning/supplier-overview/{savera}/sub-supplier", json={"material_code": "ov-steel", **body})
        self.assertEqual(edit({"sub_supplier_id": tata, "sob_percent": 50}).json(), {"parts": 1})
        refused = edit({"sub_supplier_id": jgi, "sob_percent": 60})
        self.assertEqual(refused.status_code, 422)
        self.assertIn("OV-BEAM: the other suppliers of ov-steel have 50%, so at most 50% is left", refused.json()["detail"])
        self.assertEqual(edit({"sub_supplier_id": jgi, "unit_price": 80}).json(), {"parts": 2})  # the rate goes to both parts
        after = api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-11"}).json()
        detail = {(r["code"], r["part_code"]): (r["sob_percent"], r["unit_price"]) for r in after["sub_suppliers"]}
        self.assertEqual((detail[("OV-1004", "OV-BEAM")], detail[("OV-1004", "OV-TOPB")], detail[("OV-1014", "OV-BEAM")]), ((30, 80), (100, 80), (50, 72)))
        self.assertEqual(edit({"sub_supplier_id": xyz}).status_code, 422)  # nothing to change
        self.assertEqual(api.post(f"/api/planning/supplier-overview/{savera}/sub-supplier", json={"sub_supplier_id": xyz, "material_code": "OV-STEEL", "sob_percent": 10}).status_code, 404)
        # stock typed on the overview: T2 (JGI under the beam's steel) and FG stock of the beam
        saved = api.post("/api/planning/supplier-source-stock/2026-11?workspace=trial", json={"rows": [{"t1_supplier_code": "OV-1013", "part_code": "OV-BEAM", "material_code": "OV-STEEL", "supplier_code": "OV-1004", "quantity": 5000}]})
        self.assertEqual((saved.status_code, saved.json()["rows"]), (201, 1), saved.text)
        self.assertEqual(api.post("/api/planning/supplier-stock/2026-11?workspace=trial", json={"rows": [{"supplier_part_id": beam_link, "quantity": 7500}]}).status_code, 201)
        stocked = api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-11"}).json()
        beam_view = next(p for p in stocked["parts"] if p["code"] == "OV-BEAM")
        jgi_beam = next(src for src in beam_view["lines"][0]["sources"] if src["code"] == "OV-1004")
        self.assertEqual((beam_view["stock"], jgi_beam["stock"], jgi_beam["available"]), (7500, 5000, 5000))
        self.assertIn("Row 1: OV TATA is not a supplier of OV-ORE", api.post("/api/planning/supplier-source-stock/2026-11", json={"rows": [{"t1_supplier_code": "OV-1013", "part_code": "OV-BEAM", "material_code": "OV-ORE", "supplier_code": "OV-1014", "quantity": 1}]}).json()["detail"])
        self.assertIn("stock must be a number", api.post("/api/planning/supplier-source-stock/2026-11", json={"rows": [{"t1_supplier_code": "OV-1013", "quantity": "x"}]}).json()["detail"])
        # stock typed on the blended line (JGI steel 2,000): one number for the material; the beam's 5,000 entry is replaced
        # beam net 48,000 - 7,500 = 40,500 -> JGI 30% x 2 kg = 24,300 + top beam 90,000 = 114,300; net 114,300 - 2,000 = 112,300
        api.post("/api/planning/supplier-source-stock/2026-11?workspace=trial", json={"rows": [{"t1_supplier_code": "OV-1013", "part_code": "OV-TOPB", "material_code": "OV-STEEL", "supplier_code": "OV-1004", "quantity": 1000}]})
        self.assertEqual(next(r for r in api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-11"}).json()["sub_supplier_summary"] if r["code"] == "OV-1004")["stock"], 6000)  # 5,000 beam + 1,000 top beam
        self.assertEqual(edit({"sub_supplier_id": jgi, "stock": 2000, "month": "2026-11"}).status_code, 200)
        jgi_line = next(r for r in api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-11"}).json()["sub_supplier_summary"] if r["code"] == "OV-1004")
        self.assertEqual((jgi_line["requirement"], jgi_line["stock"], jgi_line["net_requirement"]), (114300, 2000, 112300))
        self.assertEqual(edit({"sub_supplier_id": jgi, "stock": -1, "month": "2026-11"}).status_code, 422)
        # templates and uploads of the three sections
        xlsx_type = {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        def load(section):
            return load_workbook(io.BytesIO(api.get(f"/api/planning/supplier-overview/{savera}/template/{section}.xlsx", params={"month": "2026-11"}).content))
        def send(section, book):
            content = io.BytesIO()
            book.save(content)
            return api.post(f"/api/planning/supplier-overview/{savera}/upload/{section}?month=2026-11&workspace=trial", content=content.getvalue(), headers=xlsx_type)
        # 1. parts: filled with what is saved; change the beam's SOB and FG stock, add a new part
        mat("OV-BRKT", "Bracket", "Nos")
        parts_book = load("parts")
        sheet = parts_book.active
        self.assertEqual([c.value for c in sheet[3]], ["Supplier Code", "Supplier Name", "T1 Part Code", "T1 Part Name", "SOB", "Unit price", "Currency", "Lead time", "MOQ", "Buffer days", "FG stock"])
        self.assertEqual([[c.value for c in row][:5] + [row[10].value] for row in sheet.iter_rows(min_row=4)], [["OV-1013", "OV Savera", "OV-BEAM", "Base Front Beam", 60, 7500], ["OV-1013", "OV Savera", "OV-TOPB", "Top Beam", 100, None]])
        sheet.cell(4, 5).value, sheet.cell(4, 11).value = 55, 6000
        sheet.append(["OV-1013", "OV Savera", "OV-BRKT", "Bracket", 100, 12, "INR", 7, 50, 3, 400])
        up = send("parts", parts_book)
        self.assertEqual((up.status_code, up.json()["rows"]), (201, 3), up.text)
        after_parts = {p["code"]: p for p in api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-11"}).json()["parts"]}
        self.assertEqual((after_parts["OV-BEAM"]["sob_percent"], after_parts["OV-BEAM"]["stock"], after_parts["OV-BRKT"]["sob_percent"], after_parts["OV-BRKT"]["unit_price"], after_parts["OV-BRKT"]["stock"]), (55, 6000, 100, 12, 400))
        sheet.cell(4, 5).value = 70  # 70 + Supangita's 40 > 100: refused, nothing saved
        sheet.cell(5, 11).value = 1234
        refused = send("parts", parts_book)
        self.assertEqual(refused.status_code, 422)
        self.assertIn("Row 4: Supplier SOB cannot exceed 100%", refused.json()["detail"])
        self.assertIsNone(next(p for p in api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-11"}).json()["parts"] if p["code"] == "OV-TOPB")["stock"])
        # 2. BOM & sub-supplier master: the supplier-wise BOM workbook
        self.assertIn("Supplier/material coverage · OV-1013", load("bom").active["A1"].value)
        # 3. RM Master: only changed values are applied (JGI rate 81 on both parts, TATA stock 4,000)
        rm_book = load("rm")
        sheet = rm_book.active
        self.assertEqual([c.value for c in sheet[3]], ["Sub-supplier Code", "Sub-supplier Name", "Tier", "Material code", "Material name", "Blended SOB", "Rate", "Currency", "Stock"])
        rm_rows = {(row[0].value, row[3].value): row[0].row for row in sheet.iter_rows(min_row=4)}
        sheet.cell(rm_rows[("OV-1004", "OV-STEEL")], 7).value = 81
        sheet.cell(rm_rows[("OV-1014", "OV-STEEL")], 9).value = 4000
        up_rm = send("rm", rm_book)
        self.assertEqual((up_rm.status_code, up_rm.json()["rows"]), (201, 2), up_rm.text)
        final = api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-11"}).json()
        prices = {(r["code"], r["part_code"]): (r["sob_percent"], r["unit_price"]) for r in final["sub_suppliers"]}
        self.assertEqual((prices[("OV-1004", "OV-BEAM")], prices[("OV-1004", "OV-TOPB")], prices[("OV-1014", "OV-BEAM")]), ((30, 81), (100, 81), (50, 72)))  # SOB unchanged
        self.assertEqual(next(r for r in final["sub_supplier_summary"] if r["code"] == "OV-1014")["stock"], 4000)
        sheet.cell(rm_rows[("OV-1004", "OV-STEEL")], 4).value = "OV-NOPE"
        self.assertIn("does not supply OV-NOPE", send("rm", rm_book).json()["detail"])
        self.assertEqual(api.get(f"/api/planning/supplier-overview/{savera}/template/other.xlsx").status_code, 404)
        # no plan for the month: requirement 0; unknown supplier 404
        later = api.get(f"/api/planning/supplier-overview/{savera}", params={"month": "2026-10"}).json()
        self.assertEqual((later["plan"], later["parts"][0]["part_requirement"], later["sub_suppliers"][0]["requirement"]), (None, 0, 0))
        self.assertEqual(api.get("/api/planning/supplier-overview/999999").status_code, 404)

    def test_wide_month_plan_template_and_month_formats_upload(self):
        product = self.client.post("/api/planning/products", json={"code": "MONTH-PRD", "name": "Month product"})
        model = self.client.post("/api/planning/models", json={
            "product_id": product.json()["id"], "code": "MONTH-12000", "name": "Eva 8014",
        })
        self.assertEqual(model.status_code, 201, model.text)

        template = self.client.get("/api/planning/templates/plan.xlsx?issue_month=2026-04")
        self.assertEqual(template.status_code, 200, template.text)
        template_sheet = load_workbook(io.BytesIO(template.content), data_only=False).active
        # logo + title, then Model Code | Model Name | 24 months (DD-MM-YYYY)
        self.assertEqual(template_sheet["B1"].value, "Design to Cost · Production plan")
        self.assertEqual(len(template_sheet._images), 1)  # the DC logo
        self.assertEqual([template_sheet.cell(4, c).value for c in range(1, 3)], ["Model Code", "Model Name"])
        self.assertEqual((template_sheet["C4"].value.date(), template_sheet["C4"].number_format), (date(2026, 4, 1), "DD-MM-YYYY"))
        self.assertEqual(template_sheet["Z4"].value.date(), date(2028, 3, 1))  # 24th month
        self.assertIsNone(template_sheet["AA4"].value)
        self.assertEqual((template_sheet["A5"].value, template_sheet["B5"].value), ("MONTH-12000", "Eva 8014"))

        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["model_code", "model_MONTH", "Apr-26", date(2026, 5, 1), "Jun 2026", "2026/07", "08/2026", "2026-09-01"])
        sheet.append(["MONTH-12000", "Eva 8014", 100, 110, 120, 130, 140, 150])
        content = io.BytesIO()
        workbook.save(content)
        response = self.client.post(
            "/api/planning/plans/2026-04/upload?workspace=trial",
            content=content.getvalue(),
            headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        )
        self.assertEqual(response.status_code, 201, response.text)
        self.assertEqual(response.json()["rows"], 6)
        self.assertEqual(response.json()["revision"], "Rev.00")

        long_workbook = Workbook()
        long_sheet = long_workbook.active
        long_sheet.append(["model_code", "target_month", "quantity"])
        long_sheet.append(["MONTH-12000", "Oct-2026", 160])
        long_content = io.BytesIO()
        long_workbook.save(long_content)
        second = self.client.post(
            "/api/planning/plans/Apr-2026/upload?workspace=trial",
            content=long_content.getvalue(),
            headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        )
        self.assertEqual(second.status_code, 201, second.text)
        self.assertEqual(second.json()["revision"], "Rev.01")

        revisions = self.client.get("/api/planning/plans?workspace=trial&issue_month=Apr-2026").json()
        latest = revisions[0]
        self.assertIn("created_at", latest)
        detail = self.client.get(f"/api/planning/plans/{latest['id']}/rows?search=MONTH-12000")
        self.assertEqual(detail.status_code, 200, detail.text)
        self.assertTrue(detail.json()["latest"])
        self.assertEqual(detail.json()["items"][0]["quantities"]["2026-10"], 160)
        edited = self.client.post(f"/api/planning/plans/{latest['id']}/revisions", json={
            "rows": [{"model_code": "MONTH-12000", "target_month": "May-2026", "quantity": 999}],
        })
        self.assertEqual(edited.status_code, 201, edited.text)
        self.assertEqual(edited.json()["revision"], "Rev.02")
        self.assertTrue(edited.json()["created_at"])
        revised = self.client.get(f"/api/planning/plans/{edited.json()['id']}/rows?search=MONTH-12000")
        self.assertEqual(revised.json()["items"][0]["quantities"]["2026-05"], 999)
        self.assertEqual(revised.json()["items"][0]["quantities"]["2026-10"], 160)
        stale = self.client.post(f"/api/planning/plans/{latest['id']}/revisions", json={
            "rows": [{"model_code": "MONTH-12000", "target_month": "May-2026", "quantity": 1000}],
        })
        self.assertEqual(stale.status_code, 409)

    def test_bom_levels_are_model_root_then_parent_part_linked(self):
        product = self.client.post("/api/planning/products", json={"code": "BOM-PRD", "name": "BOM product"})
        model = self.client.post("/api/planning/models", json={
            "product_id": product.json()["id"], "code": "BOM-12000", "name": "BOM model",
        })
        self.assertEqual(model.status_code, 201, model.text)

        root_template = self.client.get("/api/planning/templates/bom-oem.xlsx")
        self.assertEqual(root_template.status_code, 200, root_template.text)
        root_workbook = load_workbook(io.BytesIO(root_template.content), data_only=False)
        self.assertEqual([cell.value for cell in root_workbook.active[1]], ["model_code", "material_code", "quantity", "uom"])
        self.assertTrue(root_workbook.active.data_validations.dataValidation)
        self.assertEqual(self.client.get("/api/planning/templates/bom-level-1.xlsx").status_code, 404)

        for kind in ("bom-level-2", "bom-level-3"):
            template = self.client.get(f"/api/planning/templates/{kind}.xlsx")
            self.assertEqual(template.status_code, 200, template.text)
            workbook = load_workbook(io.BytesIO(template.content), data_only=False)
            self.assertEqual([cell.value for cell in workbook.active[1]], ["parent_code", "material_code", "quantity", "uom"])

        def upload_model(tab, headers, row):
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(headers)
            sheet.append(row)
            content = io.BytesIO()
            workbook.save(content)
            return self.client.post(
                f"/api/planning/model-bom/{model.json()['id']}/upload",
                params={"tab": tab},
                content=content.getvalue(),
                headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
            )

        root = upload_model("Model BOM Level 1", ["model_code", "material_code", "quantity", "uom"], ["BOM-12000", "ASSY-001", 1, "ea"])
        self.assertEqual(root.status_code, 201, root.text)
        listed_model = next(row for row in self.client.get("/api/planning/models").json() if row["id"] == model.json()["id"])
        self.assertTrue(listed_model["has_bom"])

        materials = {row["code"]: row for row in self.client.get("/api/planning/materials?limit=500").json()["items"]}
        level2_parents = self.client.get("/api/planning/bom-part-parents?tab=Level%202%20Part%20BOM")
        self.assertEqual(level2_parents.status_code, 200, level2_parents.text)
        self.assertEqual({row["code"] for row in level2_parents.json()["items"]}, {"ASSY-001"})
        filtered_level2_parents = self.client.get(
            f"/api/planning/bom-part-parents?tab=Level%202%20Part%20BOM&model_id={model.json()['id']}"
        )
        self.assertEqual([row["code"] for row in filtered_level2_parents.json()["items"]], ["ASSY-001"])
        level2_workbook = Workbook()
        level2_sheet = level2_workbook.active
        level2_sheet.append(["parent_code", "material_code", "quantity", "uom"])
        level2_sheet.append(["ASSY-001", "SUB-001", 2, "ea"])
        level2_content = io.BytesIO()
        level2_workbook.save(level2_content)
        level2 = self.client.post(
            f"/api/planning/part-bom/{materials['ASSY-001']['id']}/upload?tab=Level%202%20Part%20BOM",
            content=level2_content.getvalue(),
            headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        )
        self.assertEqual(level2.status_code, 201, level2.text)

        level3_parents = self.client.get("/api/planning/bom-part-parents?tab=Level%203%20Part%20BOM")
        self.assertEqual(level3_parents.status_code, 200, level3_parents.text)
        self.assertEqual({row["code"] for row in level3_parents.json()["items"]}, {"SUB-001"})
        filtered_level3_parents = self.client.get(
            f"/api/planning/bom-part-parents?tab=Level%203%20Part%20BOM&model_id={model.json()['id']}"
        )
        self.assertEqual([row["code"] for row in filtered_level3_parents.json()["items"]], ["SUB-001"])
        level3_workbook = Workbook()
        level3_sheet = level3_workbook.active
        level3_sheet.append(["parent_code", "material_code", "quantity", "uom"])
        level3_sheet.append(["SUB-001", "MAT-101", -0.25, "ea"])
        level3_content = io.BytesIO()
        level3_workbook.save(level3_content)
        level3 = self.client.post(
            f"/api/planning/part-bom/{materials['SUB-001']['id']}/upload?tab=Level%203%20Part%20BOM",
            content=level3_content.getvalue(),
            headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        )
        self.assertEqual(level3.status_code, 201, level3.text)

        listed_level2 = self.client.get(
            f"/api/planning/part-bom/{materials['ASSY-001']['id']}?tab=Level%202%20Part%20BOM"
        )
        self.assertEqual(listed_level2.status_code, 200, listed_level2.text)
        self.assertEqual([row["code"] for row in listed_level2.json()["items"]], ["SUB-001"])
        self.assertTrue(listed_level2.json()["items"][0]["has_bom"])
        all_level2 = self.client.get("/api/planning/part-boms?tab=Level%202%20Part%20BOM")
        self.assertEqual(all_level2.status_code, 200, all_level2.text)
        self.assertEqual(
            [(row["parent"]["code"], row["code"]) for row in all_level2.json()["items"]],
            [("ASSY-001", "SUB-001")],
        )
        selected_level2 = self.client.get(
            f"/api/planning/part-boms?tab=Level%202%20Part%20BOM&parent_material_id={materials['ASSY-001']['id']}"
        )
        self.assertEqual([row["code"] for row in selected_level2.json()["items"]], ["SUB-001"])
        code_search = self.client.get(
            f"/api/planning/part-bom/{materials['ASSY-001']['id']}?tab=Level%202%20Part%20BOM&search_code=SUB-00"
        )
        self.assertEqual([row["code"] for row in code_search.json()["items"]], ["SUB-001"])
        name_search = self.client.get(
            f"/api/planning/part-bom/{materials['ASSY-001']['id']}?tab=Level%202%20Part%20BOM&search_name={materials['SUB-001']['name'][:4]}"
        )
        self.assertEqual([row["code"] for row in name_search.json()["items"]], ["SUB-001"])
        listed_level3 = self.client.get(
            f"/api/planning/part-bom/{materials['SUB-001']['id']}?tab=Level%203%20Part%20BOM"
        )
        self.assertEqual(listed_level3.status_code, 200, listed_level3.text)
        self.assertEqual([row["code"] for row in listed_level3.json()["items"]], ["MAT-101"])
        self.assertEqual(listed_level3.json()["items"][0]["quantity"], -0.25)
        self.assertTrue(listed_level3.json()["items"][0]["has_bom"])
        all_level3 = self.client.get("/api/planning/part-boms?tab=Level%203%20Part%20BOM")
        self.assertEqual(all_level3.status_code, 200, all_level3.text)
        self.assertEqual(
            [(row["parent"]["code"], row["code"]) for row in all_level3.json()["items"]],
            [("SUB-001", "MAT-101")],
        )
        level3_name_search = self.client.get(
            f"/api/planning/part-bom/{materials['SUB-001']['id']}?tab=Level%203%20Part%20BOM&search_name={materials['MAT-101']['name'][:4]}"
        )
        self.assertEqual([row["code"] for row in level3_name_search.json()["items"]], ["MAT-101"])

        listed_root = self.client.get(f"/api/planning/model-bom/{model.json()['id']}?tab=Model%20BOM%20Level%201")
        self.assertEqual(listed_root.status_code, 200, listed_root.text)
        self.assertEqual(listed_root.json()["model_code"], "BOM-12000")
        self.assertEqual(listed_root.json()["model_name"], "BOM model")
        bom = self.client.get(f"/api/planning/model-bom/{model.json()['id']}?tab=Level%202%20Part%20BOM&offset=0&limit=1")
        self.assertEqual(bom.status_code, 422, bom.text)
        rejected_model_upload = upload_model(
            "Level 2 Part BOM",
            ["parent_code", "material_code", "quantity", "uom"],
            ["ASSY-001", "SUB-001", 2, "ea"],
        )
        self.assertEqual(rejected_model_upload.status_code, 422, rejected_model_upload.text)

    def test_new_part_bom_parents_categories_and_validation(self):
        # Expected values below are written out by hand from the data this test creates.
        def material(code, category="Raw material", uom="ea"):
            response = self.client.post("/api/planning/materials", json={
                "code": code, "name": f"{code} name", "category": category, "uom": uom,
            })
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]

        l1a, l1b = material("NB-L1A"), material("NB-L1B")
        c1, c2, c3 = material("NB-C1", "Electrical"), material("NB-C2", "electrical"), material("NB-C3", "Packing")
        model = self.client.post("/api/planning/models", json={"code": "NB-MODEL", "name": "New BOM model"})
        self.assertEqual(model.status_code, 201, model.text)
        for part in (l1a, l1b):
            root = self.client.post(f"/api/planning/model-bom/{model.json()['id']}", json={"material_id": part, "quantity": 1, "uom": "ea"})
            self.assertEqual(root.status_code, 201, root.text)

        # Categories: distinct, case-insensitive ("Electrical" / "electrical" count once), sorted
        categories = self.client.get("/api/planning/material-categories").json()["items"]
        self.assertEqual(len([name for name in categories if name.lower() == "electrical"]), 1)
        self.assertIn("Packing", categories)
        self.assertEqual(categories, sorted(categories, key=str.casefold))

        def parents(tab):
            rows = self.client.get("/api/planning/bom-part-parents", params={"tab": tab}).json()["items"]
            return {row["id"]: row["has_bom"] for row in rows}

        level2 = "Level 2 Part BOM"
        self.assertFalse(parents(level2)[l1a])
        self.assertFalse(parents(level2)[l1b])

        def new_bom(parent, lines, tab=level2):
            return self.client.post(f"/api/planning/part-bom/{parent}/new", json={"tab": tab, "lines": lines})

        # Rejected inputs save nothing
        self.assertEqual(new_bom(l1a, []).status_code, 422)
        self.assertEqual(new_bom(l1a, [{"child_item_id": c1, "quantity": 0}]).status_code, 422)
        self.assertEqual(new_bom(l1a, [{"child_item_id": 999999, "quantity": 1}]).status_code, 422)
        # Duplicate line must give a clear message (not a server error from the database)
        plain = TestClient(self.client.app, raise_server_exceptions=False)
        plain.cookies = self.client.cookies
        duplicate = plain.post(f"/api/planning/part-bom/{l1a}/new", json={"tab": level2, "lines": [{"child_item_id": c1, "quantity": 1}, {"child_item_id": c1, "quantity": 2}]})
        self.assertEqual(duplicate.status_code, 422, duplicate.text)
        self.assertIn("entered twice", duplicate.json()["detail"])
        self.assertEqual(new_bom(l1a, [{"child_item_id": l1a, "quantity": 1}]).status_code, 422)
        self.assertEqual(new_bom(c3, [{"child_item_id": c1, "quantity": 1}]).status_code, 422)  # not a Level 1 part
        self.assertEqual(new_bom(l1a, [{"child_item_id": c1, "quantity": 1}, {"child_item_id": 999999, "quantity": 1}]).status_code, 422)
        self.assertFalse(parents(level2)[l1a], "a failed save must not leave partial lines")

        # Multi-line save, incl. negative scrap quantity
        saved = new_bom(l1a, [{"child_item_id": c1, "quantity": 2}, {"child_item_id": c2, "quantity": -0.3}])
        self.assertEqual(saved.status_code, 201, saved.text)
        self.assertEqual(saved.json()["rows"], 2)
        self.assertTrue(parents(level2)[l1a])
        self.assertFalse(parents(level2)[l1b])
        lines = self.client.get(f"/api/planning/part-bom/{l1a}", params={"tab": level2}).json()["items"]
        self.assertEqual(sorted((row["material_id"], row["quantity"]) for row in lines), sorted([(c1, 2.0), (c2, -0.3)]))

        # A parent that already has a BOM cannot get a second "new" BOM
        self.assertEqual(new_bom(l1a, [{"child_item_id": c3, "quantity": 1}]).status_code, 409)

        # Level 3: parents are the Level 2 parts (c1, c2); circular link back to l1a is rejected
        level3 = "Level 3 Part BOM"
        self.assertEqual(set(parents(level3)) & {c1, c2, c3}, {c1, c2})
        self.assertEqual(new_bom(c1, [{"child_item_id": l1a, "quantity": 1}], level3).status_code, 422)
        self.assertEqual(new_bom(c1, [{"child_item_id": c3, "quantity": 4}], level3).status_code, 201)
        self.assertTrue(parents(level3)[c1])
        self.assertFalse(parents(level3)[c2])

        # View-only roles cannot create a BOM
        user = self.client.post("/api/users", json={"username": "viewer1", "full_name": "Viewer", "password": "Viewer-pass-2026!", "role": "User"})
        self.assertEqual(user.status_code, 201, user.text)
        self.client.post("/api/auth/logout")
        login = self.client.post("/api/auth/login", json={"username": "viewer1", "password": "Viewer-pass-2026!"})
        self.assertEqual(login.status_code, 200, login.text)
        self.assertEqual(new_bom(l1b, [{"child_item_id": c3, "quantity": 1}]).status_code, 403)

    def test_bom_tree_export_has_full_paths_outline_levels_and_bold_parents(self):
        # Data: product BT-1000 -> model BT-12000 -> L1 BT-MOTOR -> L2 BT-ROTOR -> L3 BT-ADC (2), BT-SCRAP (-0.2)
        #                                                         -> L2 BT-STATOR (no L3)
        #       second model BT-13000 (no BOM) on the same product
        def material(code, name, uom="ea", cost=0):
            response = self.client.post("/api/planning/materials", json={"code": code, "name": name, "uom": uom, "unit_cost": cost, "currency": "INR"})
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]
        motor, rotor, stator = material("BT-MOTOR", "Motor", "PCS", 1200), material("BT-ROTOR", "Rotor", "PCS", 200), material("BT-STATOR", "Stator", "Nos")
        adc, scrap = material("BT-ADC", "ADC12", "kg", 350), material("BT-SCRAP", "Al scrap", "kg")
        m1 = self.client.post("/api/planning/models", json={"code": "BT-12000", "name": "Eva"}).json()["id"]
        m2 = self.client.post("/api/planning/models", json={"code": "BT-13000", "name": "Serena"}).json()["id"]
        self.assertEqual(self.client.post("/api/planning/products", json={"code": "BT-1000", "name": "FL", "model_ids": [m1, m2]}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/model-bom/{m1}", json={"material_id": motor, "quantity": 1, "uom": "PCS"}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{motor}/new", json={"tab": "Level 2 Part BOM", "lines": [{"child_item_id": rotor, "quantity": 1}, {"child_item_id": stator, "quantity": 1}]}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{rotor}/new", json={"tab": "Level 3 Part BOM", "lines": [{"child_item_id": adc, "quantity": 2}, {"child_item_id": scrap, "quantity": -0.2}]}).status_code, 201)

        def download(params):
            response = self.client.get("/api/planning/bom-tree/xlsx", params=params)
            self.assertEqual(response.status_code, 200, response.text)
            sheet = load_workbook(io.BytesIO(response.content)).active
            headers = [cell.value for cell in sheet[4]]
            rows = [r for r in range(5, sheet.max_row + 1) if sheet.cell(r, 1).value]
            return sheet, headers, rows

        sheet, headers, rows = download({"tab": "Product Master"})
        self.assertEqual(headers, ["Level", "Product code", "Product name", "Model code", "Model name", "L1 part code", "L1 part name",
                                   "L2 part code", "L2 part name", "L3 part code", "L3 part name", "Quantity", "Unit", "Part price", "Currency"])
        mine = [r for r in rows if sheet.cell(r, 2).value == "BT-1000"]
        values = [[sheet.cell(r, c).value for c in range(1, 14)] for r in mine]
        self.assertEqual(values, [
            ["Model", "BT-1000", "FL", "BT-12000", "Eva", None, None, None, None, None, None, None, None],
            ["L1", "BT-1000", "FL", "BT-12000", "Eva", "BT-MOTOR", "Motor", None, None, None, None, 1, "PCS"],
            ["L2", "BT-1000", "FL", "BT-12000", "Eva", "BT-MOTOR", "Motor", "BT-ROTOR", "Rotor", None, None, 1, "PCS"],
            ["L3", "BT-1000", "FL", "BT-12000", "Eva", "BT-MOTOR", "Motor", "BT-ROTOR", "Rotor", "BT-ADC", "ADC12", 2, "kg"],
            ["L3", "BT-1000", "FL", "BT-12000", "Eva", "BT-MOTOR", "Motor", "BT-ROTOR", "Rotor", "BT-SCRAP", "Al scrap", -0.2, "kg"],
            ["L2", "BT-1000", "FL", "BT-12000", "Eva", "BT-MOTOR", "Motor", "BT-STATOR", "Stator", None, None, 1, "Nos"],
            ["Model", "BT-1000", "FL", "BT-13000", "Serena", None, None, None, None, None, None, None, None],
        ])
        self.assertEqual([sheet.row_dimensions[r].outlineLevel for r in mine], [0, 1, 2, 3, 3, 2, 0])
        self.assertEqual([bool(sheet.cell(r, 2).font.b) for r in mine], [True, True, True, False, False, False, False])
        self.assertEqual([sheet.cell(r, 14).value for r in mine[1:5]], [1200, 200, 350, 0])
        self.assertIs(sheet.sheet_properties.outlinePr.summaryBelow, False, "+/- must sit on the parent row above its lines")
        # Opens collapsed: only top-level rows visible, parents marked collapsed ("+")
        self.assertEqual([bool(sheet.row_dimensions[r].hidden) for r in mine], [False, True, True, True, True, True, False])
        self.assertEqual([bool(sheet.row_dimensions[r].collapsed) for r in mine], [True, True, True, False, False, False, False])
        # Not locked: Excel allows the +/- buttons only on an unprotected sheet
        self.assertFalse(sheet.protection.sheet)
        self.assertEqual(sheet["D1"].value, "Design to Cost")
        self.assertEqual(sheet["D2"].value, "Supply Chain Management · Product Master BOM tree")
        self.assertEqual(len(sheet._images), 1, "logo image")

        # PDF layout: every cell fits on one line, table fits the page, parent rows bold, tidy numbers
        from reportlab.pdfbase import pdfmetrics
        from planning import _bom_tree_pdf_layout
        pdf_headers = ["Level", "L1 part code", "L1 part name", "Quantity", "Part price"]
        layout = _bom_tree_pdf_layout(pdf_headers, [
            {"has_children": True, "values": ["L1", "FL001MOTOR1", "MOTOR1", 1.0, 1200.0]},
            {"has_children": False, "values": ["L2", "FL001STEEL0.8", "Aluminimum scrap long name", -0.2, 0.0]},
        ])
        self.assertEqual(layout["data"][1], ["L1", "FL001MOTOR1", "MOTOR1", "1", "1,200.00"])
        self.assertEqual(layout["data"][2], ["L2", "FL001STEEL0.8", "Aluminimum scrap long name", "-0.2", "0.00"])
        self.assertEqual(layout["bold_rows"], [0, 1])
        for line in layout["data"]:
            for column, value in enumerate(line):
                needed = pdfmetrics.stringWidth(value, "Helvetica-Bold", layout["font_size"]) + 8
                self.assertGreaterEqual(layout["widths"][column] + 0.01, needed, f"{value!r} would wrap")
        self.assertLessEqual(sum(layout["widths"]), layout["pagesize"][0] - 56.69 + 0.01)
        # Realistic full Product Master row (15 columns): no cell wraps, table fits the page
        full = _bom_tree_pdf_layout(headers, [{"has_children": False, "values": [
            "L3", "1000", "FL", "FL12000", "Eva 8014", "FL001MOTOR1", "MOTOR1", "FL001ROTOR", "ROTOR1",
            "FL001STEEL0.8", "Aluminimum scrap", -0.2, "kg", 1200.0, "INR"]}])
        for line in full["data"]:
            for column, value in enumerate(line):
                self.assertGreaterEqual(full["widths"][column] + 0.01, pdfmetrics.stringWidth(value, "Helvetica-Bold", full["font_size"]) + 8, f"{value!r} would wrap")
        self.assertLessEqual(sum(full["widths"]), full["pagesize"][0] - 56.69 + 0.01)
        # Long names: stays on A4 with a smaller font (not wrapped, not moved to A3)
        from reportlab.lib.pagesizes import A4, landscape
        longer = _bom_tree_pdf_layout(headers, [{"has_children": False, "values": [
            "L3", "1000", "FL", "FL12000", "Eva 8014", "FL001MOTOR1", "Long motor assembly description text here", "FL001ROTOR", "ROTOR1",
            "FL001STEEL0.8", "Aluminium scrap from die casting process", -0.2, "kg", 1200.0, "INR"]}])
        self.assertEqual(longer["pagesize"], landscape(A4))
        self.assertTrue(5.0 <= longer["font_size"] < 7.0, longer["font_size"])
        # Extreme widths: page widens so every row is still one line inside the page
        wide = _bom_tree_pdf_layout(headers, [{"has_children": False, "values": ["X" * 60] * 15}])
        self.assertEqual(wide["font_size"], 5.0)
        for column, value in enumerate(wide["data"][1]):
            self.assertGreaterEqual(wide["widths"][column] + 0.01, pdfmetrics.stringWidth(value, "Helvetica-Bold", wide["font_size"]) + 8)
        self.assertLessEqual(sum(wide["widths"]), wide["pagesize"][0] - 56.69 + 0.01, "wide table must fit the page")

        # Model BOM Level 1, filtered to one model: model + L1 + L2 + L3 columns
        sheet, headers, rows = download({"tab": "Model BOM Level 1", "model_id": m1})
        self.assertEqual(headers[:3], ["Level", "Model code", "Model name"])
        self.assertEqual([sheet.cell(r, 4).value for r in rows], ["BT-MOTOR"] * 5)
        self.assertEqual([sheet.row_dimensions[r].outlineLevel for r in rows], [0, 1, 2, 2, 1])

        # Level 2, filtered to BT-MOTOR: parent row then its L2 / L3 lines
        sheet, headers, rows = download({"tab": "Level 2 Part BOM", "parent_item_id": motor})
        self.assertEqual(headers[:3], ["Level", "L1 part code", "L1 part name"])
        self.assertEqual([[sheet.cell(r, c).value for c in (1, 2, 4, 6)] for r in rows], [
            ["L1", "BT-MOTOR", None, None], ["L2", "BT-MOTOR", "BT-ROTOR", None], ["L3", "BT-MOTOR", "BT-ROTOR", "BT-ADC"],
            ["L3", "BT-MOTOR", "BT-ROTOR", "BT-SCRAP"], ["L2", "BT-MOTOR", "BT-STATOR", None]])

        # Level 3: only L2 parts that have a BOM (BT-ROTOR, not BT-STATOR)
        sheet, headers, rows = download({"tab": "Level 3 Part BOM"})
        self.assertEqual(headers[:5], ["Level", "L2 part code", "L2 part name", "L3 part code", "L3 part name"])
        self.assertEqual([[sheet.cell(r, c).value for c in (2, 4)] for r in rows if str(sheet.cell(r, 2).value).startswith("BT-")],
                         [["BT-ROTOR", None], ["BT-ROTOR", "BT-ADC"], ["BT-ROTOR", "BT-SCRAP"]])

        pdf = self.client.get("/api/planning/bom-tree/pdf", params={"tab": "Product Master"})
        self.assertEqual(pdf.status_code, 200)
        self.assertTrue(pdf.content.startswith(b"%PDF"))
        self.assertEqual(self.client.get("/api/planning/bom-tree/xlsx", params={"tab": "Bogus"}).status_code, 422)

        # Users without price permission get no price columns
        self.assertEqual(self.client.post("/api/users", json={"username": "noprice", "full_name": "No Price", "password": "No-price-pass-2026!", "role": "User"}).status_code, 201)
        self.client.post("/api/auth/logout")
        self.assertEqual(self.client.post("/api/auth/login", json={"username": "noprice", "password": "No-price-pass-2026!"}).status_code, 200)
        sheet, headers, rows = download({"tab": "Product Master"})
        self.assertNotIn("Part price", headers)
        self.assertEqual(headers[-2:], ["Quantity", "Unit"])

    def test_bom_price_is_sob_weighted_supplier_price(self):
        # Hand-worked: A 100 INR @ 60% + B 110 INR @ 40% = (6000 + 4400) / 100 = 104 INR
        part = self.client.post("/api/planning/materials", json={"code": "WP-STEEL", "name": "Steel 0.8", "uom": "kg", "unit_cost": 90, "currency": "INR"}).json()["id"]
        parent = self.client.post("/api/planning/materials", json={"code": "WP-MOTOR", "name": "Motor", "uom": "PCS"}).json()["id"]
        model = self.client.post("/api/planning/models", json={"code": "WP-12000", "name": "Eva"}).json()["id"]
        self.assertEqual(self.client.post(f"/api/planning/model-bom/{model}", json={"material_id": part, "quantity": 1, "uom": "kg"}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/model-bom/{model}", json={"material_id": parent, "quantity": 1, "uom": "PCS"}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{parent}/new", json={"tab": "Level 2 Part BOM", "lines": [{"child_item_id": part, "quantity": 2}]}).status_code, 201)
        suppliers = [row["id"] for row in self.client.get("/api/suppliers").json() if row["status"] == "Approved"]
        self.assertGreaterEqual(len(suppliers), 3)

        def l1_price():
            row = next(r for r in self.client.get(f"/api/planning/model-bom/{model}").json()["items"] if r["material_id"] == part)
            return round(row["unit_price"], 4), row["currency"], row["price_source"]

        def l2_price():
            row = self.client.get(f"/api/planning/part-bom/{parent}", params={"tab": "Level 2 Part BOM"}).json()["items"][0]
            return round(row["unit_price"], 4), row["currency"]

        self.assertEqual(l1_price(), (90, "INR", "Material Master"))
        def link(supplier, sob, price, currency="INR", tier="Direct", parent_supplier=None):
            response = self.client.post("/api/planning/supplier-parts", json={
                "material_id": part, "supplier_id": supplier, "sob_percent": sob, "unit_price": price, "currency": currency,
                "tier": tier, "parent_supplier_id": parent_supplier})
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]
        a = link(suppliers[0], 60, 100)
        self.assertEqual(l1_price(), (100, "INR", "SOB-weighted (1 supplier)"))
        b = link(suppliers[1], 40, 110)
        self.assertEqual(l1_price(), (104, "INR", "SOB-weighted (2 suppliers)"))
        self.assertEqual(l2_price(), (104, "INR"))
        # A Tier-2 link (supplies the T1 supplier, not us) does not change our price
        link(suppliers[2], 100, 50, tier="Tier 2", parent_supplier=suppliers[0])
        self.assertEqual(l1_price()[0], 104)
        # BOM tree export uses the same price
        sheet = load_workbook(io.BytesIO(self.client.get("/api/planning/bom-tree/xlsx", params={"tab": "Model BOM Level 1", "model_id": model}).content)).active
        headers = [cell.value for cell in sheet[4]]
        price_col = headers.index("Part price") + 1
        steel_rows = [r for r in range(5, sheet.max_row + 1) if sheet.cell(r, headers.index("L1 part code") + 1).value == "WP-STEEL"]
        self.assertEqual(round(sheet.cell(steel_rows[0], price_col).value, 4), 104)
        # Different currencies cannot be averaged -> Material Master price, reason shown
        self.assertEqual(self.client.patch(f"/api/planning/supplier-parts/{b}", json={"currency": "USD"}).status_code, 200)
        self.assertEqual(l1_price(), (90, "INR", "Material Master (suppliers in different currencies)"))
        # A supplier with 0 SOB or no price is ignored
        self.assertEqual(self.client.patch(f"/api/planning/supplier-parts/{b}", json={"currency": "INR", "sob_percent": 0}).status_code, 200)
        self.assertEqual(l1_price()[0], 100)
        # All suppliers at 0% SOB -> Material Master price (no divide-by-zero error)
        self.assertEqual(self.client.patch(f"/api/planning/supplier-parts/{a}", json={"sob_percent": 0}).status_code, 200)
        plain = TestClient(self.client.app, raise_server_exceptions=False)
        plain.cookies = self.client.cookies
        response = plain.get(f"/api/planning/model-bom/{model}")
        self.assertEqual(response.status_code, 200, response.text)
        row = next(r for r in response.json()["items"] if r["material_id"] == part)
        self.assertEqual((row["unit_price"], row["price_source"]), (90, "Material Master"))

    def test_purchased_parts_rule_and_supplier_relationship_upload(self):
        def material(code, name, uom="kg", category="Raw material"):
            response = self.client.post("/api/planning/materials", json={"code": code, "name": name, "uom": uom, "category": category})
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]
        motor, steel, scrap, pcb = material("PP-MOTOR", "Motor", "PCS"), material("PP-STEEL", "Steel"), material("PP-SCRAP", "Steel scrap"), material("PP-PCB", "PCB", "PCS", "Electrical components")
        model = self.client.post("/api/planning/models", json={"code": "PP-12000", "name": "Eva"}).json()["id"]
        for part in (motor, pcb):
            self.assertEqual(self.client.post(f"/api/planning/model-bom/{model}", json={"material_id": part, "quantity": 1, "uom": "PCS"}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{motor}/new", json={"tab": "Level 2 Part BOM", "lines": [
            {"child_item_id": steel, "quantity": 2}, {"child_item_id": scrap, "quantity": -0.3}]}).status_code, 201)

        # Purchased list: steel + pcb yes; motor (own BOM) and scrap (negative qty) no
        listed = {row["code"] for row in self.client.get("/api/planning/materials", params={"purchased": "true", "search": "PP-", "limit": 500}).json()["items"]}
        self.assertEqual(listed, {"PP-STEEL", "PP-PCB"})
        everything = {row["code"] for row in self.client.get("/api/planning/materials", params={"search": "PP-", "limit": 500}).json()["items"]}
        self.assertEqual(everything, {"PP-MOTOR", "PP-STEEL", "PP-SCRAP", "PP-PCB"})

        supplier = next(row["id"] for row in self.client.get("/api/suppliers").json() if row["status"] == "Approved")
        def link(material_id):
            return self.client.post("/api/planning/supplier-parts", json={"material_id": material_id, "supplier_id": supplier, "sob_percent": 50, "unit_price": 1, "currency": "INR"})
        made = link(motor)
        self.assertEqual(made.status_code, 422)
        self.assertIn("made in-house", made.json()["detail"])
        generated = link(scrap)
        self.assertEqual(generated.status_code, 422)
        self.assertIn("generated in your factory", generated.json()["detail"])
        self.assertEqual(link(steel).status_code, 201)

        # Template
        template = load_workbook(io.BytesIO(self.client.get("/api/planning/templates/supplier-parts.xlsx").content)).active
        self.assertEqual([c.value for c in template[1]][:7], ["material_code", "supplier_name", "tier", "parent_supplier_name", "sob_percent", "unit_price", "currency"])

        def upload(rows, headers=("material_code", "supplier_name", "tier", "parent_supplier_name", "sob_percent", "unit_price", "currency", "lead_time_days", "moq", "buffer_days")):
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(list(headers))
            for row in rows:
                sheet.append(list(row))
            content = io.BytesIO()
            workbook.save(content)
            return self.client.post("/api/planning/supplier-parts/upload", content=content.getvalue(),
                                    headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})

        def links_of(material_id):
            return {(row["supplier_name"], row["tier"]): row for row in self.client.get("/api/planning/supplier-parts", params={"limit": 1000}).json()["items"] if row["material_id"] == material_id}

        before_suppliers = len(self.client.get("/api/suppliers").json())
        # Rejected files change nothing (all rows or none)
        for rows, expected in [
            ([("PP-PCB", "Iljin Noida", "", "", 60, 1800, "INR", 90, 0, 0), ("PP-MOTOR", "Iljin Noida", "", "", 40, 1, "INR", 0, 0, 0)], "made in-house"),
            ([("PP-PCB", "Iljin Noida", "", "", 60, 1800, "INR", 90, 0, 0), ("PP-NOPE", "Iljin Noida", "", "", 40, 1, "INR", 0, 0, 0)], "is not in T1, T2 or T3 Material Master"),
            ([("PP-PCB", "Iljin Noida", "", "", 70, 1800, "INR", 90, 0, 0), ("PP-PCB", "LG Pune", "", "", 40, 1900, "INR", 60, 0, 0)], "more than 100%"),
            ([("PP-PCB", "Iljin Noida", "", "", 60, 1800, "INR", 90, 0, 0), ("PP-PCB", "Iljin Noida", "", "", 40, 1800, "INR", 90, 0, 0)], "repeated"),
            ([("PP-PCB", "Chip Co", "T2", "", 100, 5, "INR", 0, 0, 0)], "parent_supplier_name"),
        ]:
            response = upload(rows)
            self.assertEqual(response.status_code, 422, response.text)
            self.assertIn(expected, response.json()["detail"])
        self.assertEqual(len(self.client.get("/api/suppliers").json()), before_suppliers)
        self.assertEqual(links_of(pcb), {})

        # Good file: 2 new suppliers + a Tier-2 sub-supplier, 3 links
        response = upload([
            ("PP-PCB", "Iljin Noida", "", "", 60, 1800, "inr", 90, 500, 10),
            ("pp-pcb", "LG Pune", "Direct", "", 40, 1900, "INR", 60, 0, 5),
            ("PP-PCB", "Chip Co", "T2", "Iljin Noida", 100, 0, "INR", 30, 0, 0),
        ])
        self.assertEqual(response.status_code, 201, response.text)
        self.assertEqual(response.json(), {"rows": 3, "added": 3, "updated": 0, "suppliers_created": 3})
        found = links_of(pcb)
        self.assertEqual(found[("Iljin Noida", "T1")]["sob_percent"], 60)
        self.assertEqual(found[("Iljin Noida", "T1")]["unit_price"], 1800)
        self.assertEqual(found[("Iljin Noida", "T1")]["currency"], "INR")
        self.assertEqual((found[("Iljin Noida", "T1")]["lead_time_days"], found[("Iljin Noida", "T1")]["moq"]), (90, 500))
        self.assertEqual(found[("Chip Co", "T2")]["parent_supplier_name"], "Iljin Noida")
        # BOM price picks it up: (1800 x 60 + 1900 x 40) / 100 = 1840
        row = next(r for r in self.client.get(f"/api/planning/model-bom/{model}").json()["items"] if r["material_id"] == pcb)
        self.assertAlmostEqual(row["unit_price"], 1840)
        # Re-upload updates the same links (by name, case-insensitive), no duplicates
        response = upload([("PP-PCB", "iljin noida", "", "", 50, 1750, "INR", 80, 0, 0), ("PP-PCB", "LG PUNE", "", "", 50, 1900, "INR", 60, 0, 0)])
        self.assertEqual(response.json(), {"rows": 2, "added": 0, "updated": 2, "suppliers_created": 0})
        self.assertEqual(found.keys(), links_of(pcb).keys())
        self.assertEqual(links_of(pcb)[("Iljin Noida", "T1")]["unit_price"], 1750)

    def test_supplier_relationship_upload_accepts_users_sheet_as_is(self):
        # The user's sheet: rows 1-3 empty, heading on row 4, columns as in their screenshot
        data = [
            ("FL001ADC12", "ADC12", "GTR Bangalore", "Raw material", "kg", 20000, 15, 350, "INR"),
            ("FL001CUWIRE06", "CUWire06", "GKN Wire Bangalore", "Raw material", "kg", 10000, 60, 1200, "INR"),
            ("FL001LM6", "LM6", "GTR Bangalore", "Raw material", "kg", 15000, 15, 320, "INR"),
            ("FL001PCB1", "PCB1", "Iljin Noida", "Electrical components", "PCS", 60000, 90, 1800, "INR"),
            ("FL001PCB2", "PCB2", "Iljin Noida", "Electrical components", "PCS", 50000, 90, 2000, "INR"),
            ("FL001PPB200", "PPB200", "Reliance Goa", "Raw material", "kg", 20000, 60, 150, "INR"),
            ("FL001STEEL0.7", "STEEL0.7", "JGI Goa", "Raw material", "kg", 10000, 5, 100, "INR"),
            ("FL001STEEL0.8", "STEEL0.8", "JGI Goa", "Raw material", "kg", 12000, 5, 105, "INR"),
        ]
        for code, name, _, category, uom, *_ in data:
            self.assertEqual(self.client.post("/api/planning/materials", json={"code": code, "name": name, "category": category, "uom": uom}).status_code, 201)

        def sheet_bytes(rows, header=("Material code", "Material name", "Supplier Name", "Category", "Unit", "On hand", "Buffer days", "Unit cost", "Currency")):
            workbook = Workbook()
            sheet = workbook.active
            sheet.append([]); sheet.append([]); sheet.append([])
            sheet.append(list(header))
            for row in rows:
                sheet.append(list(row))
            content = io.BytesIO()
            workbook.save(content)
            return content.getvalue()

        def upload(content):
            return self.client.post("/api/planning/supplier-parts/upload", content=content,
                                    headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})

        response = upload(sheet_bytes(data))
        self.assertEqual(response.status_code, 201, response.text)
        self.assertEqual(response.json(), {"rows": 8, "added": 8, "updated": 0, "suppliers_created": 5})
        links = {row["material_code"]: row for row in self.client.get("/api/planning/supplier-parts", params={"limit": 1000}).json()["items"] if row["material_code"].startswith("FL001")}
        self.assertEqual(len(links), 8)
        steel = links["FL001STEEL0.8"]
        self.assertEqual((steel["supplier_name"], steel["tier"], steel["sob_percent"], steel["unit_price"], steel["currency"], steel["buffer_days"]),
                         ("JGI Goa", "T1", 100, 105, "INR", 5))
        self.assertEqual((links["FL001PCB1"]["supplier_name"], links["FL001PCB1"]["unit_price"], links["FL001PCB1"]["buffer_days"]), ("Iljin Noida", 1800, 90))
        # "On hand" is not a supplier field: Material Master stock is unchanged
        material = next(m for m in self.client.get("/api/planning/materials", params={"search": "FL001STEEL0.8"}).json()["items"] if m["code"] == "FL001STEEL0.8")
        self.assertEqual(material["stock_on_hand"], 0)

        # Errors name the real Excel row (row 4 = heading, so the 2nd data row is row 6)
        bad = list(data[:1]) + [("FL001NOPE", "x", "JGI Goa", "Raw material", "kg", 0, 0, 1, "INR")]
        response = upload(sheet_bytes(bad))
        self.assertEqual(response.status_code, 422)
        self.assertIn("Row 6", response.json()["detail"])
        # A sheet without the key columns is refused with a clear message
        response = upload(sheet_bytes(data[:1], header=("Code", "Name", "Vendor X")))
        self.assertEqual(response.status_code, 422)
        self.assertIn("heading row", response.json()["detail"])

    def test_supplier_materials_and_supplier_bom_cost_breakdown(self):
        post = self.client.post
        def material(code, name, uom="PCS", cost=0):
            response = post("/api/planning/materials", json={"code": code, "name": name, "uom": uom, "unit_cost": cost, "currency": "INR"})
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]
        pcb, chip, chip_neo, motor = material("SB-PCB", "PCB1"), material("SB-CHIP", "MicroChip", cost=500), material("SB-CHIPNEO", "MicroChipNeo"), material("SB-MOTOR", "Motor")
        steel = material("SB-STEEL", "Steel", "kg")
        model = post("/api/planning/models", json={"code": "SB-12000", "name": "Eva"}).json()["id"]
        post(f"/api/planning/model-bom/{model}", json={"material_id": motor, "quantity": 1, "uom": "PCS"})
        post(f"/api/planning/part-bom/{motor}/new", json={"tab": "Level 2 Part BOM", "lines": [{"child_item_id": steel, "quantity": 1}]})
        def supplier(name):
            response = post("/api/suppliers", json={"name": name, "category": "Electrical", "country": "India"})
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]
        iljin, sub_a, sub_b, other = supplier("Iljin Noida"), supplier("Chip Maker A"), supplier("Chip Maker B"), supplier("Other PCB Co")
        self.assertEqual(post("/api/planning/supplier-parts", json={"material_id": pcb, "supplier_id": iljin, "sob_percent": 100, "unit_price": 1800, "currency": "INR"}).status_code, 201)

        # --- supplier materials: own list per supplier
        def own(code, name, uom, price, supplier=iljin):
            response = post("/api/planning/supplier-materials", json={"supplier_id": supplier, "code": code, "name": name, "uom": uom, "unit_price": price, "currency": "INR"})
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]
        board, icset, solder, paste, flux = own("BRD", "Bare board", "PCS", 415), own("ICSET", "IC set", "SET", 0), own("SOLDER", "Solder wire", "kg", 900), own("PASTE", "Solder paste", "kg", 0), own("FLUX", "Flux", "kg", 20)
        other_own = own("BRD", "Other's board", "PCS", 1, supplier=other)  # same code at another supplier is fine
        def pick_t1(item_id, price):
            response = post("/api/planning/supplier-materials", json={"supplier_id": iljin, "item_id": item_id, "unit_price": price, "currency": "INR"})
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]
        chip_own, chip_neo_own = pick_t1(chip, 505), pick_t1(chip_neo, 0)
        self.assertEqual(post("/api/planning/supplier-materials", json={"supplier_id": iljin, "code": "brd", "name": "dup", "uom": "PCS"}).status_code, 409)
        listed = self.client.get("/api/planning/supplier-materials", params={"supplier_id": iljin}).json()
        self.assertEqual([m["code"] for m in listed["items"]], ["BRD", "FLUX", "ICSET", "PASTE", "SB-CHIP", "SB-CHIPNEO", "SOLDER"])

        # --- supplier BOM lines
        def line(component, quantity, parent=None, price=None, our=False, supplier=iljin, product=pcb):
            body = {"supplier_id": supplier, "material_id": product, "parent_line_id": parent, "quantity": quantity, "unit_price": price,
                    ("item_id" if our else "supplier_material_id"): component}
            return post("/api/planning/supplier-bom/lines", json=body)
        l_board = line(board, 1).json()["id"]
        l_icset = line(icset, 1).json()["id"]
        not_in_master = line(chip, 2, parent=l_icset, our=True)
        self.assertEqual(not_in_master.status_code, 422, "codes come only from the supplier's material master")
        self.assertIn("material master", not_in_master.json()["detail"])
        l_chip = line(chip_own, 2, parent=l_icset).json()["id"]
        l_solder = line(solder, 0.01, parent=l_icset).json()["id"]
        l_paste = line(paste, 1, parent=l_icset).json()["id"]
        l_flux = line(flux, 0.5, parent=l_paste).json()["id"]
        # rules
        self.assertEqual(line(flux, 1, parent=l_flux).status_code, 422, "no L4")
        self.assertEqual(line(other_own, 1).status_code, 422, "other supplier's code")
        self.assertEqual(line(board, 0).status_code, 422, "zero quantity")
        self.assertEqual(line(board, 1, product=motor).status_code, 422, "in-house part cannot have a supplier BOM")
        both = post("/api/planning/supplier-bom/lines", json={"supplier_id": iljin, "material_id": pcb, "quantity": 1, "item_id": chip, "supplier_material_id": board})
        self.assertEqual(both.status_code, 422)

        # sources on the MicroChip line: A 480 @ 60 %, B (alternate code) 520 @ 40 % -> 288 + 208 = 496
        put = self.client.put
        self.assertEqual(put(f"/api/planning/supplier-bom/lines/{l_chip}/sources", json=[
            {"sub_supplier_id": sub_a, "unit_price": 480, "currency": "INR", "sob_percent": 60},
            {"sub_supplier_id": sub_b, "supplier_material_id": chip_neo_own, "unit_price": 520, "currency": "INR", "sob_percent": 50}]).status_code, 422, "SOB > 100")
        self.assertEqual(put(f"/api/planning/supplier-bom/lines/{l_icset}/sources", json=[
            {"sub_supplier_id": sub_a, "unit_price": 1, "currency": "INR", "sob_percent": 100}]).status_code, 200, "a line with lines below can have its suppliers too")
        self.assertEqual(put(f"/api/planning/supplier-bom/lines/{l_icset}/sources", json=[]).status_code, 200)  # back to the sum of its lines
        self.assertEqual(put(f"/api/planning/supplier-bom/lines/{l_chip}/sources", json=[
            {"sub_supplier_id": sub_a, "unit_price": 480, "currency": "INR", "sob_percent": 60},
            {"sub_supplier_id": sub_b, "supplier_material_id": chip_neo_own, "unit_price": 520, "currency": "INR", "sob_percent": 40}]).status_code, 200)

        view = self.client.get("/api/planning/supplier-bom", params={"supplier_id": iljin, "material_id": pcb}).json()
        rows = {row["id"]: row for row in view["lines"]}
        self.assertEqual([row["component"]["code"] for row in view["lines"]], ["BRD", "ICSET", "SB-CHIP", "SOLDER", "PASTE", "FLUX"], "tree order")
        self.assertEqual([rows[i]["level"] for i in (l_board, l_icset, l_chip, l_solder, l_paste, l_flux)], ["L1", "L1", "L2", "L2", "L2", "L3"])
        self.assertAlmostEqual(rows[l_board]["line_cost"], 415)
        self.assertAlmostEqual(rows[l_chip]["unit_price"], 496)
        self.assertEqual(rows[l_chip]["price_source"], "SOB-weighted (2 sources)")
        self.assertEqual(rows[l_chip]["sources"][1]["alternate"]["code"], "SB-CHIPNEO")
        self.assertAlmostEqual(rows[l_chip]["line_cost"], 992)
        self.assertAlmostEqual(rows[l_solder]["line_cost"], 9)
        self.assertAlmostEqual(rows[l_paste]["line_cost"], 10)
        self.assertEqual(rows[l_paste]["price_source"], "Sum of lines below")
        self.assertAlmostEqual(rows[l_icset]["line_cost"], 1011)
        self.assertAlmostEqual(view["should_cost"], 1426)
        self.assertEqual(view["quoted_price"], 1800)
        self.assertAlmostEqual(view["gap"], 374)

        # entered price on a line overrides the list price; edit quantity
        self.assertEqual(self.client.patch(f"/api/planning/supplier-bom/lines/{l_board}", json={"quantity": 2, "unit_price": 400, "currency": "INR"}).status_code, 200)
        view = self.client.get("/api/planning/supplier-bom", params={"supplier_id": iljin, "material_id": pcb}).json()
        self.assertAlmostEqual(view["should_cost"], 800 + 1011)
        # a line in another currency is not summed and is listed
        self.assertEqual(self.client.patch(f"/api/planning/supplier-bom/lines/{l_solder}", json={"quantity": 0.01, "unit_price": 10, "currency": "USD"}).status_code, 200)
        view = self.client.get("/api/planning/supplier-bom", params={"supplier_id": iljin, "material_id": pcb}).json()
        self.assertEqual(view["excluded_lines"], [l_solder])
        self.assertAlmostEqual(view["should_cost"], 800 + 992 + 10)

        # supplier material used in a BOM cannot be deleted; deleting a line removes the lines below it
        self.assertEqual(self.client.delete(f"/api/planning/supplier-materials/{flux}").status_code, 409)
        self.assertEqual(self.client.delete(f"/api/planning/supplier-bom/lines/{l_icset}").status_code, 204)
        view = self.client.get("/api/planning/supplier-bom", params={"supplier_id": iljin, "material_id": pcb}).json()
        self.assertEqual([row["component"]["code"] for row in view["lines"]], ["BRD"])
        self.assertEqual(self.client.delete(f"/api/planning/supplier-materials/{flux}").status_code, 204)
        products = self.client.get("/api/planning/supplier-bom/products", params={"supplier_id": iljin}).json()["items"]
        self.assertEqual([(p["code"], p["has_bom"]) for p in products], [("SB-PCB", True)])

        # --- Excel: supplier materials (own headings) and supplier BOM upload (replaces), export
        def xlsx(header, rows, blank_rows=0):
            workbook = Workbook(); sheet = workbook.active
            for _ in range(blank_rows):
                sheet.append([])
            sheet.append(list(header))
            for row in rows:
                sheet.append(list(row))
            content = io.BytesIO(); workbook.save(content)
            return content.getvalue()
        xl = {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        response = post(f"/api/planning/supplier-materials/upload?supplier_id={iljin}", headers=xl,
                        content=xlsx(("Material code", "Material name", "Unit", "Unit cost", "Currency"), [("TERM", "Terminal", "PCS", 2, "INR"), ("SOLDER", "Solder wire 0.8", "kg", 950, "INR")], blank_rows=2))
        self.assertEqual(response.status_code, 201, response.text)
        self.assertEqual(response.json(), {"rows": 2, "added": 1, "updated": 1, "linked": 0})
        response = post(f"/api/planning/supplier-bom/upload?supplier_id={iljin}&material_id={pcb}", headers=xl, content=xlsx(
            ("Parent code", "Code", "Qty", "Unit price"),
            [("", "BRD", 1, ""), ("", "ICSET", 1, ""), ("ICSET", "SB-CHIP", 2, 500), ("ICSET", "SOLDER", 0.01, ""), ("SOLDER", "TERM", 3, "")]))
        self.assertEqual(response.status_code, 201, response.text)
        view = self.client.get("/api/planning/supplier-bom", params={"supplier_id": iljin, "material_id": pcb}).json()
        self.assertEqual([(r["level"], r["component"]["code"]) for r in view["lines"]], [("L1", "BRD"), ("L1", "ICSET"), ("L2", "SB-CHIP"), ("L2", "SOLDER"), ("L3", "TERM")])
        # BRD 415 + ICSET (2 x 500 + 0.01 x (3 x 2)) = 415 + 1000.06
        self.assertAlmostEqual(view["should_cost"], 1415.06)
        plain = TestClient(self.client.app, raise_server_exceptions=False)
        plain.cookies = self.client.cookies
        bad = plain.post(f"/api/planning/supplier-bom/upload?supplier_id={iljin}&material_id={pcb}", headers=xl, content=xlsx(
            ("Parent code", "Code", "Qty"), [("", "BRD", 1), ("NOPE", "TERM", 1)]))
        self.assertEqual(bad.status_code, 422)
        self.assertIn("earlier row", bad.json()["detail"])
        self.assertEqual(len(self.client.get("/api/planning/supplier-bom", params={"supplier_id": iljin, "material_id": pcb}).json()["lines"]), 5, "bad file changes nothing")
        export = self.client.get("/api/planning/supplier-bom/xlsx", params={"supplier_id": iljin, "material_id": pcb})
        self.assertEqual(export.status_code, 200)
        sheet = load_workbook(io.BytesIO(export.content)).active
        self.assertEqual([c.value for c in sheet[4]][:8], ["Level", "L1 code", "L1 name", "L2 code", "L2 name", "L3 code", "L3 name", "Source"])
        self.assertEqual([sheet.cell(r, 6).value for r in range(5, 10)], [None, None, None, None, "TERM"])

        # users without price permission see no prices
        self.assertEqual(post("/api/users", json={"username": "sbview", "full_name": "Viewer", "password": "Viewer-pass-2026!", "role": "User"}).status_code, 201)
        self.client.post("/api/auth/logout")
        self.client.post("/api/auth/login", json={"username": "sbview", "password": "Viewer-pass-2026!"})
        view = self.client.get("/api/planning/supplier-bom", params={"supplier_id": iljin, "material_id": pcb}).json()
        self.assertNotIn("should_cost", view)
        self.assertNotIn("line_cost", view["lines"][0])
        self.assertEqual(line(board, 1).status_code, 403, "view-only users cannot edit")

    def test_t1_t2_t3_material_masters_and_supplier_material_links(self):
        post, patch = self.client.post, self.client.patch
        steel = post("/api/planning/materials", json={"code": "TT-STEEL", "name": "Steel XYZ", "uom": "kg", "unit_cost": 70, "currency": "INR"}).json()["id"]
        def tier(t, code, name, price=0):
            return post("/api/planning/tier-materials", json={"tier": t, "code": code, "name": name, "uom": "kg", "unit_price": price, "currency": "INR"})
        coil = tier("T2", "TT-COIL", "HR coil", 55)
        self.assertEqual(coil.status_code, 201, coil.text)
        ore = tier("T3", "TT-ORE", "Iron ore", 8)
        self.assertEqual(ore.status_code, 201)
        coil, ore = coil.json()["id"], ore.json()["id"]
        # one code, one master: refused in the other masters (case-insensitive), both directions
        for t, code in (("T2", "tt-steel"), ("T3", "TT-COIL"), ("T2", "tt-ore")):
            response = tier(t, code, "dup")
            self.assertEqual(response.status_code, 409, (t, code))
            self.assertIn("Material Master", response.json()["detail"])
        dup_t1 = post("/api/planning/materials", json={"code": "TT-COIL", "name": "dup", "uom": "kg"})
        self.assertEqual(dup_t1.status_code, 409)
        self.assertIn("T2 Material Master", dup_t1.json()["detail"])
        self.assertEqual([m["code"] for m in self.client.get("/api/planning/tier-materials", params={"tier": "T2"}).json()["items"]], ["TT-COIL"])
        # T2 / T3 codes are not part of your planning lists
        self.assertNotIn("TT-COIL", {m["code"] for m in self.client.get("/api/planning/materials", params={"limit": 500}).json()["items"]})

        supplier = post("/api/suppliers", json={"name": "Part Maker", "category": "Steel parts", "country": "India"}).json()["id"]
        def pick(**body):
            return post("/api/planning/supplier-materials", json={"supplier_id": supplier, **body})
        # picked from T1 / T2 / T3: code + name come from your master; the supplier's own price,
        # lead time and buffer (your discounted 70 vs the supplier's 74.5)
        s_steel = pick(item_id=steel, unit_price=74.5, currency="INR", lead_time_days=21, buffer_days=7)
        self.assertEqual(s_steel.status_code, 201, s_steel.text)
        self.assertEqual({k: s_steel.json()[k] for k in ("code", "name", "uom", "source", "unit_price", "lead_time_days", "buffer_days")},
                         {"code": "TT-STEEL", "name": "Steel XYZ", "uom": "kg", "source": "T1", "unit_price": 74.5, "lead_time_days": 21, "buffer_days": 7})
        self.assertEqual(pick(tier_material_id=coil, unit_price=60).json()["source"], "T2")
        self.assertEqual(pick(tier_material_id=ore).json()["source"], "T3")
        self.assertEqual(pick(item_id=steel).status_code, 409, "already in this supplier's master")
        # the supplier's own code (fastener) - but not one of your master codes typed by hand
        own = pick(code="PM-FAST", name="M6 bolt", uom="PCS", unit_price=2)
        self.assertEqual((own.status_code, own.json()["source"]), (201, "Own"))
        typed = pick(code="tt-steel", name="Steel", uom="kg")
        self.assertEqual(typed.status_code, 422)
        self.assertIn("pick it from T1", typed.json()["detail"])
        # renaming the master entry flows into the supplier's master; price stays the supplier's
        self.assertEqual(patch(f"/api/planning/tier-materials/{coil}", json={"tier": "T2", "code": "TT-COIL2", "name": "HR coil 2mm", "uom": "kg", "unit_price": 50, "currency": "INR"}).status_code, 200)
        self.assertEqual(patch(f"/api/planning/materials/{steel}", json={"code": "TT-STEEL", "name": "Steel XYZ 0.8", "category": "Raw material", "uom": "kg", "stock_on_hand": 0, "buffer_days": 0, "unit_cost": 68, "currency": "INR"}).status_code, 200)
        listed = {m["code"]: m for m in self.client.get("/api/planning/supplier-materials", params={"supplier_id": supplier}).json()["items"]}
        self.assertIn("TT-COIL2", listed, "renamed master code must flow into the supplier's master")
        self.assertEqual(listed["TT-COIL2"]["name"], "HR coil 2mm")
        self.assertEqual(listed["TT-COIL2"]["unit_price"], 60)
        self.assertEqual((listed["TT-STEEL"]["name"], listed["TT-STEEL"]["unit_price"]), ("Steel XYZ 0.8", 74.5))
        # editing a picked entry changes only the supplier's terms, never code / name
        self.assertEqual(patch(f"/api/planning/supplier-materials/{s_steel.json()['id']}", json={"supplier_id": supplier, "code": "HACK", "name": "Hack", "unit_price": 80, "currency": "INR", "lead_time_days": 30, "buffer_days": 10}).status_code, 200)
        edited = next(m for m in self.client.get("/api/planning/supplier-materials", params={"supplier_id": supplier}).json()["items"] if m["id"] == s_steel.json()["id"])
        self.assertEqual((edited["code"], edited["name"], edited["unit_price"], edited["lead_time_days"], edited["buffer_days"]), ("TT-STEEL", "Steel XYZ 0.8", 80, 30, 10))
        # in use -> cannot delete the master entry
        self.assertEqual(self.client.delete(f"/api/planning/tier-materials/{ore}").status_code, 409)
        self.assertEqual(self.client.delete(f"/api/planning/materials/{steel}").status_code, 409)

        def xlsx(header, rows):
            workbook = Workbook(); sheet = workbook.active
            sheet.append(list(header))
            for row in rows:
                sheet.append(list(row))
            content = io.BytesIO(); workbook.save(content)
            return content.getvalue()
        xl = {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        # T3 upload: new code added; a code from another master refused (nothing saved)
        bad = post("/api/planning/tier-materials/upload?tier=T3", headers=xl, content=xlsx(("Material code", "Material name", "Unit"), [("TT-NICKEL", "Nickel", "kg"), ("TT-STEEL", "x", "kg")]))
        self.assertEqual(bad.status_code, 422)
        self.assertIn("T1 Material Master", bad.json()["detail"])
        good = post("/api/planning/tier-materials/upload?tier=T3", headers=xl, content=xlsx(("Material code", "Material name", "Unit", "Unit cost"), [("TT-NICKEL", "Nickel", "kg", 900), ("TT-ORE", "Iron ore fines", "kg", 9)]))
        self.assertEqual(good.json(), {"rows": 2, "added": 1, "updated": 1})
        # supplier material upload: master codes linked automatically, others become own codes
        response = post(f"/api/planning/supplier-materials/upload?supplier_id={supplier}", headers=xl, content=xlsx(
            ("Code", "Name", "Unit", "Unit price", "Lead time", "Buffer days"),
            [("TT-NICKEL", "", "", 950, 45, 15), ("PM-NUT", "M6 nut", "PCS", 1, 7, 30)]))
        self.assertEqual(response.status_code, 201, response.text)
        self.assertEqual(response.json(), {"rows": 2, "added": 2, "updated": 0, "linked": 1})
        listed = {m["code"]: m for m in self.client.get("/api/planning/supplier-materials", params={"supplier_id": supplier}).json()["items"]}
        self.assertEqual((listed["TT-NICKEL"]["source"], listed["TT-NICKEL"]["name"], listed["TT-NICKEL"]["unit_price"], listed["TT-NICKEL"]["lead_time_days"]), ("T3", "Nickel", 950, 45))
        self.assertEqual((listed["PM-NUT"]["source"], listed["PM-NUT"]["buffer_days"]), ("Own", 30))
        self.assertEqual(listed["TT-ORE"]["name"], "Iron ore fines", "upload rename flows into the supplier's master")
        # code search for the picker
        found = self.client.get("/api/planning/master-codes", params={"tier": "T2", "search": "coil"}).json()["items"]
        self.assertEqual([m["code"] for m in found], ["TT-COIL2"])
        export = self.client.get("/api/planning/tier-materials/T3.xlsx")
        self.assertEqual(export.status_code, 200)

    def test_alternate_codes_sob_per_bom_line_price_tree_and_planning(self):
        # Alternate SOB is set per BOM line, 0% until set. Hand-worked:
        #   Motor BOM line CHIP qty 2, NEO 80%: main 20% -> 0.4 / 1.6; price 0.2x480 + 0.8x520 = 512
        #   Direct L1 line CHIP qty 3, NEO 30% + X 10%: main 60% -> 1.8 / 0.9 / 0.3;
        #   price 0.6x480 + 0.3x520 + 0.1x600 = 504
        #   Plan 100 models: CHIP 100x(0.4+1.8) = 220, NEO 100x(1.6+0.9) = 250, X 30 (total 500)
        def material(code, name, uom="PCS", cost=0):
            response = self.client.post("/api/planning/materials", json={"code": code, "name": name, "uom": uom, "unit_cost": cost, "currency": "INR"})
            self.assertEqual(response.status_code, 201, response.text)
            return response.json()["id"]
        chip, neo, chipx = material("AC-CHIP", "MicroChip", cost=500), material("AC-CHIPNEO", "MicroChipNeo"), material("AC-CHIPX", "MicroChip X", cost=600)
        motor, steel, other = material("AC-MOTOR", "Motor"), material("AC-STEEL", "Steel", "kg"), material("AC-OTHER", "Other chip")
        model = self.client.post("/api/planning/models", json={"code": "AC-MODEL", "name": "Alt model"}).json()["id"]
        model2 = self.client.post("/api/planning/models", json={"code": "AC-MODEL2", "name": "Alt model 2"}).json()["id"]
        for m in (model, model2):
            self.assertEqual(self.client.post(f"/api/planning/model-bom/{m}", json={"material_id": motor, "quantity": 1, "uom": "PCS"}).status_code, 201)
        direct = self.client.post(f"/api/planning/model-bom/{model}", json={"material_id": chip, "quantity": 3, "uom": "PCS"}).json()["id"]
        direct2 = self.client.post(f"/api/planning/model-bom/{model2}", json={"material_id": chip, "quantity": 3, "uom": "PCS"}).json()["id"]
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{motor}/new", json={"tab": "Level 2 Part BOM", "lines": [{"child_item_id": chip, "quantity": 2}]}).status_code, 201)
        sup_a = self.client.post("/api/suppliers", json={"name": "Chip Maker A", "category": "Electrical", "country": "India"}).json()["id"]
        sup_b = self.client.post("/api/suppliers", json={"name": "Chip Maker B", "category": "Electrical", "country": "India"}).json()["id"]
        for material_id, supplier, price in ((chip, sup_a, 480), (neo, sup_b, 520)):
            response = self.client.post("/api/planning/supplier-parts", json={"material_id": material_id, "supplier_id": supplier, "sob_percent": 100, "unit_price": price, "currency": "INR"})
            self.assertEqual(response.status_code, 201, response.text)

        def add(main, alternate):
            return self.client.post("/api/planning/alternate-codes", json={"main_item_id": main, "alternate_item_id": alternate})
        # pair rules
        self.assertIn("its own alternate", add(chip, chip).json()["detail"])
        self.assertIn("same unit", add(chip, steel).json()["detail"])
        self.assertIn("made in-house", add(motor, chip).json()["detail"])
        created = add(chip, neo)
        self.assertEqual(created.status_code, 201, created.text)
        entry = created.json()
        self.assertNotIn("sob_percent", entry)  # no SOB on the Alternate codes tab
        self.assertEqual((entry["main_code"], entry["alternate_code"], entry["main_unit_price"], entry["unit_price"]), ("AC-CHIP", "AC-CHIPNEO", 480, 520))
        self.assertEqual(add(chip, chipx).status_code, 201)
        self.assertIn("already an alternate of AC-CHIP", add(other, neo).json()["detail"])
        self.assertIn("cannot be an alternate", add(other, chip).json()["detail"])
        self.assertIn("is itself an alternate", add(neo, other).json()["detail"])

        component = self.client.get(f"/api/planning/part-bom/{motor}", params={"tab": "Level 2 Part BOM"}).json()["items"][0]["component_id"]

        def l2_row():
            return self.client.get(f"/api/planning/part-bom/{motor}", params={"tab": "Level 2 Part BOM"}).json()["items"][0]

        def l1_row(model_id):
            return next(r for r in self.client.get(f"/api/planning/model-bom/{model_id}").json()["items"] if r["material_id"] == chip)

        def plan_demand():
            report = self.client.get("/api/planning/report", params={"workspace": "trial", "start_month": "2026-10", "end_month": "2026-10"}).json()
            return {r["material_code"]: r["plan_demand"] for r in report["items"] if r["material_code"].startswith("AC-CHIP")}

        # default 0%: main code keeps 100%, alternates listed with 0
        row = l2_row()
        self.assertEqual((row["main_sob_percent"], row["main_quantity"], row["unit_price"]), (100, 2, 480))
        self.assertEqual([(a["code"], a["sob_percent"], a["quantity"]) for a in row["alternates"]], [("AC-CHIPNEO", 0, 0), ("AC-CHIPX", 0, 0)])
        plan = self.client.post("/api/planning/plans", json={"issue_month": "2026-10", "workspace": "trial", "rows": [{"model_code": "AC-MODEL", "target_month": "2026-10", "quantity": 100}]})
        self.assertEqual(plan.status_code, 201, plan.text)
        self.assertEqual(plan_demand(), {"AC-CHIP": 500})

        def save(alternates, model_bom_id=None, component_id=None):
            return self.client.put("/api/planning/bom-line-alternates", json={"model_bom_id": model_bom_id, "component_id": component_id,
                                                                            "alternates": [{"alternate_item_id": a, "sob_percent": s} for a, s in alternates]})
        # line SOB rules
        self.assertEqual(save([(neo, 10)]).status_code, 422)  # no line given
        self.assertEqual(save([(neo, 10)], direct, component).status_code, 422)  # two lines given
        self.assertIn("not an alternate of AC-CHIP", save([(other, 10)], component_id=component).json()["detail"])
        self.assertIn("adds up to 110%", save([(neo, 70), (chipx, 40)], model_bom_id=direct).json()["detail"])
        self.assertEqual(save([(neo, -1)], model_bom_id=direct).status_code, 422)
        # Motor BOM line 80%, direct line 30% + 10%
        saved = save([(neo, 80), (chipx, 0)], component_id=component)
        self.assertEqual(saved.status_code, 200, saved.text)
        self.assertEqual(saved.json()["main_sob_percent"], 20)
        self.assertEqual(save([(neo, 30), (chipx, 10)], model_bom_id=direct).json()["main_sob_percent"], 60)

        row = l2_row()
        self.assertEqual(row["quantity"], 2)  # the BOM line itself is unchanged (Edit uses it)
        self.assertAlmostEqual(row["main_quantity"], 0.4)
        self.assertAlmostEqual(row["unit_price"], 512)
        self.assertIn("AC-CHIP 20% x 480.00 + AC-CHIPNEO 80% x 520.00", row["price_source"])
        self.assertEqual([(a["code"], a["sob_percent"]) for a in row["alternates"]], [("AC-CHIPNEO", 80), ("AC-CHIPX", 0)])
        self.assertAlmostEqual(row["alternates"][0]["quantity"], 1.6)
        row = l1_row(model)
        self.assertEqual(row["main_sob_percent"], 60)
        self.assertAlmostEqual(row["main_quantity"], 1.8)
        self.assertAlmostEqual(row["unit_price"], 504)
        self.assertEqual([round(a["quantity"], 6) for a in row["alternates"]], [0.9, 0.3])
        # the other model's direct line keeps its own SOB (0%); the shared Motor BOM line is 80% for both models
        self.assertEqual((l1_row(model2)["main_sob_percent"], l1_row(model2)["unit_price"]), (100, 480))
        many = self.client.get("/api/planning/part-boms", params={"tab": "Level 2 Part BOM"}).json()["items"]
        self.assertEqual(many[0]["main_sob_percent"], 20)
        self.assertNotIn("alternates", next(r for r in self.client.get(f"/api/planning/model-bom/{model}").json()["items"] if r["material_id"] == motor))

        # planning demand split per line
        demand = plan_demand()
        self.assertEqual({k: round(v, 6) for k, v in demand.items()}, {"AC-CHIP": 220, "AC-CHIPNEO": 250, "AC-CHIPX": 30})

        # Excel tree: per-line split, alternates right after their main line
        sheet = load_workbook(io.BytesIO(self.client.get("/api/planning/bom-tree/xlsx", params={"tab": "Model BOM Level 1", "model_id": model}).content)).active
        headers = [cell.value for cell in sheet[4]]
        rows = [[cell.value for cell in r] for r in sheet.iter_rows(min_row=5)]
        qty, price_col, l1_code, l2_code = headers.index("Quantity"), headers.index("Part price"), headers.index("L1 part code"), headers.index("L2 part code")
        main_index = next(i for i, r in enumerate(rows) if r[l2_code] == "AC-CHIP")
        self.assertAlmostEqual(rows[main_index][qty], 0.4)
        self.assertAlmostEqual(rows[main_index][price_col], 512)
        self.assertIn("(main 20% of 2)", rows[main_index][l2_code + 1])
        neo_row = rows[main_index + 1]
        self.assertEqual((neo_row[0], neo_row[l2_code], neo_row[price_col]), ("L2 alternate", "AC-CHIPNEO", 520))
        self.assertAlmostEqual(neo_row[qty], 1.6)
        self.assertIn("SOB 80%", neo_row[l2_code + 1])
        self.assertEqual(sheet.row_dimensions[5 + main_index + 1].outline_level, sheet.row_dimensions[5 + main_index].outline_level)
        direct_index = next(i for i, r in enumerate(rows) if r[l1_code] == "AC-CHIP" and r[l2_code] in (None, ""))
        self.assertAlmostEqual(rows[direct_index][qty], 1.8)
        self.assertAlmostEqual(rows[direct_index][price_col], 504)

        # a code in Alternate codes cannot be deleted from T1 (AC-CHIPX has no other use)
        blocked = self.client.delete(f"/api/planning/materials/{chipx}")
        self.assertEqual(blocked.status_code, 409)
        self.assertIn("1 alternate code entry", blocked.json()["detail"])
        # deleting a pair gives its SOB back to the main code: direct line 70% main, 0.7x480 + 0.3x520 = 492
        pair_x = next(r["id"] for r in self.client.get("/api/planning/alternate-codes").json()["items"] if r["alternate_code"] == "AC-CHIPX")
        self.assertEqual(self.client.delete(f"/api/planning/alternate-codes/{pair_x}").status_code, 204)
        row = l1_row(model)
        self.assertEqual((row["main_sob_percent"], len(row["alternates"])), (70, 1))
        self.assertAlmostEqual(row["unit_price"], 492)
        # re-adding the pair starts at 0% again on every line (its old SOB was removed)
        self.assertEqual(add(chip, chipx).status_code, 201)
        self.assertEqual({a["code"]: a["sob_percent"] for a in l1_row(model)["alternates"]}, {"AC-CHIPNEO": 30, "AC-CHIPX": 0})
        # deleting a BOM line removes its SOB: the newest line (its id is re-used) added again starts at 0%
        self.assertEqual(save([(neo, 50)], model_bom_id=direct2).status_code, 200)
        self.assertEqual(self.client.delete(f"/api/planning/model-bom/{model2}/{direct2}").status_code, 204)
        again = self.client.post(f"/api/planning/model-bom/{model2}", json={"material_id": chip, "quantity": 3, "uom": "PCS"}).json()["id"]
        self.assertEqual(again, direct2)
        self.assertEqual(l1_row(model2)["main_sob_percent"], 100)
        # re-uploading the Motor BOM keeps the line and its SOB
        workbook = Workbook()
        ws = workbook.active
        ws.append(["parent_code", "material_code", "quantity", "uom"])
        ws.append(["AC-MOTOR", "AC-CHIP", 4, "PCS"])
        buffer = io.BytesIO()
        workbook.save(buffer)
        upload = self.client.post(f"/api/planning/part-bom/{motor}/upload", params={"tab": "Level 2 Part BOM"}, content=buffer.getvalue())
        self.assertEqual(upload.status_code, 201, upload.text)
        row = l2_row()
        self.assertEqual((row["quantity"], row["main_sob_percent"]), (4, 20))
        self.assertAlmostEqual(row["main_quantity"], 0.8)

        # Excel upload of pairs with the user's own headings; SOB column not needed
        def sheet_bytes(data):
            workbook = Workbook()
            ws = workbook.active
            ws.append(["Alternate code list"])
            ws.append(["Part Code", "Alternate Code"])
            for r in data:
                ws.append(r)
            buffer = io.BytesIO()
            workbook.save(buffer)
            return buffer.getvalue()
        result = self.client.post("/api/planning/alternate-codes/upload", content=sheet_bytes([["AC-CHIP", "AC-CHIPNEO"], ["ac-chip", "AC-CHIPX"]]))
        self.assertEqual(result.status_code, 201, result.text)
        self.assertEqual((result.json()["added"], result.json()["already_there"]), (0, 2))
        bad = self.client.post("/api/planning/alternate-codes/upload", content=sheet_bytes([["AC-OTHER", "AC-STEEL"], ["AC-CHIP", "AC-NOPE"]]))
        self.assertEqual(bad.status_code, 422)
        self.assertIn("Row 3: ", bad.json()["detail"])
        self.assertEqual(self.client.get("/api/planning/alternate-codes", params={"search": "chip"}).json()["total"], 2)  # pairs unchanged by the failed upload
        for file_format in ("xlsx", "pdf"):
            self.assertEqual(self.client.get(f"/api/planning/alternate-codes.{file_format}").status_code, 200)
        template = load_workbook(io.BytesIO(self.client.get("/api/planning/templates/alternate-codes.xlsx").content)).active
        self.assertEqual([c.value for c in template[1]], ["main_code", "alternate_code"])

    def test_alternate_sob_changes_by_month_with_history_and_report_impact(self):
        # Hand-worked, per model: direct L1 CHIP 3, Motor BOM CHIP 2; plan 100 models a month.
        #   Motor line SOB NEO: 30% from the start, 80% from 2026-11 (this line only).
        #   From 2027-01 NEO replaces CHIP everywhere (100%, applied to every line).
        #   Oct-26: CHIP 3 + 1.4 = 4.4 -> 440, NEO 0.6 -> 60
        #   Nov-26 / Dec-26: CHIP 3 + 0.4 -> 340, NEO 160; change vs before Nov: CHIP -100, NEO +100
        #   Jan-27: CHIP 0, NEO 500; change vs Nov SOB: CHIP -340, NEO +340
        def material(code, name, cost=0):
            return self.client.post("/api/planning/materials", json={"code": code, "name": name, "uom": "PCS", "unit_cost": cost, "currency": "INR"}).json()["id"]
        chip, neo, motor = material("SH-CHIP", "MicroChip", 500), material("SH-NEO", "MicroChipNeo"), material("SH-MOTOR", "Motor")
        model = self.client.post("/api/planning/models", json={"code": "SH-MODEL", "name": "History model"}).json()["id"]
        self.assertEqual(self.client.post(f"/api/planning/model-bom/{model}", json={"material_id": motor, "quantity": 1, "uom": "PCS"}).status_code, 201)
        direct = self.client.post(f"/api/planning/model-bom/{model}", json={"material_id": chip, "quantity": 3, "uom": "PCS"}).json()["id"]
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{motor}/new", json={"tab": "Level 2 Part BOM", "lines": [{"child_item_id": chip, "quantity": 2}]}).status_code, 201)
        component = self.client.get(f"/api/planning/part-bom/{motor}", params={"tab": "Level 2 Part BOM"}).json()["items"][0]["component_id"]
        sup_a = self.client.post("/api/suppliers", json={"name": "SH Maker A", "category": "Electrical", "country": "India"}).json()["id"]
        sup_b = self.client.post("/api/suppliers", json={"name": "SH Maker B", "category": "Electrical", "country": "India"}).json()["id"]
        for material_id, supplier, price in ((chip, sup_a, 480), (neo, sup_b, 520)):
            self.assertEqual(self.client.post("/api/planning/supplier-parts", json={"material_id": material_id, "supplier_id": supplier, "sob_percent": 100, "unit_price": price, "currency": "INR"}).status_code, 201)
        self.assertEqual(self.client.post("/api/planning/alternate-codes", json={"main_item_id": chip, "alternate_item_id": neo}).status_code, 201)

        def save(sob, from_month="", model_bom_id=None, component_id=None, everywhere=False):
            return self.client.put("/api/planning/bom-line-alternates", json={
                "model_bom_id": model_bom_id, "component_id": component_id, "from_month": from_month, "apply_everywhere": everywhere,
                "alternates": [{"alternate_item_id": neo, "sob_percent": sob}]})
        self.assertEqual(save(30, component_id=component).status_code, 200)
        self.assertEqual(save(80, "2026-11", component_id=component).status_code, 200)
        self.assertEqual(save(50, "2026-13", component_id=component).status_code, 422)
        replaced = save(100, "Jan-2027", model_bom_id=direct, everywhere=True)
        self.assertEqual(replaced.status_code, 200, replaced.text)
        self.assertEqual((replaced.json()["from_month"], replaced.json()["lines"]), ("2027-01", 2))

        # history of the Motor line: what changed and from when
        history = self.client.get("/api/planning/bom-line-alternates", params={"component_id": component}).json()["history"]
        self.assertEqual([(h["from_month"], h["alternates"][0]["sob_percent"], h["main_sob_percent"]) for h in history],
                         [("", 30, 70), ("2026-11", 80, 20), ("2027-01", 100, 0)])
        direct_history = self.client.get("/api/planning/bom-line-alternates", params={"model_bom_id": direct}).json()["history"]
        self.assertEqual([(h["from_month"], h["main_sob_percent"]) for h in direct_history], [("2027-01", 0)])

        plan_rows = [{"model_code": "SH-MODEL", "target_month": m, "quantity": 100} for m in ("2026-10", "2026-11", "2026-12", "2027-01")]
        self.assertEqual(self.client.post("/api/planning/plans", json={"issue_month": "2026-10", "workspace": "trial", "rows": plan_rows}).status_code, 201)

        def report():
            items = self.client.get("/api/planning/report", params={"workspace": "trial", "start_month": "2026-10", "end_month": "2027-01"}).json()["items"]
            return {(r["month"], r["material_code"]): (round(r["plan_demand"], 6), r["sob_change_from"], round(r["sob_demand_change"], 6))
                    for r in items if r["material_code"] in ("SH-CHIP", "SH-NEO")}
        self.assertEqual(report(), {
            ("2026-10", "SH-CHIP"): (440, "", 0), ("2026-10", "SH-NEO"): (60, "", 0),
            ("2026-11", "SH-CHIP"): (340, "2026-11", -100), ("2026-11", "SH-NEO"): (160, "2026-11", 100),
            ("2026-12", "SH-CHIP"): (340, "2026-11", -100), ("2026-12", "SH-NEO"): (160, "2026-11", 100),
            ("2027-01", "SH-CHIP"): (0, "2027-01", -340), ("2027-01", "SH-NEO"): (500, "2027-01", 340),
        })
        sheet = load_workbook(io.BytesIO(self.client.get("/api/planning/report.xlsx", params={"workspace": "trial", "start_month": "2026-10", "end_month": "2027-01"}).content)).active
        header_row = next(r for r in sheet.iter_rows(values_only=True) if r and "Plan revision" in r)
        self.assertEqual(list(header_row[-2:]), ["SOB change from", "Demand change from SOB change"])

        # BOM rows show the SOB valid this month and the next change
        today = date.today().strftime("%Y-%m")
        expected_now = 30 if today < "2026-11" else 80 if today < "2027-01" else 100
        expected_next = ("2026-11", 20) if today < "2026-11" else ("2027-01", 0) if today < "2027-01" else None
        row = self.client.get(f"/api/planning/part-bom/{motor}", params={"tab": "Level 2 Part BOM"}).json()["items"][0]
        self.assertEqual((row["alternates"][0]["sob_percent"], row["main_sob_percent"]), (expected_now, 100 - expected_now))
        self.assertEqual((row["sob_next_change"]["from_month"], row["sob_next_change"]["main_sob_percent"]) if row["sob_next_change"] else None, expected_next)

        # undo one change: remove the 2027-01 change on the direct line -> Jan-27 CHIP 3 x 100 = 300 again
        self.assertEqual(self.client.delete("/api/planning/bom-line-alternates", params={"model_bom_id": direct, "from_month": "2027-01"}).status_code, 204)
        self.assertEqual(self.client.delete("/api/planning/bom-line-alternates", params={"model_bom_id": direct, "from_month": "2027-01"}).status_code, 404)
        self.assertEqual(report()[("2027-01", "SH-CHIP")][0], 300)
        # re-saving the same month replaces that change (no duplicate rows)
        import sqlite3
        db_file = Path(self.temp_dir.name) / "test.db"
        count = lambda: sqlite3.connect(db_file).execute("SELECT COUNT(*) FROM bom_line_alternates WHERE component_id = ?", (component,)).fetchone()[0]
        before = count()
        self.assertEqual(save(70, "2026-11", component_id=component).status_code, 200)
        self.assertEqual(count(), before)
        # removing a middle change leaves the others: Nov-26 back to 30% (start), Jan-27 still 100%
        self.assertEqual(self.client.delete("/api/planning/bom-line-alternates", params={"component_id": component, "from_month": "2026-11"}).status_code, 204)
        history = self.client.get("/api/planning/bom-line-alternates", params={"component_id": component}).json()["history"]
        self.assertEqual([(h["from_month"], h["alternates"][0]["sob_percent"]) for h in history], [("", 30), ("2027-01", 100)])
        self.assertEqual(report()[("2026-11", "SH-CHIP")][0], 440)

    def test_manage_sob_from_alternate_codes_on_all_uses_and_change_log(self):
        # STEEL is used in 3 places: L1 direct in two models and in the Motor (L2) BOM.
        # Manage SOB sets ALT 30% from the start on all 3 uses, then 80% from 2026-11 on the Motor use only.
        def material(code, name):
            return self.client.post("/api/planning/materials", json={"code": code, "name": name, "uom": "kg", "unit_cost": 100, "currency": "INR"}).json()["id"]
        steel, alt, alt2 = material("MS-STEEL", "Steel 0.8"), material("MS-ALT", "Import steel"), material("MS-ALT2", "Import steel new")
        motor = self.client.post("/api/planning/materials", json={"code": "MS-MOTOR", "name": "Motor", "uom": "PCS"}).json()["id"]
        stator = self.client.post("/api/planning/materials", json={"code": "MS-STATOR", "name": "Stator", "uom": "PCS"}).json()["id"]
        models = [self.client.post("/api/planning/models", json={"code": f"MS-M{i}", "name": f"Model {i}"}).json()["id"] for i in (1, 2)]
        roots = [self.client.post(f"/api/planning/model-bom/{m}", json={"material_id": steel, "quantity": 6, "uom": "kg"}).json()["id"] for m in models]
        self.assertEqual(self.client.post(f"/api/planning/model-bom/{models[0]}", json={"material_id": motor, "quantity": 1, "uom": "PCS"}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{motor}/new", json={"tab": "Level 2 Part BOM", "lines": [{"child_item_id": steel, "quantity": 2}, {"child_item_id": stator, "quantity": 1}]}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{stator}/new", json={"tab": "Level 3 Part BOM", "lines": [{"child_item_id": steel, "quantity": 4}]}).status_code, 201)
        for code_id in (alt, alt2):
            self.assertEqual(self.client.post("/api/planning/alternate-codes", json={"main_item_id": steel, "alternate_item_id": code_id}).status_code, 201)
        listed = self.client.get("/api/planning/alternate-codes").json()["items"]
        self.assertEqual(listed[0]["uses"], {"L1": 2, "L2": 1, "L3": 1})

        view = self.client.get("/api/planning/alternate-codes/uses", params={"main_item_id": steel, "month": "2026-10"}).json()
        self.assertEqual([a["code"] for a in view["alternates"]], ["MS-ALT", "MS-ALT2"])
        uses = view["uses"]
        self.assertEqual([(u["level"], u["used_in"], u["quantity"], u["main_sob_percent"]) for u in uses], [
            ("L1", "Model MS-M1 · Model 1", 6, 100), ("L1", "Model MS-M2 · Model 2", 6, 100),
            ("L2", "MS-MOTOR · Motor BOM", 2, 100), ("L3", "MS-STATOR · Stator BOM", 4, 100)])
        lines = [{"model_bom_id": u["model_bom_id"], "component_id": u["component_id"]} for u in uses]

        def post(body):
            return self.client.post("/api/planning/alternate-sob", json={"main_item_id": steel, **body})
        self.assertIn("Tick at least one", post({"alternates": [], "lines": []}).json()["detail"])
        self.assertIn("adds up to 110%", post({"alternates": [{"alternate_item_id": alt, "sob_percent": 60}, {"alternate_item_id": alt2, "sob_percent": 50}], "lines": lines}).json()["detail"])
        other = self.client.post(f"/api/planning/model-bom/{models[1]}", json={"material_id": motor, "quantity": 1, "uom": "PCS"}).json()["id"]
        self.assertIn("does not use MS-STEEL", post({"alternates": [], "lines": [{"model_bom_id": other}]}).json()["detail"])
        saved = post({"alternates": [{"alternate_item_id": alt, "sob_percent": 30}, {"alternate_item_id": alt2, "sob_percent": 30}], "lines": lines})
        self.assertEqual((saved.status_code, saved.json()["lines"], saved.json()["main_sob_percent"]), (200, 4, 40))
        motor_line = [{"model_bom_id": None, "component_id": uses[2]["component_id"]}]
        self.assertEqual(post({"from_month": "Nov-2026", "alternates": [{"alternate_item_id": alt, "sob_percent": 80}], "lines": motor_line}).json()["lines"], 1)

        after = {u["used_in"]: u for u in self.client.get("/api/planning/alternate-codes/uses", params={"main_item_id": steel, "month": "2026-10"}).json()["uses"]}
        motor_use = after["MS-MOTOR · Motor BOM"]
        self.assertEqual((motor_use["sob"][str(alt)], motor_use["sob"][str(alt2)], motor_use["main_sob_percent"]), (30, 30, 40))
        self.assertEqual((motor_use["next_change"]["from_month"], motor_use["next_change"]["main_sob_percent"]), ("2026-11", 20))
        self.assertEqual([h["from_month"] for h in motor_use["history"]], ["", "2026-11"])
        self.assertIsNone(after["Model MS-M1 · Model 1"]["next_change"])
        november = {u["used_in"]: u["main_sob_percent"] for u in self.client.get("/api/planning/alternate-codes/uses", params={"main_item_id": steel, "month": "2026-11"}).json()["uses"]}
        self.assertEqual(november, {"Model MS-M1 · Model 1": 40, "Model MS-M2 · Model 2": 40, "MS-MOTOR · Motor BOM": 20, "MS-STATOR · Stator BOM": 40})
        # the BOM tree uses the same SOB
        l3 = self.client.get(f"/api/planning/part-bom/{stator}", params={"tab": "Level 3 Part BOM"}).json()["items"][0]
        self.assertEqual(l3["main_sob_percent"], 40)
        self.assertAlmostEqual(l3["main_quantity"], 1.6)

        # remove the Nov change from the Motor use; removing a month nobody has is refused
        self.assertEqual(post({"from_month": "2026-11", "action": "remove", "lines": motor_line}).json()["lines"], 1)
        self.assertEqual(post({"from_month": "2026-12", "action": "remove", "lines": lines}).status_code, 404)
        self.assertEqual(post({"action": "undo", "lines": lines}).status_code, 422)

        # change log: who, when, where, from month, split (newest first); Excel download
        log = self.client.get("/api/planning/sob-change-log").json()
        self.assertEqual(log["total"], 6)  # 4 uses set + Motor Nov set + Motor Nov removed
        newest = log["items"][0]
        self.assertEqual((newest["action"], newest["main_code"], newest["level"], newest["used_in"], newest["from_month"], newest["changed_by"]),
                         ("Removed", "MS-STEEL", "L2", "MS-MOTOR · Motor BOM", "2026-11", "Test Admin"))
        self.assertEqual(log["items"][1]["sob_text"], "MS-ALT 80%, MS-ALT2 0%, main 20%")
        self.assertEqual(log["items"][-1]["sob_text"], "MS-ALT 30%, MS-ALT2 30%, main 40%")
        self.assertEqual(self.client.get("/api/planning/sob-change-log", params={"search": "stator"}).json()["total"], 1)
        sheet = load_workbook(io.BytesIO(self.client.get("/api/planning/sob-change-log.xlsx").content)).active
        rows = [r for r in sheet.iter_rows(values_only=True) if r and r[0] and str(r[0])[:2] == "20"]
        header = next(r for r in sheet.iter_rows(values_only=True) if r and r[0] == "Changed on")
        self.assertEqual(list(header[:8]), ["Changed on", "Changed by", "Action", "Main code", "Level", "Used in", "SOB from month", "SOB split"])
        self.assertEqual(len(rows), 6)
        self.assertIn("Start (all months)", [r[6] for r in rows])
        self.assertEqual(self.client.get("/api/planning/sob-change-log.pdf").status_code, 200)

    def test_supplier_rows_entry_stock_entry_and_safe_deletes(self):
        def material(code):
            return self.client.post("/api/planning/materials", json={"code": code, "name": code.title(), "uom": "kg"}).json()["id"]
        steel, copper = material("SR-STEEL"), material("SR-COPPER")
        sup = lambda name: self.client.post("/api/suppliers", json={"name": name, "category": "Steel", "country": "India"}).json()["id"]
        tata, jsw, spare = sup("SR Tata"), sup("SR JSW"), sup("SR Spare")
        row = lambda m, s_, sob, price=0: {"material_id": m, "supplier_id": s_, "tier": "Direct", "sob_percent": sob, "unit_price": price, "currency": "INR"}
        # all rows or none: row 3 takes steel over 100% (60 + 40 + 10) -> nothing saved
        bad = self.client.post("/api/planning/supplier-parts/rows", json={"rows": [row(steel, tata, 60), row(steel, jsw, 40), row(steel, spare, 10)]})
        self.assertEqual(bad.status_code, 422)
        self.assertIn("Row 3: Supplier SOB cannot exceed 100%", bad.json()["detail"])
        self.assertEqual(self.client.get("/api/planning/supplier-parts", params={"search": "SR-"}).json()["total"], 0)
        dup = self.client.post("/api/planning/supplier-parts/rows", json={"rows": [row(steel, tata, 50), row(steel, tata, 50)]})
        self.assertIn("Row 2: SR-STEEL with this supplier, tier and parent already exists", dup.json()["detail"])
        self.assertEqual(self.client.post("/api/planning/supplier-parts/rows", json={"rows": []}).status_code, 422)
        ok = self.client.post("/api/planning/supplier-parts/rows", json={"rows": [row(steel, tata, 60, 70), row(steel, jsw, 40, 74), row(copper, tata, 100, 700)]})
        self.assertEqual((ok.status_code, ok.json()["rows"]), (201, 3), ok.text)
        sheet = load_workbook(io.BytesIO(self.client.get("/api/planning/tables/supplier-parts/xlsx").content)).active
        header = next(r for r in sheet.iter_rows(values_only=True) if r and r[0] == "Supplier code")
        self.assertEqual(list(header[:4]), ["Supplier code", "Supplier name", "Material code", "Material name"])
        tata_links = self.client.get("/api/planning/supplier-parts", params={"supplier_id": tata}).json()["items"]
        self.assertEqual(sorted(l["material_code"] for l in tata_links), ["SR-COPPER", "SR-STEEL"])
        steel_tata = next(l["id"] for l in tata_links if l["material_code"] == "SR-STEEL")
        copper_tata = next(l["id"] for l in tata_links if l["material_code"] == "SR-COPPER")

        # supplier stock typed in (no Excel): save, then update the same row; errors name the row
        url = "/api/planning/supplier-stock/2026-11"
        self.assertIn("Row 2: Available quantity", self.client.post(url, json={"rows": [{"supplier_part_id": steel_tata, "quantity": 10}, {"supplier_part_id": copper_tata, "quantity": 5, "available_quantity": 6}]}).json()["detail"])
        self.assertIn("Row 2: this supplier and material is entered twice", self.client.post(url, json={"rows": [{"supplier_part_id": steel_tata, "quantity": 1}, {"supplier_part_id": steel_tata, "quantity": 2}]}).json()["detail"])
        self.assertIn("Row 1: choose", self.client.post(url, json={"rows": [{"supplier_part_id": 999999, "quantity": 1}]}).json()["detail"])
        self.assertEqual(self.client.get(url).json()["items"], [])  # nothing saved by the failed tries
        saved = self.client.post(url, json={"rows": [{"supplier_part_id": steel_tata, "quantity": 500, "available_quantity": 400, "arrival_month": "Dec-2026"}, {"supplier_part_id": copper_tata, "quantity": 20}]})
        self.assertEqual((saved.status_code, saved.json()["rows"]), (201, 2), saved.text)
        first = {r["material_code"]: r["arrival_month"] for r in self.client.get(url).json()["items"]}
        self.assertEqual(first["SR-STEEL"], "2026-12")  # "Dec-2026" stored as YYYY-MM (planning compares months)
        self.assertEqual(self.client.post(url, json={"rows": [{"supplier_part_id": steel_tata, "quantity": 450, "available_quantity": 450}]}).status_code, 201)
        stock = {r["material_code"]: (r["quantity"], r["available_quantity"], r["arrival_month"]) for r in self.client.get(url).json()["items"]}
        self.assertEqual(stock, {"SR-STEEL": (450, 450, None), "SR-COPPER": (20, 20, None)})

        # deletes are refused while in use, with the reason
        blocked = self.client.delete(f"/api/planning/supplier-parts/{steel_tata}")
        self.assertEqual(blocked.status_code, 409)
        self.assertIn("1 supplier stock row", blocked.json()["detail"])
        jsw_link = next(l["id"] for l in self.client.get("/api/planning/supplier-parts", params={"supplier_id": jsw}).json()["items"])
        self.assertEqual(self.client.delete(f"/api/planning/supplier-parts/{jsw_link}").status_code, 204)
        self.assertEqual(self.client.delete(f"/api/planning/supplier-parts/{jsw_link}").status_code, 404)
        used = self.client.delete(f"/api/suppliers/{tata}")
        self.assertEqual(used.status_code, 409)
        self.assertIn("2 supplier/material relationships", used.json()["detail"])
        self.assertEqual(self.client.delete(f"/api/suppliers/{jsw}").status_code, 204)  # its only link was deleted
        self.assertEqual(self.client.delete(f"/api/suppliers/{spare}").status_code, 204)
        self.assertNotIn(spare, [s_["id"] for s_ in self.client.get("/api/suppliers").json()])
        # supplier edit (Supplier register)
        edited = self.client.patch(f"/api/suppliers/{tata}", json={"name": "SR Tata Steel", "country": "India", "status": "On hold"})
        self.assertEqual((edited.json()["name"], edited.json()["status"]), ("SR Tata Steel", "On hold"))

    def test_sample_suppliers_can_be_deleted_and_do_not_come_back(self):
        # The install sample (Apex etc.) has old RFQ quotes, purchase orders and preferred-supplier
        # links from retired screens; those no longer block a delete and go with the supplier.
        import sqlite3
        db_file = Path(self.temp_dir.name) / "test.db"
        apex = next(s_ for s_ in self.client.get("/api/suppliers").json() if s_["name"] == "Apex Precision Works")
        old = lambda q: sqlite3.connect(db_file).execute(q, (apex["id"],)).fetchone()[0]
        self.assertGreater(old("SELECT COUNT(*) FROM purchase_orders WHERE supplier_id = ?") + old("SELECT COUNT(*) FROM quotes WHERE supplier_id = ?"), 0)
        self.assertGreater(old("SELECT COUNT(*) FROM items WHERE preferred_supplier_id = ?"), 0)
        # still blocked by real use: a supplier/material relationship
        material = self.client.post("/api/planning/materials", json={"code": "SS-STEEL", "name": "Steel", "uom": "kg"}).json()["id"]
        link = self.client.post("/api/planning/supplier-parts", json={"material_id": material, "supplier_id": apex["id"], "sob_percent": 100, "currency": "INR"}).json()["id"]
        blocked = self.client.delete(f"/api/suppliers/{apex['id']}")
        self.assertEqual(blocked.status_code, 409)
        self.assertEqual(blocked.json()["detail"], "Apex Precision Works is used: 1 supplier/material relationship. Remove those first, or set its status to On hold")
        self.assertEqual(self.client.delete(f"/api/planning/supplier-parts/{link}").status_code, 204)
        self.assertEqual(self.client.delete(f"/api/suppliers/{apex['id']}").status_code, 204)
        for query in ("SELECT COUNT(*) FROM purchase_orders WHERE supplier_id = ?", "SELECT COUNT(*) FROM quotes WHERE supplier_id = ?",
                      "SELECT COUNT(*) FROM stock_purchase_orders WHERE supplier_id = ?", "SELECT COUNT(*) FROM items WHERE preferred_supplier_id = ?",
                      "SELECT COUNT(*) FROM rfqs WHERE awarded_supplier_id = ?"):
            self.assertEqual(old(query), 0, query)
        self.assertEqual(old("SELECT COUNT(*) FROM items WHERE code = 'MAT-101' AND ? IS NOT NULL"), 1)  # the material itself stays
        # delete every other sample supplier, restart the app on the same database: none come back
        for supplier in self.client.get("/api/suppliers").json():
            self.assertEqual(self.client.delete(f"/api/suppliers/{supplier['id']}").status_code, 204, supplier["name"])
        restarted = TestClient(create_app(f"sqlite:///{db_file}"))
        restarted.__enter__()
        restarted.post("/api/auth/login", json={"username": "admin", "password": "Secure-test-password-2026!"})
        self.assertEqual(restarted.get("/api/suppliers").json(), [])
        restarted.__exit__(None, None, None)

    def test_one_category_list_everywhere(self):
        import sqlite3
        db_file = Path(self.temp_dir.name) / "test.db"
        listed = lambda: self.client.get("/api/planning/categories").json()
        names = [c["name"] for c in listed()["items"]]
        self.assertEqual(sorted(names, key=str.casefold), sorted([
            "Raw material", "Finished assembly", "Electrical components", "Electronic PCB Assy", "Sheet Metal", "Wire Harness",
            "Plastic Raw material", "Steel Raw material", "Electronic Child Part", "Casting", "Machining", "Packing Material",
            "Fasteners", "Subassembly"], key=str.casefold))
        self.assertEqual(self.client.get("/api/planning/material-categories").json()["items"], sorted(names, key=str.casefold))
        # old sample categories replaced as agreed
        suppliers = {s_["name"]: s_["category"] for s_ in self.client.get("/api/suppliers").json()}
        self.assertEqual(suppliers["Apex Precision Works"], "Machining")
        self.assertEqual(suppliers["Northstar Components"], "Sheet Metal")
        self.assertEqual(suppliers["Metro Polymer Group"], "Plastic Raw material")
        self.assertEqual(suppliers["Blue Ridge Fasteners"], "Fasteners")
        items = {i["code"]: i["category"] for i in self.client.get("/api/planning/materials?limit=500").json()["items"]}
        self.assertEqual((items["MAT-101"], items["MAT-102"], items["SUB-001"]), ("Machining", "Plastic Raw material", "Subassembly"))
        self.assertEqual(listed()["unlisted"], [])
        machining = next(c for c in listed()["items"] if c["name"] == "Machining")
        self.assertEqual((machining["parts"], machining["suppliers"]), (1, 1))

        # saves use the list's spelling; a new name is added to the list
        self.assertEqual(self.client.post("/api/planning/materials", json={"code": "CT-STEEL", "name": "Steel", "uom": "kg", "category": "steel raw MATERIAL"}).status_code, 201)
        self.assertEqual(self.client.post("/api/planning/materials", json={"code": "CT-BRG", "name": "Bearing", "uom": "PCS", "category": "Bearings"}).status_code, 201)
        saved = {i["code"]: i["category"] for i in self.client.get("/api/planning/materials", params={"search": "CT-"}).json()["items"]}
        self.assertEqual(saved, {"CT-STEEL": "Steel Raw material", "CT-BRG": "Bearings"})
        self.assertIn("Bearings", self.client.get("/api/planning/material-categories").json()["items"])
        caster = self.client.post("/api/suppliers", json={"name": "CT Foundry", "category": "casting", "country": "India"})
        self.assertEqual(caster.json()["category"], "Casting")
        workbook = Workbook()
        ws = workbook.active
        ws.append(["Code", "Name", "Category"])
        ws.append(["CT-BOX", "Carton", "packing material"])
        ws.append(["CT-GSK", "Gasket", "Gaskets"])
        buffer = io.BytesIO()
        workbook.save(buffer)
        self.assertEqual(self.client.post("/api/planning/tier-materials/upload?tier=T2", content=buffer.getvalue()).status_code, 201)
        tier = {m["code"]: m["category"] for m in self.client.get("/api/planning/tier-materials", params={"tier": "T2"}).json()["items"]}
        self.assertEqual(tier, {"CT-BOX": "Packing Material", "CT-GSK": "Gaskets"})
        self.assertIn("Gaskets", self.client.get("/api/planning/material-categories").json()["items"])

        # supplier code after deleting a sample supplier: next free code, not a duplicate
        apex = next(s_ for s_ in self.client.get("/api/suppliers").json() if s_["name"] == "Apex Precision Works")
        self.assertEqual(self.client.delete(f"/api/suppliers/{apex['id']}").status_code, 204)
        fresh = self.client.post("/api/suppliers", json={"name": "CT Harness Co", "category": "Wire Harness", "country": "India"})
        self.assertEqual(fresh.status_code, 201, fresh.text)
        codes = [s_["code"] for s_ in self.client.get("/api/suppliers").json()]
        self.assertEqual(len(codes), len(set(codes)))

        # a category used by your own old records but not in the list: shown, then replaced
        connection = sqlite3.connect(db_file)
        connection.execute("UPDATE items SET category = 'Electrical' WHERE code = 'MAT-104'")
        connection.commit()
        connection.close()
        unlisted = listed()["unlisted"]
        self.assertEqual([(u["name"], u["parts"]) for u in unlisted], [("Electrical", 1)])
        self.assertEqual(self.client.post("/api/planning/categories/replace", json={"old": "Electrical", "new": "Banana"}).status_code, 422)
        replaced = self.client.post("/api/planning/categories/replace", json={"old": "Electrical", "new": "electrical components"}).json()
        self.assertEqual((replaced["new"], replaced["records"]), ("Electrical components", 1))
        self.assertEqual(listed()["unlisted"], [])
        # add, rename (records follow), delete (refused while used)
        self.assertEqual(self.client.post("/api/planning/categories", json={"name": "fasteners"}).status_code, 409)
        self.assertEqual(self.client.post("/api/planning/categories", json={"name": "Rubber Parts"}).status_code, 201)
        fasteners = next(c for c in listed()["items"] if c["name"] == "Fasteners")
        self.assertEqual(self.client.patch(f"/api/planning/categories/{fasteners['id']}", json={"name": "Casting"}).status_code, 409)
        renamed = self.client.patch(f"/api/planning/categories/{fasteners['id']}", json={"name": "Fastener Parts"}).json()
        self.assertEqual(renamed["records"], 2)  # MAT-103 and Blue Ridge
        self.assertEqual(next(s_ for s_ in self.client.get("/api/suppliers").json() if s_["name"] == "Blue Ridge Fasteners")["category"], "Fastener Parts")
        busy = self.client.delete(f"/api/planning/categories/{fasteners['id']}")
        self.assertEqual(busy.status_code, 409)
        self.assertIn("1 parts, 1 suppliers", busy.json()["detail"])
        rubber = next(c for c in listed()["items"] if c["name"] == "Rubber Parts")
        self.assertEqual(self.client.delete(f"/api/planning/categories/{rubber['id']}").status_code, 204)
        # replacing a listed category moves its records and removes it from the list
        bearings = next(c for c in listed()["items"] if c["name"] == "Bearings")
        self.assertEqual(self.client.post("/api/planning/categories/replace", json={"old": "Bearings", "new": "Machining"}).json()["records"], 1)
        self.assertNotIn(bearings["id"], [c["id"] for c in listed()["items"]])

        # the one-time replacement does not run again: a re-added "Plastics" stays after a restart
        self.assertEqual(self.client.post("/api/planning/categories", json={"name": "Plastics"}).status_code, 201)
        metro = next(s_ for s_ in self.client.get("/api/suppliers").json() if s_["name"] == "Metro Polymer Group")
        self.assertEqual(self.client.patch(f"/api/suppliers/{metro['id']}", json={"category": "plastics"}).json()["category"], "Plastics")
        restarted = TestClient(create_app(f"sqlite:///{db_file}"))
        restarted.__enter__()
        restarted.post("/api/auth/login", json={"username": "admin", "password": "Secure-test-password-2026!"})
        self.assertEqual(next(s_ for s_ in restarted.get("/api/suppliers").json() if s_["name"] == "Metro Polymer Group")["category"], "Plastics")
        restarted.__exit__(None, None, None)

    def test_supplier_code_edit_and_bulk_supplier_upload(self):
        suppliers = lambda: {s_["code"]: s_ for s_ in self.client.get("/api/suppliers").json()}
        apex = suppliers()["SUP-1001"]
        # change a supplier code; a code already used is refused
        self.assertIn("already used by Northstar Components", self.client.patch(f"/api/suppliers/{apex['id']}", json={"code": "sup-1002"}).json()["detail"])
        self.assertIn("already used by Northstar Components", self.client.patch(f"/api/suppliers/{apex['id']}", json={"code": " sup-1002 "}).json()["detail"])
        self.assertEqual(self.client.patch(f"/api/suppliers/{apex['id']}", json={"code": "  APX-01 "}).json()["code"], "APX-01")
        self.assertIn("APX-01", suppliers())
        # Add supplier with your own code; blank code = next free SUP- code; a used code is refused
        own = self.client.post("/api/suppliers", json={"code": " ILN-02 ", "name": "ILIN Two", "category": "Casting", "country": "India"})
        self.assertEqual((own.status_code, own.json()["code"]), (201, "ILN-02"))
        self.assertIn("already used by ILIN Two", self.client.post("/api/suppliers", json={"code": "iln-02", "name": "Other", "category": "Casting", "country": "India"}).json()["detail"])
        self.assertTrue(self.client.post("/api/suppliers", json={"name": "No Code Co", "category": "Casting", "country": "India"}).json()["code"].startswith("SUP-"))

        def sheet(rows, heading=("Vendor Code", "Supplier Name", "Category", "Country", "Contact Person", "E-mail", "Mobile", "Status")):
            workbook = Workbook()
            ws = workbook.active
            ws.append(["Supplier list 2026"])
            ws.append(list(heading))
            for row in rows:
                ws.append(list(row))
            buffer = io.BytesIO()
            workbook.save(buffer)
            return buffer.getvalue()
        upload = lambda rows, **kw: self.client.post("/api/suppliers/upload", content=sheet(rows, **kw))
        # all rows or none: row 5 has a wrong approval
        bad = upload([("ILN-01", "ILIN Electronic", "Electronic PCB Assy", "India", "Amit", "", "", "Approved"),
                      ("", "Tata Steel", "steel raw material", "India", "", "", "", ""),
                      ("", "Bad Co", "", "", "", "", "", "Maybe")])
        self.assertEqual(bad.status_code, 422)
        self.assertIn("Row 5: approval must be", bad.json()["detail"])
        self.assertNotIn("ILN-01", suppliers())
        self.assertIn("Row 4: supplier tata steel is repeated", upload([("", "Tata Steel", "", "", "", "", "", ""), ("", "tata steel", "", "", "", "", "", "")]).json()["detail"])
        self.assertIn("Row 3: supplier name is required", upload([("X-1", "", "Casting", "", "", "", "", "")]).json()["detail"])
        self.assertIn("already in the register as SUP-1003", upload([("NEW-9", "Koyo Industrial Supply", "", "", "", "", "", "")]).json()["detail"])

        ok = upload([("ILN-01", "ILIN Electronic", "Electronic PCB Assy", "India", "Amit", "amit@ilin.example", "+91 99", "Approved"),
                     ("", "Tata Steel", "steel raw material", "India", "", "", "", ""),
                     ("", "Gear Castings", "Gear Casting", "India", "", "", "", "pending"),
                     ("SUP-1003", "Koyo Industrial Supply", "Electronic Child Part", "", "", "", "", "On hold"),
                     ("", "Northstar Components", "", "", "Jordan L.", "", "", "")])
        self.assertEqual(ok.status_code, 201, ok.text)
        self.assertEqual((ok.json()["added"], ok.json()["updated"]), (3, 2))
        now = suppliers()
        self.assertEqual((now["ILN-01"]["name"], now["ILN-01"]["category"], now["ILN-01"]["email"], now["ILN-01"]["status"]),
                         ("ILIN Electronic", "Electronic PCB Assy", "amit@ilin.example", "Approved"))
        auto = {s_["name"]: s_ for s_ in now.values()}
        self.assertEqual((auto["Tata Steel"]["category"], auto["Tata Steel"]["status"]), ("Steel Raw material", "Approved"))
        self.assertEqual(auto["Gear Castings"]["status"], "Pending review")
        self.assertIn("Gear Casting", self.client.get("/api/planning/material-categories").json()["items"])  # new category added
        self.assertTrue(auto["Tata Steel"]["code"].startswith("SUP-") and auto["Tata Steel"]["code"] != auto["Gear Castings"]["code"])
        self.assertEqual((now["SUP-1003"]["category"], now["SUP-1003"]["status"], now["SUP-1003"]["country"]), ("Electronic Child Part", "On hold", "Japan"))
        self.assertEqual(now["SUP-1002"]["contact_name"], "Jordan L.")  # matched by name, blank cells keep old values
        self.assertEqual(len(now), len({s_["code"].lower() for s_ in now.values()}))
        template = load_workbook(io.BytesIO(self.client.get("/api/planning/templates/suppliers.xlsx").content)).active
        self.assertEqual([c.value for c in template[1]], ["supplier_code", "supplier_name", "supplier_tier", "category", "country", "contact_person", "email", "phone", "approval"])
        # the template itself uploads
        self.assertEqual(upload([("TPL-1", "Template Supplier", "T1, T2", "Casting", "India", "A", "a@b.example", "1", "Approved")],
                                heading=("supplier_code", "supplier_name", "supplier_tier", "category", "country", "contact_person", "email", "phone", "approval")).status_code, 201)
        self.assertEqual(suppliers()["TPL-1"]["tiers"], ["T1", "T2"])

    def test_t1_t2_t3_tiers_t2_codes_full_row_edit_and_supplier_tiers(self):
        api = self.client
        def material(code, name="Steel"):
            return api.post("/api/planning/materials", json={"code": code, "name": name, "uom": "kg", "category": "Steel Raw material"}).json()["id"]
        steel = material("TR-STEEL")
        chip = api.post("/api/planning/tier-materials", json={"tier": "T2", "code": "TR-CHIP", "name": "MicroChip", "uom": "PCS", "unit_price": 480, "currency": "INR"}).json()["id"]
        sup = lambda name, tiers=None: api.post("/api/suppliers", json={"name": name, "category": "Casting", "country": "India", **({"tiers": tiers} if tiers else {})}).json()
        tata, supangita, ilin, samicon = sup("TR TATA Steel"), sup("TR Supangita"), sup("TR ILIN"), sup("TR Samicon", ["T2"])
        self.assertEqual((tata["tiers"], samicon["tiers"]), (["T1"], ["T2"]))
        self.assertEqual(api.post("/api/suppliers", json={"name": "TR Bad", "category": "Casting", "country": "India", "tiers": ["T9"]}).status_code, 422)
        def link(body):
            return api.post("/api/planning/supplier-parts", json={"sob_percent": 100, "currency": "INR", **body})
        # TATA sells steel to you (T1) and to Supangita (T2, parent Supangita); "Direct" still means T1
        t1 = link({"material_id": steel, "supplier_id": tata["id"], "tier": "Direct", "sob_percent": 60, "unit_price": 70})
        self.assertEqual((t1.status_code, t1.json()["tier"]), (201, "T1"), t1.text)
        self.assertEqual(link({"material_id": steel, "supplier_id": supangita["id"], "tier": "T1", "sob_percent": 40, "unit_price": 74}).status_code, 201)
        t2 = link({"material_id": steel, "supplier_id": tata["id"], "tier": "T2", "parent_supplier_id": supangita["id"], "unit_price": 72})
        self.assertEqual(t2.status_code, 201, t2.text)
        self.assertIn("must have its parent supplier", link({"material_id": steel, "supplier_id": tata["id"], "tier": "T3"}).json()["detail"])
        self.assertIn("Tier must be T1", link({"material_id": steel, "supplier_id": tata["id"], "tier": "X1"}).json()["detail"])
        # a T2 Material Master code: only for T2/T3 rows (Samicon supplies the chip to ILIN)
        self.assertIn("a T1 supplier needs a T1 code", link({"tier_material_id": chip, "supplier_id": samicon["id"], "tier": "T1"}).json()["detail"])
        self.assertIn("Choose one part code", link({"material_id": steel, "tier_material_id": chip, "supplier_id": samicon["id"], "tier": "T2", "parent_supplier_id": ilin["id"]}).json()["detail"])
        chip_link = link({"tier_material_id": chip, "supplier_id": samicon["id"], "tier": "T2", "parent_supplier_id": ilin["id"], "unit_price": 480})
        self.assertEqual(chip_link.status_code, 201, chip_link.text)
        only_t3 = sup("TR Only T3", ["T3"])
        extra = sup("TR Extra")
        self.assertEqual(link({"material_id": steel, "supplier_id": extra["id"], "tier": "T2", "parent_supplier_id": only_t3["id"], "sob_percent": 0}).status_code, 201)
        self.assertEqual(next(s_ for s_ in api.get("/api/suppliers").json() if s_["name"] == "TR Only T3")["tiers"], ["T1", "T3"])  # parent of a T2 row is T1
        self.assertEqual(api.patch(f"/api/planning/supplier-parts/{chip_link.json()['id']}", json={"lead_time_days": 60}).status_code, 200)  # 100% alone in its route
        rows = {(r["supplier_name"], r["tier"]): r for r in api.get("/api/planning/supplier-parts", params={"search": "TR-"}).json()["items"]}
        self.assertEqual((rows[("TR Samicon", "T2")]["material_code"], rows[("TR Samicon", "T2")]["code_master"], rows[("TR Samicon", "T2")]["parent_supplier_name"]), ("TR-CHIP", "T2", "TR ILIN"))
        self.assertEqual(rows[("TR TATA Steel", "T1")]["code_master"], "T1")
        # suppliers' tiers follow their relationships: TATA T1+T2, ILIN T1 (parent of a T2 row), Samicon T2
        tiers = {s_["name"]: s_["tiers"] for s_ in api.get("/api/suppliers").json()}
        self.assertEqual((tiers["TR TATA Steel"], tiers["TR Supangita"], tiers["TR ILIN"], tiers["TR Samicon"]), (["T1", "T2"], ["T1"], ["T1"], ["T2"]))
        # your BOM price uses T1 rows only: 0.6 x 70 + 0.4 x 74 = 71.6 (TATA's T2 price 72 not used)
        model = api.post("/api/planning/models", json={"code": "TR-M", "name": "Tier model"}).json()["id"]
        api.post(f"/api/planning/model-bom/{model}", json={"material_id": steel, "quantity": 1, "uom": "kg"})
        self.assertAlmostEqual(next(r for r in api.get(f"/api/planning/model-bom/{model}").json()["items"])["unit_price"], 71.6)
        # POs only for T1 Material Master codes; a T2 code in use cannot be deleted
        po = api.post("/api/planning/purchase-orders", json={"supplier_part_id": chip_link.json()["id"], "quantity_ordered": 10, "expected_month": "2026-12", "workspace": "trial"})
        self.assertIn("T1 Material Master codes", po.json()["detail"])
        self.assertIn("supplier/material relationship", api.delete(f"/api/planning/tier-materials/{chip}").json()["detail"])

        # full-row edit: the T2 row by mistake -> make it T3 under Samicon, then back; supplier and part can change
        t2_id = t2.json()["id"]
        edited = api.patch(f"/api/planning/supplier-parts/{t2_id}", json={"tier": "T3", "parent_supplier_id": samicon["id"]})
        self.assertEqual((edited.status_code, edited.json()["tier"]), (200, "T3"), edited.text)
        self.assertIn("must have its parent supplier", api.patch(f"/api/planning/supplier-parts/{t2_id}", json={"tier": "T2", "parent_supplier_id": None}).json()["detail"])
        self.assertIn("SOB cannot exceed 100%", api.patch(f"/api/planning/supplier-parts/{t2_id}", json={"tier": "T1"}).json()["detail"])  # T1 steel is already 60 + 40
        self.assertIn("already exists", api.patch(f"/api/planning/supplier-parts/{t2_id}", json={"tier": "T1", "sob_percent": 0}).json()["detail"])  # TATA is already T1 for steel
        moved = api.patch(f"/api/planning/supplier-parts/{t2_id}", json={"tier": "T2", "parent_supplier_id": ilin["id"], "material_id": None, "tier_material_id": chip, "sob_percent": 0})
        self.assertEqual(moved.status_code, 200, moved.text)
        row = next(r for r in api.get("/api/planning/supplier-parts", params={"search": "TR-CHIP"}).json()["items"] if r["id"] == t2_id)
        self.assertEqual((row["material_code"], row["tier"], row["parent_supplier_name"], row["sob_percent"]), ("TR-CHIP", "T2", "TR ILIN", 0))
        self.assertIn("SOB cannot exceed 100%", api.patch(f"/api/planning/supplier-parts/{t2_id}", json={"sob_percent": 10}).json()["detail"])
        self.assertIn("own parent", api.patch(f"/api/planning/supplier-parts/{t2_id}", json={"parent_supplier_id": tata["id"]}).json()["detail"])
        self.assertEqual(api.patch(f"/api/planning/supplier-parts/{t2_id}", json={"supplier_id": supangita["id"]}).json()["tier"], "T2")
        self.assertIn("T2", next(s_ for s_ in api.get("/api/suppliers").json() if s_["name"] == "TR Supangita")["tiers"])

        # Excel upload: T2 code with "Tier 2", parent; supplier list upload in the T2 tab
        def sheet(heading, data):
            workbook = Workbook()
            ws = workbook.active
            ws.append(heading)
            for r in data:
                ws.append(r)
            buffer = io.BytesIO()
            workbook.save(buffer)
            return buffer.getvalue()
        up = api.post("/api/planning/supplier-parts/upload", content=sheet(["Material code", "Supplier name", "Tier", "Parent supplier", "SOB %"], [["TR-CHIP", "TR New Chip Co", "Tier 2", "TR Supangita", 50]]))
        self.assertEqual(up.status_code, 201, up.text)
        self.assertEqual(next(s_ for s_ in api.get("/api/suppliers").json() if s_["name"] == "TR New Chip Co")["tiers"], ["T2"])
        bad = api.post("/api/planning/supplier-parts/upload", content=sheet(["Material code", "Supplier name", "Tier"], [["TR-CHIP", "TR ILIN", "T1"]]))
        self.assertIn("a T1 supplier needs a T1 code", bad.json()["detail"])
        listed = api.post("/api/suppliers/upload?tier=T2", content=sheet(["Supplier name", "Category"], [["TR Tab Two Co", "Casting"], ["TR TATA Steel", ""]]))
        self.assertEqual(listed.status_code, 201, listed.text)
        tiers = {s_["name"]: s_["tiers"] for s_ in api.get("/api/suppliers").json()}
        # TATA was made T3 above (tiers are added automatically, never removed; untick in Edit)
        self.assertEqual((tiers["TR Tab Two Co"], tiers["TR TATA Steel"]), (["T2"], ["T1", "T2", "T3"]))
        self.assertEqual(api.patch(f"/api/suppliers/{samicon['id']}", json={"tiers": ["T2", "T3"]}).json()["tiers"], ["T2", "T3"])
        self.assertEqual(api.patch(f"/api/suppliers/{samicon['id']}", json={"tiers": []}).status_code, 422)

    def test_older_database_tiers_renamed_table_updated_and_backed_up(self):
        # A database from before: supplier_parts without tier_material_id (material required),
        # tiers "Direct" and "T1" (= supplier of the direct supplier), one stock row on the T1 link.
        import sqlite3
        db_path = Path(self.temp_dir.name) / "tiers.db"
        first = TestClient(create_app(f"sqlite:///{db_path}"))
        first.__enter__()
        first.__exit__(None, None, None)
        connection = sqlite3.connect(db_path)
        connection.execute("DELETE FROM app_settings")
        connection.execute("DROP TABLE supplier_parts")
        connection.execute("""CREATE TABLE supplier_parts (id INTEGER PRIMARY KEY, material_id INTEGER NOT NULL, supplier_id INTEGER NOT NULL,
            tier VARCHAR(24) NOT NULL, parent_supplier_id INTEGER, sob_percent FLOAT NOT NULL, lead_time_days INTEGER NOT NULL, moq FLOAT NOT NULL,
            buffer_days INTEGER NOT NULL, unit_price FLOAT NOT NULL, currency VARCHAR(3) NOT NULL, price_mode VARCHAR(24) NOT NULL,
            price_interval_months INTEGER NOT NULL, price_updated_month VARCHAR(7) NOT NULL, UNIQUE (material_id, supplier_id, tier, parent_supplier_id))""")
        connection.execute("CREATE INDEX ix_supplier_parts_material_id ON supplier_parts (material_id)")
        item = connection.execute("SELECT id FROM items WHERE code = 'MAT-104'").fetchone()[0]
        direct, upstream, deeper = [row[0] for row in connection.execute("SELECT id FROM suppliers ORDER BY id LIMIT 3")]
        connection.execute("UPDATE suppliers SET tiers = ''")
        for row in ((10, item, direct, "Direct", None), (11, item, upstream, "T1", direct), (12, item, deeper, "T2", upstream)):
            connection.execute("INSERT INTO supplier_parts VALUES (?, ?, ?, ?, ?, 100, 10, 0, 5, 1.5, 'INR', 'spot', 1, '2026-10')", row)
        connection.execute("INSERT INTO supplier_stock_snapshots (workspace, report_month, supplier_part_id, quantity, available_quantity) VALUES ('trial', '2026-10', 11, 25, 20)")
        connection.commit()
        connection.close()
        client = TestClient(create_app(f"sqlite:///{db_path}"))
        client.__enter__()
        self.assertEqual(client.post("/api/auth/setup", json={"username": "admin", "full_name": "Test Admin", "password": "Secure-test-password-2026!"}).status_code, 201)
        check = sqlite3.connect(db_path)
        self.assertEqual(check.execute("SELECT id, tier, parent_supplier_id FROM supplier_parts ORDER BY id").fetchall(), [(10, "T1", None), (11, "T2", direct), (12, "T3", upstream)])
        columns = {row[1]: row[3] for row in check.execute("PRAGMA table_info(supplier_parts)")}
        self.assertIn("tier_material_id", columns)
        self.assertEqual(columns["material_id"], 0)  # may be empty now (T2/T3 codes)
        self.assertEqual(check.execute("SELECT supplier_part_id, quantity FROM supplier_stock_snapshots").fetchall(), [(11, 25.0)])
        self.assertEqual(sorted(k for (k,) in check.execute("SELECT key FROM app_settings")), ["supplier_tiers_set", "tiers_t1_is_direct"])
        tiers = dict(check.execute("SELECT id, tiers FROM suppliers WHERE id IN (?, ?, ?)", (direct, upstream, deeper)).fetchall())
        self.assertEqual((tiers[direct], tiers[upstream], tiers[deeper]), ("T1", "T2", "T3"))
        self.assertTrue(list(Path(self.temp_dir.name).glob("tiers.db.before-tier-update-*.bak")))
        listed = client.get("/api/planning/supplier-parts").json()["items"]
        self.assertEqual(sorted(r["tier"] for r in listed), ["T1", "T2", "T3"])
        client.__exit__(None, None, None)
        # second start: nothing renamed again (T2 stays T2)
        again = TestClient(create_app(f"sqlite:///{db_path}"))
        again.__enter__()
        self.assertEqual(sqlite3.connect(db_path).execute("SELECT tier FROM supplier_parts ORDER BY id").fetchall(), [("T1",), ("T2",), ("T3",)])
        again.__exit__(None, None, None)

    def test_t1_material_master_used_in_levels(self):
        # Model UI-M: L1 = MOTOR, STEEL, WIRE, GASKET; MOTOR (L2) = STATOR, STEEL, GASKET, BOLT;
        # STATOR (L3) = STEEL, COPPER, WIRE, BOLT. SPARE is in no BOM.
        def material(code):
            return self.client.post("/api/planning/materials", json={"code": f"UI-{code}", "name": code.title(), "uom": "PCS"}).json()["id"]
        ids = {code: material(code) for code in ("MOTOR", "STEEL", "WIRE", "GASKET", "STATOR", "BOLT", "COPPER", "SPARE")}
        model = self.client.post("/api/planning/models", json={"code": "UI-M", "name": "Used-in model"}).json()["id"]
        for code in ("MOTOR", "STEEL", "WIRE", "GASKET"):
            self.assertEqual(self.client.post(f"/api/planning/model-bom/{model}", json={"material_id": ids[code], "quantity": 1, "uom": "PCS"}).status_code, 201)
        lines = lambda codes: [{"child_item_id": ids[c], "quantity": 1} for c in codes]
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{ids['MOTOR']}/new", json={"tab": "Level 2 Part BOM", "lines": lines(["STATOR", "STEEL", "GASKET", "BOLT"])}).status_code, 201)
        self.assertEqual(self.client.post(f"/api/planning/part-bom/{ids['STATOR']}/new", json={"tab": "Level 3 Part BOM", "lines": lines(["STEEL", "COPPER", "WIRE", "BOLT"])}).status_code, 201)
        used = {m["code"]: m["used_in"] for m in self.client.get("/api/planning/materials", params={"search": "UI-", "limit": 50}).json()["items"]}
        self.assertEqual(used, {"UI-MOTOR": "L1", "UI-STATOR": "L2", "UI-COPPER": "L3", "UI-GASKET": "L1 & L2", "UI-BOLT": "L2 & L3",
                                "UI-WIRE": "L1 & L3", "UI-STEEL": "L1 & L2 & L3", "UI-SPARE": ""})
        movable = {m["code"]: m["movable"] for m in self.client.get("/api/planning/materials", params={"search": "UI-", "limit": 50}).json()["items"]}
        self.assertEqual([code for code, ok in movable.items() if ok], ["UI-SPARE"])  # only the code in no BOM can move
        # alternates: IMPORT 20% on STATOR's STEEL line (L3) -> "L3 (alternate)"; NEWIMP 0% -> not used;
        # BOLT is a line at L2 & L3 and 30% alternate of WIRE on the model's L1 line -> "L1 & L2 & L3 (alternate)"
        imp, newimp = material("IMPORT"), material("NEWIMP")
        for main, alternate in ((ids["STEEL"], imp), (ids["STEEL"], newimp), (ids["WIRE"], ids["BOLT"])):
            self.assertEqual(self.client.post("/api/planning/alternate-codes", json={"main_item_id": main, "alternate_item_id": alternate}).status_code, 201)
        uses = self.client.get("/api/planning/alternate-codes/uses", params={"main_item_id": ids["STEEL"]}).json()["uses"]
        l3_line = next(u for u in uses if u["level"] == "L3")
        self.assertEqual(self.client.post("/api/planning/alternate-sob", json={"main_item_id": ids["STEEL"], "lines": [{"component_id": l3_line["component_id"]}],
                                                                             "alternates": [{"alternate_item_id": imp, "sob_percent": 20}, {"alternate_item_id": newimp, "sob_percent": 0}]}).status_code, 200)
        wire_l1 = next(u for u in self.client.get("/api/planning/alternate-codes/uses", params={"main_item_id": ids["WIRE"]}).json()["uses"] if u["level"] == "L1")
        self.assertEqual(self.client.post("/api/planning/alternate-sob", json={"main_item_id": ids["WIRE"], "lines": [{"model_bom_id": wire_l1["model_bom_id"]}],
                                                                             "alternates": [{"alternate_item_id": ids["BOLT"], "sob_percent": 30}]}).status_code, 200)
        used = {m["code"]: m["used_in"] for m in self.client.get("/api/planning/materials", params={"search": "UI-", "limit": 50}).json()["items"]}
        self.assertEqual((used["UI-IMPORT"], used["UI-NEWIMP"], used["UI-BOLT"], used["UI-STEEL"]), ("L3 (alternate)", "", "L1 & L2 & L3 (alternate)", "L1 & L2 & L3"))
        movable = {m["code"]: m["movable"] for m in self.client.get("/api/planning/materials", params={"search": "UI-", "limit": 50}).json()["items"]}
        self.assertEqual((movable["UI-IMPORT"], movable["UI-NEWIMP"], movable["UI-SPARE"]), (False, False, True))  # alternates (even at 0%) stay in T1
        sheet = load_workbook(io.BytesIO(self.client.get("/api/planning/tables/materials/xlsx").content)).active
        header = next(r for r in sheet.iter_rows(values_only=True) if r and r[0] == "Material code")
        self.assertEqual(list(header[:3]), ["Material code", "Material name", "Used in"])
        values = {r[0]: r[2] for r in sheet.iter_rows(values_only=True) if r and str(r[0]).startswith("UI-")}
        self.assertEqual((values["UI-SPARE"], values["UI-STEEL"], values["UI-IMPORT"], values["UI-NEWIMP"]), ("Not used", "L1 & L2 & L3", "L3 (alternate)", "Not used"))

    def test_t1_code_put_by_mistake_moves_to_t2_and_delete_says_what_uses_it(self):
        api = self.client
        adc = api.post("/api/planning/materials", json={"code": "MV-ADC12", "name": "ADC12", "uom": "kg", "unit_cost": 350, "currency": "INR", "category": "Aluminimum"}).json()["id"]
        sup = lambda name: api.post("/api/suppliers", json={"name": name, "category": "Casting", "country": "India"}).json()["id"]
        caster, smelter = sup("MV Caster"), sup("MV Smelter")
        # used by a T1 relationship -> delete and move are refused, naming the use
        t1 = api.post("/api/planning/supplier-parts", json={"material_id": adc, "supplier_id": caster, "sob_percent": 100, "currency": "INR"}).json()["id"]
        deleted = api.delete(f"/api/planning/materials/{adc}")
        self.assertEqual(deleted.status_code, 409)
        self.assertEqual(deleted.json()["detail"], "MV-ADC12 cannot be deleted: it is used in 1 T1 supplier/material relationship. Remove those first, or use Move to T2/T3 if it belongs there")
        moved = api.post(f"/api/planning/materials/{adc}/move", params={"tier": "T2"})
        self.assertEqual(moved.status_code, 409)
        self.assertIn("used in 1 T1 supplier/material relationship, which need a T1 code", moved.json()["detail"])
        self.assertEqual(api.post(f"/api/planning/materials/{adc}/move", params={"tier": "T1"}).status_code, 422)
        # the right setup: the smelter supplies ADC12 to the caster (T2), and the caster keeps it in its material master
        self.assertEqual(api.delete(f"/api/planning/supplier-parts/{t1}").status_code, 204)
        t2 = api.post("/api/planning/supplier-parts", json={"material_id": adc, "supplier_id": smelter, "tier": "T2", "parent_supplier_id": caster, "sob_percent": 100, "currency": "INR"}).json()["id"]
        entry = api.post("/api/planning/supplier-materials", json={"supplier_id": caster, "item_id": adc, "unit_price": 360, "currency": "INR"}).json()["id"]
        self.assertIn("1 T2/T3 supplier/material relationship, 1 supplier material master entry", api.delete(f"/api/planning/materials/{adc}").json()["detail"])
        result = api.post(f"/api/planning/materials/{adc}/move", params={"tier": "T2"})
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual((result.json()["relationships"], result.json()["supplier_materials"]), (1, 1))
        # now in T2 Material Master with its name, unit, category and price; gone from T1
        t2_codes = {m["code"]: m for m in api.get("/api/planning/tier-materials", params={"tier": "T2", "search": "MV-"}).json()["items"]}
        self.assertEqual((t2_codes["MV-ADC12"]["name"], t2_codes["MV-ADC12"]["uom"], t2_codes["MV-ADC12"]["category"], t2_codes["MV-ADC12"]["unit_price"]), ("ADC12", "kg", "Aluminimum", 350))
        self.assertEqual(api.get("/api/planning/materials", params={"search": "MV-"}).json()["total"], 0)
        link = next(r for r in api.get("/api/planning/supplier-parts", params={"search": "MV-ADC12"}).json()["items"] if r["id"] == t2)
        self.assertEqual((link["code_master"], link["material_code"], link["tier"]), ("T2", "MV-ADC12", "T2"))
        material_entry = next(r for r in api.get("/api/planning/supplier-materials", params={"supplier_id": caster}).json()["items"] if r["id"] == entry)
        self.assertEqual((material_entry["source"], material_entry["code"], material_entry["unit_price"]), ("T2", "MV-ADC12", 360))
        # a BOM line blocks the move; an unused code moves to T3
        bolt = api.post("/api/planning/materials", json={"code": "MV-BOLT", "name": "Bolt", "uom": "PCS"}).json()["id"]
        model = api.post("/api/planning/models", json={"code": "MV-M", "name": "Move model"}).json()["id"]
        api.post(f"/api/planning/model-bom/{model}", json={"material_id": bolt, "quantity": 1, "uom": "PCS"})
        self.assertIn("1 Model BOM Level 1 line", api.post(f"/api/planning/materials/{bolt}/move", params={"tier": "T3"}).json()["detail"])
        ore = api.post("/api/planning/materials", json={"code": "MV-ORE", "name": "Ore", "uom": "kg"}).json()["id"]
        self.assertEqual(api.post(f"/api/planning/materials/{ore}/move", params={"tier": "T3"}).status_code, 200)
        self.assertEqual([m["code"] for m in api.get("/api/planning/tier-materials", params={"tier": "T3", "search": "MV-"}).json()["items"]], ["MV-ORE"])
        # the install sample's retired material request no longer blocks deleting MAT-101
        import sqlite3
        mat101 = next(m["id"] for m in api.get("/api/planning/materials", params={"search": "MAT-101"}).json()["items"])
        self.assertIn("1 part BOM line, 1 line in its own part BOM", api.delete(f"/api/planning/materials/{mat101}").json()["detail"])  # sample part BOMs (no model)
        connection = sqlite3.connect(Path(self.temp_dir.name) / "test.db")
        connection.execute("DELETE FROM bom_components WHERE parent_item_id = ? OR child_item_id = ?", (mat101, mat101))
        connection.commit()
        self.assertEqual(connection.execute("SELECT COUNT(*) FROM material_requests WHERE item_id = ?", (mat101,)).fetchone()[0], 1)
        connection.close()
        self.assertEqual(api.delete(f"/api/planning/materials/{mat101}").status_code, 204)
        self.assertEqual(sqlite3.connect(Path(self.temp_dir.name) / "test.db").execute("SELECT COUNT(*) FROM material_requests WHERE item_id = ?", (mat101,)).fetchone()[0], 0)

    def test_relationship_and_supplier_wise_coverage_excel(self):
        api = self.client
        xlsx_type = {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}

        def book(headers, rows, title=True):
            workbook = Workbook()
            sheet = workbook.active
            if title:
                sheet.append(["My sheet"])
            sheet.append(list(headers))
            for row in rows:
                sheet.append(list(row))
            content = io.BytesIO()
            workbook.save(content)
            return content.getvalue()

        beam = api.post("/api/planning/materials", json={"code": "XL-BEAM", "name": "Base Front Beam", "uom": "Nos"}).json()["id"]
        api.post("/api/planning/materials", json={"code": "XL-BRKT", "name": "Bracket", "uom": "Nos"})
        api.post("/api/planning/tier-materials", json={"tier": "T2", "code": "XL-STEEL", "name": "STEEL0.7", "uom": "kg"})
        for code, name in (("XL-ORE", "Iron ore"), ("XL-COAL", "Coal")):
            api.post("/api/planning/tier-materials", json={"tier": "T3", "code": code, "name": name, "uom": "kg"})
        sup = lambda code, name: api.post("/api/suppliers", json={"code": code, "name": name, "category": "Sheet Metal", "country": "India"}).json()
        savera = sup("XL-1013", "XL Savera")
        for code, name in (("XL-1004", "XL JGI"), ("XL-1014", "XL TATA"), ("XL-1030", "XL XYZ"), ("XL-1031", "XL ABC")):
            sup(code, name)
        # 1. relationship sheet with the screen's columns
        columns = ["Supplier Code", "Supplier Name", "T1 Part Code", "T1 Part Name", "SOB", "Lead time", "MOQ", "Buffer days", "Unit price", "Currency"]
        up = api.post("/api/planning/supplier-parts/upload", headers=xlsx_type, content=book(columns, [
            ["XL-1013", "any label", "XL-BEAM", "Base Front Beam", 60, 15, 100, 5, 85, "INR"], ["XL-1013", "", "XL-BRKT", "", 100, 7, 0, 0, 20, "INR"]]))
        self.assertEqual(up.status_code, 201, up.text)
        links = {r["material_code"]: r for r in api.get("/api/planning/supplier-parts", params={"supplier_id": savera["id"]}).json()["items"]}
        self.assertEqual((links["XL-BEAM"]["supplier_name"], links["XL-BEAM"]["sob_percent"], links["XL-BEAM"]["lead_time_days"], links["XL-BEAM"]["moq"], links["XL-BEAM"]["buffer_days"], links["XL-BEAM"]["unit_price"], links["XL-BEAM"]["tier"]),
                         ("XL Savera", 60, 15, 100, 5, 85, "T1"))
        bad = api.post("/api/planning/supplier-parts/upload", headers=xlsx_type, content=book(columns, [["XL-9999", "", "XL-BEAM", "", 10, 0, 0, 0, 0, "INR"]]))
        self.assertIn("XL-9999 is not in the Supplier register", bad.json()["detail"])
        template = load_workbook(io.BytesIO(api.get("/api/planning/supplier-relationships/template.xlsx", params={"supplier_id": savera["id"]}).content)).active
        self.assertEqual([c.value for c in template[3]], columns)
        self.assertEqual([[c.value for c in row][:5] for row in template.iter_rows(min_row=4)], [["XL-1013", "XL Savera", "XL-BEAM", "Base Front Beam", 60], ["XL-1013", "XL Savera", "XL-BRKT", "Bracket", 100]])
        # 2. supplier-wise coverage: beam needs 2 kg steel (JGI 30% as an Excel percent, TATA 70);
        #    per kg steel 1.7 ore (XYZ) and 0.6 coal (ABC); bracket has no BOM
        cov = ["T1 Part Code", "T1 Part Name", "Level", "Material code", "Material name", "BOM qty", "Unit", "Supplier Code", "Supplier Name", "SOB %", "Lead time days", "MOQ", "Buffer days", "Price mode", "Price interval months", "Unit price", "Currency"]
        sheet_rows = [
            ["XL-BEAM", "Base Front Beam", "T1", "XL-BEAM", "", 1, "Nos", "XL-1013", "", 60],
            ["XL-BEAM", "", "L1", "XL-STEEL", "", 2, "kg"],
            ["", "", "T2", "XL-STEEL", "", None, "", "XL-1004", "", 0.3, 10, 5000, 7, "fixed_monthly", 1, 74, "INR"],
            ["", "", "T2", "", "", None, "", "", "XL TATA", 70, 20, None, None, "spot", None, 72, "INR"],
            ["", "", "L2", "XL-ORE", "", 1.7, "kg"],
            ["", "", "T3", "", "", None, "", "XL-1030", "", 100],
            ["", "", "L2", "XL-COAL", "", 0.6, "kg"],
            ["", "", "T3", "", "", None, "", "XL-1031", "", None],
        ]
        loaded = api.post(f"/api/planning/coverage-workbook/upload?supplier_id={savera['id']}", headers=xlsx_type, content=book(cov, sheet_rows))
        self.assertEqual(loaded.status_code, 201, loaded.text)
        self.assertEqual(loaded.json(), {"parts": 1, "lines": 3, "suppliers": 4})
        view = api.get(f"/api/planning/supplier-coverage/{links['XL-BEAM']['id']}").json()
        got = [(l["depth"], l["component"]["code"], l["quantity"], [(s["sub_supplier"], s["sob_percent"], s["lead_time_days"]) for s in l["sources"]]) for l in view["lines"]]
        self.assertEqual(got, [(1, "XL-STEEL", 2, [("XL JGI", 30, 10), ("XL TATA", 70, 20)]), (2, "XL-ORE", 1.7, [("XL XYZ", 100, 0)]), (2, "XL-COAL", 0.6, [("XL ABC", 100, 0)])])
        jgi_src = view["lines"][0]["sources"][0]
        self.assertEqual((jgi_src["moq"], jgi_src["buffer_days"], jgi_src["price_mode"], jgi_src["unit_price"]), (5000, 7, "fixed_monthly", 74))
        self.assertIn("T3", next(s for s in api.get("/api/suppliers").json() if s["code"] == "XL-1030")["tiers"])  # added to the T3 tier
        # Upload supplier stock page: the whole tree, T1 supplier -> parts -> BOM and its suppliers, with the month's stock
        api.post("/api/planning/supplier-stock/2026-11?workspace=trial", json={"rows": [{"supplier_part_id": links["XL-BEAM"]["id"], "quantity": 500, "available_quantity": 400}]})
        tree = [n for n in api.get("/api/planning/supply-tree", params={"report_month": "2026-11"}).json()["items"]]
        start = next(i for i, n in enumerate(tree) if n["depth"] == 0 and n["supplier_code"] == "XL-1013")
        mine = tree[start:start + 10]
        self.assertEqual([(n["depth"], n["level"], n.get("supplier_code"), n.get("material_code")) for n in mine], [
            (0, "T1 supplier", "XL-1013", None), (1, "T1 part", "XL-1013", "XL-BEAM"), (2, "L1 material", None, "XL-STEEL"),
            (3, "T2 supplier", "XL-1004", "XL-STEEL"), (3, "T2 supplier", "XL-1014", "XL-STEEL"), (3, "L2 material", None, "XL-ORE"),
            (4, "T3 supplier", "XL-1030", "XL-ORE"), (3, "L2 material", None, "XL-COAL"), (4, "T3 supplier", "XL-1031", "XL-COAL"), (1, "T1 part", "XL-1013", "XL-BRKT")])
        self.assertEqual((mine[1]["stock"], mine[1]["available"], mine[1]["sob_percent"], mine[2]["bom_qty"], mine[2]["sob_percent"], mine[9]["stock"]), (500, 400, 60, 2, 100, None))
        # the same tree as Excel: one row per node, levels grouped (+ / - in Excel), stock on the part
        sheet = load_workbook(io.BytesIO(api.get("/api/planning/supply-tree.xlsx", params={"report_month": "2026-11"}).content)).active
        self.assertEqual([c.value for c in sheet[3]], ["Level", "Supplier Code", "Supplier Name", "Material code", "Material name", "SOB %", "BOM qty", "Unit", "Stock quantity", "Available quantity", "Arrival month"])
        first = next(r for r in range(4, sheet.max_row + 1) if sheet.cell(r, 2).value == "XL-1013" and sheet.cell(r, 1).value.strip() == "T1 supplier")
        got = [(sheet.cell(r, 1).value.strip(), sheet.cell(r, 2).value, sheet.cell(r, 4).value, sheet.cell(r, 6).value, sheet.cell(r, 7).value, sheet.cell(r, 9).value, sheet.row_dimensions[r].outline_level) for r in range(first, first + 5)]
        self.assertEqual(got, [("T1 supplier", "XL-1013", None, None, None, None, 0), ("T1 part", "XL-1013", "XL-BEAM", 60, None, 500, 1), ("L1 material", None, "XL-STEEL", 100, 2, None, 2),
                               ("T2 supplier", "XL-1004", "XL-STEEL", 30, None, None, 3), ("T2 supplier", "XL-1014", "XL-STEEL", 70, None, None, 3)])
        self.assertEqual(sheet.cell(first, 5).value, "2 parts")
        self.assertEqual(api.get("/api/planning/supply-tree.pdf").status_code, 200)
        self.assertEqual(api.get("/api/planning/supply-tree.csv").status_code, 404)
        # stock template = the tree; stock typed on a T1 part row uploads back (other rows and blank stock skipped)
        filled = load_workbook(io.BytesIO(api.get("/api/planning/supply-tree.xlsx", params={"report_month": "2026-12", "template": "true"}).content))
        tree_sheet = filled.active
        self.assertIn("Stock template", tree_sheet["A2"].value)
        beam_row = next(r for r in range(4, tree_sheet.max_row + 1) if tree_sheet.cell(r, 1).value.strip() == "T1 part" and tree_sheet.cell(r, 4).value == "XL-BEAM")
        tree_sheet.cell(beam_row, 9).value, tree_sheet.cell(beam_row, 10).value, tree_sheet.cell(beam_row, 11).value = 700, 650, date(2027, 1, 15)  # an Excel date cell
        steel_row = next(r for r in range(4, tree_sheet.max_row + 1) if tree_sheet.cell(r, 4).value == "XL-STEEL")
        tree_sheet.cell(steel_row, 9).value = 999  # an L1 material row: for reading only
        jgi_row = next(r for r in range(4, tree_sheet.max_row + 1) if tree_sheet.cell(r, 1).value.strip() == "T2 supplier" and tree_sheet.cell(r, 2).value == "XL-1004")
        tree_sheet.cell(jgi_row, 9).value, tree_sheet.cell(jgi_row, 10).value = 12000, 11000  # stock at JGI (T2) for Savera's beam
        content = io.BytesIO()
        filled.save(content)
        stock_up = api.post("/api/planning/supplier-stock/2026-12/upload?workspace=trial", headers=xlsx_type, content=content.getvalue())
        self.assertEqual((stock_up.status_code, stock_up.json().get("rows")), (201, 2), stock_up.text)
        listed_stock = [(r["supplier_code"], r["material_code"], r["quantity"], r["available_quantity"], r["arrival_month"]) for r in api.get("/api/planning/supplier-stock/2026-12", params={"workspace": "trial"}).json()["items"]]
        self.assertEqual(listed_stock, [("XL-1013", "XL-BEAM", 700, 650, "2027-01"), ("XL-1004", "XL-STEEL", 12000, 11000, None)])
        t2_node = next(n for n in api.get("/api/planning/supply-tree", params={"report_month": "2026-12"}).json()["items"] if n["level"] == "T2 supplier" and n["supplier_code"] == "XL-1004")
        self.assertEqual((t2_node["stock"], t2_node["available"]), (12000, 11000))
        # a T2 row for a supplier who is not in that BOM is refused with its row
        tree_sheet.cell(jgi_row, 2).value = "XL-1030"
        content = io.BytesIO()
        filled.save(content)
        wrong = api.post("/api/planning/supplier-stock/2026-12/upload?workspace=trial", headers=xlsx_type, content=content.getvalue())
        self.assertIn(f"Row {jgi_row}: XL XYZ is not a supplier of XL-STEEL in the BOM of XL-BEAM", wrong.json()["detail"])
        empty = load_workbook(io.BytesIO(api.get("/api/planning/supply-tree.xlsx", params={"report_month": "2026-10", "template": "true"}).content))
        content = io.BytesIO()
        empty.save(content)
        self.assertIn("No stock typed", api.post("/api/planning/supplier-stock/2026-10/upload?workspace=trial", headers=xlsx_type, content=content.getvalue()).json()["detail"])
        # relationship page download: the list in the upload's columns (one supplier or all, with its search)
        rel = load_workbook(io.BytesIO(api.get("/api/planning/supplier-relationships.xlsx", params={"supplier_id": savera["id"]}).content)).active
        self.assertEqual([c.value for c in rel[3]], columns)
        self.assertEqual([[c.value for c in row] for row in rel.iter_rows(min_row=4)], [["XL-1013", "XL Savera", "XL-BEAM", "Base Front Beam", 60, 15, 100, 5, 85, "INR"], ["XL-1013", "XL Savera", "XL-BRKT", "Bracket", 100, 7, 0, 0, 20, "INR"]])
        searched = load_workbook(io.BytesIO(api.get("/api/planning/supplier-relationships.xlsx", params={"search": "bracket"}).content)).active
        self.assertEqual([row[2].value for row in searched.iter_rows(min_row=4)], ["XL-BRKT"])
        self.assertEqual(api.get("/api/planning/supplier-relationships.pdf").status_code, 200)
        # the template gives back what is saved, so it can be edited and uploaded again
        back = load_workbook(io.BytesIO(api.get("/api/planning/coverage-workbook/template.xlsx", params={"supplier_id": savera["id"]}).content)).active
        self.assertEqual([c.value for c in back[3]], cov)
        levels = [(row[0].value, row[2].value, row[3].value, row[5].value, row[7].value, row[9].value) for row in back.iter_rows(min_row=4)]
        self.assertEqual(levels, [("XL-BEAM", "T1", "XL-BEAM", 1, "XL-1013", 60), ("XL-BEAM", "L1", "XL-STEEL", 2, None, None), ("XL-BEAM", "T2", "XL-STEEL", None, "XL-1004", 30),
                                  ("XL-BEAM", "T2", "XL-STEEL", None, "XL-1014", 70), ("XL-BEAM", "L2", "XL-ORE", 1.7, None, None), ("XL-BEAM", "T3", "XL-ORE", None, "XL-1030", 100),
                                  ("XL-BEAM", "L2", "XL-COAL", 0.6, None, None), ("XL-BEAM", "T3", "XL-COAL", None, "XL-1031", 100), ("XL-BRKT", "T1", "XL-BRKT", 1, "XL-1013", 100)])
        # errors name the row; nothing is saved
        def refused(rows):
            answer = api.post(f"/api/planning/coverage-workbook/upload?supplier_id={savera['id']}", headers=xlsx_type, content=book(cov, rows))
            self.assertEqual(answer.status_code, 422, answer.text)
            return answer.json()["detail"]
        self.assertIn("Row 3: an L2 line must come after the L1 line", refused([["XL-BEAM", "", "L2", "XL-ORE", "", 1]]))
        self.assertIn("Row 3: a T3 supplier must come after the L2 material", refused([["XL-BEAM", "", "T3", "", "", None, "", "XL-1030"]]))
        self.assertIn("is not a part XL Savera supplies", refused([["XL-OTHER", "", "L1", "XL-STEEL", "", 1]]))
        # a new L1 line closes the L2 lines of the one before: a T3 right after it has no L2 to supply
        self.assertIn("Row 6: a T3 supplier must come after the L2 material", refused([["XL-BEAM", "", "L1", "XL-STEEL", "", 2], ["", "", "L2", "XL-ORE", "", 1.7],
                                                                                      ["", "", "L1", "XL-COAL", "", 1], ["", "", "T3", "", "", None, "", "XL-1030"]]))
        self.assertIn("adds up to more than 100%", refused([["XL-BEAM", "", "L1", "XL-STEEL", "", 2], ["", "", "T2", "", "", None, "", "XL-1004", "", 60], ["", "", "T2", "", "", None, "", "XL-1014", "", 50]]))
        self.assertIn("must be XL-STEEL or blank", refused([["XL-BEAM", "", "L1", "XL-STEEL", "", 2], ["", "", "T2", "XL-ORE", "", None, "", "XL-1004", "", 60]]))
        self.assertEqual(len(api.get(f"/api/planning/supplier-coverage/{links['XL-BEAM']['id']}").json()["lines"]), 3)
        # uploading again replaces that part's BOM
        again = api.post(f"/api/planning/coverage-workbook/upload?supplier_id={savera['id']}", headers=xlsx_type, content=book(cov, [["XL-BEAM", "", "L1", "XL-STEEL", "", 2.5]]))
        self.assertEqual(again.json(), {"parts": 1, "lines": 1, "suppliers": 0})
        self.assertEqual([(l["component"]["code"], l["quantity"]) for l in api.get(f"/api/planning/supplier-coverage/{links['XL-BEAM']['id']}").json()["lines"]], [("XL-STEEL", 2.5)])

    def test_supplier_coverage_part_bom_with_t2_suppliers(self):
        # The user's example: Savera supplies Base Front Beam (T1, 60%); 1 beam needs 2 kg steel
        # FL001STEEL0.7 (a T2 Material Master code) bought from JGI 30% and TATA 70%.
        api = self.client
        beam = api.post("/api/planning/materials", json={"code": "CV-BEAM", "name": "Base Front Beam", "uom": "Nos"}).json()["id"]
        paint = api.post("/api/planning/materials", json={"code": "CV-PAINT", "name": "Paint", "uom": "L", "unit_cost": 300, "currency": "INR"}).json()["id"]
        steel = api.post("/api/planning/tier-materials", json={"tier": "T2", "code": "CV-STEEL0.7", "name": "STEEL0.7", "uom": "kg"}).json()["id"]
        sup = lambda name: api.post("/api/suppliers", json={"name": name, "category": "Sheet Metal", "country": "India"}).json()["id"]
        savera, jgi, tata = sup("CV Savera"), sup("CV JGI Steel"), sup("CV TATA Steel")
        link = api.post("/api/planning/supplier-parts", json={"material_id": beam, "supplier_id": savera, "sob_percent": 60, "lead_time_days": 15, "unit_price": 85, "currency": "INR"}).json()["id"]
        # T2 relationships typed earlier (old chain page) are NOT copied into a new line: only what is typed
        for who, sob in ((jgi, 30), (tata, 70)):
            made = api.post("/api/planning/supplier-parts", json={"tier_material_id": steel, "supplier_id": who, "tier": "T2", "parent_link_id": link, "parent_supplier_id": savera, "sob_percent": sob, "lead_time_days": 10, "moq": 5000, "unit_price": 75, "currency": "INR"})
            self.assertEqual(made.status_code, 201, made.text)
        t1_only = api.get("/api/planning/supplier-parts", params={"supplier_id": savera, "tier": "T1"}).json()["items"]
        self.assertEqual([(r["material_code"], r["tier"]) for r in t1_only], [("CV-BEAM", "T1")])
        src = lambda who, sob: {"sub_supplier_id": who, "sob_percent": sob, "lead_time_days": 10, "moq": 5000, "unit_price": 75, "currency": "INR"}
        over = api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "tier_material_id": steel, "quantity": 2, "sources": [src(jgi, 30), src(tata, 80)]})
        self.assertEqual((over.status_code, "adds up to more than 100%" in over.json()["detail"]), (422, True))
        twice = api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "tier_material_id": steel, "quantity": 2, "sources": [src(jgi, 30), src(jgi, 70)]})
        self.assertIn("entered twice", twice.json()["detail"])
        self.assertEqual(api.get(f"/api/planning/supplier-coverage/{link}").json()["lines"], [])  # nothing saved by refused requests
        bare = api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "tier_material_id": steel, "quantity": 2})
        self.assertEqual((bare.status_code, bare.json()["sources"]), (201, 0))
        self.assertEqual(api.get(f"/api/planning/supplier-coverage/{link}").json()["lines"][0]["sources"], [])  # earlier T2 rows not copied
        api.delete(f"/api/planning/supplier-bom/lines/{bare.json()['id']}")
        added = api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "tier_material_id": steel, "quantity": 2, "sources": [src(jgi, 30), src(tata, 70)]})
        self.assertEqual(added.status_code, 201, added.text)
        self.assertEqual((added.json()["code"], added.json()["sources"]), ("CV-STEEL0.7", 2))
        self.assertEqual(api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "tier_material_id": steel, "quantity": 3}).status_code, 409)  # one line per code
        self.assertEqual(api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "item_id": paint, "quantity": 0.1}).status_code, 201)  # a T1 code, other unit
        view = api.get(f"/api/planning/supplier-coverage/{link}").json()
        self.assertEqual((view["t1"]["supplier_name"], view["t1"]["material_code"], view["t1"]["sob_percent"], view["t1"]["uom"]), ("CV Savera", "CV-BEAM", 60, "Nos"))
        lines = {line["component"]["code"]: line for line in view["lines"]}
        self.assertEqual((lines["CV-STEEL0.7"]["quantity"], lines["CV-STEEL0.7"]["component"]["uom"], lines["CV-STEEL0.7"]["component"]["origin"]), (2, "kg", "T2"))
        self.assertEqual((lines["CV-PAINT"]["quantity"], lines["CV-PAINT"]["component"]["uom"], lines["CV-PAINT"]["sources"]), (0.1, "L", []))
        got = sorted((src["sub_supplier"], src["sob_percent"], src["lead_time_days"], src["moq"]) for src in lines["CV-STEEL0.7"]["sources"])
        self.assertEqual(got, [("CV JGI Steel", 30, 10, 5000), ("CV TATA Steel", 70, 10, 5000)])
        # the joined code is now in Savera's material master, linked to the T2 master
        own = [m for m in api.get("/api/planning/supplier-materials", params={"supplier_id": savera}).json()["items"] if m["code"] == "CV-STEEL0.7"]
        self.assertEqual(len(own), 1)
        # edit the T2 suppliers of the line with their own terms
        line_id = lines["CV-STEEL0.7"]["id"]
        put = api.put(f"/api/planning/supplier-bom/lines/{line_id}/sources", json=[
            {"sub_supplier_id": jgi, "sob_percent": 40, "lead_time_days": 12, "moq": 1000, "buffer_days": 7, "price_mode": "fixed_monthly", "unit_price": 74, "currency": "INR"},
            {"sub_supplier_id": tata, "sob_percent": 60, "lead_time_days": 20, "unit_price": 72, "currency": "INR"}])
        self.assertEqual(put.status_code, 200, put.text)
        src = {s["sub_supplier"]: s for s in next(l for l in api.get(f"/api/planning/supplier-coverage/{link}").json()["lines"] if l["id"] == line_id)["sources"]}
        self.assertEqual((src["CV JGI Steel"]["sob_percent"], src["CV JGI Steel"]["buffer_days"], src["CV JGI Steel"]["price_mode"], src["CV JGI Steel"]["price_status"], src["CV TATA Steel"]["lead_time_days"]), (40, 7, "fixed_monthly", "Current", 20))
        # price status: JGI's monthly price last updated 2 months ago is due; saving it again unchanged keeps
        # that month (still due); a new price is from this month (current)
        import sqlite3
        db_file = Path(self.temp_dir.name) / "test.db"
        connection = sqlite3.connect(db_file)
        connection.execute("UPDATE supplier_bom_line_sources SET price_updated_month = '2000-01' WHERE line_id = ? AND sub_supplier_id = ?", (line_id, jgi))
        connection.commit()
        connection.close()
        status = lambda: {s["sub_supplier"]: s["price_status"] for s in next(l for l in api.get(f"/api/planning/supplier-coverage/{link}").json()["lines"] if l["id"] == line_id)["sources"]}
        self.assertEqual(status(), {"CV JGI Steel": "Update due", "CV TATA Steel": "Spot"})
        same = [{"sub_supplier_id": jgi, "sob_percent": 40, "price_mode": "fixed_monthly", "unit_price": 74, "currency": "INR"}, {"sub_supplier_id": tata, "sob_percent": 60, "unit_price": 72, "currency": "INR"}]
        self.assertEqual(api.put(f"/api/planning/supplier-bom/lines/{line_id}/sources", json=same).status_code, 200)
        self.assertEqual(status()["CV JGI Steel"], "Update due")
        same[0]["unit_price"] = 76
        self.assertEqual(api.put(f"/api/planning/supplier-bom/lines/{line_id}/sources", json=same).status_code, 200)
        self.assertEqual(status()["CV JGI Steel"], "Current")
        over = api.put(f"/api/planning/supplier-bom/lines/{line_id}/sources", json=[{"sub_supplier_id": jgi, "sob_percent": 60}, {"sub_supplier_id": tata, "sob_percent": 60}])
        self.assertEqual(over.status_code, 422)
        self.assertEqual(api.get("/api/planning/supplier-coverage/999999").status_code, 404)
        # L2: steel has its own BOM (per 1 kg steel: 1.7 kg iron ore from XYZ, 0.6 kg coal from ABC)
        ore = api.post("/api/planning/tier-materials", json={"tier": "T3", "code": "CV-ORE", "name": "Iron ore", "uom": "kg"}).json()["id"]
        coal = api.post("/api/planning/tier-materials", json={"tier": "T3", "code": "CV-COAL", "name": "Coal", "uom": "kg"}).json()["id"]
        xyz, abc = sup("CV XYZ Mining"), sup("CV ABC Coal")
        ore_line = api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "parent_line_id": line_id, "tier_material_id": ore, "quantity": 1.7})
        self.assertEqual(ore_line.status_code, 201, ore_line.text)
        self.assertEqual(api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "parent_line_id": line_id, "tier_material_id": coal, "quantity": 0.6}).status_code, 201)
        self.assertEqual(api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "parent_line_id": line_id, "tier_material_id": ore, "quantity": 2}).status_code, 409)
        l3 = api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "parent_line_id": ore_line.json()["id"], "tier_material_id": coal, "quantity": 0.05})  # L3 is allowed
        self.assertEqual(l3.status_code, 201, l3.text)
        self.assertEqual(api.post("/api/planning/supplier-coverage/lines", json={"link_id": link, "parent_line_id": l3.json()["id"], "tier_material_id": ore, "quantity": 1}).status_code, 422)  # no L4
        api.delete(f"/api/planning/supplier-bom/lines/{l3.json()['id']}")
        # the steel line keeps its T2 suppliers while it has lines below
        self.assertEqual(api.put(f"/api/planning/supplier-bom/lines/{ore_line.json()['id']}/sources", json=[{"sub_supplier_id": xyz, "sob_percent": 100, "unit_price": 9, "currency": "INR"}]).status_code, 200)
        self.assertEqual(api.put(f"/api/planning/supplier-bom/lines/{line_id}/sources", json=[{"sub_supplier_id": jgi, "sob_percent": 30, "unit_price": 74, "currency": "INR"}, {"sub_supplier_id": tata, "sob_percent": 70, "unit_price": 72, "currency": "INR"}]).status_code, 200)
        view = api.get(f"/api/planning/supplier-coverage/{link}").json()
        by_code = {line["component"]["code"]: line for line in view["lines"]}
        self.assertEqual((by_code["CV-ORE"]["depth"], by_code["CV-ORE"]["parent_line_id"], by_code["CV-ORE"]["quantity"], by_code["CV-COAL"]["quantity"]), (2, line_id, 1.7, 0.6))
        self.assertEqual([src["sub_supplier"] for src in by_code["CV-ORE"]["sources"]], ["CV XYZ Mining"])
        self.assertEqual(len(by_code["CV-STEEL0.7"]["sources"]), 2)
        # steel is priced at what JGI / TATA charge (30% x 74 + 70% x 72 = 72.6), not from ore + coal
        self.assertAlmostEqual(by_code["CV-STEEL0.7"]["unit_price"], 72.6)
        # copy this BOM (steel 2 kg with JGI / TATA, its L2 ore with XYZ, paint) to another part of Savera, then edit there
        beam2 = api.post("/api/planning/materials", json={"code": "CV-BEAM2", "name": "Beam 2", "uom": "Nos"}).json()["id"]
        link2 = api.post("/api/planning/supplier-parts", json={"material_id": beam2, "supplier_id": savera, "sob_percent": 100, "currency": "INR"}).json()["id"]
        sources_list = api.get("/api/planning/bom-copy-sources").json()["items"]
        self.assertEqual([(row["part_code"], row["supplier_name"], row["lines"]) for row in sources_list if row["part_code"].startswith("CV-")], [("CV-BEAM", "CV Savera", 4)])
        copied = api.post(f"/api/planning/supplier-coverage/{link2}/copy-from", json={"source_link_id": link})
        self.assertEqual((copied.status_code, copied.json()), (201, {"lines": 4, "sources": 3}), copied.text)
        got = {(l["depth"], l["component"]["code"]): (l["quantity"], sorted((s["sub_supplier"], s["sob_percent"]) for s in l["sources"])) for l in api.get(f"/api/planning/supplier-coverage/{link2}").json()["lines"]}
        self.assertEqual(got[(1, "CV-STEEL0.7")], (2, [("CV JGI Steel", 30), ("CV TATA Steel", 70)]))
        self.assertEqual((got[(2, "CV-ORE")], got[(1, "CV-PAINT")]), ((1.7, [("CV XYZ Mining", 100)]), (0.1, [])))
        self.assertEqual(api.post(f"/api/planning/supplier-coverage/{link2}/copy-from", json={"source_link_id": link}).status_code, 409)  # has a BOM now
        self.assertEqual(api.post(f"/api/planning/supplier-coverage/{link2}/copy-from", json={"source_link_id": link, "replace": True}).json()["lines"], 4)
        self.assertEqual(len(api.get(f"/api/planning/supplier-coverage/{link2}").json()["lines"]), 4)  # replaced, not doubled
        self.assertEqual(api.post(f"/api/planning/supplier-coverage/{link}/copy-from", json={"source_link_id": link}).status_code, 422)
        # the source part is unchanged
        self.assertEqual(len(api.get(f"/api/planning/supplier-coverage/{link}").json()["lines"]), 4)
        # copy to another supplier's part: its material master gets the codes
        other = sup("CV Other Fabricator")
        link3 = api.post("/api/planning/supplier-parts", json={"material_id": beam2, "supplier_id": other, "sob_percent": 0, "currency": "INR"}).json()["id"]
        self.assertEqual(api.post(f"/api/planning/supplier-coverage/{link3}/copy-from", json={"source_link_id": link}).json()["lines"], 4)
        codes = sorted(m["code"] for m in api.get("/api/planning/supplier-materials", params={"supplier_id": other}).json()["items"])
        self.assertEqual(codes, ["CV-COAL", "CV-ORE", "CV-PAINT", "CV-STEEL0.7"])
        # one part's BOM in Excel (to copy it there): only that part's rows
        one = load_workbook(io.BytesIO(api.get("/api/planning/coverage-workbook/template.xlsx", params={"supplier_id": savera, "link_id": link}).content)).active
        self.assertEqual({row[0].value for row in one.iter_rows(min_row=4)}, {"CV-BEAM"})

    def test_supply_chain_rows_from_users_sheet(self):
        # The user's sheet: Supangita 40% / Savera 60% supply Base Front Beam; each buys steel
        # FL001STEEL0.7 from JGI 30% and TATA 70% (T2). Row 2 / 4 leave the T1 block blank.
        api = self.client
        mat = lambda code, name, uom="PCS": api.post("/api/planning/materials", json={"code": code, "name": name, "uom": uom}).json()["id"]
        beam, steel, bracket = mat("CH-BEAM", "Base Front Beam"), mat("CH-STEEL", "STEEL0.7", "kg"), mat("CH-BRKT", "Bracket")
        ore = api.post("/api/planning/tier-materials", json={"tier": "T3", "code": "CH-ORE", "name": "Iron ore", "uom": "kg"}).json()["id"]
        sup = lambda name: api.post("/api/suppliers", json={"name": name, "category": "Sheet Metal", "country": "India"}).json()["id"]
        supangita, savera, jgi, tata, mine = sup("CH Supangita"), sup("CH Savera"), sup("CH JGI Steel"), sup("CH TATA Steel"), sup("CH Mine")
        block = lambda supplier, part, sob, price=0, key="material_id": {"supplier_id": supplier, key: part, "sob_percent": sob, "unit_price": price}
        rows = [
            {"t1": block(supangita, beam, 40, 500), "t2": block(jgi, steel, 30, 70), "lead_time_days": 20, "moq": 1000, "buffer_days": 7, "currency": "INR"},
            {"t2": block(tata, steel, 70, 68), "lead_time_days": 30, "currency": "INR"},
            {"t1": block(savera, beam, 60, 520), "t2": block(jgi, steel, 30, 70), "currency": "INR"},
            {"t2": block(tata, steel, 70, 68), "currency": "INR"},
        ]
        saved = api.post("/api/planning/supplier-chain/rows", json={"rows": rows})
        self.assertEqual((saved.status_code, saved.json()["relationships"]), (201, 6), saved.text)
        chain = api.get("/api/planning/supplier-chain", params={"search": "CH-"}).json()["items"]
        flat = [(r["t1"]["supplier_name"], r["t1"]["sob_percent"], r["t2"]["supplier_name"], r["t2"]["material_code"], r["t2"]["sob_percent"]) for r in chain]
        self.assertEqual(flat, [("CH Savera", 60, "CH JGI Steel", "CH-STEEL", 30), ("CH Savera", 60, "CH TATA Steel", "CH-STEEL", 70),
                                ("CH Supangita", 40, "CH JGI Steel", "CH-STEEL", 30), ("CH Supangita", 40, "CH TATA Steel", "CH-STEEL", 70)])
        # terms belong to the last supplier in the row: JGI under Supangita has lead 20, MOQ 1000; Supangita keeps 0
        jgi_supangita = next(r for r in chain if r["t1"]["supplier_name"] == "CH Supangita" and r["t2"]["supplier_name"] == "CH JGI Steel")
        self.assertEqual((jgi_supangita["last"]["lead_time_days"], jgi_supangita["last"]["moq"], jgi_supangita["t1"]["lead_time_days"]), (20, 1000, 0))
        self.assertEqual(jgi_supangita["t2"]["parent_link_id"], jgi_supangita["t1"]["id"])
        # the row's currency is for every price in it (T1 too); other terms stay with the last supplier
        self.assertEqual((jgi_supangita["t1"]["currency"], jgi_supangita["t1"]["unit_price"], jgi_supangita["t2"]["currency"]), ("INR", 500, "INR"))
        # each level can carry its own terms in its block (T1 lead 15, MOQ 50 here; the T2 keeps the row's terms)
        own = api.post("/api/planning/supplier-chain/rows", json={"rows": [{"t1": {**block(supangita, beam, 40, 500), "lead_time_days": 15, "moq": 50, "currency": "USD"}, "t2": block(jgi, steel, 30, 70), "lead_time_days": 20, "moq": 1000, "buffer_days": 7, "currency": "INR"}]})
        self.assertEqual(own.status_code, 201, own.text)
        mine_rows = [r for r in api.get("/api/planning/supplier-chain", params={"search": "CH Supangita"}).json()["items"] if r["t2"] and r["t2"]["supplier_name"] == "CH JGI Steel"]
        self.assertEqual((mine_rows[0]["t1"]["lead_time_days"], mine_rows[0]["t1"]["moq"], mine_rows[0]["t1"]["currency"], mine_rows[0]["t2"]["lead_time_days"], mine_rows[0]["t2"]["moq"]), (15, 50, "USD", 20, 1000))
        # SOB per parent part: a third steel supplier under Supangita's beam would exceed 100%
        over = api.post("/api/planning/supplier-chain/rows", json={"rows": [{"t1": block(supangita, beam, 40, 500), "t2": block(mine, steel, 10)}]})
        self.assertEqual(over.status_code, 422)
        self.assertIn("Row 1: Supplier SOB cannot exceed 100%", over.json()["detail"])
        no_t1 = api.post("/api/planning/supplier-chain/rows", json={"rows": [{"t2": block(tata, steel, 70)}]})
        self.assertIn("Row 1: fill the T1 supplier", no_t1.json()["detail"])
        # saving again updates: JGI 20 / TATA 80 under Savera
        again = api.post("/api/planning/supplier-chain/rows", json={"rows": [{"t1": block(savera, beam, 60, 520), "t2": block(jgi, steel, 20, 70)}, {"t2": block(tata, steel, 80, 68)}]})
        self.assertEqual(again.status_code, 201, again.text)
        savera_rows = [r for r in api.get("/api/planning/supplier-chain", params={"search": "CH Savera"}).json()["items"]]
        self.assertEqual([r["t2"]["sob_percent"] for r in savera_rows], [20, 80])
        self.assertEqual(len(api.get("/api/planning/supplier-parts", params={"search": "CH-"}).json()["items"]), 6)  # updated, not added
        # the same relationship typed twice in one save keeps the first row's values (Savera stays 60% at 520)
        twice = api.post("/api/planning/supplier-chain/rows", json={"rows": [{"t1": block(savera, beam, 60, 520), "t2": block(jgi, steel, 20, 70)}, {"t1": block(savera, beam, 55, 999), "t2": block(tata, steel, 80, 68)}]})
        self.assertEqual(twice.status_code, 201, twice.text)
        savera_t1 = api.get("/api/planning/supplier-chain", params={"search": "CH Savera"}).json()["items"][0]["t1"]
        self.assertEqual((savera_t1["sob_percent"], savera_t1["unit_price"]), (60, 520))
        # a T3 under JGI (Supangita's beam) with a T3 Material Master code
        t3 = api.post("/api/planning/supplier-chain/rows", json={"rows": [{"t1": block(supangita, beam, 40, 500), "t2": block(jgi, steel, 30, 70), "t3": block(mine, ore, 100, 9, "tier_material_id")}]})
        self.assertEqual(t3.status_code, 201, t3.text)
        deep = next(r for r in api.get("/api/planning/supplier-chain", params={"search": "CH-ORE"}).json()["items"] if r["t3"])
        self.assertEqual((deep["t1"]["supplier_name"], deep["t2"]["supplier_name"], deep["t3"]["material_code"], deep["t3"]["code_master"]), ("CH Supangita", "CH JGI Steel", "CH-ORE", "T3"))
        # a T1 relationship with suppliers under it cannot be deleted
        self.assertIn("under it", api.delete(f"/api/planning/supplier-parts/{deep['t1']['id']}").json()["detail"])
        # a T2 saved the old way (parent supplier only) shows under every part of that supplier
        api.post("/api/planning/supplier-parts", json={"material_id": bracket, "supplier_id": supangita, "sob_percent": 100, "currency": "INR"})
        legacy = api.post("/api/planning/supplier-parts", json={"material_id": steel, "supplier_id": mine, "tier": "T2", "parent_supplier_id": supangita, "sob_percent": 100, "currency": "INR"})
        self.assertEqual(legacy.status_code, 201, legacy.text)
        under = {r["t1"]["material_code"] for r in api.get("/api/planning/supplier-chain", params={"search": "CH Mine"}).json()["items"] if r["t2"] and r["t2"]["id"] == legacy.json()["id"]}
        self.assertEqual(under, {"CH-BEAM", "CH-BRKT"})
        # edit: changing the parent supplier drops the old parent part
        moved = api.patch(f"/api/planning/supplier-parts/{jgi_supangita['t2']['id']}", json={"parent_supplier_id": savera, "sob_percent": 0})
        self.assertEqual(moved.status_code, 200, moved.text)
        self.assertIsNone(next(r for r in api.get("/api/planning/supplier-parts", params={"search": "CH-STEEL"}).json()["items"] if r["id"] == jgi_supangita["t2"]["id"])["parent_link_id"])
        # Savera to 50% so the part's total SOB is not 100 (the part row shows the real total)
        self.assertEqual(api.patch(f"/api/planning/supplier-parts/{savera_t1['id']}", json={"sob_percent": 50}).status_code, 200)
        # Excel in the sheet layout: group row, sub-headings, then chain rows
        sheet = load_workbook(io.BytesIO(api.get("/api/planning/supplier-chain.xlsx", params={"search": "CH-"}).content)).active
        # tree as on the screen: Level | T1 (part first) | T2 | T3 | terms of the supplier on that row
        self.assertEqual((sheet["A3"].value, sheet["B3"].value, sheet["F3"].value, sheet["J3"].value, sheet["N3"].value), ("Level", "T1 Supplier", "T2 Supplier", "T3 Supplier", "Terms of the supplier on the row"))
        self.assertEqual([sheet.cell(4, c).value for c in range(2, 10)], ["Material code", "Material name", "Supplier code", "Supplier name", "Supplier code", "Supplier name", "Material code", "Material name"])
        self.assertEqual([sheet.cell(4, c).value for c in range(14, 22)], ["SOB %", "Lead time days", "MOQ", "Buffer days", "Price mode", "Price status", "Unit price", "Currency"])
        # first rows, worked out from the chain list: the CH-BEAM part (T1 count, total SOB), its first T1, that T1's first T2
        listed = api.get("/api/planning/supplier-chain", params={"search": "CH-"}).json()["items"]
        beam_t1 = {r["t1"]["id"]: r["t1"] for r in listed if r["t1"] and r["t1"]["material_code"] == "CH-BEAM"}
        first = listed[0]
        self.assertEqual([sheet.cell(5, c).value for c in (1, 2, 5, 14)], ["Part", "CH-BEAM", f"{len(beam_t1)} T1 suppliers", sum(t["sob_percent"] for t in beam_t1.values())])
        self.assertEqual([sheet.cell(6, c).value for c in (1, 2, 4, 14, 15)], ["T1", "CH-BEAM", first["t1"]["supplier_code"], first["t1"]["sob_percent"], first["t1"]["lead_time_days"]])
        self.assertEqual([sheet.cell(7, c).value for c in (1, 6, 8, 14, 15)], ["T2", first["t2"]["supplier_code"], first["t2"]["material_code"], first["t2"]["sob_percent"], first["t2"]["lead_time_days"]])
        self.assertEqual([sheet.row_dimensions[r].outline_level for r in (5, 6, 7)], [0, 1, 2])
        self.assertEqual(api.get("/api/planning/supplier-chain.pdf").status_code, 200)

    def test_existing_database_gets_new_supplier_material_columns(self):
        # A database from the previous version: supplier_materials without the T1/T2/T3 link,
        # lead time and buffer columns, holding one entry. Starting the app adds the columns
        # and keeps the entry (as the supplier's own code).
        import sqlite3
        db_path = Path(self.temp_dir.name) / "old.db"
        TestClient(create_app(f"sqlite:///{db_path}")).__enter__()
        connection = sqlite3.connect(db_path)
        connection.execute("DROP TABLE supplier_materials")
        connection.execute("CREATE TABLE supplier_materials (id INTEGER PRIMARY KEY, supplier_id INTEGER, code VARCHAR(48), name VARCHAR(180), uom VARCHAR(24), unit_price FLOAT, currency VARCHAR(3), UNIQUE (supplier_id, code))")
        supplier_id = connection.execute("SELECT id FROM suppliers LIMIT 1").fetchone()[0]
        connection.execute("INSERT INTO supplier_materials (supplier_id, code, name, uom, unit_price, currency) VALUES (?, 'OLD-1', 'Old solder', 'kg', 900, 'INR')", (supplier_id,))
        # bom_line_alternates from before SOB history (no from_month): its SOB applies from the start
        connection.execute("DROP TABLE bom_line_alternates")
        connection.execute("CREATE TABLE bom_line_alternates (id INTEGER PRIMARY KEY, model_bom_id INTEGER, component_id INTEGER, main_item_id INTEGER, alternate_item_id INTEGER, sob_percent FLOAT)")
        connection.execute("INSERT INTO bom_line_alternates (model_bom_id, main_item_id, alternate_item_id, sob_percent) VALUES (1, 1, 2, 30)")
        connection.commit()
        connection.close()
        client = TestClient(create_app(f"sqlite:///{db_path}"))
        client.__enter__()
        self.assertEqual(client.post("/api/auth/setup", json={"username": "admin", "full_name": "Test Admin", "password": "Secure-test-password-2026!"}).status_code, 201)
        columns = {row[1] for row in sqlite3.connect(db_path).execute("PRAGMA table_info(supplier_materials)")}
        self.assertTrue({"item_id", "tier_material_id", "lead_time_days", "buffer_days"} <= columns, columns)
        self.assertEqual(sqlite3.connect(db_path).execute("SELECT from_month, sob_percent FROM bom_line_alternates").fetchall(), [("", 30.0)])
        rows = client.get("/api/planning/supplier-materials", params={"supplier_id": supplier_id}).json()["items"]
        self.assertEqual([(r["code"], r["source"], r["unit_price"], r["lead_time_days"]) for r in rows], [("OLD-1", "Own", 900, 0)])
        client.__exit__(None, None, None)

    def test_supplier_stock_upload_requires_tier_when_relationship_is_ambiguous(self):
        material = next(row for row in self.client.get("/api/planning/materials?limit=500").json()["items"] if row["code"] == "MAT-104")
        suppliers = [row for row in self.client.get("/api/suppliers").json() if row["status"] == "Approved"]
        direct_supplier, parent_supplier = suppliers[:2]
        direct = self.client.post("/api/planning/supplier-parts", json={
            "material_id": material["id"], "supplier_id": direct_supplier["id"], "tier": "Direct",
            "sob_percent": 100, "lead_time_days": 20, "unit_price": 1,
        })
        upstream = self.client.post("/api/planning/supplier-parts", json={
            "material_id": material["id"], "supplier_id": direct_supplier["id"], "tier": "T2",
            "parent_supplier_id": parent_supplier["id"], "sob_percent": 100, "lead_time_days": 40,
            "unit_price": 1,
        })
        self.assertEqual(direct.status_code, 201, direct.text)
        self.assertEqual(upstream.status_code, 201, upstream.text)
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["supplier_code", "material_code", "quantity"])
        sheet.append([direct_supplier["code"], material["code"], 10])
        content = io.BytesIO()
        workbook.save(content)
        ambiguous = self.client.post(
            "/api/planning/supplier-stock/2026-10/upload?workspace=trial",
            content=content.getvalue(),
            headers={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        )
        self.assertEqual(ambiguous.status_code, 422)
        self.assertIn("ambiguous", ambiguous.json()["detail"])

    def test_branded_excel_and_pdf_exports_cover_visible_tables(self):
        product = self.client.post("/api/planning/products", json={"code": "EXPORT-PRD", "name": "Export product"})
        model = self.client.post("/api/planning/models", json={
            "product_id": product.json()["id"], "code": "EXPORT-MODEL", "name": "Export model",
        })
        material = next(row for row in self.client.get("/api/planning/materials?limit=500").json()["items"] if row["code"] == "MAT-104")
        bom = self.client.post(f"/api/planning/model-bom/{model.json()['id']}", json={
            "material_id": material["id"], "quantity": 2, "uom": material["uom"],
        })
        self.assertEqual(bom.status_code, 201, bom.text)
        endpoints = [
            "/api/planning/tables/materials", "/api/planning/tables/products",
            "/api/planning/tables/models", "/api/planning/tables/suppliers",
            "/api/planning/tables/supplier-parts", "/api/planning/tables/plans?workspace=trial",
            "/api/planning/tables/actuals?period=2026-10&workspace=trial",
            "/api/planning/tables/purchase-orders?start_month=2026-10&end_month=2026-10&workspace=trial",
            "/api/planning/tables/transfers?workspace=trial", "/api/planning/tables/users",
            f"/api/planning/tables/bom?model_id={model.json()['id']}&tab=OEM%20BOM",
        ]
        for endpoint in endpoints:
            with self.subTest(endpoint=endpoint):
                path, separator, query = endpoint.partition("?")
                suffix = f"?{query}" if separator else ""
                excel = self.client.get(f"{path}/xlsx{suffix}")
                pdf = self.client.get(f"{path}/pdf{suffix}")
                self.assertEqual(excel.status_code, 200, excel.text)
                self.assertTrue(excel.content.startswith(b"PK"))
                self.assertEqual(pdf.status_code, 200, pdf.text)
                self.assertTrue(pdf.content.startswith(b"%PDF"))
        bom_export = self.client.get(
            f"/api/planning/tables/bom/xlsx?model_id={model.json()['id']}&tab=Model%20BOM%20Level%201"
        )
        self.assertEqual(bom_export.status_code, 200, bom_export.text)
        exported_bom = load_workbook(io.BytesIO(bom_export.content), data_only=True)
        self.assertEqual(
            [cell.value for cell in exported_bom.active[4]][:7],
            ["Model code", "Model name", "Level", "Part code", "Part name", "Quantity in model", "Unit"],
        )


if __name__ == "__main__":
    unittest.main()
