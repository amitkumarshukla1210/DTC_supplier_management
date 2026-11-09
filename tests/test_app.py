import io
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

    def test_wide_month_plan_template_and_month_formats_upload(self):
        product = self.client.post("/api/planning/products", json={"code": "MONTH-PRD", "name": "Month product"})
        model = self.client.post("/api/planning/models", json={
            "product_id": product.json()["id"], "code": "MONTH-12000", "name": "Eva 8014",
        })
        self.assertEqual(model.status_code, 201, model.text)

        template = self.client.get("/api/planning/templates/plan.xlsx?issue_month=2026-04")
        self.assertEqual(template.status_code, 200, template.text)
        template_sheet = load_workbook(io.BytesIO(template.content), data_only=False).active
        self.assertEqual(template_sheet["A1"].value, "model_code")
        self.assertEqual(template_sheet["B1"].value, "model_MONTH")
        self.assertEqual(template_sheet["C1"].value.date(), date(2026, 4, 1))
        self.assertEqual(template_sheet["C1"].number_format, "mmm-yy")
        self.assertEqual(template_sheet["Z1"].value.date(), date(2028, 3, 1))

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

    def test_supplier_stock_upload_requires_tier_when_relationship_is_ambiguous(self):
        material = next(row for row in self.client.get("/api/planning/materials?limit=500").json()["items"] if row["code"] == "MAT-104")
        suppliers = [row for row in self.client.get("/api/suppliers").json() if row["status"] == "Approved"]
        direct_supplier, parent_supplier = suppliers[:2]
        direct = self.client.post("/api/planning/supplier-parts", json={
            "material_id": material["id"], "supplier_id": direct_supplier["id"], "tier": "Direct",
            "sob_percent": 100, "lead_time_days": 20, "unit_price": 1,
        })
        upstream = self.client.post("/api/planning/supplier-parts", json={
            "material_id": material["id"], "supplier_id": direct_supplier["id"], "tier": "T1",
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
