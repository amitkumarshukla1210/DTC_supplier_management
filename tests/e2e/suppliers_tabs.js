// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/suppliers_tabs.js   (needs Node + Playwright).
// Checks: six Supplier tabs; material picker lists only purchased parts (server search);
// supplier search; relationship form + Excel upload; column widths remembered per tab.
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");
const os = require("os");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 950 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const api = async (method, p, data) => {
    const r = await page.request.fetch(BASE + p, { method, data, headers: { "content-type": "application/json" } });
    if (r.status() >= 300) throw new Error(`${method} ${p} -> ${r.status()} ${await r.text()}`);
    return r.json();
  };
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  const mat = async (code, name, uom) => (await api("POST", "/api/planning/materials", { code, name, uom, category: "Raw material" })).id;
  const motor = await mat("ST-MOTOR", "Motor", "PCS");
  const steel = await mat("ST-STEEL", "Steel 0.8", "kg");
  const scrap = await mat("ST-SCRAP", "Steel scrap", "kg");
  await mat("ST-PCB", "PCB board", "PCS");
  const model = await api("POST", "/api/planning/models", { code: "ST-12000", name: "Eva" });
  await api("POST", `/api/planning/model-bom/${model.id}`, { material_id: motor, quantity: 1, uom: "PCS" });
  await api("POST", `/api/planning/part-bom/${motor}/new`, { tab: "Level 2 Part BOM", lines: [{ child_item_id: steel, quantity: 2 }, { child_item_id: scrap, quantity: -0.3 }] });

  await page.goto(BASE);
  await page.click('[data-page="suppliers"]');
  await page.waitForSelector("[data-supplier-tab]");
  const tabNames = await page.locator("[data-supplier-tab]").allTextContents();
  check("Six Supplier tabs in the requested order", JSON.stringify(tabNames) === JSON.stringify(["Supplier/material relationship", "Upload supplier stock", "Supplier/material coverage", "Supplier register", "Supplier-to-supplier material transfer", "Transfer schedule"]), JSON.stringify(tabNames));
  check("Only the open tab's section is shown", (await page.locator("#supplierPartForm").count()) === 1 && (await page.locator("#supplierStockForm").count()) === 0 && (await page.locator("#transferForm").count()) === 0);

  // Material picker: purchased parts only
  await page.waitForFunction(() => [...document.querySelectorAll("#supplierMaterialSelect option")].some((o) => o.textContent.includes("ST-")));
  await page.fill("#supplierMaterialSearch", "ST-");
  await page.waitForFunction(() => { const t = [...document.querySelectorAll("#supplierMaterialSelect option")].map((o) => o.textContent); return t.some((x) => x.includes("ST-PCB")) && t.length <= 4; });
  const options = (await page.locator("#supplierMaterialSelect option").allTextContents()).filter((t) => t.startsWith("ST-")).map((t) => t.split(" · ")[0]).sort();
  check("Material list = purchased parts only (no Motor with BOM, no negative-qty Scrap)", JSON.stringify(options) === JSON.stringify(["ST-PCB", "ST-STEEL"]), JSON.stringify(options));

  // Supplier search filters the dropdown
  const supplierSearch = page.locator('[data-supplier-search="supplier_id"]');
  await supplierSearch.fill("apex");
  const visible = await page.locator('select[name="supplier_id"] option').evaluateAll((os) => os.filter((o) => !o.hidden).map((o) => o.textContent));
  check("Supplier search filters the supplier list", visible.length >= 1 && visible.every((t) => t.toLowerCase().includes("apex")), JSON.stringify(visible));

  // Save a relationship through the form
  await page.selectOption("#supplierMaterialSelect", String(steel));
  await page.fill('#supplierPartForm [name="sob_percent"]', "60");
  await page.fill('#supplierPartForm [name="unit_price"]', "105");
  await page.click('#supplierPartForm button[type="submit"]');
  await page.waitForTimeout(800);
  const links = await api("GET", "/api/planning/supplier-parts?limit=1000");
  const saved = links.items.find((l) => l.material_code === "ST-STEEL");
  check("Form saves the relationship (SOB 60, price 105, INR default)", saved && saved.sob_percent === 60 && saved.unit_price === 105 && saved.currency === "INR", saved && JSON.stringify([saved.sob_percent, saved.unit_price, saved.currency]));

  // Excel upload through the page
  const xlsx = await page.request.get(`${BASE}/api/planning/templates/supplier-parts.xlsx`);
  check("Relationship template downloads", xlsx.ok());
  const ExcelPath = path.join(os.tmpdir(), `rel-${Date.now()}.xlsx`);
  // Build a tiny workbook via the server-side template is not possible here; use a prepared file from env if given
  if (process.env.REL_XLSX && fs.existsSync(process.env.REL_XLSX)) {
    await page.setInputFiles('#supplierPartUploadForm input[type="file"]', process.env.REL_XLSX);
    await page.click('#supplierPartUploadForm button[type="submit"]');
    await page.waitForTimeout(1500);
    const toastText = await page.locator("#toast").textContent();
    check("Upload from the page reports added / updated / new suppliers", toastText.includes("1 added") && toastText.includes("1 new supplier"), toastText);
    const after = await api("GET", "/api/planning/supplier-parts?limit=1000");
    check("Uploaded relationship is stored", after.items.some((l) => l.material_code === "ST-PCB" && l.supplier_name === "Iljin Noida" && l.unit_price === 1800));
  }
  void ExcelPath;

  // Coverage tab shows the saved link; register tab; transfer form uses purchased picker
  await page.click('[data-supplier-tab="Supplier/material coverage"]');
  await page.waitForSelector("#supplierPartTable tbody tr");
  check("Coverage tab lists the relationship", (await page.locator("#supplierPartTable").textContent()).includes("ST-STEEL"));
  await page.click('[data-supplier-tab="Supplier-to-supplier material transfer"]');
  await page.waitForFunction(() => [...document.querySelectorAll("#transferMaterialSelect option")].some((o) => o.value));
  await page.fill("#transferMaterialSearch", "ST-MOTOR");
  await page.waitForTimeout(700);
  check("Transfer material picker excludes in-house parts", !(await page.locator("#transferMaterialSelect option").allTextContents()).some((t) => t.includes("ST-MOTOR")));

  // Column width remembered on the Supplier register tab
  await page.click('[data-supplier-tab="Supplier register"]');
  await page.waitForSelector("#supplierRegisterTable tbody tr");
  const widths = () => page.locator("#supplierRegisterTable thead th").evaluateAll((ths) => ths.map((t) => Math.round(t.getBoundingClientRect().width)));
  const before = await widths();
  const header = page.locator("#supplierRegisterTable thead th").nth(1);
  await header.hover();
  const hb = await header.locator(".column-resizer").boundingBox();
  await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2);
  await page.mouse.down();
  await page.mouse.move(hb.x + hb.width / 2 + 90, hb.y + hb.height / 2, { steps: 8 });
  await page.mouse.up();
  const after = await widths();
  check("Supplier register column can be widened", Math.abs(after[1] - before[1] - 90) <= 6, `${before[1]} -> ${after[1]}`);
  await page.click('[data-supplier-tab="Transfer schedule"]');
  await page.waitForSelector("#transferRowsTable");
  await page.click('[data-supplier-tab="Supplier register"]');
  await page.waitForSelector("#supplierRegisterTable tbody tr");
  const back = await widths();
  check("Width remembered after changing tab", Math.abs(back[1] - after[1]) <= 3, `${back[1]}px`);
  await page.screenshot({ path: __dirname + "/suppliers_tabs.png" });

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
