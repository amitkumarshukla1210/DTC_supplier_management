// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/sort_and_buttons.js   (needs Node + Playwright).
// Every table: Sort A → Z / Z → A in each column's filter (numbers as numbers, tree rows kept
// under their parent). Add buttons sit right above their table, not in the page heading.
// Add supplier asks for the supplier code.
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 950 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on("dialog", (d) => d.accept());
  const api = async (method, p, data) => {
    const r = await page.request.fetch(BASE + p, { method, data, headers: { "content-type": "application/json" } });
    if (r.status() >= 300) throw new Error(`${method} ${p} -> ${r.status()} ${await r.text()}`);
    return r.status() === 204 ? null : r.json();
  };
  const sortBy = async (tableSel, column, direction) => {
    const th = page.locator(`${tableSel} thead th`).nth(column);
    await th.locator("summary").click();
    await th.locator(`[data-column-sort="${direction}"]`).click();
  };
  const columnTexts = (tableSel, column) => page.locator(`${tableSel} tbody tr`).evaluateAll((rows, c) => rows.map((r) => r.cells[c]?.textContent.trim() || ""), column);
  // the button's right edge is at the right-hand end of its toolbar / heading row
  const atRight = (selector) => page.evaluate((sel) => {
    const button = document.querySelector(sel);
    const bar = button.closest(".table-toolbar, .entry-actions, .subtable-heading, .new-bom-actions");
    const b = button.getBoundingClientRect(), r = bar.getBoundingClientRect();
    const padding = parseFloat(getComputedStyle(bar).paddingRight) || 0;
    return Math.abs(r.right - padding - b.right) <= 4 || [...bar.querySelectorAll("button, a")].every((el) => el.getBoundingClientRect().right <= b.right + 1);
  }, selector);
  const above = (first, second) => page.evaluate(([a, b]) => !!(document.querySelector(a).compareDocumentPosition(document.querySelector(b)) & Node.DOCUMENT_POSITION_FOLLOWING), [first, second]);

  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  for (const [code, price] of [["SB-A", 520], ["SB-B", 1800], ["SB-C", 75]]) await api("POST", "/api/planning/tier-materials", { tier: "T2", code, name: code, uom: "kg", unit_price: price, currency: "INR" });
  const product = await api("POST", "/api/planning/products", { code: "SB-PRD", name: "Sort product" });
  const modelA = await api("POST", "/api/planning/models", { code: "SB-MA", name: "Model A", product_ids: [product.id] });
  await api("POST", "/api/planning/models", { code: "SB-MB", name: "Model B", product_ids: [product.id] });
  const part = await api("POST", "/api/planning/materials", { code: "SB-PART", name: "Part", uom: "PCS" });
  await api("POST", `/api/planning/model-bom/${modelA.id}`, { material_id: part.id, quantity: 1, uom: "PCS" });

  // Suppliers: Add supplier above the register table; sort by name
  await page.goto(BASE);
  await page.click('[data-page="suppliers"]');
  await page.waitForSelector("#supplierRegisterTable tbody tr");
  check("No Add button in the page heading", (await page.locator(".page-heading [data-action]").count()) === 0);
  check("Add supplier is at the right-hand end", await atRight('[data-action="new-supplier"]'));
  check("Add supplier sits right above the register table", await page.locator('.table-toolbar [data-action="new-supplier"]').isVisible() && await above('[data-action="new-supplier"]', "#supplierRegisterTable table"));
  await sortBy("#supplierRegisterTable", 1, "desc");
  let names = await columnTexts("#supplierRegisterTable", 1);
  check("Sort supplier Z → A", JSON.stringify(names) === JSON.stringify([...names].sort((a, b) => b.localeCompare(a))), JSON.stringify(names));
  check("Sorted column is marked", await page.locator("#supplierRegisterTable thead th").nth(1).evaluate((th) => th.classList.contains("sorted-desc")));
  await sortBy("#supplierRegisterTable", 1, "asc");
  names = await columnTexts("#supplierRegisterTable", 1);
  check("Sort supplier A → Z", JSON.stringify(names) === JSON.stringify([...names].sort((a, b) => a.localeCompare(b))), JSON.stringify(names));

  // Add supplier window: supplier code; no category pre-selected
  await page.click('[data-action="new-supplier"]');
  await page.waitForSelector('#modalForm [name="code"]');
  check("Add supplier asks for the supplier code", await page.locator('#modalForm [name="code"]').isVisible());
  check("Category starts at 'Select category'", (await page.locator("#modalForm #categorySelect option:checked").textContent()) === "Select category");
  await page.fill('#modalForm [name="code"]', "ZZ-01");
  await page.fill('#modalForm [name="name"]', "Zeta Castings");
  await page.selectOption("#modalForm #categorySelect", "Casting");
  await page.fill('#modalForm [name="country"]', "India");
  await page.click("#modalForm .modal-footer button");
  await page.waitForFunction(() => document.querySelector("#supplierRegisterTable").textContent.includes("ZZ-01"));
  check("Supplier added with your own code", (await page.locator('#supplierRegisterTable tbody tr:has-text("Zeta Castings") td').first().textContent()) === "ZZ-01");

  // entry tables: + Add row / Save all above the table
  await page.click('[data-supplier-tab="Supplier/material relationship"]');
  await page.waitForSelector("#relEntryTable");
  check("+ Add row / Save all are at the right-hand end", await atRight("#relSaveAll") && await page.evaluate(() => { const add = document.querySelector("#relAddRow").getBoundingClientRect(); const save = document.querySelector("#relSaveAll").getBoundingClientRect(); return add.right <= save.left && save.left - add.right < 20; }));
  check("+ Add row and Save all sit above the relationship table", await above("#relAddRow", "#relEntryTable") && await above("#relSaveAll", "#relEntryTable"));
  await page.click('[data-supplier-tab="Upload supplier stock"]');
  await page.waitForSelector("#stockEntryTable");
  check("+ Add row sits above the stock table", await above("#stockAddRow", "#stockEntryTable"));
  await page.click('[data-supplier-tab="Supplier-to-supplier material transfer"]');
  await page.waitForSelector("#transferEntryTable");
  check("+ Add row sits above the transfer table", await above("#transferAddRow", "#transferEntryTable"));

  // Product Master: Add buttons above their tables; numbers sort as numbers
  await page.click('[data-page="master"]');
  await page.waitForSelector("#productMasterTable tbody tr");
  check("Product Master: Add line at the right-hand end", await atRight('[data-action="new-product"]'));
  check("Product Master: Add line above the table", await above('[data-action="new-product"]', "#productMasterTable table") && (await page.locator(".page-heading [data-action]").count()) === 0);
  await page.click('[data-bom-tab="Material Master"]');
  await page.waitForSelector("#materialMasterTable table");
  check("T1 Material Master: Add material and Model Master Add line at the right-hand end", await atRight('[data-action="new-material"]') && await atRight('[data-action="new-model"]'));
  check("T1 Material Master: Add material above the table", await above('[data-action="new-material"]', "#materialMasterTable table"));
  await page.click('[data-bom-tab="T2 Material Master"]');
  await page.waitForSelector("#tierMaterialTable tbody tr");
  check("T2 Material Master: Add T2 material at the right-hand end", await atRight('[data-action="new-tier-material"]'));
  check("T2 Material Master: Add T2 material above the table", await above('[data-action="new-tier-material"]', "#tierMaterialTable table"));
  await sortBy("#tierMaterialTable", 4, "asc");
  const prices = await columnTexts("#tierMaterialTable", 4);
  check("Prices sort as numbers (75, 520, 1,800)", JSON.stringify(prices) === JSON.stringify(["75", "520", "1,800"]), JSON.stringify(prices));
  await sortBy("#tierMaterialTable", 4, "desc");
  check("Prices Z → A (1,800, 520, 75)", JSON.stringify(await columnTexts("#tierMaterialTable", 4)) === JSON.stringify(["1,800", "520", "75"]));

  // tree table: an expanded model's lines stay under it when sorting
  await page.click('[data-bom-tab="Product Master"]');
  await page.waitForSelector('#productMasterTable [data-tree-model-id]');
  await page.locator(`#productMasterTable [data-tree-model-id="${modelA.id}"]`).click();
  await page.waitForSelector('#productMasterTable tr[data-bom-level1-row]');
  await sortBy("#productMasterTable", 3, "desc");
  const order = await page.locator("#productMasterTable tbody tr").evaluateAll((rows) => rows.map((r) => `${r.dataset.treeDepth}:${r.cells[3]?.textContent.trim()}:${r.cells[5]?.textContent.trim()}`));
  const ma = order.findIndex((t) => t.startsWith("0:SB-MA"));
  const mb = order.findIndex((t) => t.startsWith("0:SB-MB"));
  check("Tree sort: models Z → A, Model A's part stays right under it", mb < ma && order[ma + 1]?.startsWith("1:") && order[ma + 1].endsWith("SB-PART"), JSON.stringify(order));
  await page.screenshot({ path: __dirname + "/sort_and_buttons.png" });
  await page.click('[data-bom-tab="Material Master"]');
  await page.waitForSelector("#materialMasterTable table");
  await page.screenshot({ path: __dirname + "/buttons_material.png" });
  await page.click('[data-page="suppliers"]');
  await page.click('[data-supplier-tab="Supplier/material relationship"]');
  await page.waitForSelector("#relEntryTable");
  await page.screenshot({ path: __dirname + "/buttons_relationship.png" });

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
