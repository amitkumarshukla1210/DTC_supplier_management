// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/suppliers_tabs.js   (needs Node + Playwright).
// Checks: Supplier register first, with Edit / Delete; relationship sub-tabs and the Excel-like
// entry table (purchased parts only, supplier search, Save all); coverage Edit / Delete; supplier
// stock typed in a table; transfer entry table; column widths remembered per tab.
const { chromium } = require("playwright");
const fs = require("fs");
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
  const toastText = async () => (await page.locator("#toast").textContent()) || "";
  const clearToast = () => page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
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
  const spare = (await api("POST", "/api/suppliers", { name: "ST Spare Supplier", category: "Steel", country: "India" })).id;

  await page.goto(BASE);
  await page.click('[data-page="suppliers"]');
  await page.waitForSelector("#supplierRegisterTable tbody tr");
  const tabNames = await page.locator("[data-supplier-tab]").allTextContents();
  check("Supplier register is the first tab", JSON.stringify(tabNames) === JSON.stringify(["Supplier register", "Supplier/material relationship", "Supplier/material coverage", "Upload supplier stock", "Supplier-to-supplier material transfer", "Transfer schedule", "Supplier materials", "Supplier BOM"]), JSON.stringify(tabNames));
  const registerHeads = await page.locator("#supplierRegisterTable thead th summary").allTextContents();
  check("Register shows one value per column (country, contact, email, phone)", ["Country", "Contact", "Email", "Phone", "Approval", "Action"].every((h) => registerHeads.includes(h)), JSON.stringify(registerHeads));

  // Register: Edit and Delete
  const spareRow = page.locator('#supplierRegisterTable tbody tr:has-text("ST Spare Supplier")');
  await spareRow.locator("[data-edit-supplier]").click();
  await page.fill('#supplierEditForm [name="name"]', "ST Spare Supplier Two");
  await page.selectOption('#supplierEditForm [name="status"]', "On hold");
  await page.click("#supplierEditForm button");
  await page.waitForFunction(() => document.querySelector("#supplierRegisterTable")?.textContent.includes("ST Spare Supplier Two"));
  check("Edit a supplier (name, approval)", (await page.locator('#supplierRegisterTable tbody tr:has-text("ST Spare Supplier Two")').textContent()).includes("On hold"));
  // change the supplier code; a used code is refused
  await page.locator('#supplierRegisterTable tbody tr:has-text("ST Spare Supplier Two") [data-edit-supplier]').click();
  await page.fill('#supplierEditForm [name="code"]', "SUP-1001");
  await clearToast();
  await page.click("#supplierEditForm button");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("already used"));
  check("A supplier code already used is refused", (await toastText()).includes("Apex Precision Works"), await toastText());
  await page.fill('#supplierEditForm [name="code"]', "SPARE-01");
  await page.click("#supplierEditForm button");
  await page.waitForFunction(() => document.querySelector("#supplierRegisterTable")?.textContent.includes("SPARE-01"));
  check("Change a supplier code", true);
  // bulk upload of suppliers + template
  check("Register has supplier upload and template", await page.locator('#supplierUploadForm a[href="/api/planning/templates/suppliers.xlsx"]').isVisible());
  if (process.env.SUP_XLSX && fs.existsSync(process.env.SUP_XLSX)) {
    await page.setInputFiles('#supplierUploadForm input[type="file"]', process.env.SUP_XLSX);
    await clearToast();
    await page.click('#supplierUploadForm button[type="submit"]');
    await page.waitForFunction(() => document.querySelector("#supplierRegisterTable")?.textContent.includes("Bulk Harness Co"));
    check("Upload suppliers in bulk (2 added)", (await toastText()).includes("2 supplier(s) added"), await toastText());
  }

  // Relationship sub-tabs + entry table
  await page.click('[data-supplier-tab="Supplier/material relationship"]');
  await page.waitForSelector("#relEntryTable tbody tr");
  const subTabs = await page.locator("[data-rel-subtab]").allTextContents();
  check("Relationship tab has two sub-tabs", JSON.stringify(subTabs) === JSON.stringify(["Supplier/material relationship", "Upload supplier/material relationships"]), JSON.stringify(subTabs));
  check("Entry is a table row with every field as a column", (await page.locator("#relEntryTable thead th").count()) === 16 && (await page.locator("#relEntryTable tbody tr").count()) === 1);
  const row1 = page.locator("#relEntryTable tbody tr").nth(0);
  await row1.locator("[data-row-material-search]").fill("ST-");
  await page.waitForFunction(() => { const t = [...document.querySelectorAll("#relEntryTable tbody tr")[0].querySelectorAll("[data-row-material] option")].map((o) => o.textContent); return t.some((x) => x.includes("ST-PCB")) && t.length <= 4; });
  const options = (await row1.locator("[data-row-material] option").allTextContents()).filter((t) => t.startsWith("ST-")).map((t) => t.split(" · ")[0]).sort();
  check("Material list = purchased parts only (no Motor with BOM, no negative-qty Scrap)", JSON.stringify(options) === JSON.stringify(["ST-PCB", "ST-STEEL"]), JSON.stringify(options));
  await row1.locator("[data-row-material]").selectOption(String(steel));
  await row1.locator("[data-row-supplier-search]").first().fill("apex");
  const relHeads = await page.locator("#relEntryTable thead th").allTextContents();
  check("Code and name in separate columns (part, supplier, parent supplier)", ["Part code (purchased part)", "Part name", "Supplier code", "Supplier name", "Parent supplier code (for T1+)", "Parent supplier name"].every((h) => relHeads.includes(h)), JSON.stringify(relHeads.slice(0, 7)));
  check("Choosing the part code fills the part name", (await row1.locator("[data-row-material-name] option:checked").textContent()) === "Steel 0.8");
  const visible = await row1.locator('[data-row-supplier-name="supplier"] option').evaluateAll((os) => os.filter((o) => !o.hidden && o.value).map((o) => o.textContent));
  check("Supplier search filters the row's supplier list (one match is picked)", visible.length === 1 && visible[0].toLowerCase().includes("apex") && (await row1.locator('[data-row-supplier="supplier"]').inputValue()) !== "", JSON.stringify(visible));
  check("Supplier code column shows the code, name column the name", (await row1.locator('[data-row-supplier="supplier"] option:checked').textContent()).startsWith("SUP-") && (await row1.locator('[data-row-supplier-name="supplier"] option:checked').textContent()).includes("Apex"));
  await row1.locator('[data-f="sob_percent"]').fill("60");
  await row1.locator('[data-f="unit_price"]').fill("105");
  // second row: PCB with the same supplier
  await page.click("#relAddRow");
  const row2 = page.locator("#relEntryTable tbody tr").nth(1);
  await row2.locator("[data-row-material-search]").fill("ST-PCB");
  await page.waitForFunction(() => document.querySelectorAll("#relEntryTable tbody tr")[1].querySelector("[data-row-material]").value !== "");
  const apexId = await row1.locator('[data-row-supplier="supplier"]').inputValue();
  await row2.locator('[data-row-supplier-name="supplier"]').selectOption(apexId);
  check("Choosing the supplier name fills the supplier code", (await row2.locator('[data-row-supplier="supplier"]').inputValue()) === apexId);
  await row2.locator('[data-f="unit_price"]').fill("1800");
  // third row left empty: ignored
  await page.click("#relAddRow");
  await clearToast();
  await page.click("#relSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("saved"));
  check("Save all saves the filled rows and skips the empty one", (await toastText()).includes("2 supplier/material relationships saved"), await toastText());
  check("Entry table is cleared after saving", (await page.locator("#relEntryTable tbody tr").count()) === 1);
  const links = await api("GET", "/api/planning/supplier-parts?limit=1000&search=ST-");
  const saved = links.items.find((l) => l.material_code === "ST-STEEL");
  check("Saved values (SOB 60, price 105, INR)", saved && saved.sob_percent === 60 && saved.unit_price === 105 && saved.currency === "INR", saved && JSON.stringify([saved.sob_percent, saved.unit_price, saved.currency]));
  // an error names the row and saves nothing
  const r = page.locator("#relEntryTable tbody tr").nth(0);
  await r.locator("[data-row-material-search]").fill("ST-STEEL");
  await page.waitForFunction(() => document.querySelectorAll("#relEntryTable tbody tr")[0].querySelector("[data-row-material]").value !== "");
  await r.locator("[data-row-supplier-search]").first().fill("koyo");
  await r.locator('[data-f="sob_percent"]').fill("50");
  await clearToast();
  await page.click("#relSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("Row 1"));
  check("Over 100% SOB is refused with the row number", (await toastText()).includes("Row 1: Supplier SOB cannot exceed 100%"), await toastText());

  await page.click('[data-rel-subtab="Upload supplier/material relationships"]');
  await page.waitForSelector("#supplierPartUploadForm");
  check("Upload sub-tab has the upload form and template", await page.locator('#supplierPartUploadForm a[href="/api/planning/templates/supplier-parts.xlsx"]').isVisible());
  check("Upload sub-tab: one panel, upload row first, then + Add row / Save all and the table", await page.evaluate(() => {
    const form = document.querySelector("#supplierPartUploadForm"), table = document.querySelector("#relEntryTable"), add = document.querySelector("#relAddRow");
    return form.closest("section") === table.closest("section") && !!(form.compareDocumentPosition(add) & Node.DOCUMENT_POSITION_FOLLOWING) && !!(add.compareDocumentPosition(table) & Node.DOCUMENT_POSITION_FOLLOWING);
  }));
  check("Upload sub-tab also has the typing table with + Add row and Save all", (await page.locator("#relEntryTable tbody tr").count()) === 1 && await page.locator("#relAddRow").isVisible() && await page.locator("#relSaveAll").isVisible());
  const u1 = page.locator("#relEntryTable tbody tr").nth(0);
  await u1.locator("[data-row-material-search]").fill("ST-PCB");
  await page.waitForFunction(() => document.querySelector("#relEntryTable [data-row-material]").value !== "");
  await u1.locator("[data-row-supplier-search]").first().fill("koyo");
  await u1.locator('[data-f="sob_percent"]').fill("0");
  await clearToast();
  await page.click("#relSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("saved"));
  check("Save all works on the upload sub-tab", (await toastText()).includes("1 supplier/material relationship saved"), await toastText());
  if (process.env.REL_XLSX && fs.existsSync(process.env.REL_XLSX)) {
    await page.setInputFiles('#supplierPartUploadForm input[type="file"]', process.env.REL_XLSX);
    await page.click('#supplierPartUploadForm button[type="submit"]');
    await page.waitForTimeout(1500);
    check("Upload from the page reports added / updated / new suppliers", (await toastText()).includes("1 new supplier"), await toastText());
  }

  // Coverage: one value per column, Edit and Delete
  await page.click('[data-supplier-tab="Supplier/material coverage"]');
  await page.waitForSelector('#supplierPartTable tbody tr:has-text("ST-STEEL")');
  const coverHeads = await page.locator("#supplierPartTable thead th summary").allTextContents();
  check("Coverage shows code and name in separate columns", ["Material code", "Material name", "Supplier code", "Supplier name", "Parent supplier", "Unit price", "Currency", "Price agreement", "Action"].every((h) => coverHeads.includes(h)), JSON.stringify(coverHeads));
  await page.locator('#supplierPartTable tbody tr:has-text("ST-STEEL") [data-edit-supplier-part]').click();
  await page.fill('#supplierPartEditForm [name="sob_percent"]', "70");
  await page.click("#supplierPartEditForm button");
  await page.waitForFunction(() => document.querySelector("#supplierPartTable tr:not(:first-child)") && [...document.querySelectorAll("#supplierPartTable tr")].some((tr) => tr.textContent.includes("ST-STEEL") && tr.textContent.includes("70%")));
  check("Edit a relationship (SOB 60 -> 70)", true);
  const pcbRows = await page.locator('#supplierPartTable tbody tr:has-text("ST-PCB")').count();
  await page.locator('#supplierPartTable tbody tr:has-text("ST-PCB") [data-delete-supplier-part]').first().click();
  await page.waitForFunction((n) => [...document.querySelectorAll("#supplierPartTable tbody tr")].filter((tr) => tr.textContent.includes("ST-PCB")).length === n - 1, pcbRows);
  check("Delete a relationship", true);

  // Supplier stock typed in a table
  await page.click('[data-supplier-tab="Upload supplier stock"]');
  await page.waitForSelector("#stockEntryTable tbody tr");
  const s1 = page.locator("#stockEntryTable tbody tr").nth(0);
  await s1.locator("[data-row-supplier-search]").fill("apex");
  await page.waitForFunction(() => [...document.querySelectorAll("#stockEntryTable [data-row-link] option")].some((o) => o.textContent.includes("ST-STEEL")));
  check("Stock tab: one panel, upload row first, then + Add row / Save all and the table", await page.evaluate(() => {
    const form = document.querySelector("#supplierStockForm"), table = document.querySelector("#stockEntryTable"), add = document.querySelector("#stockAddRow");
    return form.closest("section") === table.closest("section") && !!(form.compareDocumentPosition(add) & Node.DOCUMENT_POSITION_FOLLOWING) && !!(add.compareDocumentPosition(table) & Node.DOCUMENT_POSITION_FOLLOWING) && !document.body.textContent.includes("Or upload a workbook");
  }));
  check("Material list shows the chosen supplier's relationships", (await s1.locator("[data-row-link]").inputValue()) !== "");
  await s1.locator('[data-f="quantity"]').fill("500");
  await s1.locator('[data-f="available_quantity"]').fill("400");
  await clearToast();
  await page.click("#stockSaveAll");
  await page.waitForFunction(() => document.querySelector("#stockListTable")?.textContent.includes("ST-STEEL"));
  const stockCells = await page.locator('#stockListTable tr:has-text("ST-STEEL") td').allTextContents();
  check("Stock saved without Excel and listed for the month", stockCells.includes("500") && stockCells.includes("400"), JSON.stringify(stockCells));
  await s1.locator("[data-row-supplier-search]").fill("apex");
  await page.waitForFunction(() => document.querySelector("#stockEntryTable [data-row-link]").value !== "");
  await page.locator("#stockEntryTable tbody tr").nth(0).locator('[data-f="quantity"]').fill("10");
  await page.locator("#stockEntryTable tbody tr").nth(0).locator('[data-f="available_quantity"]').fill("20");
  await clearToast();
  await page.click("#stockSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("Row 1"));
  check("Wrong stock row is refused with the row number", (await toastText()).includes("Row 1: Available quantity"), await toastText());

  // a used supplier cannot be deleted; an unused one can
  await page.click('[data-supplier-tab="Supplier register"]');
  await page.waitForSelector("#supplierRegisterTable tbody tr");
  await clearToast();
  await page.locator('#supplierRegisterTable tbody tr:has-text("Apex") [data-delete-supplier]').click();
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("is used"));
  check("Deleting a used supplier is refused with where it is used", (await toastText()).includes("supplier/material relationship"), await toastText());
  await page.locator('#supplierRegisterTable tbody tr:has-text("ST Spare Supplier Two") [data-delete-supplier]').click();
  await page.waitForFunction(() => { const t = document.querySelector("#supplierRegisterTable"); return t && t.querySelector("tbody tr") && !t.textContent.includes("ST Spare Supplier Two"); });
  check("Delete an unused supplier", true);
  void spare;

  // Transfer entry table: purchased parts only
  await page.click('[data-supplier-tab="Supplier-to-supplier material transfer"]');
  await page.waitForSelector("#transferEntryTable tbody tr");
  const t1 = page.locator("#transferEntryTable tbody tr").nth(0);
  check("Transfer entry is a table row", (await page.locator("#transferEntryTable thead th").count()) === 10);
  await t1.locator("[data-row-material-search]").fill("ST-MOTOR");
  await page.waitForTimeout(800);
  check("Transfer material list excludes in-house parts", !(await t1.locator("[data-row-material] option").allTextContents()).some((t) => t.includes("ST-MOTOR")));

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
  // entry table columns can be widened too
  await page.click('[data-supplier-tab="Supplier/material relationship"]');
  await page.click('[data-rel-subtab="Supplier/material relationship"]');
  await page.waitForSelector('[data-rel-subtab="Supplier/material relationship"].active');
  await page.waitForSelector("#relEntryTable tbody tr");
  await page.waitForTimeout(300);
  const entryHead = page.locator("#relEntryTable thead th").nth(2);
  const eb0 = Math.round((await entryHead.boundingBox()).width);
  await entryHead.hover();
  const eh = await entryHead.locator(".column-resizer").boundingBox();
  await page.mouse.move(eh.x + eh.width / 2, eh.y + eh.height / 2);
  await page.mouse.down();
  await page.mouse.move(eh.x + eh.width / 2 + 60, eh.y + eh.height / 2, { steps: 6 });
  await page.mouse.up();
  check("Entry table columns can be widened", Math.abs(Math.round((await entryHead.boundingBox()).width) - eb0 - 60) <= 6);
  await page.screenshot({ path: __dirname + "/suppliers_tabs.png" });

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
