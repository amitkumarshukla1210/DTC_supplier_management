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
  // three Supplier Master tabs; Add supplier in the T2 tab makes a T2 supplier
  const tierTabs = await page.locator("[data-supplier-tier-tab]").allTextContents();
  check("Supplier register has T1 / T2 / T3 Supplier Master tabs", tierTabs.length === 3 && tierTabs[0].startsWith("T1 Supplier Master") && tierTabs[1].startsWith("T2 Supplier Master") && tierTabs[2].startsWith("T3 Supplier Master"), JSON.stringify(tierTabs));
  await page.click('[data-supplier-tier-tab="T2"]');
  await page.waitForSelector('[data-supplier-tier-tab="T2"].active');
  await page.click('[data-action="new-supplier"]');
  await page.waitForSelector('#modalForm [name="tiers"]');
  check("Add supplier in the T2 tab ticks T2", JSON.stringify(await page.locator('#modalForm [name="tiers"]:checked').evaluateAll((boxes) => boxes.map((b) => b.value))) === JSON.stringify(["T2"]));
  await page.fill('#modalForm [name="name"]', "ST Samicon");
  await page.selectOption("#modalForm #categorySelect", "Electronic Child Part");
  await page.fill('#modalForm [name="country"]', "India");
  await page.click("#modalForm .modal-footer button");
  await page.waitForFunction(() => document.querySelector("#supplierRegisterTable")?.textContent.includes("ST Samicon"));
  check("T2 supplier listed in the T2 tab with its tier", (await page.locator('#supplierRegisterTable tbody tr:has-text("ST Samicon") td').nth(2).textContent()) === "T2");
  await page.click('[data-supplier-tier-tab="T1"]');
  await page.waitForSelector('[data-supplier-tier-tab="T1"].active');
  check("...and not in the T1 tab", !(await page.locator("#supplierRegisterTable").textContent()).includes("ST Samicon"));
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

  // Relationship page: one panel (upload row, then + Add row / Save all and the typing table)
  await page.click('[data-supplier-tab="Supplier/material relationship"]');
  await page.waitForSelector("#relEntryTable tbody tr");
  check("No 'Upload supplier/material relationships' sub-tab any more", (await page.locator("[data-rel-subtab]").count()) === 0);
  check("Upload section on the relationship page: upload row, then + Add row / Save all, then the table", await page.evaluate(() => {
    const form = document.querySelector("#supplierPartUploadForm"), table = document.querySelector("#relEntryTable"), add = document.querySelector("#relAddRow");
    return form && form.closest("section") === table.closest("section") && !!(form.compareDocumentPosition(add) & Node.DOCUMENT_POSITION_FOLLOWING) && !!(add.compareDocumentPosition(table) & Node.DOCUMENT_POSITION_FOLLOWING);
  }) && await page.locator('#supplierPartUploadForm a[href="/api/planning/templates/supplier-parts.xlsx"]').isVisible());
  // chain typing table: T1 block | T2 block | T3 block | terms of the last supplier
  check("Entry is a chain row: 3 tier blocks of 12 columns (who + own terms) + Remove", (await page.locator("#relEntryTable thead tr:last-child th").count()) === 37 && (await page.locator("#relEntryTable tbody tr").count()) === 1);
  check("Group heading row: T1 / T2 / T3 Supplier", JSON.stringify((await page.locator("#relEntryTable thead tr.group-head th").allTextContents()).slice(0, 3).map((t) => t.split(" (")[0])) === JSON.stringify(["T1 Supplier", "T2 Supplier", "T3 Supplier"]));
  const relHeads = await page.locator("#relEntryTable thead tr:last-child th summary").allTextContents();
  check("T1 block starts with your part, then the supplier", JSON.stringify(relHeads.slice(0, 4)) === JSON.stringify(["T1 part code", "T1 part name", "T1 supplier code", "T1 supplier name"]), JSON.stringify(relHeads.slice(0, 4)));
  check("Each level has its own SOB, lead time, MOQ, buffer, price mode, interval, unit price, currency", JSON.stringify(relHeads.slice(4, 12)) === JSON.stringify(["T1 SOB (%)", "T1 lead time (days)", "T1 MOQ", "T1 buffer (days)", "T1 price mode", "T1 price interval (months)", "T1 unit price", "T1 currency"]) && relHeads[28] === "T3 SOB (%)" && relHeads[35] === "T3 currency", relHeads.slice(4, 12).join("|"));
  check("T2 block starts with the supplier, then its part", JSON.stringify(relHeads.slice(12, 16)) === JSON.stringify(["T2 supplier code", "T2 supplier name", "T2 part code", "T2 part name"]), JSON.stringify(relHeads.slice(12, 16)));
  check("Column headings have the filter (sort, search, Select all, Clear all)", await page.locator("#relEntryTable thead tr:last-child th").nth(0).evaluate((th) => !!th.querySelector("[data-column-sort]") && !!th.querySelector(".multi-search") && !!th.querySelector("[data-column-select-all]") && !!th.querySelector("[data-column-clear]")));
  const pick = (row, key, part) => row.locator(`[data-pick-key="${key}"][data-pick-part="${part}"]`);
  const hints = (row, key, part) => pick(row, key, part).evaluate((input) => [...input.list.options].map((o) => o.value));
  const rowN = (n) => page.locator("#relEntryTable tbody tr").nth(n);
  const row1 = rowN(0);
  // typing gives hints; T1 part = purchased T1 parts only
  await pick(row1, "m1", "code").fill("ST-");
  await page.waitForFunction(() => { const input = document.querySelector('#relEntryTable [data-pick-key="m1"][data-pick-part="code"]'); const v = [...input.list.options].map((o) => o.value); return v.includes("ST-PCB") && v.length <= 3; });
  check("T1 part hints = purchased parts only (no Motor with BOM, no negative-qty Scrap)", JSON.stringify((await hints(row1, "m1", "code")).filter((v) => v.startsWith("ST-")).sort()) === JSON.stringify(["ST-PCB", "ST-STEEL"]), JSON.stringify(await hints(row1, "m1", "code")));
  check("Each hint also shows the name", await pick(row1, "m1", "code").evaluate((input) => [...input.list.options].some((o) => o.value === "ST-STEEL" && o.textContent === "Steel 0.8")));
  await pick(row1, "m1", "code").fill("st-steel");
  await page.waitForFunction((id) => document.querySelector('#relEntryTable [data-row-material="m1"]').value === String(id), steel);
  check("Picking the part code (any capitals) fills the part name and shows the master's spelling", (await pick(row1, "m1", "name").inputValue()) === "Steel 0.8" && (await pick(row1, "m1", "code").inputValue()) === "ST-STEEL");
  await pick(row1, "s1", "name").fill("Apex Precision Works");
  check("Picking the supplier name fills the supplier code", (await pick(row1, "s1", "code").inputValue()) === "SUP-1001" && (await row1.locator('[data-row-supplier="s1"]').inputValue()) !== "");
  check("Supplier hints list code with name", await pick(row1, "s1", "code").evaluate((input) => [...input.list.options].some((o) => o.value === "SUP-1001" && o.textContent === "Apex Precision Works")));
  await row1.locator('[data-f="sob1"]').fill("60");
  await row1.locator('[data-f="price1"]').fill("105");
  // second row: picked by part name; supplier by code
  await page.click("#relAddRow");
  const row2 = rowN(1);
  await pick(row2, "m1", "name").fill("PCB board");
  await page.waitForFunction(() => document.querySelectorAll("#relEntryTable tbody tr")[1].querySelector('[data-row-material="m1"]').value !== "");
  check("Picking the part name fills the part code", (await pick(row2, "m1", "code").inputValue()) === "ST-PCB");
  await pick(row2, "s1", "code").fill("SUP-1001");
  await row2.locator('[data-f="price1"]').fill("1800");
  // column filter and sort work on the typed values
  const partHead = page.locator("#relEntryTable thead tr:last-child th").nth(0);
  await partHead.locator("summary").click();
  await page.waitForFunction(() => document.querySelectorAll("#relEntryTable thead tr:last-child th")[0].querySelectorAll("[data-column-value]").length === 2);
  await partHead.locator('[data-column-value][value="ST-PCB"]').check();
  await partHead.locator("[data-column-apply]").click();
  check("Column filter on typed part codes shows only the chosen rows", await row1.isHidden() && await row2.isVisible());
  await partHead.locator("summary").click();
  await partHead.locator("[data-column-clear]").click();
  check("Clear all shows every row again", await row1.isVisible() && await row2.isVisible());
  await partHead.locator("summary").click();
  await partHead.locator('[data-column-sort="desc"]').click();
  check("Sort Z → A on typed part codes", JSON.stringify(await page.locator('#relEntryTable tbody tr [data-pick-key="m1"][data-pick-part="code"]').evaluateAll((inputs) => inputs.map((i) => i.value))) === JSON.stringify(["ST-STEEL", "ST-PCB"]));
  // third row left empty: ignored
  await page.click("#relAddRow");
  await clearToast();
  await page.click("#relSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("saved"));
  check("Save all saves the filled rows and skips the empty one", (await toastText()).includes("2 rows saved (2 supplier/material relationships)"), await toastText());
  check("Entry table is cleared after saving", (await page.locator("#relEntryTable tbody tr").count()) === 1);
  const links = await api("GET", "/api/planning/supplier-parts?limit=1000&search=ST-");
  const saved = links.items.find((l) => l.material_code === "ST-STEEL");
  const pcbSaved = links.items.find((l) => l.material_code === "ST-PCB");
  check("SOB left blank is saved as 100%", pcbSaved && pcbSaved.sob_percent === 100, pcbSaved && String(pcbSaved.sob_percent));
  check("Saved values (T1, SOB 60, price 105, INR)", saved && saved.tier === "T1" && saved.sob_percent === 60 && saved.unit_price === 105 && saved.currency === "INR", saved && JSON.stringify([saved.tier, saved.sob_percent, saved.unit_price, saved.currency]));
  // a code that is not a purchased part is refused before saving
  const r = rowN(0);
  await pick(r, "m1", "code").fill("NOPE-1");
  await pick(r, "s1", "code").fill("SUP-1003");
  await page.waitForTimeout(400);
  check("An unknown code is marked while typing", await pick(r, "m1", "code").evaluate((input) => input.classList.contains("pick-unknown")));
  await clearToast();
  await page.click("#relSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("Row 1"));
  check("An unknown part code is refused with the row number", (await toastText()).includes("Row 1: T1 part NOPE-1 is not found as a purchased T1 code"), await toastText());
  // an error from the server names the row and saves nothing
  await pick(r, "m1", "code").fill("ST-STEEL");
  await page.waitForFunction(() => document.querySelector('#relEntryTable [data-row-material="m1"]').value !== "");
  await r.locator('[data-f="sob1"]').fill("50");
  await clearToast();
  await page.click("#relSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("Row 1"));
  check("Over 100% SOB is refused with the row number", (await toastText()).includes("Row 1: Supplier SOB cannot exceed 100%"), await toastText());
  await pick(r, "m1", "code").fill("ST-PCB");
  await page.waitForFunction(() => document.querySelector("#relEntryTable [data-pick-key='m1'][data-pick-part='name']").value === "PCB board");
  await r.locator('[data-f="sob1"]').fill("0");
  await clearToast();
  await page.click("#relSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("saved"));
  check("Save a 0% relationship (Koyo for PCB)", (await toastText()).includes("1 row saved (1 supplier/material relationship)"), await toastText());
  if (process.env.REL_XLSX && fs.existsSync(process.env.REL_XLSX)) {
    await page.setInputFiles('#supplierPartUploadForm input[type="file"]', process.env.REL_XLSX);
    await page.click('#supplierPartUploadForm button[type="submit"]');
    await page.waitForTimeout(1500);
    check("Upload from the page reports added / updated / new suppliers", (await toastText()).includes("1 new supplier"), await toastText());
  }

  // chain: Apex (T1) steel <- ST Samicon 70% + Koyo 30% (T2, a T2 Material Master code); row 2 leaves T1 blank
  await api("POST", "/api/planning/tier-materials", { tier: "T2", code: "ST-CHIP", name: "Micro chip", uom: "PCS" });
  const c1 = rowN(0);
  await pick(c1, "s1", "code").fill("SUP-1001");
  await pick(c1, "m1", "code").fill("ST-STEEL");
  await page.waitForFunction(() => document.querySelector('#relEntryTable [data-row-material="m1"]').value !== "");
  await c1.locator('[data-f="sob1"]').fill("60");
  await c1.locator('[data-f="price1"]').fill("105");
  await pick(c1, "s2", "name").fill("ST Samicon");
  await pick(c1, "m2", "code").fill("ST-CH");
  await page.waitForFunction(() => [...document.querySelector('#relEntryTable [data-pick-key="m2"][data-pick-part="code"]').list.options].some((o) => o.value === "ST-CHIP"));
  check("T2 part hints include T2 Material Master codes, labelled", await pick(c1, "m2", "code").evaluate((input) => [...input.list.options].some((o) => o.value === "ST-CHIP" && o.textContent.includes("T2 Material Master"))));
  check("T1 part hints never offer T2 codes", !(await hints(c1, "m1", "code")).includes("ST-CHIP"));
  await pick(c1, "m2", "code").fill("ST-CHIP");
  await page.waitForFunction(() => document.querySelector('#relEntryTable [data-row-tier-material="m2"]').value !== "");
  await c1.locator('[data-f="sob2"]').fill("70");
  await c1.locator('[data-f="price2"]').fill("12");
  await c1.locator('[data-f="lead_time_days2"]').fill("9");
  await c1.locator('[data-f="lead_time_days1"]').fill("4");
  await page.click("#relAddRow");
  const c2 = rowN(1);
  await pick(c2, "s2", "code").fill("SUP-1003");
  await pick(c2, "m2", "code").fill("ST-CHIP");
  await page.waitForFunction(() => document.querySelectorAll("#relEntryTable tbody tr")[1].querySelector('[data-row-tier-material="m2"]').value !== "");
  await c2.locator('[data-f="sob2"]').fill("30");
  await clearToast();
  await page.click("#relSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("saved") || document.querySelector("#toast")?.textContent.includes("Row"));
  check("Save a chain: 2 rows, T1 updated + 2 T2 relationships", (await toastText()).includes("2 rows saved (3 supplier/material relationships)"), await toastText());
  const chain = await api("GET", "/api/planning/supplier-chain?search=ST-CHIP");
  const koyoRow = chain.items.find((row) => row.t2 && row.t2.supplier_code === "SUP-1003");
  check("Row 2 (T1 blank) is under the same T1 as row 1", koyoRow && koyoRow.t1.material_code === "ST-STEEL" && koyoRow.t1.supplier_code === "SUP-1001" && koyoRow.t2.sob_percent === 30, JSON.stringify(koyoRow));
  const samiRow = chain.items.find((row) => row.t2 && row.t2.supplier_name === "ST Samicon");
  check("Each level keeps its own terms (T2 lead time 9)", samiRow && samiRow.t2.lead_time_days === 9 && samiRow.t1.lead_time_days === 4 && samiRow.t2.unit_price === 12, JSON.stringify(samiRow?.t2));

  // Coverage: chain table like the sheet, Edit last
  await page.click('[data-supplier-tab="Supplier/material coverage"]');
  await page.waitForSelector('#supplierPartTable tbody tr:has-text("ST-CHIP")');
  const groups = await page.locator("#supplierPartTable thead tr.group-head th").allTextContents();
  check("Coverage group headings: T1 / T2 / T3 Supplier, terms of the row", groups.slice(1, 5).map((t) => t.split(" (")[0]).join("|") === "T1 Supplier|T2 Supplier|T3 Supplier|Terms of the supplier on this row", groups.join("|"));
  const coverHeads = await page.locator("#supplierPartTable thead tr:last-child th summary").allTextContents();
  check("Coverage T1 block: material code, name, supplier code, name", JSON.stringify(coverHeads.slice(0, 4)) === JSON.stringify(["T1 Material code", "T1 Material name", "T1 Supplier code", "T1 Supplier name"]), JSON.stringify(coverHeads.slice(0, 4)));
  check("Coverage: each row's own terms, then Edit as the last column", JSON.stringify(coverHeads.slice(12)) === JSON.stringify(["SOB", "Lead time", "MOQ", "Buffer days", "Price mode", "Price status", "Unit price", "Currency", "Edit"]), JSON.stringify(coverHeads.slice(12)));
  const chainRows = await page.locator("#supplierPartTable tbody tr").evaluateAll((trs) => trs.filter((tr) => tr.textContent.includes("ST-CHIP")).map((tr) => ({ depth: tr.dataset.treeDepth, t1: tr.cells[3].textContent, t1faded: tr.cells[3].classList.contains("chain-repeat"), t2: tr.cells[6].textContent, chip: tr.cells[7].textContent, sob: tr.cells[13].textContent, lead: tr.cells[14].textContent, edits: [...tr.cells[21].querySelectorAll("button")].map((b) => b.textContent) })));
  check("Two T2 rows under Apex steel: Samicon 70%, Koyo 30%, Apex shown faded", chainRows.length === 2 && chainRows.every((row) => row.depth === "2" && row.t1 === "SUP-1001" && row.t1faded && row.edits.join() === "Edit T2") && chainRows.map((row) => row.sob).sort().join() === "30%,70%", JSON.stringify(chainRows));
  check("T2 code marked as from the T2 Material Master", chainRows[0].chip.includes("ST-CHIP") && chainRows[0].chip.includes("T2"));
  check("Samicon row shows its own lead time (9)", chainRows.some((row) => row.t2 === "ST Samicon" && row.lead === "9 days"));
  const apexSteel = page.locator('#supplierPartTable tbody tr[data-tree-depth="1"]', { hasText: "ST-STEEL" }).filter({ hasText: "SUP-1001" });
  const apexCells = await apexSteel.evaluate((tr) => [tr.cells[13].textContent, tr.cells[14].textContent, tr.cells[21].textContent, !!tr.querySelector("[data-sb-toggle]")]);
  check("T1 row (Apex, steel): its own SOB 60% and lead time 4, Edit T1, +/- for its T2", JSON.stringify(apexCells) === JSON.stringify(["60%", "4 days", "Edit T1", true]), JSON.stringify(apexCells));
  check("Part row above it", (await page.locator('#supplierPartTable tbody tr[data-tree-depth="0"]', { hasText: "ST-STEEL" }).count()) === 1);
  await apexSteel.locator('[data-edit-supplier-part]').click();
  await page.waitForSelector("#partEditTable tbody tr");
  const editRow = page.locator("#partEditTable tbody tr").first();
  check("Edit opens the full row filled in (supplier, part, tier, SOB)", (await editRow.locator('[data-pick-key="supplier"][data-pick-part="code"]').inputValue()) === "SUP-1001" && (await editRow.locator('[data-pick-key="material"][data-pick-part="code"]').inputValue()) === "ST-STEEL" && (await editRow.locator('[data-f="tier"]').inputValue()) === "T1" && (await editRow.locator('[data-f="sob_percent"]').inputValue()) === "60");
  // a T2 by mistake without parent is refused; then fix and save
  await editRow.locator('[data-f="tier"]').selectOption("T2");
  await clearToast();
  await page.click("#partEditSave");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("parent supplier"));
  check("Full-row edit: T2 without parent is refused", true);
  // a T1 with T2 suppliers under it cannot be deleted
  await clearToast();
  await page.click("#modalBody [data-delete-supplier-part]");
  await page.waitForFunction(() => (document.querySelector("#toast")?.textContent || "").length > 0);
  check("Delete is inside Edit; a T1 with T2 suppliers under it is not deleted", !(await toastText()).includes("Relationship deleted"), await toastText());
  await editRow.locator('[data-f="tier"]').selectOption("T1");
  await editRow.locator('[data-f="sob_percent"]').fill("70");
  await page.click("#partEditSave");
  await page.waitForFunction(() => [...document.querySelectorAll("#supplierPartTable tbody tr")].some((tr) => tr.dataset.treeDepth === "1" && tr.textContent.includes("ST-STEEL") && tr.cells[13].textContent === "70%"));
  check("Edit a relationship (SOB 60 -> 70)", true);
  const pcbRows = await page.locator('#supplierPartTable tbody tr[data-tree-depth="1"]:has-text("ST-PCB")').count();
  await page.locator('#supplierPartTable tbody tr[data-tree-depth="1"]:has-text("ST-PCB")').first().locator("[data-edit-supplier-part]").first().click();
  await page.waitForSelector("#partEditTable tbody tr");
  await page.click("#modalBody [data-delete-supplier-part]");
  await page.waitForFunction((n) => [...document.querySelectorAll("#supplierPartTable tbody tr")].filter((tr) => tr.dataset.treeDepth === "1" && tr.textContent.includes("ST-PCB")).length === n - 1, pcbRows);
  check("Delete a relationship from its Edit window", true);

  // Supplier stock typed in a table
  await page.click('[data-supplier-tab="Upload supplier stock"]');
  await page.waitForSelector("#stockEntryTable tbody tr");
  const s1 = page.locator("#stockEntryTable tbody tr").nth(0);
  await s1.locator('[data-pick-key="supplier"][data-pick-part="code"]').fill("SUP-1001");
  await page.waitForFunction(() => [...document.querySelectorAll("#stockEntryTable [data-row-link] option")].some((o) => o.textContent.includes("ST-STEEL")));
  check("Stock row: typing the supplier code fills its name", (await s1.locator('[data-pick-key="supplier"][data-pick-part="name"]').inputValue()) === "Apex Precision Works");
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
  await page.locator("#stockEntryTable tbody tr").nth(0).locator('[data-pick-key="supplier"][data-pick-part="name"]').fill("Apex Precision Works");
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
  await t1.locator('[data-pick-key="material"][data-pick-part="code"]').fill("ST-");
  await page.waitForFunction(() => document.querySelector('#transferEntryTable [data-pick-key="material"][data-pick-part="code"]').list.options.length > 0);
  check("Transfer part hints exclude in-house parts", !(await t1.locator('[data-pick-key="material"][data-pick-part="code"]').evaluate((input) => [...input.list.options].map((o) => o.value))).includes("ST-MOTOR"));

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
  await page.waitForSelector("#relEntryTable tbody tr");
  await page.waitForTimeout(300);
  const entryHead = page.locator("#relEntryTable thead tr:last-child th").nth(2);
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
