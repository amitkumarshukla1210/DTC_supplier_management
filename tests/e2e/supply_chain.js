// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/supply_chain.js   (needs Node + Playwright).
// Supplier register -> (click supplier) Supplier/material relationship -> (click part) Supplier/material
// coverage, with the user's example: 60,000 Base Front Beams; Savera makes 60%; 2 kg steel per beam from
// JGI 30% / TATA 70%; per kg steel 1.7 kg iron ore (XYZ) and 0.6 kg coal (ABC).
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1700, height: 950 } });
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
  const waitToast = (text) => page.waitForFunction((t) => document.querySelector("#toast")?.textContent.includes(t), text);
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  await api("POST", "/api/planning/materials", { code: "UF220BFB010", name: "Base Front Beam", uom: "Nos", unit_cost: 900, currency: "INR", category: "Sheet Metal" });
  await api("POST", "/api/planning/tier-materials", { tier: "T2", code: "FL001STEEL0.7", name: "STEEL0.7", uom: "kg" });
  await api("POST", "/api/planning/tier-materials", { tier: "T3", code: "ORE-01", name: "Iron ore", uom: "kg" });
  await api("POST", "/api/planning/tier-materials", { tier: "T3", code: "COAL-01", name: "Coal", uom: "kg" });
  const sup = (code, name) => api("POST", "/api/suppliers", { code, name, category: "Sheet Metal", country: "India" });
  for (const [code, name] of [["SUP-2013", "Savera"], ["SUP-2007", "Supangita"], ["SUP-2004", "JGI STEEL"], ["SUP-2014", "TATA Steel"], ["SUP-2030", "XYZ Mining"], ["SUP-2031", "ABC Coal"]]) await sup(code, name);

  // 1. Supplier register: click the supplier code
  await page.goto(BASE);
  await page.click('[data-page="suppliers"]');
  await page.waitForSelector("#supplierRegisterTable tbody tr");
  check("No separate coverage tab any more", (await page.locator('[data-supplier-tab="Supplier/material coverage"]').count()) === 0);
  await page.locator('#supplierRegisterTable td[data-open-supplier-overview]', { hasText: "SUP-2013" }).click();
  await page.waitForSelector("#ovParts table");
  check("Clicking the code opens Savera's overview (no parts yet)", (await page.textContent("#ovTitle")) === "SUP-2013 Savera" && (await page.textContent("#ovParts")).includes("supplies no part yet"));
  await page.click("#ovEditParts");
  await page.waitForSelector("#relListTable table");
  check("Edit parts opens Supplier/material relationship for Savera", (await page.locator(".panel-header, .panel").first().textContent()).includes("Supplier/material relationship · SUP-2013 Savera"));
  const relHeads = await page.locator("#relListTable thead th summary").allTextContents();
  check("Columns: Supplier Code, Supplier Name, T1 Part Code, T1 Part Name, SOB, Lead time, MOQ, Buffer days, Unit price, Currency, Action", JSON.stringify(relHeads) === JSON.stringify(["Supplier Code", "Supplier Name", "T1 Part Code", "T1 Part Name", "SOB", "Lead time", "MOQ", "Buffer days", "Unit price", "Currency", "Action"]), JSON.stringify(relHeads));
  const row = page.locator("#relEntryTable tbody tr").first();
  check("New row has the supplier filled in", (await row.locator('[data-pick-key="supplier"][data-pick-part="code"]').inputValue()) === "SUP-2013" && (await row.locator('[data-pick-key="supplier"][data-pick-part="name"]').inputValue()) === "Savera");
  await row.locator('[data-pick-key="material"][data-pick-part="code"]').fill("UF220BFB010");
  await page.waitForFunction(() => document.querySelector('#relEntryTable [data-row-material="material"]').value !== "");
  await row.locator('[data-f="sob_percent"]').fill("60");
  await row.locator('[data-f="lead_time_days"]').fill("15");
  await row.locator('[data-f="unit_price"]').fill("85");
  await clearToast();
  await page.click("#relSaveAll");
  await waitToast("saved");
  await page.waitForSelector('#relListTable tbody tr:has-text("UF220BFB010")');
  const listed = await page.locator('#relListTable tbody tr:has-text("UF220BFB010") td').allTextContents();
  check("Saved row listed (Savera, beam, 60%, 15 days)", listed[0] === "SUP-2013" && listed[1] === "Savera" && listed[4] === "60%" && listed[5] === "15 days", JSON.stringify(listed.slice(0, 6)));

  // 2. click the part: Supplier/material coverage
  await page.locator('#relListTable tbody tr:has-text("UF220BFB010") td[data-open-coverage]').nth(2).click();
  await page.waitForSelector("#coverageTable table");
  check("Coverage page with breadcrumb supplier › part", (await page.locator("#covCrumbs").textContent()).replace(/\s+/g, " ").trim().startsWith("SUP-2013 Savera › UF220BFB010 Base Front Beam"));
  check("Only the T1 row until materials are added", (await page.locator("#coverageTable tbody tr").count()) === 1);
  // everything is edited in the rows: no pop-up windows
  const editRow = () => page.locator("#coverageTable tr.cov-editing");
  const setIn = (name, value) => editRow().locator(`[data-c="${name}"]`).fill(String(value));
  const saveRow = async (expect) => { await clearToast(); await editRow().locator("[data-cov-save]").click(); await waitToast(expect); };
  const pickMaterial = async (code, name) => {
    await setIn("material", code);
    await page.waitForFunction((n) => document.querySelector("#coverageTable tr.cov-editing .cov-material-name")?.textContent === n, name);
  };
  const lineRow = (code) => page.locator("#coverageTable tbody tr.cov-line", { hasText: code });
  // L1: 2 kg steel per beam with its T2 supplier in the same new row: JGI 30% -> a new row for the other 70%
  await page.click('.table-toolbar [data-cov-add-line=""]');
  await editRow().waitFor();
  check("+ L1 material adds an empty row in the table (no window)", (await page.locator("#coverageTable tr.cov-editing.cov-new").count()) === 1 && await page.locator("#modalLayer").isHidden());
  await pickMaterial("FL001STEEL0.7", "STEEL0.7");
  check("Typing the material code fills its name and unit", (await editRow().locator(".cov-material-unit").textContent()) === "kg");
  await setIn("quantity", 2);
  await setIn("supplier", "SUP-2004");
  check("Typing the supplier code shows its name", (await editRow().locator(".cov-supplier-name").textContent()) === "JGI STEEL");
  await setIn("sob_percent", 30);
  await setIn("lead_time_days", 10);
  await setIn("unit_price", 70);
  await saveRow("remaining 70%");
  await page.waitForSelector("#coverageTable tr.cov-editing.cov-source");
  check("Saved at 30%: a new supplier row for the remaining 70% opens under steel", (await editRow().locator('[data-c="sob_percent"]').inputValue()) === "70" && (await editRow().locator('[data-c="supplier"]').inputValue()) === "" && (await editRow().textContent()).includes("FL001STEEL0.7"));
  await saveRow("Type the supplier code");
  check("The balance row needs its supplier (or Cancel)", true);
  await setIn("supplier", "SUP-2014");
  await setIn("lead_time_days", 20);
  await setIn("unit_price", 70);
  await saveRow("TATA Steel saved for FL001STEEL0.7");
  await page.waitForFunction(() => !document.querySelector("#coverageTable tr.cov-editing"));
  check("Steel saved with JGI 30% and TATA 70% (nothing else added)", (await page.locator("#coverageTable tr.cov-source").count()) === 2 && (await lineRow("FL001STEEL0.7").locator("td").nth(6).textContent()) === "100%");
  // L2: steel's BOM, per kg steel: 1.7 kg ore (XYZ, SOB blank = 100%), 0.6 kg coal (supplier added after)
  await lineRow("FL001STEEL0.7").locator("[data-cov-add-line]").click();
  check("+ L2 material: the new row sits under steel at L2", (await editRow().getAttribute("data-tree-depth")) === "2");
  await pickMaterial("ORE-01", "Iron ore");
  await setIn("quantity", 1.7);
  await setIn("supplier", "SUP-2030");
  await saveRow("ORE-01 added");
  await lineRow("FL001STEEL0.7").locator("[data-cov-add-line]").click();
  await pickMaterial("COAL-01", "Coal");
  await setIn("quantity", 0.6);
  await saveRow("COAL-01 added");
  await lineRow("COAL-01").waitFor();
  check("Ore saved with XYZ at 100% (SOB blank = 100%, no balance row)", (await page.locator('#coverageTable tr.cov-source:has-text("XYZ Mining")').locator("td").nth(6).textContent()) === "100%" && (await page.locator("#coverageTable tr.cov-editing").count()) === 0);
  check("A supplier used at T3 joins the T3 tier", (await api("GET", "/api/suppliers")).find((who) => who.code === "SUP-2030").tiers.includes("T3"));
  await lineRow("COAL-01").locator("[data-cov-add-source]").click();
  check("+ T3 supplier: new row with the SOB left (100)", (await editRow().locator('[data-c="sob_percent"]').inputValue()) === "100");
  await setIn("sob_percent", "");  // blank = what is left (100%)
  await setIn("supplier", "ABC Coal");  // by name
  await setIn("unit_price", 70);
  await saveRow("ABC Coal saved");
  await page.waitForSelector('#coverageTable tr.cov-source:has-text("ABC Coal")');
  check("Blank SOB saved as the SOB left (100%)", (await page.locator('#coverageTable tr.cov-source:has-text("ABC Coal")').locator("td").nth(6).textContent()) === "100%");
  // requirement 60,000 beams
  await page.fill("#covRequirement", "60000");
  const read = () => page.locator("#coverageTable tbody tr").evaluateAll((trs) => trs.map((tr) => [tr.cells[1].textContent, tr.cells[3].textContent || tr.cells[4].textContent, tr.cells[6].textContent, tr.cells[8].textContent, tr.cells[10].textContent]));
  const rows = await read();
  // worked out by hand: 60,000 x 60% = 36,000 beams; x 2 kg = 72,000 kg steel; JGI 30% = 21,600, TATA 70% = 50,400;
  // ore 72,000 x 1.7 = 122,400 kg (3.4 kg per beam); coal 72,000 x 0.6 = 43,200 kg (1.2 kg per beam)
  const expected = [["T1", "Savera", "60%", "1", "36,000 Nos"],
    ["L1 BOM", "FL001STEEL0.7", "100%", "2", "72,000 kg"], ["T2", "JGI STEEL", "30%", "0.6", "21,600 kg"], ["T2", "TATA Steel", "70%", "1.4", "50,400 kg"],
    ["L2 BOM", "ORE-01", "100%", "3.4", "122,400 kg"], ["T3", "XYZ Mining", "100%", "3.4", "122,400 kg"],
    ["L2 BOM", "COAL-01", "100%", "1.2", "43,200 kg"], ["T3", "ABC Coal", "100%", "1.2", "43,200 kg"]];
  check("Tree and requirement: T1 36,000; steel 72,000 kg (JGI 21,600 / TATA 50,400); ore 122,400 kg; coal 43,200 kg", JSON.stringify(rows) === JSON.stringify(expected), JSON.stringify(rows));
  check("Note says how many Savera makes", (await page.locator("#covShare").textContent()).includes("Savera makes 60% = 36,000"));
  const levels = await page.locator("#coverageTable tbody tr").evaluateAll((trs) => trs.map((tr) => tr.dataset.treeDepth).join(""));
  check("Levels like the Product Master tree (T1 0, L1 1, T2 2, L2 2, T3 3)", levels === "01222323", levels);
  check("Suppliers of a line shown like alternate rows (↳)", (await page.locator("#coverageTable tr.cov-source.tree-alternate").count()) === 4 && (await page.locator('#coverageTable tr.cov-source:has-text("JGI")').locator("td").nth(4).textContent()).startsWith("↳"));
  check("Each supplier has its own lead time (JGI 10, TATA 20)", (await page.locator('#coverageTable tr.cov-source:has-text("JGI")').locator("td").nth(11).textContent()) === "10 days" && (await page.locator('#coverageTable tr.cov-source:has-text("TATA")').locator("td").nth(11).textContent()) === "20 days");
  // edit JGI in its row: 40% is refused (only 30% left); 20% -> total 90% (a balance row is offered: Cancel)
  await page.locator('#coverageTable tr.cov-source:has-text("JGI")').locator("[data-cov-edit-source]").click();
  check("Edit turns the JGI row into boxes (no window)", (await editRow().locator('[data-c="sob_percent"]').inputValue()) === "30" && (await editRow().locator('[data-c="supplier"]').inputValue()) === "SUP-2004" && await page.locator("#modalLayer").isHidden());
  await setIn("sob_percent", 40);
  await saveRow("at most 30%");
  check("Too much SOB is refused with what is left", true);
  await setIn("sob_percent", 20);
  await saveRow("remaining 10%");
  await page.waitForSelector("#coverageTable tr.cov-editing.cov-new");
  await editRow().locator("[data-cov-cancel]").click();
  await page.waitForFunction(() => !document.querySelector("#coverageTable tr.cov-editing") && [...document.querySelectorAll("#coverageTable tr.cov-line")].some((tr) => tr.cells[6].textContent === "90%"));
  check("Steel SOB total 90% shown in red", await lineRow("FL001STEEL0.7").locator("td").nth(6).evaluate((td) => td.classList.contains("sob-off")));
  check("JGI requirement follows (72,000 x 20% = 14,400 kg)", (await page.locator('#coverageTable tr.cov-source:has-text("JGI")').locator("td").nth(10).textContent()) === "14,400 kg");
  // + / -
  const visible = () => page.locator("#coverageTable tbody tr").evaluateAll((trs) => trs.filter((tr) => !tr.hidden).length);
  await lineRow("FL001STEEL0.7").locator("[data-sb-toggle]").click();
  check("− on steel closes its suppliers and L2 BOM", (await visible()) === 2);
  await lineRow("FL001STEEL0.7").locator("[data-sb-toggle]").click();
  check("+ opens them again", (await visible()) === 8);
  await page.screenshot({ path: __dirname + "/supply_chain.png" });
  // Edit T1 in its row
  await page.click("[data-cov-edit-t1]");
  await setIn("lead_time_days", 16);
  await saveRow("T1 saved");
  await page.waitForFunction(() => document.querySelector("#coverageTable tr.cov-t1")?.cells[11].textContent === "16 days");
  check("Edit T1 in the row: lead time 16 days", true);
  // edit a quantity (Enter saves), Esc cancels, delete a line
  await lineRow("COAL-01").locator("[data-cov-edit-line]").click();
  await setIn("quantity", 0.5);
  await clearToast();
  await editRow().locator('[data-c="quantity"]').press("Enter");
  await page.waitForFunction(() => [...document.querySelectorAll("#coverageTable tr.cov-line")].some((tr) => tr.textContent.includes("COAL-01") && tr.cells[10].textContent === "36,000 kg"));
  check("Coal 0.5 kg per kg steel -> 36,000 kg (Enter saves)", true);
  await lineRow("COAL-01").locator("[data-cov-edit-line]").click();
  await editRow().locator('[data-c="quantity"]').press("Escape");
  check("Esc cancels the edit", (await page.locator("#coverageTable tr.cov-editing").count()) === 0);
  await lineRow("COAL-01").locator("[data-cov-delete-line]").click();
  await page.waitForFunction(() => document.querySelector("#coverageTable tbody") && !document.querySelector("#coverageTable tbody").textContent.includes("ABC Coal"));
  check("Delete a line removes it and its supplier", (await page.locator("#coverageTable tbody tr").count()) === 6);
  check("No pop-up window was used on the coverage page", await page.locator("#modalLayer").isHidden());
  check("Requirement kept after the redraw", (await page.inputValue("#covRequirement")) === "60000");
  // supplier-wise Excel: download the template (what is saved) and upload it back
  const templateHref = await page.getAttribute("#coverageTemplate", "href");
  check("Coverage template is for this supplier", /\/api\/planning\/coverage-workbook\/template\.xlsx\?supplier_id=\d+$/.test(templateHref), templateHref);
  const file = require("path").join(require("os").tmpdir(), `coverage-${Date.now()}.xlsx`);
  require("fs").writeFileSync(file, await (await page.request.get(BASE + templateHref)).body());
  await page.setInputFiles('#coverageUploadForm input[type="file"]', file);
  await clearToast();
  await page.click('#coverageUploadForm button[type="submit"]');
  await waitToast("saved");
  check("Uploading the template back: 1 part, 2 material lines, 3 supplier rows", (await toastText()).includes("1 part: 2 material lines, 3 supplier rows saved"), await toastText());
  const after = await read();
  check("Tree unchanged after the round trip", JSON.stringify(after) === JSON.stringify([["T1", "Savera", "60%", "1", "36,000 Nos"], ["L1 BOM", "FL001STEEL0.7", "90%", "2", "72,000 kg"], ["T2", "JGI STEEL", "20%", "0.4", "14,400 kg"], ["T2", "TATA Steel", "70%", "1.4", "50,400 kg"], ["L2 BOM", "ORE-01", "100%", "3.4", "122,400 kg"], ["T3", "XYZ Mining", "100%", "3.4", "122,400 kg"]]), JSON.stringify(after));
  // the breadcrumb's supplier opens its overview: parts (SOB, FG stock), BOM tree, sub-supplier list
  await page.locator('#covCrumbs a[data-open-supplier-overview]').click();
  await page.waitForSelector('#ovParts tbody tr:has-text("UF220BFB010")');
  const ovPart = await page.locator('#ovParts tbody tr:has-text("UF220BFB010") td').allTextContents();
  check("Overview 1: part, SOB 60%, rate, no plan -> requirement 0", ovPart[0] === "UF220BFB010" && ovPart[2] === "60%" && ovPart[3].includes("85") && ovPart[5] === "0", JSON.stringify(ovPart));
  check("Overview plan note when no plan", (await page.textContent("#ovPlan")).includes("no plan uploaded"));
  const ovTree = await page.locator("#ovTree tbody tr").evaluateAll((trs) => trs.map((tr) => [tr.cells[1].textContent, tr.cells[2].textContent, tr.cells[5].textContent, tr.cells[6].textContent]));
  check("Overview 2-3: part BOM with sub-supplier SOB", JSON.stringify(ovTree) === JSON.stringify([["T1 part", "UF220BFB010", "Savera", "60%"], ["L1 BOM", "FL001STEEL0.7", "", "90%"], ["T2 supplier", "↳ FL001STEEL0.7", "JGI STEEL", "20%"], ["T2 supplier", "↳ FL001STEEL0.7", "TATA Steel", "70%"], ["L2 BOM", "ORE-01", "", "100%"], ["T3 supplier", "↳ ORE-01", "XYZ Mining", "100%"]]), JSON.stringify(ovTree));
  check("No T2 / T3 tag under material codes (overview and coverage trees)", (await page.locator("#ovTree td small").count()) === 0);
  await page.click("#ovCloseAll");
  check("Overview tree closes", (await page.locator("#ovTree tbody tr:not([hidden])").count()) === 1);
  const ovHeads = await page.locator("#ovSubs thead th summary").allTextContents();
  check("Overview 4-6: one line per sub-supplier and material, no For part column", JSON.stringify(ovHeads) === JSON.stringify(["Sub-supplier code", "Sub-supplier name", "Tier", "Material code", "Material name", "Blended SOB", "Rate", "Blended requirement", "Stock", "Unit", "Action"]), JSON.stringify(ovHeads));
  check("Parts table has no Available column", !(await page.locator("#ovParts thead th summary").allTextContents()).includes("Available"));
  const readSubs = () => page.locator("#ovSubs tbody tr").evaluateAll((trs) => trs.map((tr) => [tr.cells[0].textContent, tr.cells[2].textContent, tr.cells[3].textContent, tr.querySelector('[data-ov-f="sob_percent"]').value, tr.querySelector('[data-ov-f="unit_price"]').value]));
  const ovSubs = await readSubs();
  // no plan -> no requirement: the blended SOB is the plain average of the parts' SOB
  check("Sub-supplier lines: SOB and rate in boxes", JSON.stringify(ovSubs) === JSON.stringify([["SUP-2004", "T2", "FL001STEEL0.7", "20", "70"], ["SUP-2014", "T2", "FL001STEEL0.7", "70", "70"], ["SUP-2030", "T3", "ORE-01", "100", "0"]]), JSON.stringify(ovSubs));
  await page.locator('#ovSubs tr[data-ov-material="FL001STEEL0.7"]').nth(1).locator('[data-ov-f="sob_percent"]').fill("80");
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.locator('#ovSubs tr[data-ov-material="FL001STEEL0.7"]').nth(1).locator("[data-ov-save]").click();
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("saved in 1 part"));
  await page.waitForFunction(() => document.querySelectorAll("#ovSubs tbody tr")[1]?.querySelector('[data-ov-f="sob_percent"]').value === "80");
  await page.waitForFunction(() => [...document.querySelectorAll("#ovTree tr.st-source")].some((tr) => tr.textContent.includes("TATA Steel") && tr.cells[6].textContent === "80%"), null, { timeout: 5000 }).catch(() => {});
  const tataTree = await page.locator('#ovTree tr.st-source:has-text("TATA Steel")').locator("td").allTextContents();
  check("Edit TATA steel SOB to 80% from the list: the BOM tree follows", tataTree[6] === "80%", JSON.stringify(tataTree));
  // stock typed in the boxes: FG stock of the beam (parts table) and TATA's stock in the BOM tree; Enter saves
  const fgBox = page.locator("#ovParts input.ov-stock").first();
  await fgBox.fill("9000");
  await clearToast();
  await fgBox.press("Enter");
  await waitToast("Stock 9,000 saved");
  await page.waitForFunction(() => document.querySelector("#ovTree tr.st-part input.ov-stock")?.value === "9000");
  check("FG stock typed in the parts table is saved (also shown in the tree)", (await page.locator("#ovParts input.ov-stock").first().inputValue()) === "9000");
  const tataBox = page.locator('#ovTree input.ov-stock[data-stock-supplier="SUP-2014"]');
  await tataBox.fill("3000");
  await clearToast();
  await tataBox.press("Enter");
  await waitToast("Stock 3,000 saved");
  await page.waitForFunction(() => [...document.querySelectorAll("#ovSubs tbody tr")].some((tr) => tr.cells[0].textContent === "SUP-2014" && tr.cells[8].textContent === "3,000"));
  check("T2 stock typed in the tree is saved; the sub-supplier list shows it", (await page.locator('#ovTree input.ov-stock[data-stock-supplier="SUP-2014"]').inputValue()) === "3000");
  const stockList = (await api("GET", `/api/planning/supplier-stock/${new Date().toISOString().slice(0, 7)}?workspace=trial`)).items.map((r) => [r.supplier_code, r.material_code, r.quantity]);
  check("Both are in the month's stock list (Upload supplier stock)", JSON.stringify(stockList.filter((r) => r[1] === "UF220BFB010" || r[0] === "SUP-2014")) === JSON.stringify([["SUP-2013", "UF220BFB010", 9000], ["SUP-2014", "FL001STEEL0.7", 3000]]), JSON.stringify(stockList));
  await page.click("#ovEditParts");
  await page.waitForSelector("#relListTable tbody tr");
  check("Breadcrumb goes back to Savera's relationships", (await page.locator("#relListTable tbody tr").count()) === 1);
  check("Relationship page: Download Excel / PDF of the list shown (Savera)", /^\/api\/planning\/supplier-relationships\.xlsx\?search=&supplier_id=\d+$/.test(await page.getAttribute("#relExcel", "href")) && (await page.request.get(BASE + await page.getAttribute("#relExcel", "href"))).status() === 200 && (await page.request.get(BASE + await page.getAttribute("#relPdf", "href"))).status() === 200, await page.getAttribute("#relExcel", "href"));
  check("Relationship template has Savera's rows", /supplier-relationships\/template\.xlsx\?supplier_id=\d+/.test(await page.getAttribute("#relTemplate", "href")) && (await page.textContent("#relTemplate")).includes("SUP-2013"));
  await page.click("[data-rel-all]");
  await page.waitForFunction(() => !document.querySelector(".crumbs") && document.querySelector("#relListTable"));
  check("Show all suppliers: no breadcrumb repeating the Supplier register tab", true);

  // Upload supplier stock: the whole supply chain as a tree, closed to the T1 suppliers at first
  await api("POST", `/api/planning/supplier-stock/${new Date().toISOString().slice(0, 7)}?workspace=trial`, { rows: [{ supplier_part_id: (await api("GET", "/api/planning/supplier-parts?tier=T1")).items[0].id, quantity: 500, available_quantity: 400 }] });
  await page.click('[data-supplier-tab="Upload supplier stock"]');
  await page.waitForSelector("#supplyTreeTable table");
  const shownTree = () => page.locator("#supplyTreeTable tbody tr").evaluateAll((trs) => trs.filter((tr) => !tr.hidden).map((tr) => [tr.cells[1].textContent, tr.cells[3].textContent || tr.cells[4].textContent]));
  check("Tree starts closed: only the T1 suppliers", JSON.stringify(await shownTree()) === JSON.stringify([["T1 supplier", "Savera"]]), JSON.stringify(await shownTree()));
  await page.locator('#supplyTreeTable tr.st-supplier [data-sb-toggle]').click();
  check("+ on Savera shows its part, and the whole branch below", (await shownTree()).length === 7);
  await page.click("#supplyTreeClose");
  check("Close all", (await shownTree()).length === 1);
  await page.click("#supplyTreeOpen");
  const all = await shownTree();
  check("Open all: supplier -> part -> L1 steel -> T2 JGI / TATA -> L2 ore -> T3 XYZ", JSON.stringify(all) === JSON.stringify([["T1 supplier", "Savera"], ["T1 part", "Savera"], ["L1 material", "FL001STEEL0.7"], ["T2 supplier", "JGI STEEL"], ["T2 supplier", "TATA Steel"], ["L2 material", "ORE-01"], ["T3 supplier", "XYZ Mining"]]), JSON.stringify(all));
  check("The part row shows the month's stock (500 / 400)", await page.locator("#supplyTreeTable tr.st-part").evaluate((tr) => tr.cells[9].textContent === "500" && tr.cells[10].textContent === "400"));
  await page.locator('#supplyTreeTable tr.st-line:has-text("FL001STEEL0.7") [data-sb-toggle]').click();
  check("− on steel closes its suppliers and L2 below", (await shownTree()).length === 3);
  await page.fill("#supplyTreeSearch", "nothing-like-this");
  await page.waitForFunction(() => document.querySelector("#supplyTreeTable")?.textContent.includes("No supplier or material matches"));
  await page.fill("#supplyTreeSearch", "XYZ");
  await page.waitForFunction(() => document.querySelectorAll("#supplyTreeTable tbody tr:not([hidden])").length === 7);
  check("Search XYZ finds Savera's tree, opened", true);
  const excelHref = await page.getAttribute("#supplyTreeExcel", "href");
  const month = await page.inputValue("#stockMonth");
  check("Download Excel follows the month and workspace", excelHref === `/api/planning/supply-tree.xlsx?report_month=${month}&workspace=trial`, excelHref);
  const download = await page.request.get(BASE + excelHref);
  check("Excel download works", download.status() === 200 && download.headers()["content-type"].includes("spreadsheetml") && download.headers()["content-disposition"].includes(`supply-chain-tree-${month}.xlsx`));
  check("PDF download works", (await page.request.get(BASE + (await page.getAttribute("#supplyTreePdf", "href")))).status() === 200);
  // Upload Excel with the template: the tree, stock typed on the part row (filled here with openpyxl from the project's .venv)
  const templateUrl = await page.getAttribute("#stockTemplate", "href");
  check("Stock template is the month's tree", templateUrl === `/api/planning/supply-tree.xlsx?report_month=${month}&workspace=trial&template=true`, templateUrl);
  const os = require("os"), path = require("path"), fs = require("fs"), { spawnSync } = require("child_process");
  const stockFile = path.join(os.tmpdir(), `stock-${Date.now()}.xlsx`);
  fs.writeFileSync(stockFile, await (await page.request.get(BASE + templateUrl)).body());
  const python = process.env.PYTHON || path.join(__dirname, "..", "..", ".venv", "bin", "python");
  const filled = spawnSync(python, ["-c", `import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1]); ws = wb.active
for r in range(4, ws.max_row + 1):
    if (ws.cell(r, 1).value or "").strip() == "T1 part":
        ws.cell(r, 9).value, ws.cell(r, 10).value, ws.cell(r, 11).value = 800, 750, "2027-02"
    if (ws.cell(r, 1).value or "").strip() == "T2 supplier" and ws.cell(r, 2).value == "SUP-2014":
        ws.cell(r, 9).value, ws.cell(r, 10).value = 12000, 11000
wb.save(sys.argv[1])`, stockFile]);
  check("Template filled", filled.status === 0, String(filled.stderr));
  await page.setInputFiles('#supplierStockForm input[type="file"]', stockFile);
  await clearToast();
  await page.click('#supplierStockForm button[type="submit"]');
  await page.waitForFunction(() => /row|Row|stock/.test(document.querySelector("#toast")?.textContent || ""));
  await page.waitForFunction(() => document.querySelector("#supplyTreeTable tr.st-part")?.cells[9].textContent === "800");
  check("Stock typed on the T2 supplier row (TATA) shows on that row", await page.locator('#supplyTreeTable tr.st-source:has-text("TATA Steel")').evaluate((tr) => tr.cells[9].textContent === "12,000" && tr.cells[10].textContent === "11,000"));
  check("Uploaded stock shows on the part row (800 / 750 / Feb 2027)", await page.locator("#supplyTreeTable tr.st-part").evaluate((tr) => tr.cells[10].textContent === "750" && tr.cells[11].textContent.includes("2027")));

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
