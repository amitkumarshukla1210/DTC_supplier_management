// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/row_click.js   (needs Node + Playwright).
// Checks: row highlight and row click stop before the Edit / Delete columns on all
// five BOM tabs; the Edit / Delete buttons still work; row click still works elsewhere.
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  let dialogs = 0;
  page.on("dialog", async (d) => { dialogs += 1; await d.dismiss(); });
  const api = async (method, path, data) => {
    const r = await page.request.fetch(BASE + path, { method, data, headers: { "content-type": "application/json" } });
    if (r.status() >= 300) throw new Error(`${method} ${path} -> ${r.status()} ${await r.text()}`);
    return r.json();
  };
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  const mat = async (code, name, uom) => (await api("POST", "/api/planning/materials", { code, name, uom, category: "Raw material" })).id;
  const motor = await mat("RC-MOTOR", "Motor", "Nos");
  const steel = await mat("RC-STEEL", "Steel", "kg");
  const wire = await mat("RC-WIRE", "Wire", "kg");
  const model = await api("POST", "/api/planning/models", { code: "RC-12000", name: "Eva 8014" });
  const model2 = await api("POST", "/api/planning/models", { code: "RC-13000", name: "Serena 8014" });
  await api("POST", "/api/planning/products", { code: "RC-1000", name: "FL", model_ids: [model.id, model2.id] });
  await api("POST", `/api/planning/model-bom/${model2.id}`, { material_id: steel, quantity: 3, uom: "kg" });
  await api("POST", `/api/planning/model-bom/${model.id}`, { material_id: motor, quantity: 1, uom: "Nos" });
  await api("POST", `/api/planning/part-bom/${motor}/new`, { tab: "Level 2 Part BOM", lines: [{ child_item_id: steel, quantity: 2 }] });
  await api("POST", `/api/planning/part-bom/${steel}/new`, { tab: "Level 3 Part BOM", lines: [{ child_item_id: wire, quantity: 0.5 }] });

  const bg = (loc) => loc.evaluate((e) => getComputedStyle(e).backgroundColor);
  const HIGHLIGHT = "rgb(231, 242, 250)";
  const activeTab = () => page.locator(".tab-button.active").textContent();
  const modalOpen = () => page.locator("#modalLayer").evaluate((e) => !e.hidden);
  const closeModal = () => page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  // Click inside the cell but away from the button text (right-hand padding)
  const clickCellPadding = async (cell) => {
    const box = await cell.boundingBox();
    await page.mouse.click(box.x + box.width - 4, box.y + box.height / 2);
    await page.waitForTimeout(400);
  };

  // Header text above the row's first filled data cell (after the toggle column)
  // Text of the row's cell under the header with this name
  const cellUnder = async (tableSel, row, headerName) => {
    const names = await page.locator(`${tableSel} > .table-wrap > table > thead th`).evaluateAll((ths) => ths.map((th) => th.querySelector("summary")?.textContent.trim() || ""));
    const index = names.indexOf(headerName);
    return row.evaluate((tr, i) => tr.cells[i]?.textContent.trim(), index);
  };
  // Each row's own (deepest) code: last filled code column of the level columns
  const rowCodes = (tableSel, levelCount) => page.locator(`${tableSel} tbody tr`).evaluateAll((trs, n) => trs.map((tr) => {
    for (let k = n - 1; k >= 0; k -= 1) { const text = tr.cells[1 + 2 * k]?.textContent.trim(); if (text) return text; }
    return "";
  }), levelCount);
  const isBold = (row) => row.locator("td").nth(1).evaluate((td) => Number(getComputedStyle(td).fontWeight) >= 600);

  async function checkRow(tabName, row, opts) {
    const cells = row.locator(":scope > td");
    const n = await cells.count();
    await cells.nth(opts.dataCell).hover();
    if (opts.noDeeperLevel) check(`${tabName}: line with no deeper BOM is not highlighted (not clickable)`, (await bg(cells.nth(opts.dataCell))) !== HIGHLIGHT);
    else check(`${tabName}: hover highlights data cells`, (await bg(cells.nth(opts.dataCell))) === HIGHLIGHT);
    check(`${tabName}: hover does NOT highlight Edit / Delete cells`, (await bg(cells.nth(n - 2))) !== HIGHLIGHT && (await bg(cells.nth(n - 1))) !== HIGHLIGHT);
    for (const [label, index] of [["Edit", n - 2], ["Delete", n - 1]]) {
      const before = await activeTab();
      const dialogsBefore = dialogs;
      await clickCellPadding(cells.nth(index));
      check(`${tabName}: click in ${label} cell (not on button) does nothing`, (await activeTab()) === before && !(await modalOpen()) && dialogs === dialogsBefore);
    }
    await cells.nth(n - 2).locator("button").click();
    await page.waitForTimeout(400);
    check(`${tabName}: Edit button opens the edit form`, await modalOpen());
    await closeModal();
    const dialogsBefore = dialogs;
    await cells.nth(n - 1).locator("button").click();
    await page.waitForTimeout(400);
    check(`${tabName}: Delete button asks for confirmation`, dialogs === dialogsBefore + 1);
  }

  await page.goto(BASE);
  await page.click('[data-page="master"]');

  // Product Master
  await page.click('[data-bom-tab="Product Master"]');
  const productRow = page.locator("#productMasterTable tbody tr[data-product-model-row]").first();
  await productRow.waitFor();
  await productRow.locator("[data-tree-expand]").click();
  const pmL1 = page.locator("#productMasterTable tr[data-bom-level1-row]").first();
  await pmL1.waitFor();
  check("Product Master: one table, no nested header", (await page.locator("#productMasterTable thead").count()) === 1);
  check("Product Master: L1 part sits under 'L1 part code'", (await cellUnder("#productMasterTable", pmL1, "L1 part code")) === "RC-MOTOR");
  check("Product Master: L1 row repeats product and model", (await cellUnder("#productMasterTable", pmL1, "Product code")) === "RC-1000" && (await cellUnder("#productMasterTable", pmL1, "Model name")) === "Eva 8014");
  await pmL1.locator("[data-tree-expand]").click();
  const pmL2 = page.locator("#productMasterTable tr[data-bom-level2-row]").first();
  await pmL2.waitFor();
  check("Product Master: L2 part sits under 'L2 part code'", (await cellUnder("#productMasterTable", pmL2, "L2 part code")) === "RC-STEEL");
  check("Product Master: L2 row repeats L1 part", (await cellUnder("#productMasterTable", pmL2, "L1 part code")) === "RC-MOTOR" && (await cellUnder("#productMasterTable", pmL2, "L1 part name")) === "Motor");
  await pmL2.locator("[data-tree-expand]").click();
  await page.waitForFunction(() => document.querySelectorAll("#productMasterTable tr[data-bom-level2-row]").length === 2);
  const pmL3 = page.locator("#productMasterTable tr[data-bom-level2-row]").nth(1);
  check("Product Master: L3 part sits under 'L3 part code'", (await cellUnder("#productMasterTable", pmL3, "L3 part code")) === "RC-WIRE");
  check("Product Master: L3 row repeats the full path", (await cellUnder("#productMasterTable", pmL3, "Product code")) === "RC-1000" && (await cellUnder("#productMasterTable", pmL3, "L1 part code")) === "RC-MOTOR" && (await cellUnder("#productMasterTable", pmL3, "L2 part code")) === "RC-STEEL");
  const order = await rowCodes("#productMasterTable", 5);
  check("Product Master: tree order Model > L1 > L2 > L3", JSON.stringify(order) === JSON.stringify(["RC-12000", "RC-MOTOR", "RC-STEEL", "RC-WIRE", "RC-13000"]), JSON.stringify(order));
  check("Product Master: rows with lines under them are bold, the last level is not", (await isBold(productRow)) && (await isBold(pmL1)) && (await isBold(pmL2)) && !(await isBold(pmL3)));
  await page.screenshot({ path: __dirname + "/look_product.png" });
  await productRow.locator("[data-tree-expand]").click();
  check("Product Master: collapsing the product removes all levels below it", (await page.locator("#productMasterTable tbody tr").count()) === 2);
  await checkRow("Product Master", productRow, { dataCell: 1 });
  await productRow.locator("td").nth(1).click();
  await page.waitForTimeout(500);
  check("Product Master: clicking product code still opens Level 1 BOM", (await activeTab()) === "Model BOM Level 1");

  // Model BOM Level 1
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  const l1Row = page.locator("#masterTable tbody tr[data-bom-level1-row]").first();
  await l1Row.waitFor();
  await checkRow("Model BOM Level 1", l1Row, { dataCell: 4 });

  // Level 2: parent row + expanded line
  await page.click('[data-bom-tab="Level 2 Part BOM"]');
  const group = page.locator("#masterTable tr[data-tree-group]").first();
  await group.waitFor();
  await group.locator("[data-tree-expand]").click();
  const l2Row = page.locator("#masterTable tr[data-bom-level2-row]").first();
  await l2Row.waitFor();
  check("Level 2: one table, no nested header", (await page.locator("#masterTable thead").count()) === 1);
  check("Level 2: parent sits under 'L1 part code'", (await cellUnder("#masterTable", group, "L1 part code")) === "RC-MOTOR");
  check("Level 2: line sits under 'L2 part code' and repeats its L1 parent", (await cellUnder("#masterTable", l2Row, "L2 part code")) === "RC-STEEL" && (await cellUnder("#masterTable", l2Row, "L1 part code")) === "RC-MOTOR");
  await page.screenshot({ path: __dirname + "/look_level2.png" });
  await checkRow("Level 2 Part BOM", l2Row, { dataCell: 3 });
  const groupCells = group.locator(":scope > td");
  await clickCellPadding(groupCells.last());
  const l2Count = await page.locator("#masterTable tr[data-bom-level2-row]").count();
  check("Level 2: click in parent row's Delete column does not collapse it", l2Count === 1, `rows=${l2Count}`);
  await groupCells.nth(1).click();
  check("Level 2: click on parent code still collapses it", (await page.locator("#masterTable tr[data-bom-level2-row]").count()) === 0);

  // Level 3
  await page.click('[data-bom-tab="Level 3 Part BOM"]');
  const g3 = page.locator("#masterTable tr[data-tree-group]").first();
  await g3.waitFor();
  await g3.locator("[data-tree-expand]").click();
  const l3Row = page.locator("#masterTable tr[data-bom-level2-row]").first();
  await l3Row.waitFor();
  await checkRow("Level 3 Part BOM", l3Row, { dataCell: 3, noDeeperLevel: true });

  // Material Master
  await page.click('[data-bom-tab="Material Master"]');
  const matRow = page.locator('#materialMasterTable tbody tr[data-edit-material]').first();
  await matRow.waitFor();
  await checkRow("Material Master", matRow, { dataCell: 1 });
  await matRow.locator("td").nth(1).click();
  await page.waitForTimeout(400);
  check("Material Master: clicking material name still opens edit form", await modalOpen());
  await closeModal();
  await page.mouse.move(5, 5);
  await page.screenshot({ path: __dirname + "/look_material.png" });

  // Download buttons on the BOM tabs give the tree file of that tab (and its selection)
  for (const [tab, file] of [["Product Master", "product-master-bom-tree.xlsx"], ["Level 2 Part BOM", "level-2-part-bom-tree.xlsx"]]) {
    await page.click(`[data-bom-tab="${tab}"]`);
    await page.waitForSelector(`.tab-button.active[data-bom-tab="${tab}"]`);
    const [download] = await Promise.all([page.waitForEvent("download"), page.click('[data-bom-tree-export="xlsx"]')]);
    check(`${tab}: Download Excel gives the BOM tree file`, download.suggestedFilename() === file && download.url().includes("/api/planning/bom-tree/xlsx"), download.suggestedFilename());
  }
  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
