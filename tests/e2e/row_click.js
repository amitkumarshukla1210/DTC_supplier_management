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
  await api("POST", "/api/planning/products", { code: "RC-1000", name: "FL", model_ids: [model.id] });
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
  await page.locator(".bom-tree-row tr[data-bom-level1-row]").first().waitFor();
  check("Product Master: expanded Level 1 shows no repeated table header", await page.locator(".bom-tree-row thead").first().isHidden());
  await page.screenshot({ path: __dirname + "/look_product.png" });
  await productRow.locator("[data-tree-expand]").click();
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
  const l2Row = page.locator(".bom-tree-row tr[data-bom-level2-row]").first();
  await l2Row.waitFor();
  check("Level 2: expanded lines show no repeated table header", await page.locator(".bom-tree-row thead").first().isHidden());
  const mainCols = await page.locator("#masterTable > .table-wrap > table > thead th").evaluateAll((ths) => ths.map((t) => Math.round(t.getBoundingClientRect().left)));
  const nestedCols = await l2Row.locator(":scope > td").evaluateAll((tds) => tds.map((t) => Math.round(t.getBoundingClientRect().left)));
  const offBy = Math.max(...mainCols.slice(1).map((x, i) => Math.abs(x - nestedCols[i + 1])));
  check("Level 2: expanded lines line up with the main header columns", offBy <= 4, `max offset ${offBy}px`);
  await page.screenshot({ path: __dirname + "/look_level2.png" });
  await checkRow("Level 2 Part BOM", l2Row, { dataCell: 3 });
  const groupCells = group.locator(":scope > td");
  await clickCellPadding(groupCells.nth(9));
  check("Level 2: click in parent row's Delete column does not collapse it", (await page.locator(".bom-tree-row").count()) === 1);
  await groupCells.nth(1).click();
  check("Level 2: click on parent code still collapses it", (await page.locator(".bom-tree-row").count()) === 0);

  // Level 3
  await page.click('[data-bom-tab="Level 3 Part BOM"]');
  const g3 = page.locator("#masterTable tr[data-tree-group]").first();
  await g3.waitFor();
  await g3.locator("[data-tree-expand]").click();
  const l3Row = page.locator(".bom-tree-row tr[data-bom-level2-row]").first();
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

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
