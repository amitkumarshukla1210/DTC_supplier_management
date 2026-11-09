// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/column_resize.js   (needs Node + Playwright).
// Checks on the five master-data tabs: dragging a header edge widens only that column,
// the width survives a tab change and a page reload, double-click resets it;
// Level 2 / 3 first (+) column is at least 96px wide.
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1800, height: 900 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const api = async (method, path, data) => {
    const r = await page.request.fetch(BASE + path, { method, data, headers: { "content-type": "application/json" } });
    if (r.status() >= 300) throw new Error(`${method} ${path} -> ${r.status()} ${await r.text()}`);
    return r.json();
  };
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  const mat = async (code, name, uom) => (await api("POST", "/api/planning/materials", { code, name, uom, category: "Raw material" })).id;
  const motor = await mat("CR-MOTOR", "Motor", "Nos");
  const steel = await mat("CR-STEEL", "Steel", "kg");
  const wire = await mat("CR-WIRE", "Wire", "kg");
  const model = await api("POST", "/api/planning/models", { code: "CR-12000", name: "Eva 8014" });
  await api("POST", "/api/planning/products", { code: "CR-1000", name: "FL", model_ids: [model.id] });
  await api("POST", `/api/planning/model-bom/${model.id}`, { material_id: motor, quantity: 1, uom: "Nos" });
  await api("POST", `/api/planning/part-bom/${motor}/new`, { tab: "Level 2 Part BOM", lines: [{ child_item_id: steel, quantity: 2 }] });
  await api("POST", `/api/planning/part-bom/${steel}/new`, { tab: "Level 3 Part BOM", lines: [{ child_item_id: wire, quantity: 0.5 }] });

  const tabs = [
    ["Product Master", "#productMasterTable"],
    ["Model BOM Level 1", "#masterTable"],
    ["Level 2 Part BOM", "#masterTable"],
    ["Level 3 Part BOM", "#masterTable"],
    ["Material Master", "#materialMasterTable"],
  ];
  const headerWidths = (sel) => page.locator(`${sel} > .table-wrap > table > thead th`).evaluateAll((ths) => ths.map((t) => Math.round(t.getBoundingClientRect().width)));
  const openTab = async (tab, sel) => {
    await page.click(`[data-bom-tab="${tab}"]`);
    await page.waitForSelector(`.tab-button.active[data-bom-tab="${tab}"]`);
    await page.waitForSelector(`${sel} > .table-wrap > table > tbody tr`);
    await page.waitForTimeout(300);
  };

  await page.goto(BASE);
  await page.click('[data-page="master"]');
  const COL = 2; // second data column (after the first)
  for (const [tab, sel] of tabs) {
    await openTab(tab, sel);
    if (tab.startsWith("Level")) {
      const first = (await headerWidths(sel))[0];
      check(`${tab}: first (+) column is wide (>= 96px)`, first >= 96, `${first}px`);
    }
    const before = await headerWidths(sel);
    const header = page.locator(`${sel} > .table-wrap > table > thead th`).nth(COL);
    const box = await header.boundingBox();
    await header.hover();
    const handle = header.locator(".column-resizer");
    check(`${tab}: drag handle visible on header hover`, (await handle.evaluate((e) => getComputedStyle(e).backgroundImage)) !== "none");
    const hb = await handle.boundingBox();
    await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2);
    await page.mouse.down();
    await page.mouse.move(hb.x + hb.width / 2 + 80, hb.y + hb.height / 2, { steps: 8 });
    await page.mouse.up();
    const after = await headerWidths(sel);
    check(`${tab}: dragged column grew by ~80px`, Math.abs(after[COL] - before[COL] - 80) <= 6, `${before[COL]} -> ${after[COL]}`);
    const others = after.map((w, i) => (i === COL ? 0 : Math.abs(w - before[i])));
    check(`${tab}: other columns keep their width`, Math.max(...others) <= 3, `max change ${Math.max(...others)}px`);
    await page.click(`[data-bom-tab="${tab === "Product Master" ? "Material Master" : "Product Master"}"]`);
    await openTab(tab, sel);
    const back = await headerWidths(sel);
    check(`${tab}: width kept after changing tab`, Math.abs(back[COL] - after[COL]) <= 3, `${back[COL]}px`);
    if (box) void box;
  }

  await page.reload();
  await page.click('[data-page="master"]');
  await openTab("Material Master", "#materialMasterTable");
  const reloaded = await headerWidths("#materialMasterTable");
  await page.locator("#materialMasterTable > .table-wrap > table > thead th").nth(COL).locator(".column-resizer").dblclick();
  const reset = await headerWidths("#materialMasterTable");
  check("Material Master: width kept after page reload", reloaded[COL] >= 150, `${reloaded[COL]}px`);
  check("Double-click on a column edge resets widths", reset[COL] < reloaded[COL] - 40, `${reloaded[COL]} -> ${reset[COL]}`);

  // Expanded tree rows follow the resized column
  await openTab("Product Master", "#productMasterTable");
  await page.locator("#productMasterTable [data-tree-expand]").first().click();
  await page.waitForSelector("#productMasterTable tr[data-bom-level1-row]");
  const head = await headerWidths("#productMasterTable");
  const childCell = await page.locator("#productMasterTable tr[data-bom-level1-row] td").nth(COL).evaluate((td) => Math.round(td.getBoundingClientRect().width));
  check("Product Master: expanded rows follow the resized column", Math.abs(childCell - head[COL]) <= 2, `${head[COL]} vs ${childCell}`);
  await page.screenshot({ path: __dirname + "/column_resize.png" });

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
