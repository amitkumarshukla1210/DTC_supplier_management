// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/bom_ui.js   (needs Node + Playwright). Screenshots/results are written next to this file.
// Browser test: L2/L3 dropdown Apply visible, parents-with-BOM only, collapsed tree,
// inline New BOM, Category "+ Add new". Expected values come from the data set up here.
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const api = async (method, path, data) => {
    const r = await page.request.fetch(BASE + path, { method, data, headers: { "content-type": "application/json" } });
    if (r.status() >= 300) throw new Error(`${method} ${path} -> ${r.status()} ${await r.text()}`);
    return r.json();
  };
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  const mat = async (code, name, uom, category = "Raw material") => (await api("POST", "/api/planning/materials", { code, name, uom, category })).id;
  const m1 = await mat("FL001MOTOR1", "MOTOR1", "Nos", "electrical");
  const m2 = await mat("FL001MOTOR2", "MOTOR2", "Nos", "electrical");
  const steel = await mat("FL001STEEL0.7", "STEEL0.7", "kg");
  const scrap = await mat("SCRAP001", "Steel Scrap", "kg");
  const cu = await mat("CU01", "Copper wire", "kg");
  const model = await api("POST", "/api/planning/models", { code: "FL001", name: "Front loader" });
  for (const id of [m1, m2]) await api("POST", `/api/planning/model-bom/${model.id}`, { material_id: id, quantity: 1, uom: "Nos" });
  // 5 more Level 1 parts with a BOM -> long dropdown list (like the real data)
  const extra = [];
  for (let i = 1; i <= 5; i += 1) extra.push(await mat(`FL001P${i}`, `PART${i}`, "Nos"));
  for (const id of extra) {
    await api("POST", `/api/planning/model-bom/${model.id}`, { material_id: id, quantity: 1, uom: "Nos" });
    await api("POST", `/api/planning/part-bom/${id}/new`, { tab: "Level 2 Part BOM", lines: [{ child_item_id: steel, quantity: 1 }] });
  }
  await api("POST", `/api/planning/part-bom/${m1}/new`, { tab: "Level 2 Part BOM", lines: [{ child_item_id: steel, quantity: 2 }, { child_item_id: scrap, quantity: -0.3 }] });

  await page.goto(BASE);
  await page.click('[data-page="master"]');
  await page.click('[data-bom-tab="Level 2 Part BOM"]');
  await page.waitForSelector("[data-tree-group]");

  // Point 1: open the dropdown; Apply must be visible and be the element actually under the cursor
  await page.click("#bomParentPickerSummary");
  const apply = page.locator("[data-parent-apply]");
  const box = await apply.boundingBox();
  const hit = box && await page.evaluate(([x, y]) => document.elementFromPoint(x, y)?.hasAttribute("data-parent-apply"), [box.x + box.width / 2, box.y + box.height / 2]);
  check("L2 dropdown: Apply button visible and clickable", hit, box ? `at y=${Math.round(box.y)}` : "no box");
  // Point 3: dropdown lists only parents that have a BOM
  const options = await page.locator("[data-parent-option]").evaluateAll((els) => els.map((e) => e.closest("label").textContent.trim()));
  check("L2 dropdown lists only parents with BOM (MOTOR1 + 5 parts), not MOTOR2", options.length === 6 && options.some((o) => o.includes("FL001MOTOR1")) && !options.some((o) => o.includes("FL001MOTOR2")), JSON.stringify(options));
  const boxes = page.locator("[data-parent-option]");
  for (let i = 0; i < await boxes.count(); i += 1) {
    const label = await boxes.nth(i).evaluate((e) => e.closest("label").textContent);
    await boxes.nth(i).setChecked(label.includes("FL001MOTOR1"));
  }
  await apply.click();
  await page.waitForSelector("[data-tree-group]");
  check("Apply closes the dropdown", !(await page.locator("#bomParentPicker").evaluate((d) => d.open)));
  // Short table (1 row) + long list: Apply must still be reachable (the reported bug)
  await page.click("#bomParentPickerSummary");
  const box2 = await apply.boundingBox();
  const hit2 = box2 && await page.evaluate(([x, y]) => document.elementFromPoint(x, y)?.hasAttribute("data-parent-apply"), [box2.x + box2.width / 2, box2.y + box2.height / 2]);
  check("Apply reachable with long list over a short table", hit2, box2 ? `at y=${Math.round(box2.y)}` : "no box");
  await page.screenshot({ path: __dirname + "/dropdown_open.png" });
  await page.click("#bomParentPickerSummary");

  // Point 3: collapsed tree, + opens, − closes
  const groups = page.locator("#masterTable > .table-wrap > table > tbody > tr[data-tree-group]");
  check("Tree starts collapsed: 1 parent row, no child rows", (await groups.count()) === 1 && (await page.locator("#masterTable tr[data-bom-level2-row]").count()) === 0);
  await groups.first().locator("[data-tree-expand]").click();
  await page.waitForSelector("#masterTable tr[data-bom-level2-row]");
  const childCodes = await page.locator("#masterTable tr[data-bom-level2-row] td:nth-child(4)").allTextContents();
  check("'+' shows MOTOR1 lines STEEL0.7 and SCRAP001", childCodes.join(",") === "FL001STEEL0.7,SCRAP001", childCodes.join(","));
  check("Toggle turns into '−'", (await groups.first().locator("[data-tree-expand]").textContent()).trim() === "−");
  await page.screenshot({ path: __dirname + "/l2_tree.png" });
  await groups.first().locator("td:nth-child(2)").click();
  check("Clicking the parent row collapses it again", (await page.locator("#masterTable tr[data-bom-level2-row]").count()) === 0);

  // Point 4: New BOM for MOTOR2 with two lines picked from Material Master
  await page.click('[data-action="new-part-bom"]');
  const parentOpts = await page.locator("#newBomParent option").allTextContents();
  check("New BOM parent list = Level 1 parts without BOM (MOTOR2 only)", parentOpts.length === 2 && parentOpts[1].includes("FL001MOTOR2"), JSON.stringify(parentOpts));
  await page.selectOption("#newBomParent", String(m2));
  const row1 = page.locator("[data-new-bom-line]").nth(0);
  await row1.locator("[data-line-name]").selectOption(String(cu));
  check("Picking part name fills part code and unit", (await row1.locator("[data-line-code]").inputValue()) === String(cu) && (await row1.locator("[data-line-unit]").inputValue()) === "kg");
  await row1.locator("[data-line-qty]").fill("0.45");
  await page.click("[data-new-bom-add]");
  const row2 = page.locator("[data-new-bom-line]").nth(1);
  await row2.locator("[data-line-search]").fill("SCRAP");
  await page.waitForFunction(() => document.querySelectorAll("[data-new-bom-line]")[1].querySelector("[data-line-code]").options.length === 2);
  check("Search narrows the part list and auto-picks the single match", (await row2.locator("[data-line-code]").inputValue()) === String(scrap));
  await page.click("[data-new-bom-save]");
  check("Empty quantity is stopped before saving", (await page.locator("#toast").textContent()).includes("Line 2"));
  await row2.locator("[data-line-qty]").fill("-0.05");
  await page.screenshot({ path: __dirname + "/new_bom.png" });
  await page.click("[data-new-bom-save]");
  await page.waitForFunction(() => document.querySelectorAll("#masterTable > .table-wrap > table > tbody > tr[data-tree-group]").length === 7);
  const saved = await api("GET", `/api/planning/part-bom/${m2}?tab=Level%202%20Part%20BOM`);
  check("MOTOR2 BOM saved with CU01 0.45 and SCRAP001 -0.05", saved.items.map((r) => `${r.code}:${r.quantity}`).sort().join(",") === "CU01:0.45,SCRAP001:-0.05", saved.items.map((r) => `${r.code}:${r.quantity}`).join(","));
  check("MOTOR2 now in the tree", (await groups.allTextContents()).some((t) => t.includes("FL001MOTOR2")));

  // Level 3: STEEL/SCRAP/CU have no L3 BOM -> dropdown empty, New BOM offers them
  await page.click('[data-bom-tab="Level 3 Part BOM"]');
  await page.waitForSelector(".tab-button.active[data-bom-tab=\"Level 3 Part BOM\"]");
  await page.waitForSelector("#newBomEditor", { state: "attached" });
  check("L3 dropdown empty when no Level 3 BOM exists", (await page.locator("[data-parent-option]").count()) === 0);
  await page.click('[data-action="new-part-bom"]');
  const l3Opts = (await page.locator("#newBomParent option").allTextContents()).slice(1).map((t) => t.split(" · ")[0]).sort();
  check("L3 New BOM offers Level 2 parts", l3Opts.join(",") === "CU01,FL001STEEL0.7,SCRAP001", l3Opts.join(","));

  // Point 2: Category dropdown with "+ Add new category"
  await page.click('[data-bom-tab="Material Master"]');
  await page.click('[data-action="new-material"]');
  const cats = await page.locator("#categorySelect option").allTextContents();
  check("Category dropdown lists existing categories + Add new", cats.includes("electrical") && cats.includes("Raw material") && cats.at(-1) === "+ Add new category", JSON.stringify(cats));
  check("New-category box hidden until chosen", await page.locator("#categoryNew").isHidden());
  await page.fill('[name="code"]', "HYD01");
  await page.fill('[name="name"]', "Hydraulic hose");
  await page.selectOption("#categorySelect", "__new__");
  check("Choosing '+ Add new category' shows the text box", await page.locator("#categoryNew").isVisible());
  await page.fill("#categoryNew", "Hydraulic");
  await page.screenshot({ path: __dirname + "/category.png" });
  await page.click("#modalForm button.button-primary");
  await page.waitForSelector("#modalLayer[hidden]", { state: "attached" });
  const hyd = (await api("GET", "/api/planning/materials?search=HYD01")).items[0];
  check("Material saved with new category 'Hydraulic'", hyd && hyd.category === "Hydraulic", hyd && hyd.category);
  await page.click('[data-action="new-material"]');
  check("New category now listed in dropdown", (await page.locator("#categorySelect option").allTextContents()).includes("Hydraulic"));
  await page.selectOption("#categorySelect", "__new__");
  await page.fill('[name="code"]', "ELE02"); await page.fill('[name="name"]', "Switch");
  await page.fill("#categoryNew", "ELECTRICAL");
  await page.click("#modalForm button.button-primary");
  await page.waitForSelector("#modalLayer[hidden]", { state: "attached" });
  const ele = (await api("GET", "/api/planning/materials?search=ELE02")).items[0];
  check("Typing an existing category in other case reuses it ('electrical')", ele && ele.category === "electrical", ele && ele.category);

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  require("fs").writeFileSync(__dirname + "/bom_ui.json", JSON.stringify(results, null, 1));
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
