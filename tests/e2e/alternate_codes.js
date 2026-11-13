// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/alternate_codes.js   (needs Node + Playwright).
// Alternate codes tab and alternate rows in the BOM tree. Hand-worked values:
// CHIP 480 (A) / CHIPNEO 520 (B). At 40%: price 0.6 x 480 + 0.4 x 520 = 496; at 50%: 500.
// L1 line CHIP qty 3 -> alternate 1.5 at 50%; L2 Motor -> CHIP qty 2 -> alternate 1.
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
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  const mat = async (code, name, uom = "PCS", cost = 0) => (await api("POST", "/api/planning/materials", { code, name, uom, unit_cost: cost, currency: "INR", category: "Electrical" })).id;
  const chip = await mat("AC-CHIP", "MicroChip", "PCS", 500);
  const neo = await mat("AC-CHIPNEO", "MicroChipNeo");
  const motor = await mat("AC-MOTOR", "Motor");
  await mat("AC-STEEL", "Steel", "kg");
  const model = (await api("POST", "/api/planning/models", { code: "AC-MODEL", name: "Alt model" })).id;
  await api("POST", `/api/planning/model-bom/${model}`, { material_id: motor, quantity: 1, uom: "PCS" });
  await api("POST", `/api/planning/model-bom/${model}`, { material_id: chip, quantity: 3, uom: "PCS" });
  await api("POST", `/api/planning/part-bom/${motor}/new`, { tab: "Level 2 Part BOM", lines: [{ child_item_id: chip, quantity: 2 }] });
  const supA = (await api("POST", "/api/suppliers", { name: "Chip Maker A", category: "Electrical", country: "India" })).id;
  const supB = (await api("POST", "/api/suppliers", { name: "Chip Maker B", category: "Electrical", country: "India" })).id;
  await api("POST", "/api/planning/supplier-parts", { material_id: chip, supplier_id: supA, sob_percent: 100, unit_price: 480, currency: "INR" });
  await api("POST", "/api/planning/supplier-parts", { material_id: neo, supplier_id: supB, sob_percent: 100, unit_price: 520, currency: "INR" });

  await page.goto(BASE);
  await page.click('[data-page="master"]');
  await page.waitForSelector("[data-bom-tab]");
  check("Product Master has an Alternate codes tab", (await page.locator("[data-bom-tab]").allTextContents()).includes("Alternate codes"));
  await page.click('[data-bom-tab="Alternate codes"]');
  await page.waitForSelector("#altTable table");
  check("Empty list says so", (await page.locator("#altTable").textContent()).includes("No alternate codes yet"));
  const mainOptions = await page.locator("#altMainSelect option").allTextContents();
  check("Main code list = bought parts only (no in-house Motor)", mainOptions.some((t) => t.includes("AC-CHIP")) && !mainOptions.some((t) => t.includes("AC-MOTOR")), JSON.stringify(mainOptions.slice(0, 6)));

  // a wrong pick is refused with the reason (different unit)
  await page.fill("#altMainSearch", "AC-CHIP");
  await page.waitForFunction(() => [...document.querySelectorAll("#altMainSelect option")].some((o) => o.textContent.startsWith("AC-CHIP ·")) && document.querySelectorAll("#altMainSelect option").length <= 4);
  await page.selectOption("#altMainSelect", String(chip));
  await page.fill("#altAltSearch", "AC-STEEL");
  await page.waitForFunction(() => document.querySelectorAll("#altAltSelect option").length === 2);
  await page.selectOption("#altAltSelect", { index: 1 });
  await page.fill('#altAddForm [name="sob_percent"]', "40");
  await page.click('#altAddForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("same unit"));
  check("Different unit is refused with the reason", true, await toastText());

  await page.fill("#altAltSearch", "AC-CHIPNEO");
  await page.waitForFunction(() => document.querySelectorAll("#altAltSelect option").length === 2 && document.querySelectorAll("#altAltSelect option")[1].textContent.includes("AC-CHIPNEO"));
  await page.selectOption("#altAltSelect", String(neo));
  await page.click('#altAddForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#altTable").textContent.includes("AC-CHIPNEO"));
  const cells = await page.locator("#altTable tbody tr").first().locator("td").allTextContents();
  check("Row shows main 60% / alternate 40% with both prices", cells[0] === "AC-CHIP" && cells[2] === "60%" && cells[3] === "AC-CHIPNEO" && cells[5] === "40%" && cells[7].includes("480") && cells[8].includes("520"), JSON.stringify(cells));

  // Level 2 tree at 40%: alternate row below CHIP, qty 2 x 40% = 0.8, main price 496
  await page.click('[data-bom-tab="Level 2 Part BOM"]');
  await page.waitForSelector('#masterTable [data-tree-expand]');
  await page.locator('#masterTable [data-tree-expand]').first().click();
  await page.waitForSelector("#masterTable tr.tree-alternate");
  const mainRow = page.locator('#masterTable tr[data-bom-level2-row]').first();
  const altRow = page.locator("#masterTable tr.tree-alternate").first();
  check("Alternate row directly follows its main code row", await mainRow.evaluate((tr) => tr.nextElementSibling?.classList.contains("tree-alternate")));
  const altText = await altRow.locator("td").allTextContents();
  check("Alternate row: code, SOB, quantity 0.8, own price 520", altText.some((t) => t === "↳ AC-CHIPNEO") && altText.some((t) => t.includes("alternate 40%")) && altText.includes("0.8") && altText.some((t) => t.includes("520")), JSON.stringify(altText));
  check("Alternate row repeats the parent (Motor)", altText.includes("AC-MOTOR"));
  const mainText = await mainRow.textContent();
  check("Main row shows main 60% and blended price 496", mainText.includes("main 60%") && mainText.includes("496"), mainText);
  const mainCells = await mainRow.locator("td").allTextContents();
  check("Main row quantity = its share 2 x 60% = 1.2, BOM quantity shown in the name", mainCells.includes("1.2") && mainText.includes("main 60% of 2"), JSON.stringify(mainCells));
  check("Edit on the main row still carries the full BOM quantity 2", (await mainRow.getAttribute("data-quantity")) === "2");

  // edit SOB to 50% in the tab
  await page.click('[data-bom-tab="Alternate codes"]');
  await page.waitForSelector("[data-alt-edit]");
  await page.click("[data-alt-edit]");
  await page.fill('#altEditForm [name="sob_percent"]', "50");
  await page.click("#altEditForm button");
  await page.waitForFunction(() => document.querySelector("#altTable tbody tr td:nth-child(3)")?.textContent === "50%");
  check("Edit SOB to 50% leaves main 50%", true);

  // Level 1 tree at 50%: CHIP qty 3 -> alternate 1.5, price 500
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  await page.waitForSelector("#masterTable tr.tree-alternate");
  const l1Alt = await page.locator("#masterTable tr.tree-alternate").first().locator("td").allTextContents();
  check("Level 1: alternate qty 1.5 and repeats the model code", l1Alt.includes("1.5") && l1Alt.includes("AC-MODEL"), JSON.stringify(l1Alt));
  const l1Main = await page.locator('#masterTable tr[data-bom-level1-row][data-material-id="' + chip + '"]').textContent();
  check("Level 1: main price 500 (50% x 480 + 50% x 520)", l1Main.includes("500") && l1Main.includes("main 50%"), l1Main);
  check("Level 1: main quantity 3 x 50% = 1.5", (await page.locator('#masterTable tr[data-bom-level1-row][data-material-id="' + chip + '"] td').allTextContents()).includes("1.5"));
  await page.screenshot({ path: __dirname + "/alternate_codes.png" });

  // delete: tree back to the main code alone
  await page.click('[data-bom-tab="Alternate codes"]');
  await page.waitForSelector("[data-alt-delete]");
  await page.click("[data-alt-delete]");
  await page.waitForFunction(() => document.querySelector("#altTable").textContent.includes("No alternate codes yet"));
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  await page.waitForSelector("#masterTable tr[data-bom-level1-row]");
  check("After delete the tree has no alternate rows", (await page.locator("#masterTable tr.tree-alternate").count()) === 0);

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
