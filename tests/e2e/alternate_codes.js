// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/alternate_codes.js   (needs Node + Playwright).
// Alternate codes tab (pairs only) and alternate SOB set per BOM line by clicking the row.
// Hand-worked: CHIP 480 (A) / CHIPNEO 520 (B). Motor BOM line qty 2 at NEO 80%: main 0.4,
// alternate 1.6, price 0.2 x 480 + 0.8 x 520 = 512. Direct L1 line qty 3 at NEO 30%:
// main 2.1, alternate 0.9, price 0.7 x 480 + 0.3 x 520 = 492.
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
  check("No SOB field on the Alternate codes tab", (await page.locator('#altAddForm [name="sob_percent"]').count()) === 0);
  await page.fill("#altMainSearch", "AC-CHIP");
  await page.waitForFunction(() => [...document.querySelectorAll("#altMainSelect option")].some((o) => o.textContent.startsWith("AC-CHIP ·")) && document.querySelectorAll("#altMainSelect option").length <= 4);
  await page.selectOption("#altMainSelect", String(chip));
  await page.fill("#altAltSearch", "AC-STEEL");
  await page.waitForFunction(() => document.querySelectorAll("#altAltSelect option").length === 2);
  await page.selectOption("#altAltSelect", { index: 1 });
  await page.click('#altAddForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("same unit"));
  check("Different unit is refused with the reason", true, await toastText());

  await page.fill("#altAltSearch", "AC-CHIPNEO");
  await page.waitForFunction(() => document.querySelectorAll("#altAltSelect option").length === 2 && document.querySelectorAll("#altAltSelect option")[1].textContent.includes("AC-CHIPNEO"));
  await page.selectOption("#altAltSelect", String(neo));
  await page.click('#altAddForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#altTable").textContent.includes("AC-CHIPNEO"));
  const headers = await page.locator("#altTable thead th").allTextContents();
  const cells = await page.locator("#altTable tbody tr").first().locator("td").allTextContents();
  check("Pair listed with both prices and no SOB column", cells[0] === "AC-CHIP" && cells[2] === "AC-CHIPNEO" && cells[5].includes("480") && cells[6].includes("520") && !headers.some((h) => h.includes("SOB")), JSON.stringify(cells));

  // Level 2 (Motor BOM): alternate at 0% by default
  const openMotor = async () => {
    await page.click('[data-bom-tab="Level 2 Part BOM"]');
    await page.waitForSelector('#masterTable [data-tree-expand]');
    await page.locator('#masterTable [data-tree-expand]').first().click();
    await page.waitForSelector("#masterTable tr.tree-alternate");
    return page.locator('#masterTable tr[data-bom-level2-row]').first();
  };
  let mainRow = await openMotor();
  let altText = await page.locator("#masterTable tr.tree-alternate").first().locator("td").allTextContents();
  check("Default: alternate 0%, quantity 0; main 100% of 2", altText.includes("0") && altText.some((t) => t.includes("alternate 0%")) && (await mainRow.textContent()).includes("main 100% of 2") && (await mainRow.textContent()).includes("480"), JSON.stringify(altText));

  // click the Motor BOM row, set NEO 80%: main SOB shows 20% at once
  await mainRow.locator("td").nth(3).click();
  await page.waitForSelector("#modalForm [data-line-alt]");
  check("Row click opens the line with an alternate SOB field", await page.locator("#modalForm .line-alternates").isVisible());
  await page.fill(`#modalForm [data-line-alt="${neo}"]`, "80");
  check("Main code SOB updates to 20 while typing", (await page.inputValue("#lineMainSob")) === "20");
  await page.click("#modalForm .modal-footer button");
  await page.waitForFunction(() => document.querySelector("#modalLayer").hidden);
  mainRow = await openMotor();
  altText = await page.locator("#masterTable tr.tree-alternate").first().locator("td").allTextContents();
  const mainCells = await mainRow.locator("td").allTextContents();
  check("Motor line: main 20% = 0.4, price 512", mainCells.includes("0.4") && mainCells.some((t) => t.includes("main 20% of 2")) && mainCells.some((t) => t.includes("512")), JSON.stringify(mainCells));
  check("Motor line: alternate 80% = 1.6 at 520, repeats Motor", altText.includes("1.6") && altText.some((t) => t.includes("alternate 80%")) && altText.some((t) => t.includes("520")) && altText.includes("AC-MOTOR"), JSON.stringify(altText));
  check("Alternate row directly follows its main code row", await mainRow.evaluate((tr) => tr.nextElementSibling?.classList.contains("tree-alternate")));
  check("Edit still carries the full BOM quantity 2", (await mainRow.getAttribute("data-quantity")) === "2");

  // Level 1 direct line is independent: still 0%; set 30% with the Edit button
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  await page.waitForSelector("#masterTable tr.tree-alternate");
  const l1Main = () => page.locator(`#masterTable tr[data-bom-level1-row][data-material-id="${chip}"]`);
  check("Direct line keeps its own SOB (main 100%)", (await l1Main().textContent()).includes("main 100% of 3"));
  await l1Main().locator("[data-edit-bom-row]").first().click();
  await page.waitForSelector("#modalForm [data-line-alt]");
  await page.fill(`#modalForm [data-line-alt="${neo}"]`, "30");
  await page.click("#modalForm .modal-footer button");
  await page.waitForFunction(() => document.querySelector("#modalLayer").hidden);
  await page.waitForFunction(() => [...document.querySelectorAll("#masterTable tr.tree-alternate")].some((tr) => tr.textContent.includes("alternate 30%")));
  const l1Cells = await l1Main().locator("td").allTextContents();
  check("Direct line: main 70% = 2.1, price 492", l1Cells.includes("2.1") && l1Cells.some((t) => t.includes("492")), JSON.stringify(l1Cells));
  const l1Alt = await page.locator("#masterTable tr.tree-alternate").first().locator("td").allTextContents();
  check("Direct line: alternate 0.9, repeats the model code", l1Alt.includes("0.9") && l1Alt.includes("AC-MODEL"), JSON.stringify(l1Alt));
  await page.screenshot({ path: __dirname + "/alternate_codes.png" });

  // dated SOB change (a replacement): NEO 100% from Jan 2099 on every line using AC-CHIP
  mainRow = await openMotor();
  await mainRow.locator("td").nth(3).click();
  await page.waitForSelector("#lineSobHistory table");
  check("History lists the start SOB (80% / main 20%)", (await page.locator("#lineSobHistory tbody tr").count()) === 1 && (await page.locator("#lineSobHistory tbody tr").textContent()).includes("80%"), await page.locator("#lineSobHistory").textContent());
  check("From month defaults to this month once a line has history", (await page.inputValue("#lineSobFrom")) === await page.evaluate(() => monthNow()));
  await page.fill("#lineSobFrom", "2099-01");
  await page.fill(`#modalForm [data-line-alt="${neo}"]`, "100");
  await page.check("#lineSobEverywhere");
  await page.click("#modalForm .modal-footer button");
  await page.waitForFunction(() => document.querySelector("#modalLayer").hidden);
  mainRow = await openMotor();
  const motorMain = await mainRow.textContent();
  const motorAlt = await page.locator("#masterTable tr.tree-alternate").first().textContent();
  check("Motor line keeps this month's SOB and shows the dated change", motorMain.includes("main 20% of 2 → 0% from Jan 2099") && motorAlt.includes("alternate 80% → 100% from Jan 2099") && (await mainRow.locator("td").allTextContents()).includes("0.4"), motorMain + " | " + motorAlt);
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  await page.waitForSelector("#masterTable tr.tree-alternate");
  check("Apply-everywhere also dated the direct L1 line", (await l1Main().textContent()).includes("main 70% of 3 → 0% from Jan 2099"), await l1Main().textContent());
  // history shows both changes; remove the Jan 2099 change on the Motor line
  mainRow = await openMotor();
  await mainRow.locator("td").nth(3).click();
  await page.waitForFunction(() => document.querySelectorAll("#lineSobHistory tbody tr").length === 2);
  check("History shows Start and Jan 2099", (await page.locator("#lineSobHistory tbody tr").nth(1).textContent()).includes("Jan 2099"));
  await page.locator("#modalBody .line-alternates").screenshot({ path: __dirname + "/sob_history.png" });
  await page.click('#lineSobHistory [data-sob-change-delete="2099-01"]');
  await page.waitForFunction(() => document.querySelectorAll("#lineSobHistory tbody tr").length === 1);
  check("Remove a change from the history", true);
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  mainRow = await openMotor();
  check("Motor line no longer shows the removed change", !(await mainRow.textContent()).includes("2099"));

  // delete the pair: tree back to the main code alone
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
