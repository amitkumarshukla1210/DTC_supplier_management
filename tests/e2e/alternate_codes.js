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
  // Manage SOB opens by itself after adding: three panels L1 / L2 / L3 with the main code's uses
  await page.waitForSelector("#sobForm");
  const panel = (level) => page.locator(`#sobForm [data-sob-level="${level}"]`);
  check("Manage SOB opens with L1, L2 and L3 panels", (await page.locator("#modalTitle").textContent()).includes("AC-CHIP") && await panel("L1").isVisible() && await panel("L2").isVisible() && await panel("L3").isVisible());
  check("L1 panel: direct use in the model; L2 panel: Motor BOM; L3: none", (await panel("L1").textContent()).includes("Model AC-MODEL") && (await panel("L2").textContent()).includes("AC-MOTOR · Motor BOM") && (await panel("L3").textContent()).includes("not used at L3"));
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  const cells = await page.locator("#altTable tbody tr").first().locator("td").allTextContents();
  check("Pair listed with uses and both prices", cells[0] === "AC-CHIP" && cells[2] === "AC-CHIPNEO" && cells[5] === "L1 × 1 · L2 × 1" && cells[6].includes("480") && cells[7].includes("520"), JSON.stringify(cells));

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
  // the BOM line window shows the SOB but does not change it
  await mainRow.locator("td").nth(3).click();
  await page.waitForSelector("#modalForm .line-alternates");
  check("BOM line window: SOB read-only with a link to Manage SOB", (await page.locator("#modalForm input[data-line-alt], #modalForm input[data-sob-alt]").count()) === 0 && await page.locator("#modalForm [data-alt-manage]").isVisible());

  // Manage SOB from the BOM window: L2 panel only, NEO 80% -> main 20%
  await page.click("#modalForm [data-alt-manage]");
  await page.waitForSelector("#sobForm");
  await page.uncheck('#sobForm [data-sob-level-all="L1"]');
  check("Unticking the L1 panel unticks its uses", !(await page.locator('#sobForm [data-sob-line][data-level="L1"]').isChecked()));
  await page.fill(`#sobForm [data-sob-alt="${neo}"]`, "80");
  check("Main code SOB updates to 20 while typing", (await page.inputValue("#sobMain")) === "20");
  check("From month empty for the first SOB (all months)", (await page.inputValue("#sobFrom")) === "");
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.click('#sobForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("on 1 use"));
  await page.waitForFunction(() => document.querySelector('#sobForm [data-sob-level="L2"] tbody tr')?.textContent.includes("80% / 20%"));
  const motorUse = await page.locator('#sobForm [data-sob-level="L2"] tbody tr').first().locator("td").allTextContents();
  check("Manage SOB shows the Motor use at 80% / 20% after saving", motorUse[3] === "80% / 20%", JSON.stringify(motorUse));
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  mainRow = await openMotor();
  altText = await page.locator("#masterTable tr.tree-alternate").first().locator("td").allTextContents();
  const mainCells = await mainRow.locator("td").allTextContents();
  check("Motor line: main 20% = 0.4, price 512", mainCells.includes("0.4") && mainCells.some((t) => t.includes("main 20% of 2")) && mainCells.some((t) => t.includes("512")), JSON.stringify(mainCells));
  check("Motor line: alternate 80% = 1.6 at 520, repeats Motor", altText.includes("1.6") && altText.some((t) => t.includes("alternate 80%")) && altText.some((t) => t.includes("520")) && altText.includes("AC-MOTOR"), JSON.stringify(altText));
  check("Alternate row directly follows its main code row", await mainRow.evaluate((tr) => tr.nextElementSibling?.classList.contains("tree-alternate")));
  check("Edit still carries the full BOM quantity 2", (await mainRow.getAttribute("data-quantity")) === "2");

  // L1 direct use is independent (still 0%); set 30% on the L1 panel only
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  await page.waitForSelector("#masterTable tr.tree-alternate");
  const l1Main = () => page.locator(`#masterTable tr[data-bom-level1-row][data-material-id="${chip}"]`);
  check("Direct line keeps its own SOB (main 100%)", (await l1Main().textContent()).includes("main 100% of 3"));
  const manage = async () => {
    await page.click('[data-bom-tab="Alternate codes"]');
    await page.waitForSelector("[data-alt-manage]");
    await page.click("#altTable [data-alt-manage]");
    await page.waitForSelector("#sobForm");
  };
  await manage();
  await page.uncheck('#sobForm [data-sob-level-all="L2"]');
  await page.fill(`#sobForm [data-sob-alt="${neo}"]`, "30");
  await page.fill("#sobFrom", "");
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.click('#sobForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("on 1 use"));
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  await page.waitForFunction(() => [...document.querySelectorAll("#masterTable tr.tree-alternate")].some((tr) => tr.textContent.includes("alternate 30%")));
  const l1Cells = await l1Main().locator("td").allTextContents();
  check("Direct line: main 70% = 2.1, price 492", l1Cells.includes("2.1") && l1Cells.some((t) => t.includes("492")), JSON.stringify(l1Cells));
  const l1Alt = await page.locator("#masterTable tr.tree-alternate").first().locator("td").allTextContents();
  check("Direct line: alternate 0.9, repeats the model code", l1Alt.includes("0.9") && l1Alt.includes("AC-MODEL"), JSON.stringify(l1Alt));
  await page.screenshot({ path: __dirname + "/alternate_codes.png" });

  // dated change (a replacement): NEO 100% from Jan 2099 on all uses (L1 and L2 ticked)
  await manage();
  check("From month defaults to this month once SOB has history", (await page.inputValue("#sobFrom")) === await page.evaluate(() => monthNow()));
  await page.fill("#sobFrom", "2099-01");
  await page.fill(`#sobForm [data-sob-alt="${neo}"]`, "100");
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.click('#sobForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("on 2 uses"));
  await page.waitForFunction(() => document.querySelector('#sobForm [data-sob-level="L2"] tbody tr')?.textContent.includes("Jan 2099"));
  check("Manage SOB lists the next change on each use", (await page.locator('#sobForm [data-sob-level="L2"] tbody tr').first().textContent()).includes("Jan 2099: main 0%"));
  await page.locator("#modalBody").screenshot({ path: __dirname + "/manage_sob.png" });
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  mainRow = await openMotor();
  const motorMain = await mainRow.textContent();
  const motorAlt = await page.locator("#masterTable tr.tree-alternate").first().textContent();
  check("Motor line keeps this month's SOB and shows the dated change", motorMain.includes("main 20% of 2 → 0% from Jan 2099") && motorAlt.includes("alternate 80% → 100% from Jan 2099") && (await mainRow.locator("td").allTextContents()).includes("0.4"), motorMain + " | " + motorAlt);
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  await page.waitForSelector("#masterTable tr.tree-alternate");
  check("The L1 use got the same dated change", (await l1Main().textContent()).includes("main 70% of 3 → 0% from Jan 2099"), await l1Main().textContent());
  // the BOM line window shows the history (read-only)
  mainRow = await openMotor();
  await mainRow.locator("td").nth(3).click();
  await page.waitForFunction(() => document.querySelectorAll("#lineSobHistory tbody tr").length === 2);
  check("BOM line history shows Start and Jan 2099", (await page.locator("#lineSobHistory tbody tr").nth(1).textContent()).includes("Jan 2099"));
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  // remove the Jan 2099 change from the L2 use only
  await manage();
  await page.uncheck('#sobForm [data-sob-level-all="L1"]');
  await page.fill("#sobFrom", "2099-01");
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.click("#sobRemove");
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("removed from 1 use"));
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  mainRow = await openMotor();
  check("Motor line no longer shows the removed change", !(await mainRow.textContent()).includes("2099"));
  await page.click('[data-bom-tab="Model BOM Level 1"]');
  await page.waitForSelector("#masterTable tr.tree-alternate");
  check("The L1 use keeps its Jan 2099 change", (await l1Main().textContent()).includes("from Jan 2099"));

  // SOB change log tab: 1 (L2) + 1 (L1) + 2 (Jan 2099) + 1 removed = 5 rows, newest first
  await page.click('[data-bom-tab="SOB change log"]');
  await page.waitForSelector("#sobLogTable tbody tr");
  const logRows = await page.locator("#sobLogTable tbody tr").count();
  const newest = await page.locator("#sobLogTable tbody tr").first().locator("td").allTextContents();
  check("SOB change log lists every change, newest first", logRows === 5 && newest[2] === "Removed" && newest[3] === "AC-CHIP" && newest[4] === "L2" && newest[6] === "Jan 2099" && newest[1] === "Admin", `${logRows} ${JSON.stringify(newest)}`);
  const [logDownload] = await Promise.all([page.waitForEvent("download"), page.click("#sobLogExcel")]);
  check("SOB change log downloads as Excel", logDownload.suggestedFilename() === "sob-change-log.xlsx", logDownload.suggestedFilename());
  await page.click('[data-bom-tab="Alternate codes"]');
  await page.waitForSelector("#altTable table");
  check("Alternate codes has a Download SOB history button", (await page.locator('#pageContent a:has-text("Download SOB history")').getAttribute("href")) === "/api/planning/sob-change-log.xlsx");

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
