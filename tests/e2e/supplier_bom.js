// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/supplier_bom.js   (needs Node + Playwright).
// Supplier materials + Supplier BOM tabs, used through the screen. Expected values are
// worked out by hand: BRD 1 x 415 = 415; ICSET = CHIP 2 x 500 (480 @50% + 520 @50%)
// + SOLDER 0.01 x 900 = 1009; should-cost 1424; quoted 1800; gap 376.
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
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  const mat = async (code, name, uom, cost = 0) => (await api("POST", "/api/planning/materials", { code, name, uom, unit_cost: cost, currency: "INR", category: "Electrical" })).id;
  const pcb = await mat("FL001PCB1", "PCB1", "PCS");
  const chip = await mat("FL001MICROCHIP", "MicroChip", "PCS", 500);
  await mat("FL001MICROCHIPNEO", "MicroChipNeo", "PCS");
  const sup = async (name) => (await api("POST", "/api/suppliers", { name, category: "Electrical", country: "India" })).id;
  const iljin = await sup("Iljin Noida");
  await sup("Chip Maker A");
  await sup("Chip Maker B");
  await api("POST", "/api/planning/supplier-parts", { material_id: pcb, supplier_id: iljin, sob_percent: 100, unit_price: 1800, currency: "INR" });
  const own = async (code, name, uom, price) => (await api("POST", "/api/planning/supplier-materials", { supplier_id: iljin, code, name, uom, unit_price: price, currency: "INR" })).id;
  const board = await own("BRD", "Bare board", "PCS", 415);
  const icset = await own("ICSET", "IC set", "SET", 0);
  await own("SOLDER", "Solder wire", "kg", 900);
  const lBoard = (await api("POST", "/api/planning/supplier-bom/lines", { supplier_id: iljin, material_id: pcb, supplier_material_id: board, quantity: 1 })).id;
  const lIc = (await api("POST", "/api/planning/supplier-bom/lines", { supplier_id: iljin, material_id: pcb, supplier_material_id: icset, quantity: 1 })).id;
  const lChip = (await api("POST", "/api/planning/supplier-bom/lines", { supplier_id: iljin, material_id: pcb, parent_line_id: lIc, item_id: chip, quantity: 2 })).id;
  void lBoard;

  await page.goto(BASE);
  await page.click('[data-page="suppliers"]');
  const tabNames = await page.locator("[data-supplier-tab]").allTextContents();
  check("Tabs include Supplier materials and Supplier BOM", tabNames.includes("Supplier materials") && tabNames.includes("Supplier BOM"), JSON.stringify(tabNames.slice(-2)));

  // ---- Supplier materials tab
  await page.click('[data-supplier-tab="Supplier materials"]');
  await page.selectOption("#smSupplier", String(iljin));
  await page.waitForFunction(() => document.querySelectorAll("#smTable tbody tr").length === 3);
  await page.fill('#smAddForm [name="code"]', "TERM");
  await page.fill('#smAddForm [name="name"]', "Terminal");
  await page.fill('#smAddForm [name="unit_price"]', "2");
  await page.click('#smAddForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelectorAll("#smTable tbody tr").length === 4);
  check("Add supplier material from the form", (await page.locator("#smTable").textContent()).includes("Terminal"));
  await page.locator('#smTable tr:has-text("TERM") [data-sm-edit]').click();
  await page.fill('#smEditForm [name="name"]', "Terminal 6.3");
  await page.click('#smEditForm button');
  await page.waitForFunction(() => document.querySelector("#smTable").textContent.includes("Terminal 6.3"));
  check("Edit supplier material", true);
  await page.locator('#smTable tr:has-text("TERM") [data-sm-delete]').click();
  await page.waitForFunction(() => document.querySelectorAll("#smTable tbody tr").length === 3);
  check("Delete supplier material", true);

  // ---- Supplier BOM tab
  await page.click('[data-supplier-tab="Supplier BOM"]');
  await page.selectOption("#sbSupplier", String(iljin));
  await page.waitForSelector('#sbTable tr[data-sb-line]');
  const productText = await page.locator("#sbProduct option:checked").textContent();
  check("Part list shows parts bought from the supplier", productText.includes("FL001PCB1"), productText);

  // add SOLDER under ICSET through the screen
  await page.click(`[data-sb-add="${lIc}"]`);
  await page.waitForFunction(() => [...document.querySelectorAll("#sbPart option")].some((o) => o.textContent.includes("SOLDER")));
  await page.fill("#sbPartSearch", "SOLD");
  await page.waitForFunction(() => document.querySelectorAll("#sbPart option").length === 1);
  await page.fill('#sbLineForm [name="quantity"]', "0.01");
  await page.click("#sbLineForm button");
  await page.waitForFunction(() => document.querySelectorAll("#sbTable tr[data-sb-line]").length === 4);
  check("Add L2 line below IC set from the screen", (await page.locator("#sbTable").textContent()).includes("SOLDER"));

  // sources on MicroChip: A 480 @ 50, B with alternate code 520 @ 50
  await page.click(`[data-sb-sources="${lChip}"]`);
  await page.waitForSelector("#sbSourceRows tr");
  const first = page.locator("#sbSourceRows tr").nth(0);
  await first.locator("[data-src-supplier]").selectOption({ label: (await first.locator("[data-src-supplier] option").allTextContents()).find((t) => t.includes("Chip Maker A")) });
  await first.locator("[data-src-price]").fill("480");
  await first.locator("[data-src-sob]").fill("50");
  await page.click("#sbAddSource");
  const second = page.locator("#sbSourceRows tr").nth(1);
  await second.locator("[data-src-supplier]").selectOption({ label: (await second.locator("[data-src-supplier] option").allTextContents()).find((t) => t.includes("Chip Maker B")) });
  await second.locator("[data-src-kind]").selectOption("our");
  await second.locator("[data-src-search]").fill("MICROCHIPNEO");
  await page.waitForFunction(() => { const o = [...document.querySelectorAll("#sbSourceRows tr")[1].querySelectorAll("[data-src-alt] option")]; return o.length === 2 && o[1].textContent.includes("FL001MICROCHIPNEO"); });
  await second.locator("[data-src-alt]").selectOption({ index: 1 });
  await second.locator("[data-src-price]").fill("520");
  await second.locator("[data-src-sob]").fill("50");
  await page.click("#sbSaveSources");
  await page.waitForFunction(() => document.querySelector("#sbTable").textContent.includes("SOB-weighted (2 sources)"));
  check("Sources saved and shown with the alternate code", (await page.locator("#sbTable").textContent()).includes("FL001MICROCHIPNEO"));

  const summary = await page.locator("#sbSummary").textContent();
  const amount = (label) => { const m = summary.match(new RegExp(`${label}\\s*(?:INR|₹)?\\s*([\\d,]+\\.\\d{2})`)); return m ? Number(m[1].replace(/,/g, "")) : NaN; };
  check("Should-cost on screen = 1,424.00 (hand-worked)", amount("Should-cost") === 1424, summary);
  check("Quoted price = 1,800.00", amount("Quoted price") === 1800);
  check("Gap = 376.00", amount("Quoted − should-cost") === 376);
  const icRow = page.locator(`#sbTable tr[data-sb-line="${lIc}"]`);
  check("IC set line is bold and costs 1,009.00", (await icRow.evaluate((tr) => tr.classList.contains("tree-has-children"))) && (await icRow.textContent()).includes("1,009.00"));
  const chipRow = page.locator(`#sbTable tr[data-sb-line="${lChip}"]`);
  check("MicroChip row repeats its L1 parent (IC set)", (await chipRow.locator("td").nth(1).textContent()).trim() === "ICSET");

  // collapse / expand
  await icRow.locator("[data-sb-toggle]").click();
  check("Collapse hides the lines below IC set", await chipRow.isHidden());
  await icRow.locator("[data-sb-toggle]").click();
  check("Expand shows them again", await chipRow.isVisible());

  // edit quantity through the screen: board 2 -> should-cost 830 + 1009 = 1839
  await page.click(`[data-sb-edit="${lBoard}"]`);
  await page.fill('#sbLineForm [name="quantity"]', "2");
  await page.click("#sbLineForm button");
  await page.waitForFunction(() => document.querySelector("#sbSummary").textContent.includes("1,839.00"));
  check("Edit line quantity updates the should-cost (1,839.00)", true);

  const [download] = await Promise.all([page.waitForEvent("download"), page.locator('#sbActions a:has-text("Download Excel")').click()]);
  check("Download Excel of the supplier BOM", download.suggestedFilename().endsWith(".xlsx"), download.suggestedFilename());
  await page.screenshot({ path: __dirname + "/supplier_bom.png", fullPage: true });

  // delete IC set removes the lines below it
  await page.click(`[data-sb-delete="${lIc}"]`);
  await page.waitForFunction(() => document.querySelectorAll("#sbTable tr[data-sb-line]").length === 1);
  check("Delete a line removes the lines below it", true);

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
