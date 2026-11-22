// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/demand_plan.js   (needs Node + Playwright).
// Demand & PO Plan: T1 / T2 / T3 Supplier part Planning, month-wise requirement and suggested PO.
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1700, height: 950 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const api = async (method, p, data) => {
    const r = await page.request.fetch(BASE + p, { method, data, headers: { "content-type": "application/json" } });
    if (r.status() >= 300) throw new Error(`${method} ${p} -> ${r.status()} ${await r.text()}`);
    return r.json();
  };
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  const beam = (await api("POST", "/api/planning/materials", { code: "UF220BFB010", name: "Base Front Beam", uom: "Nos", stock_on_hand: 10000, buffer_days: 15 })).id;
  const steel = (await api("POST", "/api/planning/tier-materials", { tier: "T2", code: "FL001STEEL0.7", name: "STEEL0.7", uom: "kg" })).id;
  const ore = (await api("POST", "/api/planning/tier-materials", { tier: "T3", code: "ORE-01", name: "Iron ore", uom: "kg" })).id;
  const product = (await api("POST", "/api/planning/products", { code: "FL", name: "Fans" })).id;
  const eva = (await api("POST", "/api/planning/models", { product_id: product, code: "FL12000", name: "Eva 8014" })).id;
  await api("POST", `/api/planning/model-bom/${eva}`, { material_id: beam, quantity: 1, uom: "Nos" });
  const sup = async (code, name) => (await api("POST", "/api/suppliers", { code, name, category: "Sheet Metal", country: "India" })).id;
  const savera = await sup("SUP-2013", "Savera"), jgi = await sup("SUP-2004", "JGI STEEL"), tata = await sup("SUP-2014", "TATA Steel"), xyz = await sup("SUP-2030", "XYZ Mining");
  const link = (await api("POST", "/api/planning/supplier-parts", { material_id: beam, supplier_id: savera, sob_percent: 60, moq: 5000, lead_time_days: 45, unit_price: 85, currency: "INR" })).id;
  const steelLine = (await api("POST", "/api/planning/supplier-coverage/lines", { link_id: link, tier_material_id: steel, quantity: 2, sources: [{ sub_supplier_id: jgi, sob_percent: 30, moq: 10000, lead_time_days: 20, unit_price: 74, currency: "INR" }, { sub_supplier_id: tata, sob_percent: 70, unit_price: 72, currency: "INR" }] })).id;
  await api("POST", "/api/planning/supplier-coverage/lines", { link_id: link, parent_line_id: steelLine, tier_material_id: ore, quantity: 1.7, sources: [{ sub_supplier_id: xyz, sob_percent: 100 }] });
  await api("POST", "/api/planning/plans", { issue_month: "2026-11", workspace: "trial", release_date: "2026-11-10", rows: [{ model_code: "FL12000", target_month: "2026-11", quantity: 60000 }, { model_code: "FL12000", target_month: "2026-12", quantity: 30000 }] });
  await api("POST", "/api/planning/supplier-stock/2026-11?workspace=trial", { rows: [{ supplier_part_id: link, quantity: 9000 }] });

  await page.goto(BASE);
  check("Menu: Demand & PO Plan (was PO schedule)", (await page.locator('[data-page="purchase-orders"]').textContent()).trim() === "Demand & PO Plan");
  await page.click('[data-page="purchase-orders"]');
  await page.waitForSelector("#dpTable table");
  const tabs = await page.locator("[data-dp-tab]").allTextContents();
  check("Three tabs: T1 / T2 / T3 Supplier part Planning", JSON.stringify(tabs) === JSON.stringify(["T1 Supplier part Planning", "T2 Supplier part Planning", "T3 Supplier part Planning"]), JSON.stringify(tabs));
  await page.fill("#dpStart", "2026-11");
  await page.dispatchEvent("#dpStart", "change");
  await page.selectOption("#dpMonths", "3");
  await page.waitForFunction(() => document.querySelectorAll("#dpTable thead th summary").length === 21);
  const heads = await page.locator("#dpTable thead th summary").allTextContents();
  check("T1 columns: Opening Stock 01st Month, Buffer norm, Supplier FG stock (not counted)", heads.slice(8, 11).join("|") === "Opening Stock 01st Month|Buffer norm|Supplier FG stock (not counted)", heads.join("|"));
  check("Month blocks: 1 Requirement x 3, 2 PO demand x 3, 3 Actual PO x 3", heads.slice(12).join("|") === "Requirement Nov-26|Requirement Dec-26|Requirement Jan-27|PO demand Nov-26|PO demand Dec-26|PO demand Jan-27|Actual PO Nov-26|Actual PO Dec-26|Actual PO Jan-27", heads.join("|"));
  const band = await page.locator("#dpTable tr.dp-band th").allTextContents();
  check("Band row over the blocks", band.slice(1).join("|") === "1 · Requirement|2 · PO demand (to issue)|3 · Actual PO (editable)", band.join("|"));
  const t1 = await page.locator("#dpTable tbody tr").first().locator("td").allTextContents();
  // your PO (Savera's FG 9,000 NOT counted): Nov need 60,000 + buffer 60,000/30.4375x15 (29,568.8) - opening 10,000 = 79,568.8 x 60% = 47,741.3 -> 50,000 (MOQ 5,000) by Sep-26;
  // arrives 50,000 + the 40% not linked (31,827.5) -> Nov closing 31,827.5; Dec 30,000 + 14,784.4 - 31,827.5 = 12,956.9 x 60% = 7,774.1 -> 10,000 by Oct-26; Jan no plan
  check("T1 Savera beam: Requirement Nov 36,000 / Dec 18,000; PO demand 50,000 by Sep-26 / 10,000 by Oct-26", t1[0] === "SUP-2013" && t1[4] === "OEM" && t1[8] === "10,000" && t1[9] === "15 days" && t1[10] === "9,000" && t1[12] === "36,000" && t1[13] === "18,000" && t1[14] === "0" && t1[15] === "50,000by Sep-26" && t1[16] === "10,000by Oct-26" && t1[17] === "—", JSON.stringify(t1));
  const actuals = await page.locator("#dpTable .dp-actual-input").evaluateAll((inputs) => inputs.map((input) => input.value));
  check("Actual PO = PO demand until you change it", JSON.stringify(actuals) === JSON.stringify(["50000", "10000", "0"]), JSON.stringify(actuals));
  // you issue 40,000 in Nov: closing 21,827.5; Dec 30,000 + 14,784.4 - 21,827.5 = 22,956.9 x 60% = 13,774.1 -> 15,000
  await page.fill('#dpTable .dp-actual-input[data-dp-month="2026-11"]', "40000");
  await page.press('#dpTable .dp-actual-input[data-dp-month="2026-11"]', "Enter");
  await page.waitForFunction(() => document.querySelectorAll("#dpTable tbody tr td")[16]?.textContent === "15,000by Oct-26", null, { timeout: 5000 }).catch(() => {});
  const after = await page.locator("#dpTable tbody tr").first().locator("td").allTextContents();
  check("Actual PO 40,000 in Nov -> Dec PO demand becomes 15,000 (saved box marked)", after[16] === "15,000by Oct-26" && await page.locator('#dpTable .dp-actual-input[data-dp-month="2026-11"].dp-actual-saved').count() === 1 && (await page.inputValue('#dpTable .dp-actual-input[data-dp-month="2026-12"]')) === "15000", JSON.stringify(after));
  await page.fill('#dpTable .dp-actual-input[data-dp-month="2026-11"]', "");
  await page.press('#dpTable .dp-actual-input[data-dp-month="2026-11"]', "Enter");
  await page.waitForFunction(() => document.querySelectorAll("#dpTable tbody tr td")[16]?.textContent === "10,000by Oct-26", null, { timeout: 5000 }).catch(() => {});
  check("Clear the box -> back to the PO demand", (await page.locator("#dpTable tbody tr").first().locator("td").nth(16).textContent()) === "10,000by Oct-26" && (await page.inputValue('#dpTable .dp-actual-input[data-dp-month="2026-11"]')) === "50000");
  check("T1 tab keeps the PO issue and register", await page.locator("#poForm").isVisible() && await page.locator("#poTable").isVisible());
  check("Excel download follows the tab and months", /demand-plan\.xlsx\?tier=T1&start_month=2026-11&months=3/.test(await page.getAttribute("#dpExcel", "href")));
  await page.click('[data-dp-tab="T2"]');
  await page.waitForSelector("#dpTable tbody tr");
  const t2 = await page.locator("#dpTable tbody tr").evaluateAll((trs) => trs.map((tr) => [tr.cells[0].textContent, tr.cells[2].textContent, tr.cells[4].textContent, tr.cells[10].textContent, tr.cells[13].textContent]));
  // Savera's balance demand = your PO 50,000 - its FG 9,000 = 41,000 x 2 kg: JGI 30% = 24,600 -> PO 30,000 (MOQ 10,000) by Oct-26; TATA 70% = 57,400
  check("T2: JGI 24,600 kg -> PO 30,000 by Oct-26; TATA 57,400 kg, for Savera", JSON.stringify(t2) === JSON.stringify([["SUP-2004", "FL001STEEL0.7", "Savera", "24,600", "30,000by Oct-26"], ["SUP-2014", "FL001STEEL0.7", "Savera", "57,400", "57,400by Nov-26"]]), JSON.stringify(t2));
  check("T2 tab has no PO form", (await page.locator("#poForm").count()) === 0);
  await page.click('[data-dp-tab="T3"]');
  await page.waitForSelector("#dpTable tbody tr");
  const t3 = await page.locator("#dpTable tbody tr").first().locator("td").allTextContents();
  check("T3: XYZ ore 41,000 x 3.4 = 139,400 kg", t3[0] === "SUP-2030" && t3[2] === "ORE-01" && t3[10] === "139,400", JSON.stringify(t3));
  await page.fill("#dpSearch", "nothing");
  await page.waitForFunction(() => document.querySelector("#dpTable")?.textContent.includes("Nothing matches"));
  check("Search filters the rows", true);

  // Rev. Change Impact on the Shortage & excess page: Nov Rev.01 drops Dec to 15,000
  await api("POST", "/api/planning/plans", { issue_month: "2026-11", workspace: "trial", release_date: "2026-11-20", rows: [{ model_code: "FL12000", target_month: "2026-11", quantity: 60000 }, { model_code: "FL12000", target_month: "2026-12", quantity: 15000 }] });
  await page.click('[data-page="reports"]');
  await page.click('[data-report-tab="impact"]');
  const subTabs = await page.locator("[data-impact-tier]").allTextContents();
  check("Sub-tabs: T1 Supplier / T2 Supplier / T3 Supplier", JSON.stringify(subTabs) === JSON.stringify(["T1 Supplier", "T2 Supplier", "T3 Supplier"]), JSON.stringify(subTabs));
  await page.waitForSelector("#impTable table");
  const picked = [await page.locator("#impOld option:checked").textContent(), await page.locator("#impNew option:checked").textContent()];
  check("Defaults: old Rev.00, new Rev.01", /Rev\.00/.test(picked[0]) && /Rev\.01/.test(picked[1]), JSON.stringify(picked));
  // Savera lead 45 days: Nov / Dec POs fixed (50,000 / 10,000). Dec: 31,827.5 + 10,000 - 15,000 = 26,827.5 > 15,000 + buffer 7,392.2 -> Over stock 19,435.3
  const imp = await page.locator("#impTable tbody tr").evaluateAll((trs) => trs.map((tr) => [...tr.cells].map((td) => td.textContent)));
  check("T1: only Dec shows: -15,000 (-50%), PO fixed 10,000, Over stock 19,435.32", imp.length === 1 && imp[0][0] === "Savera" && imp[0][5] === "Dec-26" && imp[0][6] === "30,000" && imp[0][7] === "15,000" && imp[0][8] === "-15,000" && imp[0][9] === "-50%" && imp[0][10] === "10,000" && imp[0][12] === "26,827.52" && imp[0][14] === "Over stock" && imp[0][16] === "19,435.32", JSON.stringify(imp));
  check("Summary counts (months Nov-Dec: the plan's horizon)", /: 0 shortage, 1 over stock, 0 below buffer, 1 big change/.test(await page.textContent("#impSummary")), await page.textContent("#impSummary"));
  await page.uncheck("#impIssues");
  await page.waitForFunction(() => document.querySelectorAll("#impTable tbody tr").length === 2);
  check("All months when unticked (Nov, Dec: up to the last planned month)", true);
  await page.click('[data-impact-tier="T2"]');
  await page.waitForSelector("#impTable table");
  await page.waitForFunction(() => document.querySelector("#impTable thead th summary")?.textContent === "T2 Supplier");
  const imp2 = await page.locator("#impTable tbody tr").evaluateAll((trs) => trs.map((tr) => [tr.cells[0].textContent, tr.cells[5].textContent, tr.cells[8].textContent, tr.cells[14].textContent]).filter((r) => r[1] !== "Jan-27" || r[3] !== "OK").slice(0, 4));
  // Savera's POs are fixed inside its lead time: JGI's requirement does not change
  check("T2: JGI / TATA requirement unchanged (POs fixed in lead time)", imp2.length >= 2 && imp2.every((r) => r[2] === "0"), JSON.stringify(imp2));
  check("Excel link follows the tier", /rev-impact\.xlsx\?.*tier=T2/.test(await page.getAttribute("#impExcel", "href")));
  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
