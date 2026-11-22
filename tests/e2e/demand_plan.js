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
  await page.waitForFunction(() => document.querySelectorAll("#dpTable thead th summary").length === 18);
  const heads = await page.locator("#dpTable thead th summary").allTextContents();
  check("T1 columns: Opening Stock 01st Month, Buffer norm, Supplier FG stock (not counted)", heads.slice(8, 11).join("|") === "Opening Stock 01st Month|Buffer norm|Supplier FG stock (not counted)", heads.join("|"));
  check("Month columns: Requirement and Suggested PO", heads.slice(12).join("|") === "Nov-26 Requirement|Nov-26 Suggested PO|Dec-26 Requirement|Dec-26 Suggested PO|Jan-27 Requirement|Jan-27 Suggested PO", heads.join("|"));
  const t1 = await page.locator("#dpTable tbody tr").first().locator("td").allTextContents();
  // your PO (Savera's FG 9,000 NOT counted): Nov need 60,000 + buffer 60,000/30.4375x15 = 29,568 - opening 10,000 = 79,568 x 60% = 47,740.8 -> 50,000 (MOQ 5,000) by Sep-26;
  // Dec 30,000 + 14,784 - 29,568 left = 15,216 x 60% = 9,129.6 -> 10,000 by Oct-26; Jan no plan
  check("T1 Savera beam: Nov 36,000 -> PO 50,000 by Sep-26; Dec 18,000 -> 10,000 by Oct-26", t1[0] === "SUP-2013" && t1[4] === "OEM" && t1[8] === "10,000" && t1[9] === "15 days" && t1[10] === "9,000" && t1[12] === "36,000" && t1[13] === "50,000by Sep-26" && t1[14] === "18,000" && t1[15] === "10,000by Oct-26" && t1[17] === "—", JSON.stringify(t1));
  check("T1 tab keeps the PO issue and register", await page.locator("#poForm").isVisible() && await page.locator("#poTable").isVisible());
  check("Excel download follows the tab and months", /demand-plan\.xlsx\?tier=T1&start_month=2026-11&months=3/.test(await page.getAttribute("#dpExcel", "href")));
  await page.click('[data-dp-tab="T2"]');
  await page.waitForSelector("#dpTable tbody tr");
  const t2 = await page.locator("#dpTable tbody tr").evaluateAll((trs) => trs.map((tr) => [tr.cells[0].textContent, tr.cells[2].textContent, tr.cells[4].textContent, tr.cells[10].textContent, tr.cells[11].textContent]));
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
  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
