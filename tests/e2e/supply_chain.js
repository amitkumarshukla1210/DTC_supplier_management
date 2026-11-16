// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/supply_chain.js   (needs Node + Playwright).
// Supplier/material coverage as supply chains, with the user's own example sheet:
//   Supangita 40% / Savera 60% make Base Front Beam (T1); each buys steel from JGI 30% + TATA 70% (T2).
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1700, height: 900 } });
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
  const beam = (await api("POST", "/api/planning/materials", { code: "UF220BFB010", name: "Base Front Beam", uom: "PCS", unit_cost: 900, currency: "INR", category: "Sheet Metal" })).id;
  const steel = (await api("POST", "/api/planning/tier-materials", { tier: "T2", code: "FL001STEEL0.7", name: "STEEL0.7", uom: "kg" })).id;
  const sup = async (code, name) => (await api("POST", "/api/suppliers", { code, name, category: "Sheet Metal", country: "India" })).id;
  const supangita = await sup("SUP-2007", "Supangita"), savera = await sup("SUP-2013", "Savera"), jgi = await sup("SUP-2004", "JGI STEEL"), tata = await sup("SUP-2014", "TATA Steel");
  await sup("SUP-2020", "Steel Trader");
  const t1 = (supplier, sob, price) => ({ supplier_id: supplier, material_id: beam, sob_percent: sob, unit_price: price });
  const t2 = (supplier, sob, price) => ({ supplier_id: supplier, tier_material_id: steel, sob_percent: sob, unit_price: price });
  const saved = await api("POST", "/api/planning/supplier-chain/rows", { rows: [
    { t1: t1(supangita, 40, 950), t2: t2(jgi, 30, 72), lead_time_days: 12, currency: "INR", price_mode: "spot" },
    { t2: t2(tata, 70, 70), lead_time_days: 20, currency: "INR", price_mode: "spot" },
    { t1: t1(savera, 60, 940), t2: t2(jgi, 30, 72), lead_time_days: 12, currency: "INR", price_mode: "spot" },
    { t2: t2(tata, 70, 70), lead_time_days: 20, currency: "INR", price_mode: "spot" },
  ] });
  check("The sheet's 4 rows save as 6 relationships", saved.rows === 4 && saved.relationships === 6, JSON.stringify(saved));

  await page.goto(BASE);
  await page.click('[data-page="suppliers"]');
  await page.click('[data-supplier-tab="Supplier/material coverage"]');
  await page.waitForSelector('#supplierPartTable tbody tr:has-text("TATA")');
  const read = () => page.locator("#supplierPartTable tbody tr").evaluateAll((trs) => trs.map((tr) => [Number(tr.dataset.treeDepth), tr.cells[1].textContent, tr.cells[4].textContent, tr.cells[6].textContent, tr.cells[13].textContent, tr.cells[14].textContent, tr.cells[21].textContent]));
  const rows = await read();
  // expected, worked out from the sheet: the part, then its T1 suppliers by name (Savera < Supangita), each with its T2 by name (JGI < TATA);
  // T1 rows carry no terms (the terms were typed for the T2 at the end of each row)
  const expected = [[0, "UF220BFB010", "2 T1 suppliers", "", "100%", "", ""],
    [1, "UF220BFB010", "Savera", "", "60%", "0 days", "Edit T1+ Add T2"], [2, "UF220BFB010", "Savera", "JGI STEEL", "30%", "12 days", "Edit T2+ Add T3"], [2, "UF220BFB010", "Savera", "TATA Steel", "70%", "20 days", "Edit T2+ Add T3"],
    [1, "UF220BFB010", "Supangita", "", "40%", "0 days", "Edit T1+ Add T2"], [2, "UF220BFB010", "Supangita", "JGI STEEL", "30%", "12 days", "Edit T2+ Add T3"], [2, "UF220BFB010", "Supangita", "TATA Steel", "70%", "20 days", "Edit T2+ Add T3"]];
  check("Coverage tree matches the sheet: part -> T1 suppliers (SOB) -> T2 suppliers (SOB), each level with its own SOB and terms", JSON.stringify(rows) === JSON.stringify(expected), JSON.stringify(rows));
  const fade = await page.locator("#supplierPartTable tbody tr").evaluateAll((trs) => trs.slice(0, 3).map((tr) => [tr.cells[1].classList.contains("chain-repeat"), tr.cells[4].classList.contains("chain-repeat")]));
  check("Parent values repeat faded (T1 row: part faded; T2 row: part and T1 supplier faded)", JSON.stringify(fade) === JSON.stringify([[false, false], [true, false], [true, true]]), JSON.stringify(fade));
  check("Terms columns after the T3 block: SOB, Lead time, MOQ, Buffer days, Price mode, Price status, Unit price, Currency, then Edit", JSON.stringify((await page.locator("#supplierPartTable thead tr:last-child th summary").allTextContents()).slice(12)) === JSON.stringify(["SOB", "Lead time", "MOQ", "Buffer days", "Price mode", "Price status", "Unit price", "Currency", "Edit"]));
  check("T1 block starts with your part code and name", JSON.stringify((await page.locator("#supplierPartTable thead tr:last-child th summary").allTextContents()).slice(0, 4)) === JSON.stringify(["T1 Material code", "T1 Material name", "T1 Supplier code", "T1 Supplier name"]));
  check("Footer counts parts and relationships", (await page.locator("#supplierPartTable .pagination-toolbar").textContent()).includes("1 part · 6 relationships"));
  // + / − like the BOM tree panel
  const visible = () => page.locator("#supplierPartTable tbody tr").evaluateAll((trs) => trs.filter((tr) => !tr.hidden).length);
  await page.locator("#supplierPartTable tbody tr").nth(1).locator("[data-sb-toggle]").click();
  check("− on Savera closes its T2 rows", (await visible()) === 5);
  await page.locator("#supplierPartTable tbody tr").nth(0).locator("[data-sb-toggle]").click();
  check("− on the part closes the whole tree", (await visible()) === 1 && (await page.locator("#supplierPartTable tbody tr").nth(0).locator("[data-sb-toggle]").textContent()) === "+");
  await page.locator("#supplierPartTable tbody tr").nth(0).locator("[data-sb-toggle]").click();
  check("+ on the part opens every level again", (await visible()) === 7);
  check("Download Excel is the chain sheet", (await page.locator("#chainExcel").getAttribute("href")) === "/api/planning/supplier-chain.xlsx");
  // search narrows to chains that mention the text
  await page.fill("#supplierPartSearch", "Savera");
  await page.waitForFunction(() => document.querySelectorAll("#supplierPartTable tbody tr").length === 4);
  check("Search 'Savera' keeps the part, Savera and its 2 T2 rows", true);
  await page.fill("#supplierPartSearch", "");
  await page.waitForFunction(() => document.querySelectorAll("#supplierPartTable tbody tr").length === 7);
  // header filter on a block column (T2 Supplier name = cell 6)
  const t2Head = page.locator("#supplierPartTable thead tr:last-child th").nth(6);
  await t2Head.locator("summary").click();
  await t2Head.locator('[data-column-value][value="TATA Steel"]').check();
  await t2Head.locator("[data-column-apply]").click();
  check("Column filter on T2 Supplier name shows TATA rows only", (await page.locator("#supplierPartTable tbody tr:visible").count()) === 2);
  await page.click('[data-action="clear-table-filters"]');
  await page.waitForFunction(() => [...document.querySelectorAll("#supplierPartTable tbody tr")].filter((tr) => tr.offsetParent).length === 7);
  // column resize works under the two-row heading and is kept after a redraw
  await page.waitForTimeout(500);
  const head = page.locator("#supplierPartTable thead tr:last-child th").nth(1);
  await head.waitFor();
  const w0 = Math.round((await head.boundingBox()).width);
  await head.hover();
  const h = await head.locator(".column-resizer").boundingBox();
  await page.mouse.move(h.x + h.width / 2, h.y + h.height / 2);
  await page.mouse.down();
  await page.mouse.move(h.x + h.width / 2 + 80, h.y + h.height / 2, { steps: 6 });
  await page.mouse.up();
  const w1 = Math.round((await head.boundingBox()).width);
  check("A column under the group heading can be widened", Math.abs(w1 - w0 - 80) <= 8, `${w0} -> ${w1}`);
  await page.click('[data-supplier-tab="Supplier register"]');
  await page.click('[data-supplier-tab="Supplier/material coverage"]');
  await page.waitForSelector('#supplierPartTable tbody tr:has-text("TATA")');
  await page.waitForTimeout(200);
  check("Width kept after changing tab", Math.abs(Math.round((await page.locator("#supplierPartTable thead tr:last-child th").nth(1).boundingBox()).width) - w1) <= 3);
  await page.screenshot({ path: __dirname + "/supply_chain.png" });

  // T2 edit from the coverage: TATA under Savera to 60% keeps it under Savera's beam
  const tataSavera = page.locator("#supplierPartTable tbody tr").nth(3);
  await tataSavera.locator("[data-edit-supplier-part]", { hasText: "Edit T2" }).click();
  await page.waitForSelector("#partEditTable tbody tr");
  await page.locator('#partEditTable [data-f="sob_percent"]').fill("60");
  await page.click("#partEditSave");
  await page.waitForFunction(() => document.querySelectorAll("#supplierPartTable tbody tr")[3]?.cells[13].textContent === "60%");
  const after = await api("GET", "/api/planning/supplier-chain");
  check("Edited T2 stays under the same T1 part (Savera) and Supangita's TATA is unchanged", after.items[1].t1.supplier_name === "Savera" && after.items[1].t2.sob_percent === 60 && after.items[3].t2.sob_percent === 70, JSON.stringify(after.items.map((r) => [r.t1.supplier_name, r.t2.supplier_name, r.t2.sob_percent])));

  // "+ Add T2" on a T1 row: a further steel supplier under Savera's beam, linked to that T1
  const saveraRow = () => page.locator('#supplierPartTable tbody tr[data-tree-depth="1"]', { hasText: "Savera" });
  const addT2 = async () => {
    await saveraRow().locator("[data-add-under-part]").click();
    await page.waitForSelector("#underTable tbody tr");
    const row = page.locator("#underTable tbody tr").first();
    await row.locator('[data-pick-key="supplier"][data-pick-part="code"]').fill("SUP-2020");
    await row.locator('[data-pick-key="material"][data-pick-part="code"]').fill("FL001STEEL0.7");
    await page.waitForFunction(() => document.querySelector('#underTable [data-row-tier-material="material"]').value !== "");
    return row;
  };
  check("T1 rows have + Add T2, T2 rows have + Add T3", (await saveraRow().locator("[data-add-under-part]").textContent()) === "+ Add T2" && (await page.locator('#supplierPartTable tbody tr[data-tree-depth="2"] [data-add-under-part]').first().textContent()) === "+ Add T3");
  let row = await addT2();
  // Savera's steel: JGI 30% + TATA 60% (edited above) = 90% -> 10% left
  check("The window lists the T2 suppliers already there and the SOB left", (await page.locator("#modalBody").textContent()).includes("SOB left: 10%"), await page.locator("#modalBody p").textContent());
  await row.locator('[data-f="sob_percent"]').fill("20");
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.click("#underSave");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("Row 1"));
  check("Too much SOB says how much is left", (await page.locator("#toast").textContent()).includes("at most 10% is left"), await page.locator("#toast").textContent());
  row = page.locator("#underTable tbody tr").first();
  if (!(await page.locator("#underTable").count())) row = await addT2();
  await row.locator('[data-f="sob_percent"]').fill("10");
  await row.locator('[data-f="lead_time_days"]').fill("25");
  await page.click("#underSave");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("added under Savera"));
  await page.waitForSelector('#supplierPartTable tbody tr:has-text("Steel Trader")');
  const trader = (await api("GET", "/api/planning/supplier-chain?search=Steel Trader")).items.find((r) => r.t2 && r.t2.supplier_name === "Steel Trader");
  check("Added as T2 under Savera's beam (10%, lead 25)", trader && trader.t1.supplier_name === "Savera" && trader.t2.tier === "T2" && trader.t2.sob_percent === 10 && trader.t2.lead_time_days === 25 && trader.t2.parent_link_id === trader.t1.id, JSON.stringify(trader?.t2));
  check("Not added under Supangita", !(await api("GET", "/api/planning/supplier-chain?search=Steel Trader")).items.some((r) => r.t1?.supplier_name === "Supangita" && r.t2?.supplier_name === "Steel Trader"));

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
