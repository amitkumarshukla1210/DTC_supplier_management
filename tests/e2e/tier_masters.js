// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/tier_masters.js   (needs Node + Playwright).
// T1 / T2 / T3 Material Master tabs and picking codes into a supplier's material master.
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 950 } });
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
  await api("POST", "/api/planning/materials", { code: "XYZ-STEEL", name: "Steel XYZ", uom: "kg", unit_cost: 70, currency: "INR", category: "Raw material" });
  const supplier = (await api("POST", "/api/suppliers", { name: "Part Maker Pune", category: "Steel parts", country: "India" })).id;

  await page.goto(BASE);
  await page.click('[data-page="master"]');
  await page.waitForSelector("[data-bom-tab]");
  const tabs = await page.locator("[data-bom-tab]").allTextContents();
  check("Tabs: T1 Material Master (renamed), T2 and T3 Material Master", tabs.includes("T1 Material Master") && tabs.includes("T2 Material Master") && tabs.includes("T3 Material Master") && !tabs.includes("Material Master"), JSON.stringify(tabs));
  await page.click('[data-bom-tab="Material Master"]');
  check("T1 tab shows the T1 Material Master heading", (await page.locator("#pageContent").textContent()).includes("T1 Material Master"));

  // T2 tab: add, duplicate refused, edit, search, delete
  await page.click('[data-bom-tab="T2 Material Master"]');
  await page.waitForSelector("#tierMaterialTable");
  await page.click('[data-action="new-tier-material"]');
  await page.fill('#tierForm [name="code"]', "HR-COIL");
  await page.fill('#tierForm [name="name"]', "HR coil");
  await page.fill('#tierForm [name="unit_price"]', "55");
  await page.click("#tierForm button");
  await page.waitForFunction(() => document.querySelector("#tierMaterialTable").textContent.includes("HR-COIL"));
  check("Add a T2 code from the form", true);
  await page.click('[data-action="new-tier-material"]');
  await page.fill('#tierForm [name="code"]', "xyz-steel");
  await page.fill('#tierForm [name="name"]', "dup");
  await page.click("#tierForm button");
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("Material Master"));
  check("A code already in T1 is refused in T2", (await toastText()).includes("T1 Material Master"), await toastText());
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });
  await page.click('[data-tier-edit]');
  await page.fill('#tierForm [name="name"]', "HR coil 2mm");
  await page.click("#tierForm button");
  await page.waitForFunction(() => document.querySelector("#tierMaterialTable").textContent.includes("HR coil 2mm"));
  check("Edit a T2 code", true);
  await page.fill("#masterSearch", "nothing-like-this");
  await page.click("[data-execute-master-search]");
  await page.waitForFunction(() => document.querySelector("#tierMaterialTable").textContent.includes("No codes match"));
  check("Search the T2 master", true);
  await page.fill("#masterSearch", "");
  await page.click("[data-execute-master-search]");
  await page.waitForFunction(() => document.querySelector("#tierMaterialTable").textContent.includes("HR-COIL"));
  const t3 = await api("POST", "/api/planning/tier-materials", { tier: "T3", code: "ORE-1", name: "Iron ore", uom: "kg", unit_price: 8, currency: "INR" });
  void t3;

  // Supplier materials: pick from T2, own code, T1 code typed as own is refused
  await page.click('[data-page="suppliers"]');
  await page.click('[data-supplier-tab="Supplier materials"]');
  await page.selectOption("#smSupplier", String(supplier));
  await page.selectOption("#smFrom", "T2");
  await page.waitForFunction(() => [...document.querySelectorAll("#smPick option")].some((o) => o.textContent.includes("HR-COIL")));
  check("Own-code fields hidden while picking from a master", await page.locator('#smAddForm [name="code"]').isHidden());
  await page.fill('#smAddForm [name="unit_price"]', "60");
  await page.fill('#smAddForm [name="lead_time_days"]', "21");
  await page.fill('#smAddForm [name="buffer_days"]', "7");
  await page.click('#smAddForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#smTable").textContent.includes("HR-COIL"));
  const row = await page.locator('#smTable tr:has-text("HR-COIL") td').allTextContents();
  check("Picked T2 code shows From T2, its name, the supplier's own price, lead time and buffer", row[0] === "T2" && row[2] === "HR coil 2mm" && row[4] === "60" && row[6] === "21 days" && row[7] === "7 days", JSON.stringify(row.slice(0, 8)));
  await page.selectOption("#smFrom", "Own");
  await page.fill('#smAddForm [name="code"]', "XYZ-STEEL");
  await page.fill('#smAddForm [name="name"]', "Steel");
  await page.click('#smAddForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("pick it from"));
  check("Typing your T1 code as the supplier's own code is refused", (await toastText()).includes("T1 Material Master"), await toastText());
  await page.fill('#smAddForm [name="code"]', "PM-BOLT");
  await page.fill('#smAddForm [name="name"]', "M6 bolt");
  await page.click('#smAddForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#smTable").textContent.includes("PM-BOLT"));
  check("Add the supplier's own code", (await page.locator('#smTable tr:has-text("PM-BOLT") td').first().textContent()) === "Supplier's own");
  check("After saving, 'Code from' stays on the chosen source with its fields shown", (await page.locator("#smFrom").inputValue()) === "Own" && await page.locator('#smAddForm [name="code"]').isVisible() && await page.locator("#smPick").isHidden());
  // edit a picked entry: code / name read-only, terms editable
  await page.locator('#smTable tr:has-text("HR-COIL") [data-sm-edit]').click();
  check("Editing a picked code shows code and name read-only", (await page.locator('#smEditForm [name="code"]').count()) === 0);
  await page.fill('#smEditForm [name="unit_price"]', "62");
  await page.click("#smEditForm button");
  await page.waitForFunction(() => [...document.querySelectorAll("#smTable td")].some((td) => td.textContent === "62"));
  check("Supplier's price updated", true);
  await page.screenshot({ path: __dirname + "/tier_masters.png" });
  // T2 code in use cannot be deleted
  await page.click('[data-page="master"]');
  await page.click('[data-bom-tab="T2 Material Master"]');
  await page.waitForSelector("[data-tier-delete]");
  await page.click("[data-tier-delete]");
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("supplier's material master"));
  check("A T2 code used by a supplier cannot be deleted", true);

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((r) => !r.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
