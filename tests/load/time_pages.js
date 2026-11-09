// Time how long each screen takes to show its data with the big database (port 8766).
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8766";
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 900 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(BASE);
  await page.request.post(`${BASE}/api/auth/login`, { data: { username: "admin", password: "Secure-test-password-2026!" } });
  await page.goto(BASE);
  await page.waitForLoadState("networkidle");
  await page.waitForTimeout(1500);
  const timeIt = async (label, action, ready) => {
    const start = Date.now();
    await action();
    await page.waitForFunction(ready, null, { timeout: 120000 });
    const ms = Date.now() - start;
    console.log(`${(ms / 1000).toFixed(2).padStart(6)}s  ${label}`);
    return ms;
  };
  const rows = (sel) => `document.querySelectorAll('${sel} tbody tr').length`;
  await timeIt("Product Master (1,000 product rows)", () => page.click('[data-page="master"]'), () => document.querySelectorAll("#productMasterTable tbody tr").length >= 1000);
  await timeIt("Expand one product (+)", () => page.locator("#productMasterTable [data-tree-expand]").first().click(), () => document.querySelectorAll("#productMasterTable tr[data-bom-level1-row]").length > 0);
  await timeIt("Model BOM Level 1 (page of 50 models)", () => page.click('[data-bom-tab="Model BOM Level 1"]'), () => document.querySelectorAll("#masterTable tr[data-bom-level1-row]").length >= 900 && document.querySelector("[data-l1-page]"));
  await timeIt("Level 1: Next page", () => page.locator('[data-l1-page]:has-text("Next")').click(), () => document.querySelector(".pagination-toolbar span")?.textContent.includes("Models 51"));
  await timeIt("Level 2 Part BOM (300 parents)", () => page.click('[data-bom-tab="Level 2 Part BOM"]'), () => document.querySelectorAll("#masterTable tr[data-tree-group]").length >= 300);
  await timeIt("Level 3 Part BOM (300 parents)", () => page.click('[data-bom-tab="Level 3 Part BOM"]'), () => document.querySelectorAll("#masterTable tr[data-tree-group]").length >= 300);
  await timeIt("Material Master (page of 500 of 10,000)", () => page.click('[data-bom-tab="Material Master"]'), () => document.querySelector("#materialMasterTable .pagination-toolbar")?.textContent.includes("of 10,"));
  await page.fill("#masterSearch", "LM-0912");
  await timeIt("Material Master search over all codes", () => page.click("[data-execute-master-search]"), () => document.querySelector("#materialMasterTable .pagination-toolbar")?.textContent.includes("(search)"));
  await timeIt("Suppliers: relationship tab + material picker", () => page.click('[data-page="suppliers"]'), () => [...document.querySelectorAll("#supplierMaterialSelect option")].some((o) => o.textContent.includes("more: type to search")));
  await timeIt("Suppliers: coverage (500 of 9,400 links)", () => page.click('[data-supplier-tab="Supplier/material coverage"]'), () => document.querySelectorAll("#supplierPartTable tbody tr").length >= 500);
  await timeIt("Suppliers: register (1,000 suppliers)", () => page.click('[data-supplier-tab="Supplier register"]'), () => document.querySelectorAll("#supplierRegisterTable tbody tr").length >= 1000);
  await timeIt("Suppliers: Supplier BOM tab", () => page.click('[data-supplier-tab="Supplier BOM"]'), () => document.querySelector("#sbSupplier"));
  console.log(errors.length ? `PAGE ERRORS: ${errors.join(" | ")}` : "No page errors");
  await browser.close();
})().catch((e) => { console.error(e); process.exit(2); });
