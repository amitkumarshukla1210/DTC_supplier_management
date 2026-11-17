// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/filter_close.js   (needs Node + Playwright).
// A column filter opened and left without a choice closes when you click elsewhere or press Esc.
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(BASE);
  await page.request.post(BASE + "/api/auth/setup", { data: { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" } });
  await page.goto(BASE);
  await page.click('[data-page="suppliers"]');
  await page.waitForSelector("#supplierRegisterTable tbody tr");
  const rows = await page.locator("#supplierRegisterTable tbody tr:visible").count();
  const filter = (n) => page.locator("#supplierRegisterTable thead th").nth(n).locator("details");
  const isOpen = (n) => filter(n).evaluate((d) => d.open);
  await filter(3).locator("summary").click();
  check("Category filter opens", await isOpen(3));
  await page.locator("h1, .page-title, #pageContent h2").first().click();
  check("Click outside closes it (nothing chosen)", !(await isOpen(3)));
  check("No rows hidden", (await page.locator("#supplierRegisterTable tbody tr:visible").count()) === rows);
  await filter(3).locator("summary").click();
  await page.keyboard.press("Escape");
  check("Esc closes it", !(await isOpen(3)));
  await filter(3).locator("summary").click();
  await filter(4).locator("summary").click();
  check("Opening another filter closes the first", !(await isOpen(3)) && await isOpen(4));
  await filter(4).locator(".multi-search").fill("x");
  check("Typing in its search box keeps it open", await isOpen(4));
  await filter(4).locator("summary").click();
  check("Its own heading closes it", !(await isOpen(4)));
  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
