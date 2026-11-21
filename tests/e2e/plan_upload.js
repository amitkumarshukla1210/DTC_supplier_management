// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/plan_upload.js   (needs Node + Playwright, and the project's .venv for openpyxl).
// Plan template (logo, header, Model Code / Model Name / 24 months DD-MM-YYYY) and uploads in one month on three
// plan dates (01-11, 10-11, 15-11-2026): each upload is the next revision.
const { chromium } = require("playwright");
const os = require("os"), path = require("path"), fs = require("fs"), { spawnSync } = require("child_process");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };
const python = process.env.PYTHON || path.join(__dirname, "..", "..", ".venv", "bin", "python");

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const api = async (method, p, data) => {
    const r = await page.request.fetch(BASE + p, { method, data, headers: { "content-type": "application/json" } });
    if (r.status() >= 300) throw new Error(`${method} ${p} -> ${r.status()} ${await r.text()}`);
    return r.json();
  };
  const toastText = async () => (await page.locator("#toast").textContent()) || "";
  await page.goto(BASE);
  await api("POST", "/api/auth/setup", { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" });
  const product = (await api("POST", "/api/planning/products", { code: "PL-FAN", name: "Fans" })).id;
  await api("POST", "/api/planning/models", { product_id: product, code: "PL12000", name: "Eva 8014" });
  await api("POST", "/api/planning/models", { product_id: product, code: "PL12001", name: "Serena 8014" });

  await page.goto(BASE);
  await page.click('[data-page="plans"]');
  await page.waitForSelector("#planUploadForm");
  check("Upload has one Plan date (calendar), no start-month field", (await page.locator('#planUploadForm input[type="date"][name="release_date"]').count()) === 1 && (await page.locator('#planUploadForm [name="issue_month"]').count()) === 0);
  await page.fill('#planUploadForm [name="release_date"]', "2026-11-01");
  await page.dispatchEvent('#planUploadForm [name="release_date"]', "change");
  const href = await page.getAttribute("#planTemplateLink", "href");
  check("Template link follows the plan date", href === "/api/planning/templates/plan.xlsx?release_date=2026-11-01", href);
  const templateFile = path.join(os.tmpdir(), `plan-${Date.now()}.xlsx`);
  fs.writeFileSync(templateFile, await (await page.request.get(BASE + href)).body());
  const filled = (qty) => {
    const out = `${templateFile}.${qty}.xlsx`;
    const made = spawnSync(python, ["-c", `import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1]); ws = wb.active
heads = [ws.cell(4, c).value for c in range(1, 28)]
print("HEAD", heads[:2], heads[2].strftime("%d-%m-%Y"), heads[25].strftime("%d-%m-%Y"), heads[26], ws.cell(4, 3).number_format, len(ws._images))
for r in (5, 6):
    for c in range(3, 27):
        ws.cell(r, c).value = int(sys.argv[3]) * (1 if r == 5 else 2)
wb.save(sys.argv[2])`, templateFile, out, String(qty)]);
    return { out, text: String(made.stdout) + String(made.stderr) };
  };
  const first = filled(100);
  check("Template: Model Code, Model Name, 24 months 01-11-2026 .. 01-10-2028 (DD-MM-YYYY), no Release date / Plan revision, logo", first.text.includes("HEAD ['Model Code', 'Model Name'] 01-11-2026 01-10-2028 None DD-MM-YYYY 1"), first.text);
  const uploadOn = async (day, file) => {
    await page.fill('#planUploadForm [name="release_date"]', day);
    await page.setInputFiles('#planUploadForm input[type="file"]', file);
    await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
    await page.click('#planUploadForm button[type="submit"]');
    await page.waitForFunction(() => /Plan saved|Row|plan/.test(document.querySelector("#toast")?.textContent || ""));
    return toastText();
  };
  const t1 = await uploadOn("2026-11-01", first.out);
  check("01-11-2026: Rev.00", t1.includes("Plan saved as Rev.00 for") && t1.includes("plan date 01-11-2026"), t1);
  const t2 = await uploadOn("2026-11-10", filled(200).out);
  check("10-11-2026: Rev.01", t2.includes("Plan saved as Rev.01") && t2.includes("plan date 10-11-2026"), t2);
  const t3 = await uploadOn("2026-11-15", filled(300).out);
  check("15-11-2026: Rev.02", t3.includes("Plan saved as Rev.02") && t3.includes("plan date 15-11-2026"), t3);
  await page.waitForSelector('#planHistory tbody tr:has-text("Rev.02")');
  const history = await page.locator("#planHistory tbody tr").evaluateAll((trs) => trs.map((tr) => [tr.cells[1].textContent, tr.cells[2].textContent, tr.cells[3].textContent]));
  check("Revision history: plan dates DD-MM-YYYY, 48 rows each", JSON.stringify(history) === JSON.stringify([["Rev.02", "15-11-2026", "48"], ["Rev.01", "10-11-2026", "48"], ["Rev.00", "01-11-2026", "48"]]), JSON.stringify(history));
  check("Uploaded on is DD-MM-YYYY too", /^\d{2}-\d{2}-\d{4} \d{2}:\d{2}$/.test(await page.locator("#planHistory tbody tr").first().locator("td").nth(4).textContent()));
  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
