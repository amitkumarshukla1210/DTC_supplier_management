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
  // actual production till date: plan of the month (latest revision Rev.02: Eva 300, Serena 600) beside the actual
  await page.fill('#actualHistoryForm [name="history_as_of"]', "2026-11-15");
  await page.click('#actualHistoryForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#actualPlanNote")?.textContent.includes("Rev.02"));
  const actualHeads = await page.locator("#actualHistoryTable thead th summary").allTextContents();
  check("Actual table: plan (latest revision), actual till date, OEM stock, balance, % achieved, as on", JSON.stringify(actualHeads) === JSON.stringify(["Model code", "Model name", "Plan qty (Rev.02)", "Actual (till date)", "OEM stock", "Balance", "% achieved", "As on"]), JSON.stringify(actualHeads));
  check("Plan note: Nov 2026 · plan Rev.02 of 15-11-2026", (await page.textContent("#actualPlanNote")).includes("plan Rev.02 of 15-11-2026"));
  const evaRow = page.locator('#actualHistoryTable tr[data-actual-model="PL12000"]');
  await evaRow.locator('[data-f="actual_quantity"]').fill("150");
  await evaRow.locator('[data-f="oem_stock"]').fill("20");
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.click("#actualSaveAll");
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("actual till 15-11-2026"));
  await page.waitForFunction(() => document.querySelector('#actualHistoryTable tr[data-actual-model="PL12000"]')?.cells[7].textContent === "15-11-2026");
  const eva = await evaRow.locator("td").allTextContents();
  check("Typed in the table: Eva plan 300, actual 150 -> balance 150, 50%, as on 15-11-2026", eva[2] === "300" && eva[5] === "150" && eva[6] === "50%" && eva[7] === "15-11-2026", JSON.stringify(eva));
  // Actual template for 25-11-2026: models of the month's plan; Serena typed in Excel, uploaded on 25-11
  await page.fill('#actualUploadForm [name="as_of"]', "2026-11-25");
  await page.dispatchEvent('#actualUploadForm [name="as_of"]', "change");
  const actualHref = await page.getAttribute("#actualTemplateLink", "href");
  check("Actual template follows the as-on date", actualHref === "/api/planning/actuals/template.xlsx?as_of=2026-11-25", actualHref);
  const actualFile = path.join(os.tmpdir(), `actual-${Date.now()}.xlsx`);
  fs.writeFileSync(actualFile, await (await page.request.get(BASE + actualHref)).body());
  const actualFilled = spawnSync(python, ["-c", `import sys, openpyxl
wb = openpyxl.load_workbook(sys.argv[1]); ws = wb.active
print("TOP", ws["B1"].value, "|", ws["B2"].value, "|", len(ws._images))
print("HEAD", [ws.cell(4, c).value for c in range(1, 6)])
print("ROWS", [[ws.cell(r, c).value for c in range(1, 6)] for r in (5, 6)])
ws.cell(6, 4).value, ws.cell(6, 5).value = 400, 50
wb.save(sys.argv[1])`, actualFile]);
  const out = String(actualFilled.stdout) + String(actualFilled.stderr);
  check("Template: logo, 'Design to Cost · Actual production', month and as-on date", out.includes("TOP Design to Cost · Actual production | Month: Nov-2026 · as on 25-11-2026 · plan Rev.02 of 15-11-2026.") && out.includes("| 1\n"), out);
  check("Template rows: the plan's models with Plan qty and the actual so far", out.includes("HEAD ['Model Code', 'Model Name', 'Plan qty (Rev.02)', 'Actual production (till date)', 'OEM stock']") && out.includes("ROWS [['PL12000', 'Eva 8014', 300, 150, 20], ['PL12001', 'Serena 8014', 600, None, None]]"), out);
  await page.setInputFiles('#actualUploadForm input[type="file"]', actualFile);
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.click('#actualUploadForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("actual till 25-11-2026"));
  await page.waitForFunction(() => document.querySelector('#actualHistoryTable tr[data-actual-model="PL12001"]')?.cells[7].textContent === "25-11-2026");
  const serena = await page.locator('#actualHistoryTable tr[data-actual-model="PL12001"] td').allTextContents();
  check("Uploaded: Serena plan 600, actual 400 -> balance 200, 66.7%, as on 25-11-2026", serena[5] === "200" && serena[6] === "66.7%" && (await page.locator('#actualHistoryTable tr[data-actual-model="PL12001"] [data-f="actual_quantity"]').inputValue()) === "400", JSON.stringify(serena));
  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
