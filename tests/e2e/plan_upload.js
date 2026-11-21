// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/plan_upload.js   (needs Node + Playwright, and the project's .venv for openpyxl).
// Plan template (logo, header, 20 months DD-MM-YYYY) and uploading one revision at a time by its release date.
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
  await page.fill('#planUploadForm [name="issue_month"]', "2026-01");
  await page.dispatchEvent('#planUploadForm [name="issue_month"]', "change");
  await page.fill('#planUploadForm [name="release_date"]', "2026-01-15");
  await page.dispatchEvent('#planUploadForm [name="release_date"]', "change");
  const href = await page.getAttribute("#planTemplateLink", "href");
  check("Template link follows start month and release date", href === "/api/planning/templates/plan.xlsx?issue_month=2026-01&release_date=2026-01-15", href);
  const templateFile = path.join(os.tmpdir(), `plan-${Date.now()}.xlsx`);
  fs.writeFileSync(templateFile, await (await page.request.get(BASE + href)).body());
  // read the template, then build the user's sheet: three blocks Rev.00 / Rev.01 / Rev.02
  const made = spawnSync(python, ["-c", `import sys, openpyxl, datetime
wb = openpyxl.load_workbook(sys.argv[1]); ws = wb.active
heads = [ws.cell(4, c).value for c in range(1, 25)]
print("HEAD", heads[:4], heads[4].strftime("%d-%m-%Y"), heads[23].strftime("%d-%m-%Y"), ws.cell(4, 5).number_format, len(ws._images))
print("ROWS", [(ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value.strftime("%d-%m-%Y"), ws.cell(r, 4).value) for r in (5, 6)])
out = openpyxl.Workbook(); sheet = out.active
months = ["01-%02d-2026" % m for m in range(1, 11)]
for released, rev, eva, serena in (("01-01-2026", "Rev.00", 20000, 20000), ("15-01-2026", "Rev.01", 15000, 35000), ("17-01-2026", "Rev.02", 17000, 30000)):
    sheet.append([]); sheet.append([])
    sheet.append(["Model Code", "Model Name", "Release date", "Plan revision"] + months)
    sheet.append(["PL12000", "Eva 8014", released, rev] + [eva] * 10)
    sheet.append(["PL12001", "Serena 8014", released, rev] + [serena] * 10)
out.save(sys.argv[2])`, templateFile, templateFile + ".plans.xlsx"]);
  const lines = String(made.stdout);
  check("Template: header row Model Code / Model Name / Release date / Plan revision, months 01-01-2026 .. 01-08-2027 (20), DD-MM-YYYY, logo", lines.includes("HEAD ['Model Code', 'Model Name', 'Release date', 'Plan revision'] 01-01-2026 01-08-2027 DD-MM-YYYY 1"), lines + made.stderr);
  check("Template rows: every model with the release date and next revision", lines.includes("ROWS [('PL12000', 'Eva 8014', '15-01-2026', 'Rev.00'), ('PL12001', 'Serena 8014', '15-01-2026', 'Rev.00')]"), lines);
  // upload the sheet with release date 15-01-2026: Rev.01 only
  await page.setInputFiles('#planUploadForm input[type="file"]', templateFile + ".plans.xlsx");
  await page.click('#planUploadForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("Plan saved"));
  check("Upload loads the block released on the picked date", (await toastText()).includes("Plan saved as Rev.01 for") && (await toastText()).includes("released 15-01-2026"), await toastText());
  await page.waitForSelector('#planHistory tbody tr:has-text("Rev.01")');
  const row = await page.locator('#planHistory tbody tr:has-text("Rev.01") td').allTextContents();
  check("Revision history shows the release date (DD-MM-YYYY)", row[2] === "15-01-2026" && row[3] === "20", JSON.stringify(row));
  check("Revision date also DD-MM-YYYY", /^\d{2}-\d{2}-\d{4} \d{2}:\d{2}$/.test(row[4]), row[4]);
  // a date that is not in the file lists the dates there
  await page.fill('#planUploadForm [name="release_date"]', "2026-01-16");
  await page.setInputFiles('#planUploadForm input[type="file"]', templateFile + ".plans.xlsx");
  await page.click('#planUploadForm button[type="submit"]');
  await page.waitForFunction(() => document.querySelector("#toast")?.textContent.includes("Plans in the file"));
  check("Another date: the file's plans are listed", (await toastText()).includes("01-01-2026 (Rev.00), 15-01-2026 (Rev.01), 17-01-2026 (Rev.02)"), await toastText());
  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
