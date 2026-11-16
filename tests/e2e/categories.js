// Run: start the app on port 8765 with an EMPTY test database, e.g.
//   SUPPLIER_DATABASE_URL=sqlite:///test-e2e.db uvicorn app:app --port 8765
// then: node tests/e2e/categories.js   (needs Node + Playwright).
// One category list: Product Master → Category tab (add, rename, delete) and the category
// dropdown (+ Add new category) in material, T2/T3 and supplier forms.
const { chromium } = require("playwright");
const BASE = "http://127.0.0.1:8765";
const results = [];
const check = (name, ok, detail = "") => { results.push({ name, ok: !!ok, detail }); console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? "  -- " + detail : ""}`); };
const MY_LIST = ["Raw material", "Finished assembly", "Electrical components", "Electronic PCB Assy", "Sheet Metal", "Wire Harness", "Plastic Raw material", "Steel Raw material", "Electronic Child Part", "Casting", "Machining", "Packing Material", "Fasteners", "Subassembly"];

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1600, height: 950 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  let promptAnswer = "";
  page.on("dialog", (d) => (d.type() === "prompt" ? d.accept(promptAnswer) : d.accept()));
  const toastText = async () => (await page.locator("#toast").textContent()) || "";
  await page.goto(BASE);
  await page.request.post(BASE + "/api/auth/setup", { data: { username: "admin", full_name: "Admin", password: "Secure-test-password-2026!" } });
  await page.goto(BASE);
  await page.click('[data-page="master"]');
  await page.waitForSelector("[data-bom-tab]");
  await page.click('[data-bom-tab="Category"]');
  await page.waitForSelector("#categoryTable tbody tr");
  const names = await page.locator("#categoryTable tbody tr td:first-child").allTextContents();
  check("Category tab lists your 14 categories", names.length === 14 && MY_LIST.every((n) => names.includes(n)), JSON.stringify(names));
  const machining = await page.locator('#categoryTable tbody tr:has-text("Machining") td').allTextContents();
  check("Old 'CNC machining' records now count under Machining (1 part, 1 supplier)", machining[1] === "1" && machining[3] === "1", JSON.stringify(machining));
  check("No categories outside the list after the agreed replacements", (await page.locator("#categoryUnlisted").textContent()).trim() === "");

  await page.fill('#categoryAddForm [name="name"]', "Rubber Parts");
  await page.click("#categoryAddForm button");
  await page.waitForFunction(() => document.querySelector("#categoryTable").textContent.includes("Rubber Parts"));
  check("Add a category", true);

  // material form: dropdown = the list + "+ Add new category"; a new name is saved to the list
  await page.click('[data-bom-tab="Material Master"]');
  await page.click('[data-action="new-material"]');
  await page.waitForSelector("#categorySelect");
  const options = await page.locator("#categorySelect option").allTextContents();
  check("Material form: category dropdown = your list + Add new", options.includes("Rubber Parts") && options.includes("Machining") && options[options.length - 1] === "+ Add new category" && !options.includes("CNC machining"), JSON.stringify(options));
  await page.fill('#modalForm [name="code"]', "CAT-BRG");
  await page.fill('#modalForm [name="name"]', "Ball bearing");
  await page.selectOption("#categorySelect", "__new__");
  check("'+ Add new category' shows a name box", await page.locator("#categoryNew").isVisible());
  await page.fill("#categoryNew", "Bearings");
  await page.click("#modalForm .modal-footer button");
  await page.waitForFunction(() => document.querySelector("#modalLayer").hidden);
  // the new category is offered at once in another page's form (T2 master)
  await page.click('[data-bom-tab="T2 Material Master"]');
  await page.click('[data-action="new-tier-material"]');
  await page.waitForSelector("#tierForm #categorySelect");
  check("New category is offered next time on another page (T2 form)", (await page.locator("#tierForm #categorySelect option").allTextContents()).includes("Bearings"));
  await page.evaluate(() => { document.querySelector("#modalLayer").hidden = true; });

  // suppliers: add and edit use the same dropdown
  await page.click('[data-page="suppliers"]');
  await page.waitForSelector("#supplierRegisterTable tbody tr");
  await page.click('[data-action="new-supplier"]');
  await page.waitForSelector("#modalForm #categorySelect");
  check("Add supplier: category dropdown from the list", (await page.locator("#modalForm #categorySelect option").allTextContents()).includes("Wire Harness"));
  await page.fill('#modalForm [name="name"]', "CAT Harness Works");
  await page.selectOption("#modalForm #categorySelect", "Wire Harness");
  await page.fill('#modalForm [name="country"]', "India");
  await page.click("#modalForm .modal-footer button");
  await page.waitForFunction(() => document.querySelector("#supplierRegisterTable")?.textContent.includes("CAT Harness Works"));
  check("Supplier added with category Wire Harness", (await page.locator('#supplierRegisterTable tbody tr:has-text("CAT Harness Works")').textContent()).includes("Wire Harness"));
  await page.locator('#supplierRegisterTable tbody tr:has-text("Koyo") [data-edit-supplier]').click();
  await page.waitForSelector("#supplierEditForm #categorySelect");
  check("Edit supplier: current category selected", (await page.inputValue("#supplierEditForm #categorySelect")) === "Electrical components");
  await page.selectOption("#supplierEditForm #categorySelect", "Electronic Child Part");
  await page.click("#supplierEditForm button");
  await page.waitForFunction(() => document.querySelector('#supplierRegisterTable').textContent.includes("Electronic Child Part"));
  check("Supplier category changed from the dropdown", true);

  // rename (records follow) and delete (refused while used)
  await page.click('[data-page="master"]');
  await page.click('[data-bom-tab="Category"]');
  await page.waitForSelector('#categoryTable tbody tr:has-text("Bearings")');
  promptAnswer = "Bearings & Bushes";
  await page.locator('#categoryTable tbody tr:has-text("Bearings") [data-category-rename]').click();
  await page.waitForFunction(() => document.querySelector("#categoryTable").textContent.includes("Bearings & Bushes"));
  check("Rename a category; its part follows", (await page.locator('#categoryTable tbody tr:has-text("Bearings & Bushes") td').nth(1).textContent()) === "1");
  await page.evaluate(() => { document.querySelector("#toast").textContent = ""; });
  await page.locator('#categoryTable tbody tr:has-text("Bearings & Bushes") [data-category-delete]').click();
  await page.waitForFunction(() => document.querySelector("#toast").textContent.includes("is used by"));
  check("Deleting a used category is refused with the reason", (await toastText()).includes("1 parts"), await toastText());
  await page.locator('#categoryTable tbody tr:has-text("Rubber Parts") [data-category-delete]').click();
  await page.waitForFunction(() => !document.querySelector("#categoryTable").textContent.includes("Rubber Parts"));
  check("Delete an unused category", true);
  await page.screenshot({ path: __dirname + "/categories.png", fullPage: true });

  check("No page errors", errors.length === 0, errors.join(" | "));
  await browser.close();
  const failed = results.filter((x) => !x.ok).length;
  console.log(`\n${results.length - failed}/${results.length} passed`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(2); });
