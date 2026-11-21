const state = {
  page: "dashboard",
  user: null,
  products: [],
  models: [],
  materials: [],
  suppliers: [],
  supplierParts: [],
  supplierOffset: 0,
  supplierTotal: 0,
  supplierSearch: "",
  plans: [],
  planEditor: null,
  report: [],
  reportTotal: 0,
  reportOffset: 0,
  reportQuery: null,
  bomTab: "Product Master",
  selectedModel: "",
  selectedModelIds: [],
  selectedBomParent: "",
  selectedBomParentIds: [],
  bomParentOptions: [],
  bomParentAll: [],
  materialCategories: [],
  supplierTab: "Supplier register",
  smSupplier: "",
  sbSupplier: "",
  sbProduct: "",
  sbView: null,
  bomOffset: 0,
  searchTimer: null,
  workspace: "trial",
  toastTimer: null,
};

const $ = (selector, root = document) => root.querySelector(selector);
const ROW_CLICK_SELECTOR = "[data-product-model-row], [data-product-id], [data-bom-level1-row], [data-bom-level2-row], [data-tree-group], tr[data-edit-material]";
const content = $("#pageContent");
const modalLayer = $("#modalLayer");
const modalBody = $("#modalBody");

document.querySelectorAll("[data-icon]").forEach((node) => {
  node.innerHTML = ({
    grid: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3.5" y="3.5" width="7" height="7" rx="1.4"/><rect x="13.5" y="3.5" width="7" height="7" rx="1.4"/><rect x="3.5" y="13.5" width="7" height="7" rx="1.4"/><rect x="13.5" y="13.5" width="7" height="7" rx="1.4"/></svg>',
    users: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="9" cy="8" r="3.5"/><path d="M2.8 20a6.2 6.2 0 0 1 12.4 0M16 5a3.5 3.5 0 0 1 0 6.8M18 14a5.8 5.8 0 0 1 3.2 5.2"/></svg>',
    box: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m12 3 9 4.7v8.6L12 21l-9-4.7V7.7L12 3Z"/><path d="m3.3 7.9 8.7 4.6 8.7-4.6M12 12.5V21"/></svg>',
    chart: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M3 3v18h18M8 15l4-4 3 3 6-7"/></svg>',
    calendar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/></svg>',
    layers: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m12 3 9 5-9 5-9-5 9-5Z"/><path d="m3 12 9 5 9-5M3 16l9 5 9-5"/></svg>',
    money: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="M15.5 8.5c-.7-.8-1.8-1.2-3.2-1.2-1.6 0-2.8.8-2.8 2 0 3.3 6.2 1.1 6.2 4.5 0 1.3-1.2 2.4-3.2 2.4-1.5 0-2.8-.5-3.7-1.5M12 5.5v13"/></svg>',
    refresh: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M20 7v5h-5M4 17v-5h5"/><path d="M5.5 9A7 7 0 0 1 18 6l2 2M4 16l2 2a7 7 0 0 0 12.5-3"/></svg>',
    close: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg>',
  })[node.dataset.icon] || "";
});
$("#todayLabel").textContent = new Intl.DateTimeFormat("en", { weekday: "short", month: "short", day: "numeric", year: "numeric" }).format(new Date());

function esc(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[char]);
}
function monthNow() {
  return new Date().toISOString().slice(0, 7);
}

function monthOffset(value, offset) {
  const [year, month] = value.split("-").map(Number);
  const date = new Date(year, month - 1 + offset, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthLabel(value) {
  if (!value) return "—";
  return new Intl.DateTimeFormat("en", { month: "short", year: "numeric" }).format(new Date(`${value}-01T00:00:00`));
}

function number(value, digits = 2) {
  return Number(value || 0).toLocaleString(undefined, { maximumFractionDigits: digits });
}

function money(value, currency = "USD") {
  try {
    return new Intl.NumberFormat("en", { style: "currency", currency, maximumFractionDigits: 2 }).format(Number(value || 0));
  } catch {
    return `${currency} ${number(value)}`;
  }
}
async function api(path, options = {}) {
  const response = await fetch(path, {
    ...options,
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
  });
  if (response.status === 204) return null;
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const detail = payload.detail;
    throw new Error(Array.isArray(detail) ? detail.map((item) => item.msg).join(", ") : (detail || `Request failed (${response.status})`));
  }
  return payload;
}

async function uploadExcel(path, file) {
  if (!file) throw new Error("Choose an Excel workbook first.");
  const response = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": file.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
    body: file,
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.detail || `Upload failed (${response.status})`);
  return payload;
}

function toast(message, error = false) {
  const node = $("#toast");
  node.textContent = message;
  node.classList.toggle("error", error);
  node.classList.add("show");
  clearTimeout(state.toastTimer);
  state.toastTimer = setTimeout(() => node.classList.remove("show"), 3500);
}

function heading(title, description, actions = "") {
  return `<div class="page-heading"><div><p class="eyebrow">SUPPLY CHAIN MANAGEMENT</p><h1>${esc(title)}</h1><p>${esc(description)}</p></div><div class="heading-actions">${actions}</div></div>`;
}

function table(headers, rows, empty = "No data found.") {
  const template = document.createElement("template");
  template.innerHTML = rows;
  const dataRows = [...template.content.querySelectorAll("tbody tr, tr")].filter((row) => !row.querySelector(".table-empty"));
  const headings = headers.map((header, index) => {
    if (!header) return '<th class="tree-column-heading"><span class="column-resizer" aria-hidden="true"></span></th>';
    const values = [...new Set(dataRows.map((row) => row.cells[index]?.textContent.trim() || "").filter(Boolean))].sort((a, b) => a.localeCompare(b));
    const options = values.map((value) => `<label><input type="checkbox" data-column-value="true" value="${esc(value)}" />${esc(value)}</label>`).join("");
    return `<th><div class="table-heading"><details class="multi-filter table-column-filter"><summary aria-label="Filter ${esc(header)}">${esc(header)}</summary><div class="multi-popover"><div class="column-sort"><button type="button" class="text-button" data-column-sort="asc">Sort A → Z</button><button type="button" class="text-button" data-column-sort="desc">Sort Z → A</button></div><input class="form-control multi-search" type="search" placeholder="Search ${esc(header.toLowerCase())}" /><div class="column-filter-actions"><button type="button" class="text-button" data-column-select-all>Select all</button><button type="button" class="text-button" data-column-clear>Clear all</button></div><div class="multi-options">${options || '<span class="subtle">No values</span>'}</div><button type="button" class="button button-small button-primary column-filter-apply" data-column-apply>Apply</button></div></details><span class="column-resizer" aria-hidden="true"></span></div></th>`;
  }).join("");
  return `<div class="table-wrap"><table><thead><tr>${headings}</tr></thead><tbody>${rows || `<tr><td class="table-empty" colspan="${headers.length}">${esc(empty)}</td></tr>`}</tbody></table></div>`;
}

function tableExportLinks(kind, parameters = {}) {
  const query = new URLSearchParams(Object.entries(parameters).filter(([, value]) => value !== undefined && value !== null && value !== ""));
  const suffix = query.size ? `?${query}` : "";
  return `<a class="button button-small" href="/api/planning/tables/${encodeURIComponent(kind)}/xlsx${suffix}">Download Excel</a><a class="button button-small" href="/api/planning/tables/${encodeURIComponent(kind)}/pdf${suffix}">Download PDF</a>`;
}

// BOM tabs download the full tree breakdown (Excel opens with +/- outline buttons).
// The link is built on click so it uses the models / parent parts selected at that moment.
function bomTreeExportLinks() {
  return '<a class="button button-small" href="#" data-bom-tree-export="xlsx">Download Excel</a><a class="button button-small" href="#" data-bom-tree-export="pdf">Download PDF</a>';
}

function bomTreeExportUrl(fileFormat) {
  const query = new URLSearchParams({ tab: state.bomTab });
  if (state.bomTab === "Model BOM Level 1") state.selectedModelIds.forEach((id) => query.append("model_id", String(id)));
  if (state.bomTab === "Level 2 Part BOM" || state.bomTab === "Level 3 Part BOM") state.selectedBomParentIds.forEach((id) => query.append("parent_item_id", String(id)));
  return `/api/planning/bom-tree/${fileFormat}?${query}`;
}

function panelHeader(title, description = "") {
  return `<header class="panel-header brand-panel-heading"><img src="/static/logo.png" alt="" /><div><h2>Supply Chain Management · ${esc(title)}</h2>${description ? `<p>${esc(description)}</p>` : ""}</div></header>`;
}

function modelPickerMarkup() {
  const selectedIds = state.selectedModelIds.map(String);
  const allSelected = state.models.length > 0 && selectedIds.length === state.models.length;
  const summary = !selectedIds.length || allSelected ? "All models" : `${selectedIds.length} model${selectedIds.length === 1 ? "" : "s"} selected`;
  return `<details class="model-picker"><summary id="modelPickerSummary">${summary}</summary><div class="multi-popover model-picker-popover"><input class="form-control" type="search" data-model-search placeholder="Search models" /><div class="column-filter-actions"><button class="text-button" type="button" data-model-select-all>Select all</button><button class="text-button" type="button" data-model-clear>Clear all</button></div><div class="multi-options">${state.models.map((model) => `<label><input type="checkbox" data-model-option value="${model.id}" ${selectedIds.includes(String(model.id)) ? "checked" : ""} />${esc(model.code)} · ${esc(model.name)}</label>`).join("")}</div><button class="button button-small button-primary model-picker-apply" type="button" data-model-apply>Apply models</button></div></details>`;
}

function bomParentPickerMarkup() {
  const selectedIds = new Set(state.selectedBomParentIds.map(String));
  const parentLabel = state.bomTab === "Level 2 Part BOM" ? "Level 1 parent parts" : "Level 2 parent parts";
  const allSelected = selectedIds.size === state.bomParentOptions.length;
  const summary = allSelected ? `All ${parentLabel.toLowerCase()}` : `${selectedIds.size} parent part${selectedIds.size === 1 ? "" : "s"} selected`;
  return `<details class="model-picker bom-parent-picker" id="bomParentPicker"><summary id="bomParentPickerSummary">${esc(summary)}</summary><div class="multi-popover model-picker-popover"><input class="form-control" type="search" data-parent-search placeholder="Search parent parts" /><div class="column-filter-actions"><button class="text-button" type="button" data-parent-select-all>Select all</button><button class="text-button" type="button" data-parent-clear>Clear all</button></div><div class="multi-options">${state.bomParentOptions.map((item) => `<label><input type="checkbox" data-parent-option value="${item.id}" ${selectedIds.has(String(item.id)) ? "checked" : ""} />${esc(item.code)} · ${esc(item.name)}</label>`).join("")}</div><button class="button button-small button-primary model-picker-apply" type="button" data-parent-apply>Apply parent parts</button></div></details>`;
}

function field(label, name, type = "text", options = {}) {
  if (options.choices) {
    return `<label class="field"><span>${esc(label)}</span><select class="form-control" name="${esc(name)}" ${options.required ? "required" : ""}>${options.choices.map(([value, text]) => `<option value="${esc(value)}">${esc(text)}</option>`).join("")}</select></label>`;
  }
  return `<label class="field"><span>${esc(label)}</span><input class="form-control" name="${esc(name)}" type="${type}" ${options.required ? "required" : ""} ${options.min !== undefined ? `min="${options.min}"` : ""} ${options.step ? `step="${options.step}"` : ""} ${options.maxlength ? `maxlength="${options.maxlength}"` : ""} ${options.value !== undefined ? `value="${esc(options.value)}"` : ""} ${options.placeholder ? `placeholder="${esc(options.placeholder)}"` : ""} /></label>`;
}

function categoryField(selected) {
  const names = [...state.materialCategories];
  const outside = selected && !names.some((name) => name.toLowerCase() === String(selected).toLowerCase());
  return `<label class="field"><span>Category</span><select class="form-control" name="category_choice" id="categorySelect" required>${!selected ? '<option value="" selected disabled>Select category</option>' : ""}${outside ? `<option value="" selected disabled>${esc(selected)} (not in list: choose one)</option>` : ""}${names.map((name) => `<option value="${esc(name)}" ${name.toLowerCase() === String(selected).toLowerCase() ? "selected" : ""}>${esc(name)}</option>`).join("")}<option value="__new__">+ Add new category</option></select><input class="form-control category-new" id="categoryNew" name="category_new" placeholder="Type new category name" minlength="2" maxlength="80" hidden /></label>`;
}

// Category dropdown (the one list from Product Master → Category) with "+ Add new category".
// A record whose old category is not in the list shows it marked, until you pick one from the list.
function bindCategoryField(root = document) {
  const choice = root.querySelector("#categorySelect");
  const fresh = root.querySelector("#categoryNew");
  if (!choice || !fresh) return;
  const sync = () => {
    const adding = choice.value === "__new__";
    fresh.hidden = !adding;
    fresh.required = adding;
    if (adding) fresh.focus();
  };
  choice.addEventListener("change", sync);
  sync();
}
function takeCategory(values) {
  values.category = values.category_choice === "__new__" ? (values.category_new || "").trim() : values.category_choice;
  const known = state.materialCategories.find((name) => name.toLowerCase() === String(values.category).toLowerCase());
  if (known) values.category = known;
  delete values.category_choice;
  delete values.category_new;
  return values;
}
async function reloadCategories() {
  try { state.materialCategories = (await api("/api/planning/material-categories")).items; } catch { /* keep the old list */ }
}

function unitField(label = "Unit", name = "uom", selected = "ea") {
  const units = [...new Set([...state.materials.map((material) => material.uom), "ea", "PCS", "Nos", "kg", "g", "grm", "m", "mtr", "mm", "cm", "L", "ml", "set", "pair", "roll", "box", "pack", "lot"].filter(Boolean))];
  return `<label class="field"><span>${esc(label)}</span><select class="form-control" name="${esc(name)}" required>${units.map((unit) => `<option value="${esc(unit)}" ${unit === selected ? "selected" : ""}>${esc(unit)}</option>`).join("")}</select></label>`;
}

function workspaceSelect() {
  return `<select class="form-control" name="workspace"><option value="trial" ${state.workspace === "trial" ? "selected" : ""}>Trial</option><option value="production" ${state.workspace === "production" ? "selected" : ""}>Production</option></select>`;
}

function bindMaterialSearch(searchId, selectId, nameId = null, extra = {}) {
  const search = $(`#${searchId}`);
  const select = $(`#${selectId}`);
  const nameControl = nameId ? $(`#${nameId}`) : null;
  const nameSelect = nameControl?.tagName === "SELECT" ? nameControl : null;
  function renderMaterialOptions(items) {
    const previousValue = select.value;
    select.innerHTML = `<option value="">Select part code</option>${items.map((item) => `<option value="${item.id}" data-code="${esc(item.code)}" data-name="${esc(item.name)}" data-uom="${esc(item.uom)}">${nameSelect ? esc(item.code) : nameControl ? esc(item.code) : `${esc(item.code)} · ${esc(item.name)}`}</option>`).join("")}`;
    select.value = items.some((item) => String(item.id) === previousValue) ? previousValue : "";
    if (nameSelect) {
      nameSelect.innerHTML = `<option value="">Select part name</option>${items.map((item) => `<option value="${item.id}" data-code="${esc(item.code)}" data-name="${esc(item.name)}" data-uom="${esc(item.uom)}">${esc(item.name)}</option>`).join("")}`;
      nameSelect.value = select.value;
    } else if (nameControl) {
      const selected = items.find((item) => String(item.id) === select.value);
      nameControl.value = selected?.name || "";
    }
  }
  async function run() {
    try {
      const query = new URLSearchParams({ offset: "0", limit: "100", search: search.value });
      if (extra.purchased) query.set("purchased", "true");
      const result = await api(`/api/planning/materials?${query}`);
      renderMaterialOptions(result.items);
      if (extra.purchased && result.total > result.items.length) {
        select.insertAdjacentHTML("beforeend", `<option value="" disabled>… ${(result.total - result.items.length).toLocaleString()} more: type to search</option>`);
      }
      select.dispatchEvent(new Event("change"));
    } catch (error) { toast(error.message, true); }
  }
  search.addEventListener("input", () => {
    clearTimeout(state.searchTimer);
    state.searchTimer = setTimeout(run, 250);
  });
  if (extra.loadNow) run();
}

function authScreen(setupRequired) {
  $("#appShell").hidden = true;
  const title = setupRequired ? "Set up your Admin account" : "Sign in to your workspace";
  const submitText = setupRequired ? "Create Admin account" : "Sign in";
  $("#authView").hidden = false;
  $("#authView").innerHTML = `<form id="authForm" class="auth-card">
    <img src="/static/logo.png" alt="Design to Cost logo" />
    <p class="eyebrow">SUPPLY CHAIN MANAGEMENT</p><h1>${title}</h1>
    <p class="subtle">${setupRequired ? "Create the first Admin account. You can add other Admins and users after setup." : "Use your assigned workspace account."}</p>
    ${setupRequired ? field("Full name", "full_name", "text", { required: true }) : ""}
    ${field("Username", "username", "text", { required: true })}
    ${field("Password", "password", "password", { required: true })}
    ${setupRequired ? '<small class="field-hint">Use at least 12 characters. Passwords are stored as salted hashes.</small>' : ""}
    <div id="authError" class="form-error"></div><button class="button button-primary" type="submit">${submitText}</button>
  </form>`;
  $("#authForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    const error = $("#authError");
    try {
      state.user = (await api(setupRequired ? "/api/auth/setup" : "/api/auth/login", { method: "POST", body: JSON.stringify(data) })).user;
      $("#authView").hidden = true;
      $("#appShell").hidden = false;
      initUserShell();
      await renderPage("dashboard");
    } catch (requestError) {
      error.textContent = requestError.message;
      error.classList.add("show");
    }
  });
}

function initUserShell() {
  $("#userRole").textContent = state.user.role;
  $("#userName").textContent = state.user.full_name;
  $("#usersNav").hidden = state.user.role !== "Admin";
  const supplierControl = ["Admin", "HOD", "Sourcing", "Supply Chain"].includes(state.user.role);
  document.querySelectorAll('[data-page="suppliers"], [data-page="purchase-orders"], [data-page="business"]').forEach((item) => { item.hidden = !supplierControl; });
  document.querySelectorAll("[data-icon]").forEach((node) => {
    const key = node.dataset.icon;
    if (node.innerHTML.trim()) return;
    node.innerHTML = key === "refresh" ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M20 7v5h-5M4 17v-5h5"/></svg>' : "";
  });
}

function setNav(page) {
  state.page = page;
  document.querySelectorAll(".nav-link").forEach((link) => link.classList.toggle("active", link.dataset.page === page));
}

async function loadCoreData() {
  const [products, models, materials] = await Promise.all([
    api("/api/planning/products"), api("/api/planning/models"), api("/api/planning/materials?offset=0&limit=500"),
  ]);
  try { state.materialCategories = (await api("/api/planning/material-categories")).items; } catch { state.materialCategories = []; }
  state.products = products;
  state.models = models;
  state.materials = materials.items;
}

async function renderPage(page = state.page) {
  setNav(page);
  content.dataset.view = page;
  content.innerHTML = '<div class="loading-state"><span class="spinner"></span>Loading your workspace…</div>';
  try {
    if (page === "dashboard") return await renderDashboard();
    if (page === "plans") return await renderPlans();
    if (page === "master") return await renderMaster();
    if (page === "suppliers") return await renderSupplierControl();
    if (page === "purchase-orders") return await renderPurchaseOrders();
    if (page === "reports") return await renderReports();
    if (page === "business") return await renderSupplierBusiness();
    if (page === "users") return await renderUsers();
  } catch (error) {
    content.innerHTML = `<section class="panel"><div class="panel-body"><p class="primary-cell">Could not load this view</p><p class="subtle">${esc(error.message)}</p><button class="button" data-action="refresh">Try again</button></div></section>`;
  }
}

async function renderDashboard() {
  const [models, materials, plans] = await Promise.all([
    api("/api/planning/models"), api("/api/planning/materials?offset=0&limit=1"), api(`/api/planning/plans?workspace=${state.workspace}`),
  ]);
  const cards = [
    ["Products & models", `${state.products.length} / ${models.length}`, "Product and model master"],
    ["Material master", number(materials.total, 0), "Searchable, paginated material list"],
    ["Plan revisions", number(plans.length, 0), `Stored ${state.workspace} monthly cycles`],
    ["Plan horizon", "24 months", "Each monthly upload creates a preserved revision"],
  ];
  content.innerHTML = `${heading("Monthly supply planning", "Plan demand by model, explode the BOM to material, coordinate supplier stock and act on shortages or surplus.", `<button class="button button-primary" data-page="plans">Open monthly planning</button>`)}
    <section class="metric-grid">${cards.map(([label, value, note]) => `<article class="metric-card"><div class="metric-top">${esc(label)}</div><div class="metric-value">${esc(value)}</div><div class="metric-note">${esc(note)}</div></article>`).join("")}</section>
    <section class="dashboard-grid">
      <article class="panel">${panelHeader("Monthly planning cycle", "Every upload is retained as Rev.00, Rev.01, Rev.02…")}
        <div class="panel-body"><ol class="workflow-list"><li>Maintain product, model, material codes, names, units and nested BOM quantities.</li><li>Upload each month’s 24-month model plan and actual production / OEM stock.</li><li>Upload supplier stock and manage supplier share, lead time, MOQ, buffer and price.</li><li>Review shortage / excess by supplier, then adjust POs or plan transfers.</li></ol></div>
      </article>
      <article class="panel">${panelHeader("Trial and production", "Keep historical trial data separate from production data")}
        <div class="panel-body"><label class="field"><span>Active workspace</span><select id="workspacePicker" class="form-control"><option value="trial" ${state.workspace === "trial" ? "selected" : ""}>Trial</option><option value="production" ${state.workspace === "production" ? "selected" : ""}>Production</option></select></label>
        <p class="subtle">Admin can clear trial imports without deleting production records. Historical records are retained until an Admin removes trial data.</p>
        <button class="button button-primary" data-page="reports">Open shortage &amp; excess report</button></div>
      </article>
    </section>`;
  $("#workspacePicker").addEventListener("change", (event) => {
    state.workspace = event.target.value;
    renderPage("dashboard");
  });
}

async function renderPlans() {
  await loadCoreData();
  state.planEditor = null;
  const month = monthNow();
  state.plans = await api(`/api/planning/plans?workspace=${state.workspace}`);
  const actuals = await api(`/api/planning/actuals?period=${month}&workspace=${state.workspace}`);
  const ddmmyyyy = (iso) => (iso ? iso.split("-").reverse().join("-") : "");
  const today = new Date().toISOString().slice(0, 10);
  const planRows = state.plans.map((plan) => `<tr class="clickable-row" data-plan-id="${plan.id}" tabindex="0" title="Open ${esc(plan.revision)} to review or edit monthly quantities"><td>${monthLabel(plan.issue_month)}</td><td>${esc(plan.revision)}</td><td>${esc(ddmmyyyy(plan.release_date))}</td><td>${number(plan.row_count, 0)}</td><td>${(() => { const d = new Date(plan.created_at); const two = (n) => String(n).padStart(2, "0"); return `${two(d.getDate())}-${two(d.getMonth() + 1)}-${d.getFullYear()} ${two(d.getHours())}:${two(d.getMinutes())}`; })()}</td><td>${esc(plan.workspace)}</td></tr>`).join("");
  content.innerHTML = `${heading("Monthly plans & actuals", "Import a monthly 24-month plan snapshot. Revisions are immutable; missing model/month rows are treated as zero demand.", `<a class="button" id="planTemplateLink" href="/api/planning/templates/plan.xlsx?issue_month=${month}&release_date=${today}">Plan template</a><a class="button" href="/api/planning/templates/actuals.xlsx">Actuals template</a>`)}
    <div class="planning-grid">
      <section class="panel">${panelHeader("Upload plan revision", "Plan template: logo and header, then Model Code, Model Name, Release date, Plan revision and 20 months (DD-MM-YYYY), one row per model. One revision per upload: pick its release date; if the file holds several blocks (Rev.00, Rev.01 ...), the block with that release date is loaded, and its first month is the plan start. Older sheets (one model per row, one month per column, from the start month) are also accepted.")}
        <form id="planUploadForm" class="panel-body form-grid">
          ${field("Plan start month", "issue_month", "month", { required: true, value: month })}
          ${field("Plan release date", "release_date", "date", { required: true, value: today })}
          <label class="field"><span>Workspace</span>${workspaceSelect()}</label>
          <label class="field full"><span>Plan workbook (.xlsx)</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label>
          <button class="button button-primary" type="submit">Upload as next revision</button>
        </form></section>
      <section class="panel">${panelHeader("Upload actual production & OEM stock", "Columns: model_code, actual_quantity, oem_stock. Actuals update that model and month; plans remain unchanged.")}
        <form id="actualUploadForm" class="panel-body form-grid">
          ${field("Production month", "period", "month", { required: true, value: month })}
          <label class="field"><span>Workspace</span>${workspaceSelect()}</label>
          <label class="field full"><span>Actual production workbook (.xlsx)</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label>
          <button class="button button-primary" type="submit">Upload actuals</button>
        </form></section>
    </div>
    <section class="panel table-panel">${panelHeader("Actual production & OEM stock history", "Actual production and OEM stock are retained by model and month; re-uploading a month updates that month's values.")}
        <form id="actualHistoryForm" class="table-toolbar"><div class="toolbar-group">${field("Production month", "history_period", "month", { value: month })}</div><button class="button button-small" type="submit">View month</button><span id="actualHistoryExports">${tableExportLinks("actuals", { period: month, workspace: state.workspace })}</span><div class="toolbar-group" id="actualFilters"></div><button class="button button-small" type="button" data-action="clear-table-filters">Clear all filters</button></form>
        <div id="actualHistoryTable">${table(["Model code", "Model name", "Actual production", "OEM stock"], actuals.map((row) => `<tr><td>${esc(row.model_code)}</td><td>${esc(row.model_name)}</td><td>${number(row.actual_quantity)}</td><td>${number(row.oem_stock)}</td></tr>`).join(""), `No actual production data for ${monthLabel(month)}.`)}</div>
    </section>
    <section id="planRevisionEditor" class="panel table-panel" hidden></section>
    <section class="panel table-panel">${panelHeader("Revision history", "Each issue month begins at Rev.00. Prior uploads are preserved for plan-versus-actual review.")}
      <div class="table-toolbar"><div class="toolbar-group" id="planFilters"></div><div class="toolbar-group">${tableExportLinks("plans", { workspace: state.workspace })}<button class="button button-small" data-action="clear-table-filters" data-table="planHistory">Clear all filters</button></div></div>
      <div id="planHistory">${table(["Issue month", "Revision", "Release date", "Rows", "Revision date", "Workspace"], planRows, "No monthly plans uploaded yet.")}</div>
    </section>`;
  $("#planUploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const data = new FormData(form);
    const uploadMonth = data.get("issue_month");
    state.workspace = data.get("workspace");
    try {
      const result = await uploadExcel(`/api/planning/plans/${uploadMonth}/upload?workspace=${state.workspace}&release_date=${data.get("release_date") || ""}`, data.get("workbook"));
      toast(`Plan saved as ${result.revision} for ${monthLabel(result.issue_month)}${result.release_date ? `, released ${ddmmyyyy(result.release_date)}` : ""}`);
      await renderPlans();
    } catch (error) { toast(error.message, true); }
  });
  const templateHref = () => { $("#planTemplateLink").href = `/api/planning/templates/plan.xlsx?issue_month=${encodeURIComponent($('#planUploadForm [name="issue_month"]').value)}&release_date=${encodeURIComponent($('#planUploadForm [name="release_date"]').value)}`; };
  $('#planUploadForm [name="issue_month"]').addEventListener("change", templateHref);
  $('#planUploadForm [name="release_date"]').addEventListener("change", templateHref);
  $("#actualUploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const period = data.get("period");
    state.workspace = data.get("workspace");
    try {
      const result = await uploadExcel(`/api/planning/actuals/${period}/upload?workspace=${state.workspace}`, data.get("workbook"));
      toast(`${result.rows} actual production row(s) saved for ${period}`);
      await renderPlans();
    } catch (error) { toast(error.message, true); }
  });
  $("#actualHistoryForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const period = new FormData(event.currentTarget).get("history_period");
    try {
      const rows = await api(`/api/planning/actuals?period=${period}&workspace=${state.workspace}`);
      $("#actualHistoryTable").innerHTML = table(["Model code", "Model name", "Actual production", "OEM stock"], rows.map((row) => `<tr><td>${esc(row.model_code)}</td><td>${esc(row.model_name)}</td><td>${number(row.actual_quantity)}</td><td>${number(row.oem_stock)}</td></tr>`).join(""), `No actual production data for ${monthLabel(period)}.`);
      $("#actualHistoryExports").innerHTML = tableExportLinks("actuals", { period, workspace: state.workspace });
      setupTableFilters("#actualHistoryTable", "#actualFilters", null, [{ title: "Model", index: 0 }]);
    } catch (error) { toast(error.message, true); }
  });
  setupTableFilters("#planHistory", "#planFilters", null, [{ title: "Issue month", index: 0 }, { title: "Revision", index: 1 }, { title: "Release date", index: 2 }, { title: "Workspace", index: 5 }]);
  setupTableFilters("#actualHistoryTable", "#actualFilters", null, [{ title: "Model", index: 0 }]);
}

async function openPlanEditor(cycleId) {
  const plan = state.plans.find((item) => item.id === cycleId);
  if (!plan) return;
  state.workspace = plan.workspace;
  state.planEditor = { cycleId, edits: new Map(), search: "" };
  await loadPlanEditorPage(0);
}

async function loadPlanEditorPage(offset) {
  const editor = state.planEditor;
  if (!editor) return;
  const query = new URLSearchParams({ offset: String(offset), limit: "25", search: editor.search });
  const result = await api(`/api/planning/plans/${editor.cycleId}/rows?${query}`);
  state.planEditor = { ...editor, ...result, edits: editor.edits, page: offset };
  renderPlanEditor();
}

function renderPlanEditor() {
  const editor = state.planEditor;
  const root = $("#planRevisionEditor");
  if (!editor || !root) return;
  root.hidden = false;
  const monthHeaders = editor.months.map((month) => `<th title="${esc(month)}">${esc(monthLabel(month))}</th>`).join("");
  const modelRows = editor.items.map((item) => `<tr><td class="primary-cell">${esc(item.model_code)}</td><td>${esc(item.model_name)}</td>${editor.months.map((month) => {
    const key = `${item.model_code}\u0000${month}`;
    const value = editor.edits.has(key) ? editor.edits.get(key).quantity : (item.quantities[month] ?? 0);
    return `<td><input class="form-control plan-quantity-input" type="number" min="0" max="1000000000" step="any" required aria-label="${esc(`${item.model_code}, ${monthLabel(month)} quantity`)}" data-plan-quantity data-model-code="${esc(item.model_code)}" data-month="${esc(month)}" value="${esc(value)}" /></td>`;
  }).join("")}</tr>`).join("");
  const canSave = editor.latest && editor.edits.size > 0;
  root.innerHTML = `${panelHeader(`Edit ${editor.revision}`, `Issue month ${monthLabel(editor.issue_month)} · ${editor.workspace} workspace · original saved ${new Date(editor.created_at).toLocaleString()}. Saving creates a new revision and keeps this one unchanged.`)}
    <div class="table-toolbar plan-editor-toolbar"><label class="search-field"><input class="form-control" id="planEditorSearch" type="search" placeholder="Search model code or name" value="${esc(editor.search)}" /></label><span>${editor.total_models.toLocaleString()} models · edit month quantities below</span><button class="button button-small" type="button" data-plan-editor-close>Close editor</button></div>
    ${editor.latest ? "" : '<p class="plan-editor-warning">A newer revision already exists. Open the latest revision to save further changes.</p>'}
    <form id="planEditorForm"><div class="table-wrap plan-editor-table"><table><thead><tr><th>Model code</th><th>Model name</th>${monthHeaders}</tr></thead><tbody>${modelRows || `<tr><td class="table-empty" colspan="${editor.months.length + 2}">No models found.</td></tr>`}</tbody></table></div>
      <div class="table-toolbar pagination-toolbar"><span>${editor.total_models ? `${(editor.offset + 1).toLocaleString()}–${Math.min(editor.offset + editor.items.length, editor.total_models).toLocaleString()} of ${editor.total_models.toLocaleString()} models` : "No models"}</span><div><button class="button button-small" type="button" data-plan-editor-page="${Math.max(0, editor.offset - editor.limit)}" ${editor.offset === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" type="button" data-plan-editor-page="${editor.offset + editor.limit}" ${editor.offset + editor.items.length >= editor.total_models ? "disabled" : ""}>Next</button><button class="button button-primary" id="savePlanRevision" type="submit" ${canSave ? "" : "disabled"}>Save changes as new revision</button></div></div>
    </form>`;
  $("#planEditorSearch").addEventListener("input", (event) => {
    clearTimeout(state.searchTimer);
    state.searchTimer = setTimeout(async () => {
      try {
        state.planEditor.search = event.target.value.trim();
        await loadPlanEditorPage(0);
      } catch (error) { toast(error.message, true); }
    }, 250);
  });
  $("#planEditorForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const rows = [...editor.edits.values()].map((change) => ({
      model_code: change.model_code, target_month: change.target_month, quantity: Number(change.quantity),
    }));
    if (!rows.length) return toast("Change at least one monthly quantity before saving.", true);
    if (rows.some((row) => !Number.isFinite(row.quantity) || row.quantity < 0)) {
      return toast("Edited quantities must be finite numbers greater than or equal to zero.", true);
    }
    try {
      const saved = await api(`/api/planning/plans/${editor.cycleId}/revisions`, {
        method: "POST", body: JSON.stringify({ rows }),
      });
      toast(`Saved ${saved.revision} on ${new Date(saved.created_at).toLocaleString()}`);
      await renderPlans();
      await openPlanEditor(saved.id);
    } catch (error) { toast(error.message, true); }
  });
}

async function renderMaster() {
  await loadCoreData();
  if (state.bomTab === "Model BOM Level 1" && !state.selectedModelIds.length) {
    state.selectedModelIds = state.models.map((model) => model.id);
  }
  if (state.bomTab === "Model BOM Level 1") {
    state.selectedModel = state.selectedModelIds.length === 1 ? String(state.selectedModelIds[0]) : "";
  }
  if (state.bomTab === "Level 2 Part BOM" || state.bomTab === "Level 3 Part BOM") {
    await loadBomParentOptions();
  }
  const productRows = state.products.flatMap((item) => {
    const assigned = state.models
      .filter((model) => (item.model_ids || []).includes(model.id))
      .sort((left, right) => left.code.localeCompare(right.code));
    const display = assigned.length ? assigned : [null];
    return display.map((model) => treeRowHtml("Product Master", {
      depth: 0,
      className: model?.has_bom ? "clickable-row" : "",
      attributes: model ? `data-product-model-row="${model.id}" data-has-bom="${model.has_bom}"` : `data-product-id="${item.id}"`,
      cells: { product: [item.code, item.name], model: [model?.code || "", model?.name || (assigned.length ? "" : "No models assigned")] },
      toggle: model?.has_bom ? treeToggle(model.code, `data-tree-expand="model" data-tree-model-id="${model.id}"`) : "",
      edit: `<button class="text-button" data-edit-product="${item.id}" data-edit-model-id="${model?.id || ""}">Edit</button>`,
      remove: `<button class="text-button button-danger" data-delete-product="${item.id}">Delete</button>`,
    }));
  }).join("");
  const modelRows = state.models.map((item) => `<tr data-edit-model="${item.id}" tabindex="0"><td>${esc(item.code)}</td><td>${esc(item.name)}</td><td>${esc(item.product_names?.join(", ") || "Unassigned")}</td><td><button class="text-button" data-edit-model="${item.id}">Edit</button><button class="text-button button-danger" data-delete-model="${item.id}">Delete</button></td></tr>`).join("");
  const tabs = ["Product Master", "Model BOM Level 1", "Level 2 Part BOM", "Level 3 Part BOM", "Material Master", "T2 Material Master", "T3 Material Master", "Alternate codes", "SOB change log", "Category"];
  const categoryTab = state.bomTab === "Category";
  const altTab = state.bomTab === "Alternate codes";
  const logTab = state.bomTab === "SOB change log";
  const tierTab = TIER_TABS[state.bomTab];
  const bomTemplateKind = {
    "Model BOM Level 1": "bom-oem", "Level 2 Part BOM": "bom-level-2", "Level 3 Part BOM": "bom-level-3",
  }[state.bomTab];
  const actions = state.bomTab === "Product Master"
    ? '<button class="button button-primary" data-action="new-product">Add line</button>'
    : state.bomTab === "Material Master"
      ? '<button class="button" data-action="new-model">Add model</button><button class="button button-primary" data-action="new-material">Add material</button>'
      : tierTab ? `<button class="button button-primary" data-action="new-tier-material">Add ${tierTab} material</button>` : "";
  const categorySection = `<section class="panel table-panel">${panelHeader("Category", "The one category list for parts (T1, T2, T3) and suppliers: every category dropdown uses it. Add, rename (every record using it follows) or delete (only when nothing uses it). A new category typed in a dropdown or found in an Excel upload is added here automatically.")}
      <form id="categoryAddForm" class="panel-body inline-upload"><label class="field"><span>New category</span><input class="form-control" name="name" minlength="2" maxlength="80" required placeholder="e.g. Rubber Parts" /></label><button class="button button-primary" type="submit">Add category</button></form>
      <div id="categoryTable"></div>
      <div id="categoryUnlisted"></div></section>`;
  const logSection = `<section class="panel table-panel">${panelHeader("SOB change log", "Every alternate SOB change: when, who, which code, where it is used (L1 / L2 / L3), from which month and the split. Set or remove SOB in Alternate codes → Manage SOB.")}
      <div class="table-toolbar"><label class="search-field"><input class="form-control" id="masterSearch" type="search" placeholder="Search code, use, split or user" /></label><button class="button button-small" type="button" data-execute-master-search>Execute</button><a class="button button-small" href="/api/planning/sob-change-log.xlsx" id="sobLogExcel">Download Excel</a><a class="button button-small" href="/api/planning/sob-change-log.pdf">Download PDF</a><button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <div id="sobLogTable"></div></section>`;
  const altSection = `<section class="panel table-panel">${panelHeader("Alternate codes", "Which code is an alternate of which (e.g. FL101STEEL0.8 for FL001STEEL0.8). Both codes come from T1 Material Master, are bought parts and use the same unit. Set the SOB here with Manage SOB: it lists every use of the main code (L1 direct per model, Level 2 and Level 3 part BOMs) and updates all ticked uses at once, from a month. 0% (main code 100%) until set.")}
      <form id="altAddForm" class="panel-body form-grid">
        <label class="field"><span>Search main code</span><input class="form-control" id="altMainSearch" type="search" placeholder="Type a code or name" /></label>
        <label class="field"><span>Main code (purchased parts)</span><select class="form-control" id="altMainSelect" required><option value="">Loading…</option></select></label>
        <label class="field"><span>Search alternate code</span><input class="form-control" id="altAltSearch" type="search" placeholder="Type a code or name" /></label>
        <label class="field"><span>Alternate code (purchased parts)</span><select class="form-control" id="altAltSelect" required><option value="">Loading…</option></select></label>
        <div class="field"><span>&nbsp;</span><button class="button button-primary" type="submit">Add alternate code</button></div>
      </form>
      <div class="table-toolbar"><label class="search-field"><input class="form-control" id="masterSearch" type="search" placeholder="Search main or alternate code / name" /></label><button class="button button-small" type="button" data-execute-master-search>Execute</button><a class="button button-small" href="/api/planning/templates/alternate-codes.xlsx">Template</a><a class="button button-small" href="/api/planning/alternate-codes.xlsx">Download Excel</a><a class="button button-small" href="/api/planning/alternate-codes.pdf">Download PDF</a><a class="button button-small" href="/api/planning/sob-change-log.xlsx">Download SOB history</a><button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <form id="altUploadForm" class="panel-body inline-upload"><label class="field"><span>Upload alternate codes (Main code, Alternate code; your own headings are fine)</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label><button class="button" type="submit">Upload</button></form>
      <div id="altTable"></div></section>`;
  void actions;  // the Add buttons sit in each panel, right above its table
  content.innerHTML = `${heading("Supply Chain Management", "Build products from model codes in T1 Material Master, define units once, and manage each model’s BOM by level.", "")}
    <div class="tab-strip">${tabs.map((tab) => `<button class="tab-button ${state.bomTab === tab ? "active" : ""}" data-bom-tab="${esc(tab)}">${esc(tab === "Material Master" ? "T1 Material Master" : tab)}</button>`).join("")}</div>
    ${state.bomTab === "Product Master" ? `<section class="panel table-panel">${panelHeader("Product Master")}
      <div class="table-toolbar"><span class="toolbar-add"><button class="button button-small button-primary" data-action="new-product">Add line</button></span><label class="search-field"><input class="form-control" id="masterSearch" type="search" placeholder="Search product or model" /></label><button class="button button-small" type="button" data-execute-master-search>Execute</button>${bomTreeExportLinks()}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <div id="productMasterTable">${table(treeHeaders("Product Master"), productRows, "No products yet.")}</div></section>`
    : state.bomTab === "Material Master" ? `<section class="panel table-panel">${panelHeader("T1 Material Master", "Codes your company uses directly or indirectly and buys from T1 suppliers. Define each code once; the unit selected here is used throughout its BOM and supply chain records. A code can be in only one of T1, T2 or T3.")}
      <div class="master-summaries"><article class="mini-stat"><span>Materials</span><strong id="materialCount">…</strong></article><article class="mini-stat"><span>Models</span><strong>${state.models.length}</strong></article><article class="mini-stat"><span>Products</span><strong>${state.products.length}</strong></article></div>
      <div class="table-toolbar"><span class="toolbar-add"><button class="button button-small button-primary" data-action="new-material">Add material</button></span><label class="search-field"><input class="form-control" id="masterSearch" type="search" placeholder="Search material code or name" /></label><button class="button button-small" type="button" data-execute-master-search>Execute</button><a class="button button-small" href="/api/planning/templates/own-stock.xlsx">Stock template</a>${tableExportLinks("materials")}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <div id="materialMasterTable"></div>
      <h3 class="subtable-heading subtable-with-add">Model Master ${tableExportLinks("models")}<span class="toolbar-add"><button class="button button-small button-primary" data-action="new-model">Add line</button></span></h3>
      <div class="table-toolbar"><label class="search-field"><input class="form-control" id="modelMasterSearch" type="search" placeholder="Search model code or name" /></label><button class="button button-small" type="button" data-execute-model-search>Execute</button><button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <div id="modelMasterTable">${table(["Model code", "Model name", "Assigned products", "Action"], modelRows, "No models yet.")}</div></section>`
    : altTab ? altSection
    : logTab ? logSection
    : categoryTab ? categorySection
    : tierTab ? `<section class="panel table-panel">${panelHeader(`${tierTab} Material Master`, `Codes used further down the supply chain (${tierTab === "T2" ? "by your T1 suppliers' suppliers" : "by your T2 suppliers' suppliers"}). Not part of your planning; suppliers pick them into their own material master. A code can be in only one of T1, T2 or T3.`)}
      <div class="table-toolbar"><span class="toolbar-add"><button class="button button-small button-primary" data-action="new-tier-material">Add ${tierTab} material</button></span><label class="search-field"><input class="form-control" id="masterSearch" type="search" placeholder="Search code or name" /></label><button class="button button-small" type="button" data-execute-master-search>Execute</button><a class="button button-small" href="/api/planning/tier-materials/${tierTab}.xlsx">Download Excel</a><a class="button button-small" href="/api/planning/tier-materials/${tierTab}.pdf">Download PDF</a><button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <form id="tierUploadForm" class="panel-body inline-upload"><label class="field"><span>Upload ${tierTab} Material Master (Code, Name, Category, Unit, Unit price / Unit cost, Currency; your own headings are fine)</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label><button class="button" type="submit">Upload</button></form>
      <div id="tierMaterialTable"></div></section>`
    : `<section class="panel table-panel ${state.bomTab === "Model BOM Level 1" ? "" : "part-bom-page"}">${panelHeader(state.bomTab)}
      <div class="table-toolbar bom-controls ${state.bomTab === "Model BOM Level 1" ? "" : "bom-controls-part"}"><div class="toolbar-group">
        ${state.bomTab === "Model BOM Level 1"
          ? modelPickerMarkup()
          : bomParentPickerMarkup()}
      </div>
      ${bomTemplateKind ? `<form id="bomUploadForm" class="bom-upload-inline"><input class="form-control bom-file-input" type="file" name="workbook" accept=".xlsx" aria-label="Select ${esc(state.bomTab)} Excel workbook" required /><button class="button button-small button-primary" id="bomUploadButton" type="submit" ${state.bomTab === "Model BOM Level 1" ? (state.selectedModelIds.length === 1 ? "" : "disabled") : (state.selectedBomParent ? "" : "disabled")}>Upload BOM</button></form><a class="button button-small" href="/api/planning/templates/${bomTemplateKind}.xlsx">Level ${state.bomTab === "Model BOM Level 1" ? "1" : state.bomTab === "Level 2 Part BOM" ? "2" : "3"} template</a>` : ""}
      ${bomTreeExportLinks()}
      <button class="button button-small" data-action="clear-table-filters">Clear all filters</button>
      <span class="toolbar-add">${state.bomTab === "Model BOM Level 1" ? '<button class="button button-small button-primary" data-action="new-bom">Add line</button>' : '<button class="button button-small button-primary" id="newPartBomButton" data-action="new-bom" disabled>Select one parent part</button><button class="button button-small button-primary" data-action="new-part-bom">New BOM</button>'}</span>
      </div>
      ${state.bomTab === "Model BOM Level 1" ? "" : '<div id="newBomEditor"></div>'}
      <div id="masterTable">${state.bomTab === "Model BOM Level 1"
        ? table(treeHeaders("Model BOM Level 1"), "", "Select one or more models to view their BOM.")
        : table(treeHeaders(state.bomTab), "", "Select a parent part to view this BOM level.")}</div>
    </section>`}
    `;
  if (state.bomTab === "Product Master") {
    $("#masterSearch").dataset.tableFilterRoot = "#productMasterTable";
  } else if (state.bomTab === "Material Master") {
    $("#masterSearch").dataset.tableFilterRoot = "#materialMasterTable";
    $("#modelMasterSearch").dataset.tableFilterRoot = "#modelMasterTable";
  } else if (altTab) {
    initAlternateCodesTab();
  } else if (logTab) {
    loadSobLog(0);
  } else if (categoryTab) {
    $("#categoryAddForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      try {
        const saved = await api("/api/planning/categories", { method: "POST", body: JSON.stringify({ name: form.elements.name.value }) });
        toast(`${saved.name} added to the category list`);
        form.reset();
        await reloadCategories();
        loadCategoryTab();
      } catch (error) { toast(error.message, true); }
    });
    loadCategoryTab();
  } else if (tierTab) {
    await loadTierMaterialPage(0);
    $("#tierUploadForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      try {
        const result = await uploadExcel(`/api/planning/tier-materials/upload?tier=${tierTab}`, new FormData(form).get("workbook"));
        toast(`${result.rows} row(s): ${result.added} added, ${result.updated} updated`);
        form.reset();
        loadTierMaterialPage(0);
      } catch (error) { toast(error.message, true); }
    });
  } else {
    $("#masterTable").dataset.filtertable = "true";
  }
  if (state.bomTab === "Material Master") await loadMaterialMasterPage(0);
  if (state.bomTab === "Model BOM Level 1") await loadBomTab(0);
  else if ($("#bomParentPicker")) updatePartBomActions();
  if ($("#bomUploadForm")) {
    $("#bomUploadForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      const file = new FormData(form).get("workbook");
      try {
        const query = new URLSearchParams({ tab: state.bomTab });
        const uploadPath = state.bomTab === "Model BOM Level 1"
          ? `/api/planning/model-bom/${state.selectedModel}/upload?${query}`
          : `/api/planning/part-bom/${state.selectedBomParent}/upload?${query}`;
        const result = await uploadExcel(uploadPath, file);
        toast(`${result.rows} ${state.bomTab} row(s) uploaded`);
        if (state.bomTab === "Model BOM Level 1") await loadBomTab(0);
        else await loadPartBomTab();
        form.reset();
      } catch (error) { toast(error.message, true); }
    });
  }
  const modelSearch = $('[data-model-search]');
  if (modelSearch) modelSearch.addEventListener("input", () => {
    const needle = modelSearch.value.trim().toLowerCase();
    modelSearch.closest(".model-picker").querySelectorAll("[data-model-option]").forEach((input) => {
      input.closest("label").hidden = !input.closest("label").textContent.toLowerCase().includes(needle);
    });
  });
  const parentSearch = $('[data-parent-search]');
  if (parentSearch) parentSearch.addEventListener("input", () => {
    const needle = parentSearch.value.trim().toLowerCase();
    parentSearch.closest(".model-picker").querySelectorAll("[data-parent-option]").forEach((input) => {
      input.closest("label").hidden = !input.closest("label").textContent.toLowerCase().includes(needle);
    });
  });
  if (parentSearch) await loadPartBomTab(0);
}

async function loadBomParentOptions() {
  const query = new URLSearchParams({ tab: state.bomTab });
  const result = await api(`/api/planning/bom-part-parents?${query}`);
  // Dropdown and tree show only parents that already have a BOM; parents without one
  // are offered by "New BOM".
  state.bomParentAll = result.items;
  state.bomParentOptions = result.items.filter((item) => item.has_bom);
  const validIds = new Set(result.items.map((item) => String(item.id)));
  state.selectedBomParentIds = state.selectedBomParentIds.filter((id) => validIds.has(String(id)));
  state.selectedBomParent = state.selectedBomParentIds.length === 1 ? String(state.selectedBomParentIds[0]) : "";
  updatePartBomActions();
}

function updatePartBomActions() {
  const selectedIds = state.selectedBomParentIds.length
    ? state.selectedBomParentIds
    : state.bomParentOptions.map((item) => item.id);
  const parent = selectedIds.length === 1
    ? state.bomParentOptions.find((item) => String(item.id) === String(selectedIds[0]))
    : null;
  state.selectedBomParent = parent ? String(parent.id) : "";
  if ($("#bomUploadButton")) $("#bomUploadButton").disabled = !parent;
  const button = $("#newPartBomButton");
  if (!button) return;
  button.disabled = !parent;
  button.textContent = parent ? `Add component to ${parent.code}` : "Select one parent part";
}

async function applyParentSelection() {
  const picker = $("#bomParentPicker");
  state.selectedBomParentIds = [...picker.querySelectorAll("[data-parent-option]:checked")].map((input) => Number(input.value));
  picker.open = false;
  const summary = $("#bomParentPickerSummary");
  if (summary) {
    const allSelected = state.selectedBomParentIds.length === state.bomParentOptions.length;
    summary.textContent = !state.selectedBomParentIds.length || allSelected
      ? `All ${state.bomTab === "Level 2 Part BOM" ? "level 1 parent parts" : "level 2 parent parts"}`
      : `${state.selectedBomParentIds.length} parent part${state.selectedBomParentIds.length === 1 ? "" : "s"} selected`;
  }
  updatePartBomActions();
  state.bomOffset = 0;
  return renderMaster();
}

async function loadPartBomTab() {
  const headers = treeHeaders(state.bomTab);
  if (!state.bomParentOptions.length) {
    $("#masterTable").innerHTML = table(headers, "", `No ${esc(state.bomTab)} yet. Use New BOM to create one.`);
    return;
  }
  const selected = state.selectedBomParentIds.length ? new Set(state.selectedBomParentIds.map(String)) : null;
  const parents = state.bomParentOptions.filter((item) => !selected || selected.has(String(item.id)));
  const parentLevel = state.bomTab === "Level 2 Part BOM" ? "l1" : "l2";
  const rows = parents.map((item) => treeRowHtml(state.bomTab, {
    depth: 0, className: "clickable-row", attributes: "data-tree-group",
    cells: { [parentLevel]: [item.code, item.name] },
    toggle: treeToggle(item.code, `data-tree-expand="part" data-tree-parent-id="${item.id}" data-tree-level="${esc(state.bomTab)}"`),
  })).join("");
  $("#masterTable").innerHTML = `${table(headers, rows, `No ${esc(state.bomTab)} lines found.`)}<div class="table-toolbar pagination-toolbar"><span>${parents.length.toLocaleString()} parent part${parents.length === 1 ? "" : "s"} with BOM</span></div>`;
}

// One-table BOM tree: every level has its own code + name columns, so items of the
// same level line up in one column; a child row leaves its parents' columns blank.
const TREE_LEVELS = {
  "Product Master": ["product", "model", "l1", "l2", "l3"],
  "Model BOM Level 1": ["model", "l1", "l2", "l3"],
  "Level 2 Part BOM": ["l1", "l2", "l3"],
  "Level 3 Part BOM": ["l2", "l3"],
};
const TREE_LEVEL_LABELS = {
  product: ["Product code", "Product name"], model: ["Model code", "Model name"],
  l1: ["L1 part code", "L1 part name"], l2: ["L2 part code", "L2 part name"], l3: ["L3 part code", "L3 part name"],
};

function treeHeaders(tab) {
  return ["", ...TREE_LEVELS[tab].flatMap((level) => TREE_LEVEL_LABELS[level]), "Quantity", "Unit", "Part price", "Edit", "Delete"];
}

function treeToggle(label, attributes) {
  return `<button type="button" class="tree-toggle" aria-label="Expand ${esc(label)} BOM" ${attributes} aria-expanded="false">+</button>`;
}

function treeRowHtml(tab, row) {
  // Columns this row does not fill itself repeat the parent row's code / name (row.inherit).
  const levelCells = TREE_LEVELS[tab].map((level, index) => row.cells[level]
    ? `<td>${esc(row.cells[level][0])}</td><td>${esc(row.cells[level][1])}</td>`
    : [0, 1].map((part) => `<td class="tree-inherited">${esc(row.inherit?.[index * 2 + part] || "")}</td>`).join("")).join("");
  const className = `${row.className || ""}${row.toggle ? " tree-has-children" : ""}`;
  return `<tr class="${className}" ${row.attributes || ""} data-tree-depth="${row.depth}" tabindex="0"><td class="tree-toggle-cell">${row.toggle || ""}</td>${levelCells}<td>${row.quantity ?? ""}</td><td>${esc(row.uom || "")}</td><td>${row.price ?? ""}</td><td>${row.edit || ""}</td><td>${row.remove || ""}</td></tr>`;
}

function treePrice(item) {
  if (item.unit_price === undefined) return "Restricted";
  return `<span title="${esc(item.price_source || "")}">${money(item.unit_price, item.currency)}</span>`;
}

// A main code's alternate codes follow its row: same level column, SOB share of the quantity,
// the alternate's own price. Columns to the left repeat the main row's parents.
function withAlternates(tab, row, item, level) {
  if (!item.alternates?.length) return treeRowHtml(tab, row);
  // Main row shows its own SOB share of the BOM quantity; Edit still changes the full BOM quantity.
  const lineAlternates = item.alternates.map((alt) => ({ id: alt.material_id, code: alt.code, name: alt.name, sob: alt.sob_percent }));
  const next = item.sob_next_change;  // next dated SOB change (e.g. a replacement from Jan 2027)
  const later = (sob) => (next ? ` → ${number(sob, 2)}% from ${monthLabel(next.from_month)}` : "");
  const main = { ...row, attributes: `${row.attributes || ""} data-alternates="${esc(JSON.stringify(lineAlternates))}"`, quantity: number(item.main_quantity, 4), cells: { ...row.cells, [level]: [row.cells[level][0], `${row.cells[level][1]} · main ${number(item.main_sob_percent, 2)}% of ${number(item.quantity, 4)}${later(next?.main_sob_percent)}`] } };
  const inherit = TREE_LEVELS[tab].flatMap((name, index) => row.cells[name] || [row.inherit?.[index * 2] || "", row.inherit?.[index * 2 + 1] || ""]);
  return treeRowHtml(tab, main) + item.alternates.map((alt) => treeRowHtml(tab, {
    depth: row.depth, inherit, className: "tree-alternate", attributes: `data-alternate-of="${item.material_id}"`,
    cells: { [level]: [`↳ ${alt.code}`, `${alt.name} · alternate ${number(alt.sob_percent, 2)}%${later(alt.next_sob_percent)}`] },
    quantity: number(alt.quantity, 4), uom: alt.uom, price: treePrice(alt),
  })).join("");
}

function level1RowHtml(tab, item, modelId, depth, showModel, inherit) {
  const data = `data-model-id="${modelId}" data-material-id="${item.material_id}" data-quantity="${item.edit_quantity}" data-uom="${esc(item.uom)}"`;
  return withAlternates(tab, {
    depth, inherit, className: "clickable-row",
    attributes: `data-bom-level1-row data-has-bom="${item.has_bom}" data-edit-bom-row="${item.root_bom_id || ""}" ${data}`,
    cells: { ...(showModel ? { model: [item.model_code, item.model_name] } : {}), l1: [item.code, item.name] },
    toggle: item.has_bom ? treeToggle(item.code, `data-tree-expand="part" data-tree-parent-id="${item.material_id}" data-tree-level="Level 2 Part BOM"`) : "",
    quantity: number(item.quantity), uom: item.uom, price: treePrice(item),
    edit: `<button class="text-button" data-edit-bom-row="${item.root_bom_id || ""}" ${data}>Edit</button>`,
    remove: `<button class="text-button button-danger" data-delete-bom-row="${item.root_bom_id || ""}" data-model-id="${modelId}">Delete</button>`,
  }, item, "l1");
}

function partRowHtml(tab, item, parent, lineLevel, depth, inherit) {
  const level = lineLevel === "Level 2 Part BOM" ? "l2" : "l3";
  const expandable = level === "l2" && item.has_bom;
  const data = `data-parent-item="${parent.id}" data-material-id="${item.material_id}" data-quantity="${item.quantity}" data-uom="${esc(item.uom)}"`;
  return withAlternates(tab, {
    depth, inherit, className: "clickable-row",
    attributes: `data-bom-level2-row data-has-bom="${expandable}" data-edit-component-row="${item.component_id}" ${data}`,
    cells: { [level]: [item.code, item.name] },
    toggle: expandable ? treeToggle(item.code, `data-tree-expand="part" data-tree-parent-id="${item.material_id}" data-tree-level="Level 3 Part BOM"`) : "",
    quantity: number(item.quantity), uom: item.uom, price: treePrice(item),
    edit: `<button class="text-button" data-edit-component-row="${item.component_id}" ${data}>Edit</button>`,
    remove: `<button class="text-button button-danger" data-delete-component-row="${item.component_id}" data-parent-item="${parent.id}">Delete</button>`,
  }, item, level);
}

async function expandBomTree(button) {
  const row = button.closest("tr");
  const depth = Number(row.dataset.treeDepth || 0);
  if (button.getAttribute("aria-expanded") === "true") {
    while (row.nextElementSibling && Number(row.nextElementSibling.dataset.treeDepth) > depth) row.nextElementSibling.remove();
    button.textContent = "+";
    button.setAttribute("aria-expanded", "false");
    return;
  }
  button.disabled = true;
  const inherit = [...row.cells].slice(1, 1 + TREE_LEVELS[state.bomTab].length * 2).map((cell) => cell.textContent);
  try {
    let rows;
    if (button.dataset.treeExpand === "model") {
      const modelId = Number(button.dataset.treeModelId);
      const query = new URLSearchParams({ tab: "Model BOM Level 1", offset: "0", limit: "500" });
      const result = await api(`/api/planning/model-bom/${modelId}?${query}`);
      rows = result.items.map((item) => level1RowHtml(state.bomTab, item, modelId, depth + 1, false, inherit));
    } else {
      const level = button.dataset.treeLevel;
      const query = new URLSearchParams({ tab: level, offset: "0", limit: "500" });
      const result = await api(`/api/planning/part-bom/${Number(button.dataset.treeParentId)}?${query}`);
      rows = result.items.map((item) => partRowHtml(state.bomTab, item, result.parent, level, depth + 1, inherit));
    }
    if (!rows.length) {
      button.remove();
      return;
    }
    row.insertAdjacentHTML("afterend", rows.join(""));
    button.textContent = "−";
    button.setAttribute("aria-expanded", "true");
  } finally {
    button.disabled = false;
  }
}

function materialOptions(items, labelKey, selectedId, placeholder) {
  return `<option value="">${placeholder}</option>${items.map((item) => `<option value="${item.id}" data-uom="${esc(item.uom)}" ${String(item.id) === String(selectedId) ? "selected" : ""}>${esc(item[labelKey])}</option>`).join("")}`;
}

function newBomLineRow() {
  return `<tr data-new-bom-line><td><input class="form-control" type="search" data-line-search placeholder="Search code / name" /></td><td><select class="form-control" data-line-code>${materialOptions(state.materials, "code", "", "Select part code")}</select></td><td><select class="form-control" data-line-name>${materialOptions(state.materials, "name", "", "Select part name")}</select></td><td><input class="form-control" type="number" step="any" data-line-qty placeholder="e.g. 2 or -0.3" /></td><td><input class="form-control" data-line-unit readonly /></td><td><button type="button" class="text-button button-danger" data-new-bom-remove aria-label="Remove line">Remove</button></td></tr>`;
}

function openNewBomEditor() {
  const holder = $("#newBomEditor");
  if (!holder) return;
  const candidates = state.bomParentAll.filter((item) => !item.has_bom);
  const levelName = state.bomTab === "Level 2 Part BOM" ? "Level 1" : "Level 2";
  if (!candidates.length) return toast(`Every ${levelName} part already has a BOM. Add lines to an existing BOM instead.`, true);
  holder.innerHTML = `<div class="new-bom-editor"><div class="new-bom-head"><label class="field"><span>Parent part (${levelName} part without BOM)</span><select class="form-control" id="newBomParent"><option value="">Select parent part</option>${candidates.map((item) => `<option value="${item.id}">${esc(item.code)} · ${esc(item.name)}</option>`).join("")}</select></label><p class="subtle">Parts come from Material Master only. Use a negative quantity for scrap or a by-product; zero is not allowed.</p></div><div class="new-bom-actions"><button type="button" class="button button-small" data-new-bom-add>+ Add line</button><button type="button" class="button button-small button-primary" data-new-bom-save>Save BOM</button><button type="button" class="button button-small" data-new-bom-cancel>Cancel</button></div><div class="table-wrap"><table class="new-bom-lines"><thead><tr><th>Search</th><th>Part code</th><th>Part name</th><th>Quantity per parent</th><th>Unit</th><th></th></tr></thead><tbody id="newBomLines">${newBomLineRow()}</tbody></table></div></div>`;
  const lines = $("#newBomLines");
  const syncLine = (row, id) => {
    row.querySelector("[data-line-code]").value = id;
    row.querySelector("[data-line-name]").value = id;
    const option = row.querySelector("[data-line-code]").selectedOptions[0];
    row.querySelector("[data-line-unit]").value = id ? option?.dataset.uom || "" : "";
  };
  lines.addEventListener("change", (event) => {
    const select = event.target.closest("[data-line-code], [data-line-name]");
    if (select) syncLine(select.closest("tr"), select.value);
  });
  lines.addEventListener("input", (event) => {
    const search = event.target.closest("[data-line-search]");
    if (!search) return;
    const row = search.closest("tr");
    clearTimeout(row.searchTimer);
    row.searchTimer = setTimeout(async () => {
      try {
        const query = new URLSearchParams({ offset: "0", limit: "100", search: search.value });
        const result = await api(`/api/planning/materials?${query}`);
        const current = row.querySelector("[data-line-code]").value;
        const keep = result.items.some((item) => String(item.id) === current) ? current : "";
        row.querySelector("[data-line-code]").innerHTML = materialOptions(result.items, "code", keep, "Select part code");
        row.querySelector("[data-line-name]").innerHTML = materialOptions(result.items, "name", keep, "Select part name");
        syncLine(row, keep || (result.items.length === 1 ? String(result.items[0].id) : ""));
      } catch (error) { toast(error.message, true); }
    }, 250);
  });
  holder.querySelector("[data-new-bom-add]").addEventListener("click", () => lines.insertAdjacentHTML("beforeend", newBomLineRow()));
  lines.addEventListener("click", (event) => {
    if (!event.target.closest("[data-new-bom-remove]")) return;
    if (lines.rows.length > 1) event.target.closest("tr").remove();
    else toast("A BOM needs at least one line.", true);
  });
  holder.querySelector("[data-new-bom-cancel]").addEventListener("click", () => { holder.innerHTML = ""; });
  holder.querySelector("[data-new-bom-save]").addEventListener("click", async () => {
    const parentId = Number($("#newBomParent").value);
    if (!parentId) return toast("Select the parent part.", true);
    const payload = [];
    for (const [index, row] of [...lines.rows].entries()) {
      const childId = Number(row.querySelector("[data-line-code]").value);
      const quantity = Number(row.querySelector("[data-line-qty]").value);
      if (!childId) return toast(`Line ${index + 1}: select a part.`, true);
      if (!quantity || !Number.isFinite(quantity)) return toast(`Line ${index + 1}: quantity must be a non-zero number.`, true);
      payload.push({ child_item_id: childId, quantity });
    }
    try {
      const result = await api(`/api/planning/part-bom/${parentId}/new`, { method: "POST", body: JSON.stringify({ tab: state.bomTab, lines: payload }) });
      // Clear the parent filter so the new BOM is visible in the tree.
      state.selectedBomParentIds = [];
      await renderMaster();
      toast(`New BOM saved for ${result.parent_code} (${result.rows} line${result.rows === 1 ? "" : "s"})`);
    } catch (error) { toast(error.message, true); }
  });
}

async function applyModelSelection() {
  const picker = $(".model-picker");
  if (picker) state.selectedModelIds = [...picker.querySelectorAll("[data-model-option]:checked")].map((input) => Number(input.value));
  if (!state.selectedModelIds.length) {
    state.selectedModelIds = state.models.map((model) => model.id);
    picker?.querySelectorAll("[data-model-option]").forEach((input) => { input.checked = true; });
  }
  state.selectedModel = state.selectedModelIds.length === 1 ? String(state.selectedModelIds[0]) : "";
  state.l1Page = 0;
  const summary = $("#modelPickerSummary");
  if (summary) summary.textContent = state.selectedModelIds.length
    ? `${state.selectedModelIds.length} model${state.selectedModelIds.length === 1 ? "" : "s"} selected`
    : "All models";
  if ($("#bomUploadButton")) $("#bomUploadButton").disabled = state.selectedModelIds.length !== 1;
  state.bomOffset = 0;
  if (state.bomTab === "Model BOM Level 1") return loadBomTab(0);
  await loadBomParentOptions();
  return loadPartBomTab(0);
}

const L1_PAGE_SIZE = 50;

async function loadBomTab() {
  const modelIds = state.selectedModelIds.length ? state.selectedModelIds : state.models.map((model) => model.id);
  if (!modelIds.length) {
    $("#masterTable").innerHTML = table(treeHeaders("Model BOM Level 1"), "", "Select one or more models to view their BOM.");
    return;
  }
  // 1,000+ models: show 50 models a page, fetched in one request
  const pages = Math.ceil(modelIds.length / L1_PAGE_SIZE);
  state.l1Page = Math.min(Math.max(0, state.l1Page || 0), pages - 1);
  const pageIds = modelIds.slice(state.l1Page * L1_PAGE_SIZE, (state.l1Page + 1) * L1_PAGE_SIZE);
  const query = new URLSearchParams();
  pageIds.forEach((id) => query.append("model_id", String(id)));
  const result = await api(`/api/planning/model-boms?${query}`);
  const items = result.models.flatMap((model) => model.items.map((row) => ({ ...row, model_id: model.model_id, model_code: model.model_code, model_name: model.model_name })));
  const first = state.l1Page * L1_PAGE_SIZE + 1;
  const pager = pages > 1 ? `<div class="table-toolbar pagination-toolbar"><span>Models ${first.toLocaleString()}–${(first + pageIds.length - 1).toLocaleString()} of ${modelIds.length.toLocaleString()}</span><div><button class="button button-small" data-l1-page="${state.l1Page - 1}" ${state.l1Page === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" data-l1-page="${state.l1Page + 1}" ${state.l1Page >= pages - 1 ? "disabled" : ""}>Next</button></div></div>` : "";
  $("#masterTable").innerHTML = table(treeHeaders("Model BOM Level 1"), items.map((item) => level1RowHtml("Model BOM Level 1", item, item.model_id, 0, true)).join(""), "No Model BOM Level 1 lines found.") + pager;
}

const TIER_TABS = { "T2 Material Master": "T2", "T3 Material Master": "T3" };

async function loadTierMaterialPage(offset = 0) {
  const target = $("#tierMaterialTable");
  const tier = TIER_TABS[state.bomTab];
  if (!target || !tier) return;
  const search = $("#masterSearch")?.value.trim() || "";
  const result = await api(`/api/planning/tier-materials?${new URLSearchParams({ tier, offset: String(offset), limit: "500", search })}`);
  state.tierItems = result.items;
  const rows = result.items.map((m) => `<tr><td>${esc(m.code)}</td><td>${esc(m.name)}</td><td>${esc(m.category)}</td><td>${esc(m.uom)}</td><td>${m.unit_price === undefined ? "Restricted" : number(m.unit_price, 2)}</td><td>${esc(m.currency || "")}</td><td><button class="text-button" data-tier-edit="${m.id}">Edit</button></td><td><button class="text-button button-danger" data-tier-delete="${m.id}">Delete</button></td></tr>`).join("");
  target.innerHTML = `${table(["Code", "Name", "Category", "Unit", "Reference price", "Currency", "Edit", "Delete"], rows, search ? "No codes match this search." : `No ${tier} codes yet.`)}<div class="table-toolbar pagination-toolbar"><span>${result.total ? `${(offset + 1).toLocaleString()}–${(offset + result.items.length).toLocaleString()} of ${result.total.toLocaleString()}` : "No codes"}</span><div><button class="button button-small" data-tier-page="${Math.max(0, offset - 500)}" ${offset === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" data-tier-page="${offset + 500}" ${offset + result.items.length >= result.total ? "disabled" : ""}>Next</button></div></div>`;
}

function openTierMaterial(material = null) {
  const tier = TIER_TABS[state.bomTab];
  $("#modalTitle").textContent = material ? `Edit ${material.code}` : `Add ${tier} material`;
  $("#modalBody").innerHTML = `<form id="tierForm" class="form-grid">${field("Code", "code", "text", { required: true, maxlength: 48, value: material?.code })}${field("Name", "name", "text", { required: true, maxlength: 180, value: material?.name })}${categoryField(material?.category || "Raw material")}${unitField("Unit", "uom", material?.uom || "kg")}${canSeePrices() ? `${field("Reference price", "unit_price", "number", { min: 0, step: "0.0001", value: material?.unit_price ?? 0 })}<label class="field"><span>Currency</span><select class="form-control" name="currency">${["INR", "USD", "EUR", "GBP", "JPY"].map((c) => `<option ${c === (material?.currency || "INR") ? "selected" : ""}>${c}</option>`).join("")}</select></label>` : ""}<div class="modal-footer field full"><button class="button button-primary">Save ${tier} material</button></div></form>`;
  modalLayer.hidden = false;
  bindCategoryField($("#tierForm"));
  $("#tierForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const values = takeCategory(Object.fromEntries(new FormData(event.currentTarget).entries()));
    const body = { tier, code: values.code, name: values.name, category: values.category, uom: values.uom, unit_price: Number(values.unit_price ?? material?.unit_price ?? 0), currency: values.currency || material?.currency || "INR" };
    try {
      await api(material ? `/api/planning/tier-materials/${material.id}` : "/api/planning/tier-materials", { method: material ? "PATCH" : "POST", body: JSON.stringify(body) });
      modalLayer.hidden = true;
      toast(`${tier} material saved`);
      reloadCategories();
      loadTierMaterialPage(0);
    } catch (error) { toast(error.message, true); }
  });
}

// Alternate codes (Product Master tab): main code + alternate code + alternate SOB %
function initAlternateCodesTab() {
  bindMaterialSearch("altMainSearch", "altMainSelect", null, { purchased: true, loadNow: true });
  bindMaterialSearch("altAltSearch", "altAltSelect", null, { purchased: true, loadNow: true });
  $("#altAddForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const body = { main_item_id: Number($("#altMainSelect").value), alternate_item_id: Number($("#altAltSelect").value) };
    try {
      const saved = await api("/api/planning/alternate-codes", { method: "POST", body: JSON.stringify(body) });
      toast(`${saved.alternate_code} added as alternate of ${saved.main_code}. Set its SOB on each use below.`);
      await loadAlternateCodes(0);
      openManageSob(saved.main_item_id);
    } catch (error) { toast(error.message, true); }
  });
  $("#altUploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    try {
      const result = await uploadExcel("/api/planning/alternate-codes/upload", new FormData(form).get("workbook"));
      toast(`${result.rows} row(s): ${result.added} added, ${result.already_there} already there`);
      form.reset();
      loadAlternateCodes(0);
    } catch (error) { toast(error.message, true); }
  });
  loadAlternateCodes(0);
}

async function loadAlternateCodes(offset = 0) {
  const target = $("#altTable");
  if (!target) return;
  const search = $("#masterSearch")?.value.trim() || "";
  const result = await api(`/api/planning/alternate-codes?${new URLSearchParams({ offset: String(offset), limit: "500", search })}`);
  state.altItems = result.items;
  const priced = canSeePrices();
  const rows = result.items.map((a) => `<tr><td>${esc(a.main_code)}</td><td>${esc(a.main_name)}</td><td>${esc(a.alternate_code)}</td><td>${esc(a.alternate_name)}</td><td>${esc(a.uom)}</td><td>${esc(["L1", "L2", "L3"].filter((level) => a.uses?.[level]).map((level) => `${level} × ${a.uses[level]}`).join(" · ") || "Not used in a BOM")}</td><td>${priced && a.main_unit_price !== undefined ? money(a.main_unit_price, a.main_currency) : "Restricted"}</td><td>${priced && a.unit_price !== undefined ? money(a.unit_price, a.currency) : "Restricted"}</td><td><button class="text-button" data-alt-manage="${a.main_item_id}">Manage SOB</button></td><td><button class="text-button button-danger" data-alt-delete="${a.id}">Delete</button></td></tr>`).join("");
  target.innerHTML = `${table(["Main code", "Main name", "Alternate code", "Alternate name", "Unit", "Main code used in", "Main price", "Alternate price", "SOB", "Delete"], rows, search ? "No alternate codes match this search." : "No alternate codes yet.")}<div class="table-toolbar pagination-toolbar"><span>${result.total ? `${(offset + 1).toLocaleString()}–${(offset + result.items.length).toLocaleString()} of ${result.total.toLocaleString()}` : "No alternate codes"}</span><div><button class="button button-small" data-alt-page="${Math.max(0, offset - 500)}" ${offset === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" data-alt-page="${offset + 500}" ${offset + result.items.length >= result.total ? "disabled" : ""}>Next</button></div></div>`;
}

const SOB_LEVELS = [["L1", "direct use in Model BOM Level 1"], ["L2", "Level 2 part BOMs"], ["L3", "Level 3 part BOMs"]];

async function loadCategoryTab() {
  const target = $("#categoryTable");
  if (!target) return;
  const { items, unlisted } = await api("/api/planning/categories");
  state.categoryItems = items;
  target.innerHTML = table(["Category", "Parts (T1)", "T2 / T3 codes", "Suppliers", "Rename", "Delete"],
    items.map((c) => `<tr><td>${esc(c.name)}</td><td>${number(c.parts, 0)}</td><td>${number(c.tier_codes, 0)}</td><td>${number(c.suppliers, 0)}</td><td><button class="text-button" data-category-rename="${c.id}">Rename</button></td><td><button class="text-button button-danger" data-category-delete="${c.id}">Delete</button></td></tr>`).join(""), "No categories yet.");
  const options = items.map((c) => `<option value="${esc(c.name)}">${esc(c.name)}</option>`).join("");
  $("#categoryUnlisted").innerHTML = unlisted.length ? `<h3 class="subtable-heading">Categories used by records but not in your list</h3><p class="subtle panel-body">Replace each with a category from your list (the records move to it), or add it to your list.</p>${table(["Old category", "Parts (T1)", "T2 / T3 codes", "Suppliers", "Replace with", "Action"],
    unlisted.map((u) => `<tr data-unlisted="${esc(u.name)}"><td>${esc(u.name)}</td><td>${number(u.parts, 0)}</td><td>${number(u.tier_codes, 0)}</td><td>${number(u.suppliers, 0)}</td><td><select class="form-control" data-replace-with><option value="">Choose from your list</option>${options}</select></td><td><button class="text-button" data-category-replace>Replace</button><button class="text-button" data-category-adopt>Add to my list</button></td></tr>`).join(""))}` : "";
}

// Manage SOB (Alternate codes): every use of the main code with its SOB; one change, from a
// month, goes on all ticked uses at once and is written to the SOB change log.
async function openManageSob(mainId) {
  let view;
  try { view = await api(`/api/planning/alternate-codes/uses?${new URLSearchParams({ main_item_id: String(mainId) })}`); }
  catch (error) { return toast(error.message, true); }
  const { main, alternates, uses } = view;
  const sobCells = (sob, mainShare) => `${alternates.map((alt) => `${number(sob[alt.id] || 0, 2)}%`).join(" / ")} / ${number(mainShare, 2)}%`;
  const historyText = (use) => use.history.map((change) => `${change.from_month ? monthLabel(change.from_month) : "Start"}: ${alternates.map((alt) => `${number(change.alternates.find((a) => a.alternate_item_id === alt.id)?.sob_percent || 0, 2)}%`).join(" / ")} / ${number(change.main_sob_percent, 2)}%`).join(" · ") || "Not set (main 100%)";
  const anyHistory = uses.some((use) => use.history.length);
  $("#modalTitle").textContent = `Manage SOB · ${main.code} · ${main.name}`;
  $("#modalBody").innerHTML = `<form id="sobForm" class="form-grid">
    <div class="field full subtle">${uses.length} use${uses.length === 1 ? "" : "s"} of ${esc(main.code)}. The split is shown as ${alternates.map((alt) => esc(alt.code)).join(" / ")} / main.</div>
    ${alternates.map((alt) => `<label class="field"><span>${esc(alt.code)} · ${esc(alt.name)} (%)</span><input class="form-control" type="number" min="0" max="100" step="any" data-sob-alt="${alt.id}" value="0" /></label>`).join("")}
    <label class="field"><span>Main code ${esc(main.code)} SOB (%)</span><input class="form-control" id="sobMain" readonly /></label>
    <label class="field"><span>SOB valid from month (empty = all months)</span><input class="form-control" type="month" id="sobFrom" value="${anyHistory ? monthNow() : ""}" /></label>
    ${SOB_LEVELS.map(([level, label]) => { const rows = uses.filter((use) => use.level === level); return `<section class="field full sob-level-panel" data-sob-level="${level}"><h4><label><input type="checkbox" data-sob-level-all="${level}" ${rows.length ? "checked" : "disabled"} /> ${level} · ${label} <span class="subtle">(${rows.length} use${rows.length === 1 ? "" : "s"})</span></label></h4>${rows.length ? `<div class="table-wrap"><table class="sob-uses"><thead><tr><th></th><th>Used in</th><th>BOM qty</th><th>SOB this month</th><th>Next change</th><th>History</th></tr></thead><tbody>${rows.map((use) => `<tr data-sob-use="${use.key}"><td><input type="checkbox" data-sob-line data-level="${level}" checked data-model-bom-id="${use.model_bom_id || ""}" data-component-id="${use.component_id || ""}" aria-label="Tick ${esc(use.used_in)}" /></td><td>${esc(use.used_in)}</td><td>${number(use.quantity, 4)} ${esc(main.uom)}</td><td>${sobCells(use.sob, use.main_sob_percent)}</td><td>${use.next_change ? `${esc(monthLabel(use.next_change.from_month))}: main ${number(use.next_change.main_sob_percent, 2)}%` : "—"}</td><td class="sob-history-cell">${esc(historyText(use))}</td></tr>`).join("")}</tbody></table></div>` : `<p class="subtle">${esc(main.code)} is not used at ${level}.</p>`}</section>`; }).join("")}
    <div class="modal-footer field full"><button type="button" class="button" id="sobRemove">Remove this month's change from ticked uses</button><button class="button button-primary" type="submit">Save SOB for ticked uses</button></div></form>`;
  modalLayer.hidden = false;
  const form = $("#sobForm");
  const syncMain = () => { $("#sobMain").value = number(100 - [...form.querySelectorAll("[data-sob-alt]")].reduce((sum, input) => sum + Number(input.value || 0), 0), 2); };
  syncMain();
  form.addEventListener("input", (event) => { if (event.target.matches("[data-sob-alt]")) syncMain(); });
  form.querySelectorAll("[data-sob-level-all]").forEach((all) => all.addEventListener("change", () => {
    form.querySelectorAll(`[data-sob-line][data-level="${all.dataset.sobLevelAll}"]`).forEach((box) => { box.checked = all.checked; });
  }));
  const ticked = () => [...form.querySelectorAll("[data-sob-line]:checked")].map((box) => ({ model_bom_id: Number(box.dataset.modelBomId) || null, component_id: Number(box.dataset.componentId) || null }));
  const send = async (action) => {
    const body = { main_item_id: main.id, from_month: $("#sobFrom").value, action, lines: ticked(),
      alternates: [...form.querySelectorAll("[data-sob-alt]")].map((input) => ({ alternate_item_id: Number(input.dataset.sobAlt), sob_percent: Number(input.value || 0) })) };
    try {
      const result = await api("/api/planning/alternate-sob", { method: "POST", body: JSON.stringify(body) });
      const when = result.from_month ? `from ${monthLabel(result.from_month)}` : "for all months";
      toast(action === "set" ? `SOB saved ${when} on ${result.lines} use${result.lines === 1 ? "" : "s"} of ${main.code}` : `SOB change ${when} removed from ${result.lines} use${result.lines === 1 ? "" : "s"}`);
      openManageSob(main.id);
      loadAlternateCodes(0);
    } catch (error) { toast(error.message, true); }
  };
  form.addEventListener("submit", (event) => { event.preventDefault(); send("set"); });
  $("#sobRemove").addEventListener("click", () => {
    const when = $("#sobFrom").value ? monthLabel($("#sobFrom").value) : "the start";
    if (window.confirm(`Remove the SOB change from ${when} on the ticked uses?`)) send("remove");
  });
}

async function loadSobLog(offset = 0) {
  const target = $("#sobLogTable");
  if (!target) return;
  const search = $("#masterSearch")?.value.trim() || "";
  const result = await api(`/api/planning/sob-change-log?${new URLSearchParams({ offset: String(offset), limit: "500", search })}`);
  const excel = $("#sobLogExcel");
  if (excel) excel.href = `/api/planning/sob-change-log.xlsx?${new URLSearchParams({ search })}`;
  const rows = result.items.map((r) => `<tr><td>${esc(r.changed_at)}</td><td>${esc(r.changed_by)}</td><td>${esc(r.action)}</td><td>${esc(r.main_code)}</td><td>${esc(r.level)}</td><td>${esc(r.used_in)}</td><td>${r.from_month ? esc(monthLabel(r.from_month)) : "Start (all months)"}</td><td>${esc(r.sob_text || "—")}</td></tr>`).join("");
  target.innerHTML = `${table(["Changed on", "Changed by", "Action", "Main code", "Level", "Used in", "SOB from month", "SOB split"], rows, search ? "No SOB changes match this search." : "No SOB changes yet.")}<div class="table-toolbar pagination-toolbar"><span>${result.total ? `${(offset + 1).toLocaleString()}–${(offset + result.items.length).toLocaleString()} of ${result.total.toLocaleString()}` : "No changes"}</span><div><button class="button button-small" data-sob-log-page="${Math.max(0, offset - 500)}" ${offset === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" data-sob-log-page="${offset + 500}" ${offset + result.items.length >= result.total ? "disabled" : ""}>Next</button></div></div>`;
}

// A code put in T1 Material Master by mistake moves to T2 or T3 Material Master
function openMoveMaterial(item) {
  if (!item) return;
  $("#modalTitle").textContent = `Move ${item.code} · ${item.name}`;
  $("#modalBody").innerHTML = `<div class="move-material"><p>Move this code out of T1 Material Master into:</p><div class="move-buttons"><button class="button button-primary" data-move-to="T2">Move to T2 Material Master</button><button class="button button-primary" data-move-to="T3">Move to T3 Material Master</button></div><p class="subtle">Code, name, category, unit and price move with it. Supplier material-master entries and T2/T3 supplier relationships switch to the moved code. If it is still in your BOMs, a T1 supplier relationship, Alternate codes, own stock or a transfer, the move is refused and the message lists what to remove first.</p></div>`;
  modalLayer.hidden = false;
  $("#modalBody").querySelectorAll("[data-move-to]").forEach((button) => button.addEventListener("click", async () => {
    try {
      const result = await api(`/api/planning/materials/${item.id}/move?tier=${button.dataset.moveTo}`, { method: "POST" });
      modalLayer.hidden = true;
      toast(`${result.code} moved to ${result.tier} Material Master (${result.relationships} relationship${result.relationships === 1 ? "" : "s"}, ${result.supplier_materials} supplier material entr${result.supplier_materials === 1 ? "y" : "ies"} switched)`);
      await loadCoreData();
      renderPage();
    } catch (error) { toast(error.message, true); }
  }));
}

// Material Master: 500 a page from the server; search covers all codes
async function loadMaterialMasterPage(offset = 0) {
  const target = $("#materialMasterTable");
  if (!target) return;
  const search = $("#masterSearch")?.value.trim() || "";
  const query = new URLSearchParams({ offset: String(offset), limit: "500", search });
  const result = await api(`/api/planning/materials?${query}`);
  state.materialPageItems = result.items;
  const canPrice = state.user.can_view_prices || state.user.role === "Admin";
  // Used in = BOM levels the code is a line of; a code used in no BOM is shown in red
  const rows = result.items.map((item) => `<tr data-edit-material="${item.id}" tabindex="0" class="${item.used_in ? "" : "material-unused"}" title="${item.used_in ? "" : "Not used in Model BOM Level 1, Level 2 or Level 3 Part BOM"}"><td>${esc(item.code)}</td><td>${esc(item.name)}</td><td>${esc(item.used_in || "Not used")}</td><td>${esc(item.category)}</td><td>${esc(item.uom)}</td><td>${number(item.stock_on_hand)}</td><td>${number(item.buffer_days, 0)} days</td><td>${canPrice ? money(item.unit_cost, item.currency) : "Restricted"}</td><td><button class="text-button" data-edit-material="${item.id}">Edit</button></td><td>${item.movable ? `<button class="text-button" data-move-material="${item.id}">Move to T2/T3</button>` : `<span class="subtle" title="Used in a BOM (L1 / L2 / L3) or in Alternate codes: it must stay in T1">—</span>`}</td><td><button class="text-button button-danger" data-delete-material="${item.id}">Delete</button></td></tr>`).join("");
  const size = result.limit;
  target.innerHTML = `${table(["Material code", "Material name", "Used in", "Category", "Unit", "On hand", "Buffer norm", "Part price", "Edit", "Move", "Delete"], rows, search ? "No materials match this search." : "No materials yet.")}<div class="table-toolbar pagination-toolbar"><span>${result.total ? `${(offset + 1).toLocaleString()}–${(offset + result.items.length).toLocaleString()} of ${result.total.toLocaleString()}${search ? " (search)" : ""}` : "No materials"}</span><div><button class="button button-small" data-material-page="${Math.max(0, offset - size)}" ${offset === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" data-material-page="${offset + size}" ${offset + result.items.length >= result.total ? "disabled" : ""}>Next</button></div></div>`;
  const count = $("#materialCount");
  if (count && !search) count.textContent = result.total.toLocaleString();
}


// Supplier/material relationship (T1: parts a supplier supplies you) and Supplier/material coverage
// (one T1 relationship: the supplier's part BOM, L1 -> L3, with the suppliers of every line).
// Whichever of the two pages is open is refreshed after an edit.
async function loadSupplierPartPage() {
  if ($("#coverageTable")) return loadCoverage();
  if ($("#relListTable")) return loadRelList();
  return null;
}

async function loadRelList() {
  const params = new URLSearchParams({ tier: "T1", limit: "1000", search: state.relSearch || "" });
  if (state.relSupplierId) params.set("supplier_id", state.relSupplierId);
  const exportQuery = new URLSearchParams({ search: state.relSearch || "" });  // the list shown: one supplier or all, with the search
  if (state.relSupplierId) exportQuery.set("supplier_id", state.relSupplierId);
  if ($("#relExcel")) { $("#relExcel").href = `/api/planning/supplier-relationships.xlsx?${exportQuery}`; $("#relPdf").href = `/api/planning/supplier-relationships.pdf?${exportQuery}`; }
  const result = await api(`/api/planning/supplier-parts?${params}`);
  state.supplierParts = result.items;
  const target = $("#relListTable");
  if (!target) return;
  const priced = canSeePrices();
  const open = (link, text) => `<td class="cell-link" data-open-coverage="${link.id}" title="Open Supplier/material coverage">${text}</td>`;
  const rows = result.items.map((link) => `<tr>${open(link, esc(link.supplier_code))}${open(link, esc(link.supplier_name))}${open(link, esc(link.material_code))}${open(link, esc(link.material_name))}${open(link, `${number(link.sob_percent, 1)}%`)}<td>${number(link.lead_time_days, 0)} days</td><td>${number(link.moq)}</td><td>${number(link.buffer_days, 0)} days</td><td>${priced && link.unit_price !== undefined ? money(link.unit_price, link.currency) : "Restricted"}</td><td>${priced ? esc(link.currency || "") : ""}</td><td><button class="text-button" data-open-coverage="${link.id}">Coverage</button><button class="text-button" data-edit-supplier-part="${link.id}">Edit</button><button class="text-button button-danger" data-delete-supplier-part="${link.id}">Delete</button></td></tr>`).join("");
  target.innerHTML = table(["Supplier Code", "Supplier Name", "T1 Part Code", "T1 Part Name", "SOB", "Lead time", "MOQ", "Buffer days", "Unit price", "Currency", "Action"], rows,
    state.relSupplierId ? "This supplier supplies no part yet: add rows above." : "No supplier/material relationships yet.");
}

function t1EntryRow() {
  const priced = canSeePrices();
  return `<tr>${entrySupplierCell("supplier")}${entryMaterialCell(false, "material")}<td><input class="form-control" type="number" min="0" max="100" step="0.01" data-f="sob_percent" placeholder="100" /></td><td><input class="form-control" type="number" min="0" data-f="lead_time_days" placeholder="0" /></td><td><input class="form-control" type="number" min="0" step="any" data-f="moq" placeholder="0" /></td><td><input class="form-control" type="number" min="0" data-f="buffer_days" placeholder="0" /></td><td>${priced ? '<input class="form-control" type="number" min="0" step="0.0001" data-f="unit_price" placeholder="0" />' : '<span class="subtle">Restricted</span>'}</td><td><select class="form-control" data-f="currency">${["INR", "USD", "EUR", "GBP", "JPY"].map((c) => `<option>${c}</option>`).join("")}</select></td>${entryRemoveCell}</tr>`;
}

// ---- coverage page
const PRICE_MODE_LABEL = { spot: "Spot", fixed_monthly: "Fixed monthly", every_n_months: "Every N months" };
async function loadCoverage() {
  const data = await api(`/api/planning/supplier-coverage/${state.coverageLinkId}`);
  state.coverage = data;
  state.supplierParts = [data.t1];
  renderCoverageTable();
}

function renderCoverageTable() {
  const target = $("#coverageTable");
  const data = state.coverage;
  if (!target || !data) return;
  const { t1, lines } = data;
  const priced = canSeePrices();
  const needed = state.covRequirement?.[t1.id];
  const hasNeed = needed !== undefined && needed !== "" && !Number.isNaN(Number(needed));
  const share = hasNeed ? Number(needed) * Number(t1.sob_percent || 0) / 100 : null;  // pieces this supplier makes
  const show = (value, unit = "") => (value === null ? "" : `${number(value, 2)}${unit ? ` ${esc(unit)}` : ""}`);
  const byParent = new Map();
  lines.forEach((line) => { const key = line.parent_line_id || 0; if (!byParent.has(key)) byParent.set(key, []); byParent.get(key).push(line); });
  const toggle = '<button type="button" class="tree-toggle" data-sb-toggle aria-expanded="true">−</button>';
  const price = (value, currency) => (priced && value !== undefined ? money(value, currency) : "Restricted");
  const terms = (item) => `<td>${number(item.lead_time_days, 0)} days</td><td>${number(item.moq)}</td><td>${number(item.buffer_days, 0)} days</td><td>${esc(PRICE_MODE_LABEL[item.price_mode] || item.price_mode || "")}</td><td>${esc(item.price_status || "")}</td><td>${price(item.unit_price, item.currency)}</td><td>${priced ? esc(item.currency || "") : ""}</td>`;
  const html = [`<tr class="tree-has-children cov-t1" data-tree-depth="0"><td class="tree-toggle-cell">${lines.length ? toggle : ""}</td><td>T1</td><td>${esc(t1.supplier_code)}</td><td>${esc(t1.supplier_name)}</td><td>${esc(t1.material_code)}</td><td>${esc(t1.material_name)}</td><td>${number(t1.sob_percent, 1)}%</td><td>1</td><td>1</td><td>${esc(t1.uom || "")}</td><td>${show(share, t1.uom)}</td>${terms(t1)}<td><button class="text-button" data-edit-supplier-part="${t1.id}">Edit T1</button><button class="text-button" data-cov-add-line="">+ L1 material</button></td></tr>`];
  const walk = (parentId, perParent, needParent) => {
    (byParent.get(parentId) || []).forEach((line) => {
      const depth = line.depth;  // 1 = L1
      const perPiece = perParent * line.quantity;  // per one T1 part
      const need = needParent === null ? null : needParent * line.quantity;
      const unit = line.component.uom || "";
      const sob = line.sources.reduce((sum, src) => sum + Number(src.sob_percent || 0), 0);
      const kids = (byParent.get(line.id) || []).length || line.sources.length;
      html.push(`<tr class="tree-has-children cov-line" data-tree-depth="${depth}"><td class="tree-toggle-cell">${kids ? toggle : ""}</td><td>L${depth} BOM</td><td></td><td></td><td>${esc(line.component.code)}${line.component.origin && line.component.origin !== "T1" ? ` <small class="subtle">${esc(line.component.origin)}</small>` : ""}</td><td>${esc(line.component.name)}</td><td class="${line.sources.length && Math.abs(sob - 100) > 0.001 ? "sob-off" : ""}">${line.sources.length ? `${number(sob, 1)}%` : "No supplier"}</td><td>${number(line.quantity, 4)}</td><td>${number(perPiece, 4)}</td><td>${esc(unit)}</td><td>${show(need, unit)}</td><td colspan="7"></td><td><button class="text-button" data-cov-add-source="${line.id}">+ T${depth + 1} supplier</button>${depth < 3 ? `<button class="text-button" data-cov-add-line="${line.id}">+ L${depth + 1} material</button>` : ""}<button class="text-button" data-cov-edit-line="${line.id}">Edit</button><button class="text-button button-danger" data-cov-delete-line="${line.id}">Delete</button></td></tr>`);
      line.sources.forEach((src, index) => {
        const code = src.alternate ? src.alternate.code : line.component.code;
        const name = src.alternate ? `${src.alternate.name} · alternate code` : line.component.name;
        html.push(`<tr class="tree-alternate cov-source" data-tree-depth="${depth + 1}"><td class="tree-toggle-cell"></td><td>T${depth + 1}</td><td>${esc(src.sub_supplier_code || "")}</td><td>${esc(src.sub_supplier || "")}</td><td>↳ ${esc(code)}</td><td>${esc(name)}</td><td>${number(src.sob_percent, 1)}%</td><td></td><td>${number(perPiece * src.sob_percent / 100, 4)}</td><td>${esc(unit)}</td><td>${show(need === null ? null : need * src.sob_percent / 100, unit)}</td>${terms(src)}<td><button class="text-button" data-cov-edit-source="${line.id}:${index}">Edit</button><button class="text-button button-danger" data-cov-delete-source="${line.id}:${index}">Delete</button></td></tr>`);
      });
      walk(line.id, perPiece, need);
    });
  };
  walk(0, 1, share);
  target.innerHTML = table(["", "Level", "Supplier Code", "Supplier Name", "Material code", "Material name", "SOB", "BOM qty", `Qty per ${t1.material_code}`, "Unit", "Requirement", "Lead time", "MOQ", "Buffer days", "Price mode", "Price status", "Unit price", "Currency", "Action"], html.join(""), "");
  const note = $("#covShare");
  if (note) note.textContent = hasNeed ? `${t1.supplier_name} makes ${number(t1.sob_percent, 1)}% = ${number(share, 2)} ${t1.uom || "pcs"} of ${number(Number(needed), 2)}` : "Type the requirement of the part to see the quantity of every material and supplier.";
}

// Upload supplier stock: the whole supply chain as a tree (closed to T1 suppliers at first; + / − per level,
// Open all / Close all; search keeps the suppliers whose tree mentions the text, opened)
function renderSupplyTree(open = false) {
  const target = $("#supplyTreeTable");
  const nodes = state.supplyTree || [];
  if (!target) return;
  const needle = ($("#supplyTreeSearch")?.value || "").trim().toLowerCase();
  const groups = [];
  nodes.forEach((node) => { if (node.depth === 0) groups.push([node]); else if (groups.length) groups[groups.length - 1].push(node); });
  const shown = needle ? groups.filter((group) => group.some((node) => `${node.supplier_code || ""} ${node.supplier_name || ""} ${node.material_code || ""} ${node.material_name || ""}`.toLowerCase().includes(needle))) : groups;
  const flat = shown.flat();
  const hasKids = (index) => flat[index + 1] && flat[index + 1].depth > flat[index].depth;
  const rows = flat.map((node, index) => {
    const kids = hasKids(index);
    const toggle = kids ? `<button type="button" class="tree-toggle" data-sb-toggle aria-expanded="${open}">${open ? "−" : "+"}</button>` : "";
    const kind = node.depth === 0 ? "tree-has-children st-supplier" : node.level.endsWith("supplier") ? "tree-alternate st-source" : node.level === "T1 part" ? "st-part" : "st-line";
    const material = node.depth === 0 ? `${node.parts} part${node.parts === 1 ? "" : "s"}` : esc(node.material_name || "");
    const stockCell = (value) => (value === null || value === undefined ? "" : number(value));
    return `<tr class="${kind}" data-tree-depth="${node.depth}" ${node.depth > 0 && !open ? "hidden" : ""}><td class="tree-toggle-cell">${toggle}</td><td>${esc(node.level)}</td><td>${esc(node.supplier_code || "")}</td><td>${esc(node.supplier_name || "")}</td><td>${node.level.endsWith("supplier") && node.depth > 0 ? "↳ " : ""}${esc(node.material_code || "")}</td><td>${material}</td><td>${node.sob_percent === null || node.sob_percent === undefined ? "" : `${number(node.sob_percent, 1)}%`}</td><td>${node.bom_qty === undefined ? "" : number(node.bom_qty, 4)}</td><td>${esc(node.uom || "")}</td><td>${stockCell(node.stock)}</td><td>${stockCell(node.available)}</td><td>${node.arrival_month ? esc(monthLabel(node.arrival_month)) : ""}</td></tr>`;
  }).join("");
  target.innerHTML = table(["", "Level", "Supplier Code", "Supplier Name", "Material code", "Material name", "SOB", "BOM qty", "Unit", "Stock quantity", "Available quantity", "Arrival month"], rows,
    needle ? "No supplier or material matches this search." : "No supplier/material relationships yet.");
}

function coverageLine(id) { return state.coverage.lines.find((line) => String(line.id) === String(id)); }

// + L1 / L2 / L3 material: each row = material, quantity, and (optional) one supplier with SOB and terms.
// More suppliers of the same material: more rows with the same code. SOB blank = what is left (100% for
// one supplier). If the suppliers of a material add up to less than 100%, Save first adds a row for the
// balance (same material, SOB left) to pick its supplier; Remove it and Save again to keep less than 100%.
const SUPPLIER_TERM_LABELS = ["SOB (%)", "Lead time (days)", "MOQ", "Buffer (days)", "Unit price", "Currency"];
function supplierTermCells() {
  return `<td><input class="form-control" type="number" min="0" max="100" step="0.01" data-f="sob_percent" placeholder="100" /></td><td><input class="form-control" type="number" min="0" data-f="lead_time_days" placeholder="0" /></td><td><input class="form-control" type="number" min="0" step="any" data-f="moq" placeholder="0" /></td><td><input class="form-control" type="number" min="0" data-f="buffer_days" placeholder="0" /></td><td>${canSeePrices() ? '<input class="form-control" type="number" min="0" step="0.0001" data-f="unit_price" placeholder="0" />' : '<span class="subtle">Restricted</span>'}</td><td><select class="form-control" data-f="currency">${["INR", "USD", "EUR", "GBP", "JPY"].map((c) => `<option>${c}</option>`).join("")}</select></td>`;
}
function supplierTerms(row) {
  const value = (key) => cellValue(row, `[data-f="${key}"]`);
  const out = { lead_time_days: Number(value("lead_time_days") || 0), moq: Number(value("moq") || 0), buffer_days: Number(value("buffer_days") || 0), price_mode: "spot", price_interval_months: 1, currency: value("currency") || "INR" };
  if (row.querySelector('[data-f="unit_price"]')) out.unit_price = Number(value("unit_price") || 0);
  return out;
}
function openCoverageAddLine(parentId) {
  const parent = parentId ? coverageLine(parentId) : null;
  const t1 = state.coverage.t1;
  const level = parent ? parent.depth + 1 : 1;
  const tier = `T${level + 1}`;
  const per = parent ? `1 ${parent.component.uom || "unit"} of ${parent.component.code}` : `1 ${t1.uom || "piece"} of ${t1.material_code}`;
  $("#modalTitle").textContent = `Add L${level} material ${parent ? `under ${parent.component.code}` : `to ${t1.material_code}`}`;
  $("#modalBody").innerHTML = `<div class="part-edit"><p class="subtle">The BOM of ${esc(parent ? `${parent.component.code} ${parent.component.name}` : `${t1.material_code} ${t1.material_name}`)}: each material, the quantity needed for ${esc(per)} (e.g. 2 kg steel per beam; 1.7 kg iron ore per kg steel), and its ${tier} supplier with SOB (blank = 100%). Codes come from your T1, T2 or T3 Material Master. A second supplier of the same material: one more row with the same code. If the suppliers of a material add up to less than 100%, Save adds a row for the rest to pick its supplier.</p>${entryTable("covLineTable", ["Material code", "Material name", `Quantity per ${parent ? parent.component.code : t1.material_code}`, `${tier} Supplier Code`, `${tier} Supplier Name`, ...SUPPLIER_TERM_LABELS])}<div class="modal-footer"><button type="button" class="button button-small" id="covLineAddRow">+ Add row</button><button type="button" class="button button-primary" id="covLineSave">Save</button></div></div>`;
  modalLayer.hidden = false;
  const entries = bindEntryTable("covLineTable", () => `<tr>${entryMaterialCell(true, "material")}<td><input class="form-control" type="number" min="0" step="any" data-f="quantity" placeholder="0" /></td>${entrySupplierCell("supplier")}${supplierTermCells()}${entryRemoveCell}</tr>`);
  $("#covLineAddRow").addEventListener("click", () => entries.add());
  let balanceOffered = false;
  $("#covLineSave").addEventListener("click", async () => {
    const groups = new Map();  // material -> {rows, item, quantity, sources}
    for (const [index, row] of entries.rows().entries()) {
      const typedCode = cellValue(row, '[data-pick-key="material"][data-pick-part="code"]').trim();
      const item = Number(cellValue(row, '[data-row-material="material"]'));
      const tierItem = Number(cellValue(row, '[data-row-tier-material="material"]'));
      const supplierId = Number(cellValue(row, '[data-row-supplier="supplier"]'));
      const typedSupplier = cellValue(row, '[data-pick-key="supplier"][data-pick-part="code"]').trim();
      const sobText = cellValue(row, '[data-f="sob_percent"]');
      const qtyText = cellValue(row, '[data-f="quantity"]');
      if (!typedCode && !item && !tierItem && !typedSupplier && !qtyText) continue;  // empty row
      if (!item && !tierItem) return toast(`Row ${index + 1}: ${typedCode ? `${typedCode} is not found; pick the code from the hint list (T1, T2 or T3 Material Master).` : "type the material code."}`, true);
      const key = item ? `i${item}` : `t${tierItem}`;
      if (!groups.has(key)) {
        if (!(Number(qtyText) > 0)) return toast(`Row ${index + 1}: type the quantity (more than 0).`, true);
        groups.set(key, { rows: [], code: typedCode, quantity: Number(qtyText), item: item || null, tierItem: tierItem || null, sources: [], blanks: [] });
      }
      const group = groups.get(key);
      if (qtyText !== "" && Number(qtyText) !== group.quantity) return toast(`Row ${index + 1}: ${group.code} has quantity ${number(group.quantity, 4)} on an earlier row; use the same quantity (or leave it blank).`, true);
      group.rows.push(row);
      if (!supplierId) {
        if (typedSupplier) return toast(`Row ${index + 1}: ${typedSupplier} is not in the Supplier register; pick the supplier from the hint list.`, true);
        if (sobText !== "" && group.rows.length > 1) return toast(`Row ${index + 1}: pick the supplier for this ${number(Number(sobText), 2)}% of ${group.code}, or Remove the row.`, true);
        continue;  // a material without a supplier yet
      }
      if (group.sources.some((src) => src.sub_supplier_id === supplierId)) return toast(`Row ${index + 1}: this supplier is already on an earlier row of ${group.code}.`, true);
      const source = { sub_supplier_id: supplierId, sob_percent: sobText === "" ? null : Number(sobText), ...supplierTerms(row) };
      if (source.sob_percent !== null && !(source.sob_percent > 0)) return toast(`Row ${index + 1}: SOB must be more than 0%.`, true);
      group.sources.push(source);
    }
    if (!groups.size) return toast("Fill in at least one row.", true);
    for (const group of groups.values()) {
      const typed = group.sources.filter((src) => src.sob_percent !== null).reduce((sum, src) => sum + src.sob_percent, 0);
      const blanks = group.sources.filter((src) => src.sob_percent === null);
      blanks.forEach((src) => { src.sob_percent = Math.max(0, 100 - typed) / blanks.length; });  // blank = what is left
      if (blanks.some((src) => !(src.sob_percent > 0))) return toast(`${group.code}: no SOB is left for a supplier with blank SOB; type its SOB.`, true);
      const total = group.sources.reduce((sum, src) => sum + src.sob_percent, 0);
      if (total > 100.0001) return toast(`${group.code}: SOB of its suppliers adds up to ${number(total, 2)}%; at most 100%.`, true);
      group.total = total;
    }
    // suppliers below 100%: first time, add a row for the balance (same material) to pick its supplier
    const short = [...groups.values()].filter((group) => group.sources.length && group.total < 99.9999);
    if (short.length && !balanceOffered) {
      balanceOffered = true;
      for (const group of short) {
        const last = group.rows[group.rows.length - 1];
        const pickItem = (group.rows[0].pickItems?.["material:material"] || []).find((hint) => (group.item ? hint.kind === "T1" && hint.id === group.item : hint.kind !== "T1" && hint.id === group.tierItem));
        const row = entries.add();
        last.after(row);
        if (pickItem) setRowPick(row, "material", "material", pickItem);
        row.querySelector('[data-f="quantity"]').value = group.quantity;
        row.querySelector('[data-f="sob_percent"]').value = Number((100 - group.total).toFixed(4));
        row.classList.add("balance-row");
      }
      return toast(`${short.map((group) => `${group.code}: ${number(group.total, 2)}%`).join(", ")}. A row for the remaining SOB is added: pick its ${tier} supplier and Save, or Remove that row and Save to keep less than 100%.`);
    }
    let saved = 0;
    for (const group of groups.values()) {
      const body = { link_id: t1.id, parent_line_id: parent ? parent.id : null, item_id: group.item, tier_material_id: group.tierItem, quantity: group.quantity, sources: group.sources };
      try { await api("/api/planning/supplier-coverage/lines", { method: "POST", body: JSON.stringify(body) }); saved += 1; }
      catch (error) { toast(`${group.code}: ${error.message}${saved ? ` (${saved} material${saved === 1 ? "" : "s"} above saved)` : ""}`, true); state.suppliers = await api("/api/suppliers"); return loadCoverage(); }
    }
    modalLayer.hidden = true;
    toast(`${saved} material line${saved === 1 ? "" : "s"} added`);
    state.suppliers = await api("/api/suppliers");
    loadCoverage();
  });
}

function openCoverageEditLine(line) {
  $("#modalTitle").textContent = `Edit L${line.depth} ${line.component.code}`;
  $("#modalBody").innerHTML = `<form id="covQtyForm" class="form-grid"><label class="field"><span>Quantity per unit of the line above (${esc(line.component.uom || "")})</span><input class="form-control" type="number" min="0" step="any" name="quantity" value="${line.quantity}" required /></label><div class="modal-footer"><button class="button button-primary" type="submit">Save</button></div></form>`;
  modalLayer.hidden = false;
  $("#covQtyForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const quantity = Number(new FormData(event.currentTarget).get("quantity"));
    if (!(quantity > 0)) return toast("Quantity must be more than 0.", true);
    try {
      await api(`/api/planning/supplier-bom/lines/${line.id}`, { method: "PATCH", body: JSON.stringify({ quantity, unit_price: line.entered_price ?? null, currency: line.entered_currency ?? null }) });
      modalLayer.hidden = true;
      toast("Quantity saved");
      loadCoverage();
    } catch (error) { toast(error.message, true); }
  });
}

// suppliers of a line are saved together (SOB up to 100%)
function sourcePayload(src) {
  const out = { sub_supplier_id: src.sub_supplier_id, sob_percent: Number(src.sob_percent), lead_time_days: Number(src.lead_time_days || 0), moq: Number(src.moq || 0), buffer_days: Number(src.buffer_days || 0),
    price_mode: src.price_mode || "spot", price_interval_months: Number(src.price_interval_months || 1), unit_price: Number(src.unit_price || 0), currency: src.currency || "INR" };
  if (src.alternate) out[src.alternate.kind === "our" ? "item_id" : "supplier_material_id"] = src.alternate.id;
  return out;
}
async function saveSources(line, sources, message) {
  try {
    await api(`/api/planning/supplier-bom/lines/${line.id}/sources`, { method: "PUT", body: JSON.stringify(sources) });
    modalLayer.hidden = true;
    toast(message);
    state.suppliers = await api("/api/suppliers");
    loadCoverage();
  } catch (error) { toast(error.message, true); }
}

function openCoverageSource(line, index = null) {
  if (!canSeePrices()) return toast("Price access is needed to change the suppliers of a line (their prices are saved with them).", true);
  const tier = `T${line.depth + 1}`;
  const editing = index === null ? null : line.sources[index];
  const others = line.sources.filter((_, i) => i !== index);
  const used = others.reduce((sum, src) => sum + Number(src.sob_percent || 0), 0);
  $("#modalTitle").textContent = `${editing ? "Edit" : "Add"} ${tier} supplier of ${line.component.code}`;
  $("#modalBody").innerHTML = `<div class="part-edit"><p class="subtle">${tier} suppliers of ${esc(line.component.code)} ${esc(line.component.name)}: ${others.length ? others.map((src) => `${esc(src.sub_supplier)} ${number(src.sob_percent, 1)}%`).join(", ") : "none yet"}${editing ? " (besides this one)" : ""}. SOB left: <strong>${number(Math.max(0, 100 - used), 1)}%</strong>.${editing ? "" : " Add one or more rows."}</p>${entryTable("covSourceTable", ["Supplier code", "Supplier name", "SOB (%)", "Lead time (days)", "MOQ", "Buffer (days)", "Price mode", "Price interval (months)", "Unit price", "Currency"])}<div class="modal-footer">${editing ? "" : '<button type="button" class="button button-small" id="covSourceAddRow">+ Add row</button>'}<button type="button" class="button button-primary" id="covSourceSave">Save</button></div></div>`;
  modalLayer.hidden = false;
  const entries = bindEntryTable("covSourceTable", () => `<tr>${entrySupplierCell("supplier")}<td><input class="form-control" type="number" min="0" max="100" step="0.01" data-f="sob_percent" placeholder="100" /></td><td><input class="form-control" type="number" min="0" data-f="lead_time_days" placeholder="0" /></td><td><input class="form-control" type="number" min="0" step="any" data-f="moq" placeholder="0" /></td><td><input class="form-control" type="number" min="0" data-f="buffer_days" placeholder="0" /></td><td><select class="form-control" data-f="price_mode"><option value="spot">Spot</option><option value="fixed_monthly">Fixed monthly</option><option value="every_n_months">Every N months</option></select></td><td><input class="form-control" type="number" min="1" data-f="price_interval_months" value="1" /></td><td><input class="form-control" type="number" min="0" step="0.0001" data-f="unit_price" placeholder="0" /></td><td><select class="form-control" data-f="currency">${["INR", "USD", "EUR", "GBP", "JPY"].map((c) => `<option>${c}</option>`).join("")}</select></td>${entryRemoveCell}</tr>`);
  if (editing) {
    const row = entries.rows()[0];
    row.querySelector("[data-entry-remove]")?.remove();
    const supplier = state.suppliers.find((item) => item.id === editing.sub_supplier_id);
    if (supplier) setRowPick(row, "supplier", "supplier", supplier);
    ["sob_percent", "lead_time_days", "moq", "buffer_days", "price_mode", "price_interval_months", "unit_price", "currency"].forEach((key) => { const input = row.querySelector(`[data-f="${key}"]`); if (input && editing[key] !== undefined && editing[key] !== null) input.value = editing[key]; });
  } else {
    $("#covSourceAddRow").addEventListener("click", () => entries.add());
  }
  let balanceOffered = false;
  $("#covSourceSave").addEventListener("click", async () => {
    const typed = [];
    for (const [rowIndex, row] of entries.rows().entries()) {
      const supplierId = Number(cellValue(row, '[data-row-supplier="supplier"]'));
      const code = cellValue(row, '[data-pick-key="supplier"][data-pick-part="code"]').trim();
      if (!supplierId && !code && row.classList.contains("balance-row")) return toast(`Row ${rowIndex + 1}: pick the supplier for the remaining SOB, or Remove the row.`, true);
      if (!supplierId && !code) continue;
      if (!supplierId) return toast(`Row ${rowIndex + 1}: ${code} is not in the Supplier register; pick the supplier from the hint list.`, true);
      if (supplierId === line.sources.find((_, i) => i !== index && line.sources[i].sub_supplier_id === supplierId)?.sub_supplier_id) return toast(`Row ${rowIndex + 1}: this supplier is already a ${tier} supplier of ${line.component.code}; use its Edit.`, true);
      const value = (key) => cellValue(row, `[data-f="${key}"]`);
      const sob = value("sob_percent") === "" ? 100 : Number(value("sob_percent"));
      if (!(sob > 0)) return toast(`Row ${rowIndex + 1}: SOB must be more than 0% (delete the supplier instead).`, true);
      typed.push({ sub_supplier_id: supplierId, sob_percent: sob, lead_time_days: Number(value("lead_time_days") || 0), moq: Number(value("moq") || 0), buffer_days: Number(value("buffer_days") || 0),
        price_mode: value("price_mode"), price_interval_months: Number(value("price_interval_months") || 1), unit_price: Number(value("unit_price") || 0), currency: value("currency"), alternate: editing?.alternate || null });
    }
    if (!typed.length) return toast("Fill in at least one row.", true);
    const total = used + typed.reduce((sum, src) => sum + src.sob_percent, 0);
    if (total > 100.0001) return toast(`SOB adds up to ${number(total, 2)}%: at most ${number(Math.max(0, 100 - used), 2)}% is left for ${line.component.code}. Lower another supplier's SOB with its Edit first.`, true);
    if (index === null && total < 99.9999 && !balanceOffered) {  // first time below 100%: a row for the balance
      balanceOffered = true;
      const row = entries.add();
      row.querySelector('[data-f="sob_percent"]').value = Number((100 - total).toFixed(4));
      row.classList.add("balance-row");
      return toast(`${line.component.code}: ${number(total, 2)}%. A row for the remaining ${number(100 - total, 2)}% is added: pick its ${tier} supplier and Save, or Remove that row and Save to keep less than 100%.`);
    }
    const list = index === null ? [...line.sources.map(sourcePayload), ...typed.map(sourcePayload)] : line.sources.map((src, i) => sourcePayload(i === index ? typed[0] : src));
    saveSources(line, list, `${tier} supplier${typed.length === 1 ? "" : "s"} saved for ${line.component.code}`);
  });
}

const SUPPLIER_TABS = ["Supplier register", "Supplier/material relationship", "Upload supplier stock", "Supplier-to-supplier material transfer", "Transfer schedule", "Supplier materials", "Supplier BOM"];

// ---- Excel-like entry tables (one row per entry, every field a column, + Add row, Save all)
// Column headings have the same filter as other tables (sort, search, Select all, Clear all);
// the list of values is made from what is typed in the rows when the filter opens.
function entryHeading(label) {
  if (!label) return '<th><div class="table-heading"><span class="column-resizer" aria-hidden="true"></span></div></th>';
  return `<th><div class="table-heading"><details class="multi-filter table-column-filter entry-column-filter"><summary aria-label="Filter ${esc(label)}">${esc(label)}</summary><div class="multi-popover"><div class="column-sort"><button type="button" class="text-button" data-column-sort="asc">Sort A → Z</button><button type="button" class="text-button" data-column-sort="desc">Sort Z → A</button></div><input class="form-control multi-search" type="search" placeholder="Search ${esc(label.toLowerCase())}" /><div class="column-filter-actions"><button type="button" class="text-button" data-column-select-all>Select all</button><button type="button" class="text-button" data-column-clear>Clear all</button></div><div class="multi-options"></div><button type="button" class="button button-small button-primary column-filter-apply" data-column-apply>Apply</button></div></details><span class="column-resizer" aria-hidden="true"></span></div></th>`;
}
function entryTable(id, labels) {
  return `<div class="table-wrap entry-wrap"><table class="entry-table" id="${id}"><thead><tr>${labels.map(entryHeading).join("")}${entryHeading("")}</tr></thead><tbody></tbody></table></div>`;
}
// What a cell shows: typed text, the chosen option, or plain text
function cellText(cell) {
  if (!cell) return "";
  const control = cell.querySelector("input:not([type=hidden]), select");
  if (!control) return cell.textContent.trim();
  if (control.tagName === "SELECT") return control.value ? (control.selectedOptions[0]?.textContent || "").trim() : "";
  return control.value.trim();
}
// Code and name in their own columns, each one box: type and pick from the hint list; choosing
// either the code or the name fills the other. The chosen id is kept in a hidden field.
let entryListCounter = 0;
function pickCell(kind, key, part, placeholder) {
  entryListCounter += 1;
  const list = `entry-list-${entryListCounter}`;
  return `<input class="form-control" list="${list}" data-pick="${kind}" data-pick-key="${key}" data-pick-part="${part}" placeholder="${esc(placeholder)}" autocomplete="off" /><datalist id="${list}"></datalist>`;
}
function entryMaterialCell(withTierCodes = false, key = "material") {
  const extra = withTierCodes ? `<input type="hidden" data-row-tier-material="${key}" /><input type="hidden" data-pick-tier-codes="${key}" value="1" />` : "";
  return `<td class="entry-pick">${pickCell("material", key, "code", withTierCodes ? "Part code (T1, T2 or T3)" : "Type part code")}<input type="hidden" data-row-material="${key}" />${extra}</td><td class="entry-name">${pickCell("material", key, "name", "Type part name")}</td>`;
}
function entrySupplierCell(key, optional = false) {
  return `<td class="entry-pick">${pickCell("supplier", key, "code", optional ? "Parent code (T2/T3 rows)" : "Type supplier code")}<input type="hidden" data-row-supplier="${key}" /></td><td class="entry-name">${pickCell("supplier", key, "name", optional ? "Parent name (T2/T3 rows)" : "Type supplier name")}</td>`;
}
const entryRemoveCell = '<td><button type="button" class="text-button button-danger" data-entry-remove>Remove</button></td>';
function fillHints(row, kind, key, items) {
  row.pickItems = row.pickItems || {};
  row.pickItems[`${kind}:${key}`] = items;
  row.querySelectorAll(`[data-pick="${kind}"][data-pick-key="${key}"]`).forEach((input) => {
    const byCode = input.dataset.pickPart === "code";
    input.nextElementSibling.innerHTML = items.map((item) => `<option value="${esc(byCode ? item.code : item.name)}">${esc(byCode ? item.name : item.code)}${item.kind && item.kind !== "T1" ? ` · ${item.kind} Material Master` : ""}</option>`).join("");
  });
}
async function loadRowMaterials(row, search = "", onlyKey = null) {
  const keys = [...new Set([...row.querySelectorAll('[data-pick="material"]')].map((input) => input.dataset.pickKey))].filter((key) => !onlyKey || key === onlyKey);
  for (const key of keys) {
    try {
      const result = await api(`/api/planning/materials?${new URLSearchParams({ offset: "0", limit: "100", search, purchased: "true" })}`);
      let items = result.items.map((item) => ({ ...item, kind: "T1" }));
      if (row.querySelector(`[data-pick-tier-codes="${key}"]`)) {  // T2/T3 rows: T2/T3 Material Master codes too
        const [t2, t3] = await Promise.all(["T2", "T3"].map((tier) => api(`/api/planning/master-codes?${new URLSearchParams({ tier, search, limit: "50" })}`)));
        items = items.concat(t2.items.map((item) => ({ ...item, kind: "T2" })), t3.items.map((item) => ({ ...item, kind: "T3" })));
      }
      fillHints(row, "material", key, items);
    } catch (error) { toast(error.message, true); }
  }
}
// Set the hidden id from the typed code / name (exact match, capitals ignored) and fill the other box
function resolvePick(row, input) {
  const { pick: kind, pickKey: key, pickPart: part } = input.dataset;
  const items = row.pickItems?.[`${kind}:${key}`] || [];
  const text = input.value.trim().toLowerCase();
  const found = text ? items.find((item) => String(part === "code" ? item.code : item.name).toLowerCase() === text) : null;
  const tierHidden = kind === "material" ? row.querySelector(`[data-row-tier-material="${key}"]`) : null;
  const hidden = row.querySelector(kind === "material" ? `[data-row-material="${key}"]` : `[data-row-supplier="${key}"]`);
  const other = row.querySelector(`[data-pick="${kind}"][data-pick-key="${key}"][data-pick-part="${part === "code" ? "name" : "code"}"]`);
  const before = hidden.value + (tierHidden ? `|${tierHidden.value}` : "");
  if (tierHidden) tierHidden.value = "";
  if (found && found.kind && found.kind !== "T1" && tierHidden) {
    hidden.value = "";
    tierHidden.value = String(found.id);
    const spelled = part === "code" ? found.code : found.name;
    if (input.value !== spelled) input.value = spelled;
    other.value = part === "code" ? found.name : found.code;
  } else if (found) {
    hidden.value = String(found.id);
    const spelled = part === "code" ? found.code : found.name;  // the master's spelling ("st-steel" -> "ST-STEEL")
    if (input.value !== spelled) input.value = spelled;
    other.value = part === "code" ? found.name : found.code;
  } else {
    hidden.value = "";
    if (!text) other.value = "";
  }
  input.classList.toggle("pick-unknown", Boolean(text) && !found);
  if (hidden.value + (tierHidden ? `|${tierHidden.value}` : "") !== before) hidden.dispatchEvent(new Event("change", { bubbles: true }));
}
// Pick a row's item by id (used by tests and when filling a row from code)
function setRowPick(row, kind, key, item) {
  fillHints(row, kind, key, [item, ...(row.pickItems?.[`${kind}:${key}`] || []).filter((other) => !(other.id === item.id && other.kind === item.kind))]);
  const input = row.querySelector(`[data-pick="${kind}"][data-pick-key="${key}"][data-pick-part="code"]`);
  input.value = item.code;
  resolvePick(row, input);
}
// One listener set per entry table: hints, remove row, column filters; onChange(row, target) for extra wiring
function bindEntryTable(tableId, newRow, onChange = null) {
  const tableNode = $(`#${tableId}`);
  const body = tableNode.tBodies[0];
  const add = () => {
    body.insertAdjacentHTML("beforeend", newRow());
    const row = body.lastElementChild;
    row.querySelectorAll('[data-pick="supplier"][data-pick-part="code"]').forEach((input) => fillHints(row, "supplier", input.dataset.pickKey, state.suppliers));
    loadRowMaterials(row);
    return row;
  };
  body.addEventListener("input", (event) => {
    const input = event.target.closest("[data-pick]");
    if (!input) return;
    const row = input.closest("tr");
    resolvePick(row, input);
    const key = input.dataset.pickKey;
    if (input.dataset.pick === "material" && !row.querySelector(`[data-row-material="${key}"]`).value && !row.querySelector(`[data-row-tier-material="${key}"]`)?.value) {
      clearTimeout(row.searchTimer);
      row.searchTimer = setTimeout(async () => { await loadRowMaterials(row, input.value.trim(), key); resolvePick(row, input); }, 250);
    }
  });
  body.addEventListener("click", (event) => {
    if (!event.target.closest("[data-entry-remove]")) return;
    if (body.rows.length > 1) event.target.closest("tr").remove();
    else toast("Keep at least one row (leave it empty if not needed).", true);
  });
  if (onChange) body.addEventListener("change", (event) => { if (!event.target.matches("[data-pick]")) onChange(event.target.closest("tr"), event.target); });
  // column filters: options from the typed values when a filter opens; Apply hides other rows
  tableNode.querySelectorAll("thead .entry-column-filter").forEach((details) => details.addEventListener("toggle", () => {
    if (!details.open) return;
    const column = details.closest("th").cellIndex;
    const chosen = new Set([...details.querySelectorAll("[data-column-value]:checked")].map((input) => input.value));
    const values = [...new Set([...body.rows].map((row) => cellText(row.cells[column])).filter(Boolean))].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    details.querySelector(".multi-options").innerHTML = values.length ? values.map((value) => `<label><input type="checkbox" data-column-value="true" value="${esc(value)}" ${chosen.has(value) ? "checked" : ""} />${esc(value)}</label>`).join("") : '<span class="subtle">No values typed yet</span>';
  }));
  tableNode.closest(".entry-wrap").applyTableFilters = () => {
    const selections = [...tableNode.querySelectorAll("thead th")].map((th) => [th.cellIndex, [...th.querySelectorAll("[data-column-value]:checked")].map((input) => input.value)]).filter(([, values]) => values.length);
    [...body.rows].forEach((row) => { row.hidden = selections.some(([column, values]) => !values.includes(cellText(row.cells[column]))); });
  };
  add();
  return { add, rows: () => [...body.rows], reset: () => { body.innerHTML = ""; add(); } };
}
const cellValue = (row, selector) => row.querySelector(selector)?.value ?? "";

function relationshipEntryRow() {
  const priced = canSeePrices();
  return `<tr>${entrySupplierCell("supplier")}${entryMaterialCell(true)}<td><select class="form-control" data-f="tier"><option value="T1">T1 (supplies you)</option><option value="T2">T2 (supplies a T1 supplier)</option><option value="T3">T3 (supplies a T2 supplier)</option></select></td>${entrySupplierCell("parent", true)}<td><input class="form-control" type="number" min="0" max="100" step="0.01" data-f="sob_percent" value="100" /></td><td><input class="form-control" type="number" min="0" data-f="lead_time_days" value="14" /></td><td><input class="form-control" type="number" min="0" step="any" data-f="moq" value="0" /></td><td><input class="form-control" type="number" min="0" data-f="buffer_days" value="0" /></td><td>${priced ? '<input class="form-control" type="number" min="0" step="0.0001" data-f="unit_price" value="0" />' : '<span class="subtle">Restricted</span>'}</td><td><select class="form-control" data-f="currency">${["INR", "USD", "EUR", "GBP", "JPY"].map((c) => `<option>${c}</option>`).join("")}</select></td><td><select class="form-control" data-f="price_mode"><option value="spot">Spot</option><option value="fixed_monthly">Fixed monthly</option><option value="every_n_months">Every N months</option></select></td><td><input class="form-control" type="number" min="1" data-f="price_interval_months" value="1" /></td>${entryRemoveCell}</tr>`;
}
function stockEntryRow() {
  return `<tr>${entrySupplierCell("supplier")}<td><select class="form-control" data-row-link><option value="">Choose supplier first</option></select></td><td><input class="form-control" type="number" min="0" step="any" data-f="quantity" /></td><td><input class="form-control" type="number" min="0" step="any" data-f="available_quantity" placeholder="= quantity" /></td><td><input class="form-control" type="month" data-f="arrival_month" /></td>${entryRemoveCell}</tr>`;
}
function transferEntryRow() {
  return `<tr>${entryMaterialCell()}${entrySupplierCell("from")}${entrySupplierCell("to")}<td><input class="form-control" type="number" min="0.001" step="any" data-f="quantity" /></td><td><input class="form-control" type="month" data-f="transfer_month" value="${monthOffset(monthNow(), 1)}" /></td><td><input class="form-control" data-f="note" /></td>${entryRemoveCell}</tr>`;
}

const SUPPLIER_TIER_TABS = [["T1", "T1 Supplier Master"], ["T2", "T2 Supplier Master"], ["T3", "T3 Supplier Master"]];
// Tick boxes T1 / T2 / T3: a supplier can be more than one (e.g. TATA Steel is T1 for you and T2 for Supangita)
function supplierTierField(selected) {
  return `<fieldset class="field full tier-ticks"><legend>Supplier tier (one or more)</legend>${SUPPLIER_TIER_TABS.map(([tier]) => `<label class="check-option"><input type="checkbox" name="tiers" value="${tier}" ${selected.includes(tier) ? "checked" : ""} /> ${tier}${tier === "T1" ? " (supplies you)" : tier === "T2" ? " (supplies a T1 supplier)" : " (supplies a T2 supplier)"}</label>`).join("")}</fieldset>`;
}

function openSupplierEdit(supplier) {
  if (!supplier) return;
  $("#modalTitle").textContent = `Edit ${supplier.code} · ${supplier.name}`;
  $("#modalBody").innerHTML = `<form id="supplierEditForm" class="form-grid">${field("Supplier code", "code", "text", { required: true, maxlength: 24, value: supplier.code })}${field("Supplier name", "name", "text", { required: true, value: supplier.name })}${supplierTierField(supplier.tiers || ["T1"])}${categoryField(supplier.category)}${field("Country", "country", "text", { required: true, value: supplier.country || "" })}${field("Contact person", "contact_name", "text", { value: supplier.contact_name || "" })}${field("Email", "email", "email", { value: supplier.email || "" })}${field("Phone", "phone", "text", { value: supplier.phone || "" })}<label class="field"><span>Approval</span><select class="form-control" name="status">${["Pending review", "Approved", "On hold"].map((status) => `<option ${status === supplier.status ? "selected" : ""}>${status}</option>`).join("")}</select></label><div class="modal-footer field full"><button class="button button-primary">Save supplier</button></div></form>`;
  modalLayer.hidden = false;
  bindCategoryField($("#supplierEditForm"));
  $("#supplierEditForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const values = takeCategory(Object.fromEntries(data.entries()));
    values.tiers = data.getAll("tiers");
    if (!values.tiers.length) return toast("Tick at least one tier: T1, T2 or T3.", true);
    try {
      await api(`/api/suppliers/${supplier.id}`, { method: "PATCH", body: JSON.stringify(values) });
      await reloadCategories();
      modalLayer.hidden = true;
      toast("Supplier saved");
      renderSupplierControl();
    } catch (error) { toast(error.message, true); }
  });
}

// Edit a saved relationship: the whole row (supplier, part, tier, parent, SOB, lead time, MOQ,
// buffer, price, agreement) in the same one-row table as Supplier/material relationship.
function openSupplierPartEdit(link) {
  if (!link) return;
  $("#modalTitle").textContent = `Edit ${link.supplier_name} · ${link.material_code} (${link.tier})`;
  $("#modalBody").innerHTML = `<div class="part-edit"><p class="subtle">Change any column, e.g. a T2 entered by mistake. T2/T3 rows need the parent supplier.</p>${entryTable("partEditTable", ["Supplier code", "Supplier name", "Part code", "Part name", "Tier", "Parent supplier code (T2/T3)", "Parent supplier name", "SOB (%)", "Lead time (days)", "MOQ", "Buffer (days)", "Unit price", "Currency", "Price agreement", "Price interval (months)"])}<div class="modal-footer"><button type="button" class="button button-danger" data-delete-supplier-part="${link.id}">Delete relationship</button><button type="button" class="button button-primary" id="partEditSave">Save relationship</button></div></div>`;
  modalLayer.hidden = false;
  const entries = bindEntryTable("partEditTable", relationshipEntryRow);
  const row = entries.rows()[0];
  row.querySelector("[data-entry-remove]")?.remove();
  const supplierOf = (id) => state.suppliers.find((supplier) => supplier.id === id);
  if (supplierOf(link.supplier_id)) setRowPick(row, "supplier", "supplier", supplierOf(link.supplier_id));
  if (link.parent_supplier_id && supplierOf(link.parent_supplier_id)) setRowPick(row, "supplier", "parent", supplierOf(link.parent_supplier_id));
  setRowPick(row, "material", "material", { id: link.material_id || link.tier_material_id, code: link.material_code, name: link.material_name, kind: link.code_master || "T1" });
  const set = (key, value) => { const input = row.querySelector(`[data-f="${key}"]`); if (input && value !== undefined && value !== null) input.value = value; };
  set("tier", link.tier); set("sob_percent", link.sob_percent); set("lead_time_days", link.lead_time_days); set("moq", link.moq); set("buffer_days", link.buffer_days);
  set("unit_price", link.unit_price); set("currency", link.currency); set("price_mode", link.price_mode); set("price_interval_months", link.price_interval_months);
  $("#partEditSave").addEventListener("click", async () => {
    const value = (key) => cellValue(row, `[data-f="${key}"]`);
    const body = {
      supplier_id: Number(cellValue(row, '[data-row-supplier="supplier"]')) || null,
      material_id: Number(cellValue(row, "[data-row-material]")) || null, tier_material_id: Number(cellValue(row, "[data-row-tier-material]")) || null,
      tier: value("tier"), parent_supplier_id: Number(cellValue(row, '[data-row-supplier="parent"]')) || null,
      sob_percent: Number(value("sob_percent")), lead_time_days: Number(value("lead_time_days")), moq: Number(value("moq")), buffer_days: Number(value("buffer_days")),
      currency: value("currency"), price_mode: value("price_mode"), price_interval_months: Number(value("price_interval_months") || 1),
    };
    if (row.querySelector('[data-f="unit_price"]')) body.unit_price = Number(value("unit_price"));
    if (!body.supplier_id) return toast("Pick the supplier from the hint list.", true);
    if (!body.material_id && !body.tier_material_id) return toast("Pick the part code from the hint list.", true);
    try {
      await api(`/api/planning/supplier-parts/${link.id}`, { method: "PATCH", body: JSON.stringify(body) });
      modalLayer.hidden = true;
      toast("Relationship saved");
      state.suppliers = await api("/api/suppliers");
      loadSupplierPartPage(state.supplierOffset || 0);
    } catch (error) { toast(error.message, true); }
  });
}

// Material picker for supplier forms: type-to-search on the server, purchased parts only
// (no in-house BOM, never a negative BOM quantity), so it works with 10,000+ codes.
function purchasedMaterialPicker(prefix) {
  return `<label class="field"><span>Search purchased materials</span><input class="form-control" id="${prefix}MaterialSearch" type="search" placeholder="Type a material code or name" /></label>
    <label class="field"><span>Material (purchased parts)</span><select class="form-control" name="material_id" id="${prefix}MaterialSelect" required><option value="">Loading…</option></select></label>`;
}

// Supplier picker: search box filters the supplier list
function supplierPicker(name, label, optional = false) {
  const options = state.suppliers.map((supplier) => `<option value="${supplier.id}">${esc(supplier.code)} · ${esc(supplier.name)}</option>`).join("");
  return `<label class="field"><span>${esc(label)}</span><input class="form-control supplier-search" type="search" data-supplier-search="${esc(name)}" placeholder="Search supplier" /><select class="form-control" name="${esc(name)}" ${optional ? "" : "required"}>${optional ? '<option value="">None (T1)</option>' : ""}${options}</select></label>`;
}

function bindSupplierSearch(root) {
  root.querySelectorAll("[data-supplier-search]").forEach((input) => {
    const select = input.nextElementSibling;
    input.addEventListener("input", () => {
      const needle = input.value.trim().toLowerCase();
      [...select.options].forEach((option) => { option.hidden = Boolean(option.value) && !option.textContent.toLowerCase().includes(needle); });
      const first = [...select.options].find((option) => !option.hidden && option.value);
      if (first && select.selectedOptions[0]?.hidden) select.value = first.value;
    });
  });
}

// Supplier chooser for the Supplier materials / Supplier BOM tabs (remembers the choice)
function supplierChooser(id, label) {
  const chosen = String(state[id] || "");
  const options = state.suppliers.map((supplier) => `<option value="${supplier.id}" ${String(supplier.id) === chosen ? "selected" : ""}>${esc(supplier.code)} · ${esc(supplier.name)}</option>`).join("");
  return `<label class="field"><span>${esc(label)}</span><input class="form-control supplier-search" type="search" data-supplier-search="${id}" placeholder="Search supplier" /><select class="form-control" id="${id}"><option value="">Choose supplier</option>${options}</select></label>`;
}

function canSeePrices() {
  return Boolean(state.user.can_view_prices || state.user.role === "Admin");
}

// ---------------- Supplier materials tab
function initSupplierMaterialsTab() {
  const choose = $("#smSupplier");
  const reload = () => loadSupplierMaterials();
  choose.addEventListener("change", () => { state.smSupplier = choose.value; reload(); });
  $("#smSearch").addEventListener("input", () => { clearTimeout(state.searchTimer); state.searchTimer = setTimeout(reload, 250); });
  const from = $("#smFrom");
  const pickSearch = $("#smPickSearch");
  const refreshPick = async () => {
    if (from.value === "Own") return;
    try {
      const result = await api(`/api/planning/master-codes?${new URLSearchParams({ tier: from.value, search: pickSearch.value, limit: "100" })}`);
      const previous = $("#smPick").value;
      $("#smPick").innerHTML = result.items.map((m) => `<option value="${m.id}">${esc(m.code)} · ${esc(m.name)}</option>`).join("") || '<option value="">No codes found</option>';
      if ([...$("#smPick").options].some((o) => o.value === previous)) $("#smPick").value = previous;
    } catch (error) { toast(error.message, true); }
  };
  const switchFrom = () => {
    const own = from.value === "Own";
    content.querySelectorAll(".sm-pick").forEach((node) => { node.hidden = own; });
    content.querySelectorAll(".sm-own").forEach((node) => { node.hidden = !own; });
    refreshPick();
  };
  from.addEventListener("change", switchFrom);
  pickSearch.addEventListener("input", () => { clearTimeout(state.searchTimer); state.searchTimer = setTimeout(refreshPick, 250); });
  switchFrom();
  $("#smAddForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    if (!state.smSupplier) return toast("Choose a supplier first.", true);
    const values = Object.fromEntries(new FormData(form).entries());
    const body = { supplier_id: Number(state.smSupplier), unit_price: Number(values.unit_price || 0), currency: values.currency,
      lead_time_days: Number(values.lead_time_days || 0), buffer_days: Number(values.buffer_days || 0) };
    if (from.value === "Own") Object.assign(body, { code: values.code, name: values.name, uom: values.uom });
    else {
      if (!$("#smPick").value) return toast("Choose a code.", true);
      body[from.value === "T1" ? "item_id" : "tier_material_id"] = Number($("#smPick").value);
    }
    try {
      const saved = await api("/api/planning/supplier-materials", { method: "POST", body: JSON.stringify(body) });
      toast(`${saved.code} added to the supplier's master`);
      const keep = from.value;
      form.reset();
      from.value = keep;  // keep the chosen source and show its fields
      switchFrom();
      reload();
    } catch (error) { toast(error.message, true); }
  });
  $("#smUploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    if (!state.smSupplier) return toast("Choose a supplier first.", true);
    try {
      const result = await uploadExcel(`/api/planning/supplier-materials/upload?supplier_id=${state.smSupplier}`, new FormData(form).get("workbook"));
      toast(`${result.rows} row(s): ${result.added} added, ${result.updated} updated`);
      form.reset();
      reload();
    } catch (error) { toast(error.message, true); }
  });
  reload();
}

async function loadSupplierMaterials() {
  const target = $("#smTable");
  if (!target) return;
  const headers = ["From", "Code", "Name", "Unit", "Supplier's price", "Currency", "Lead time", "Buffer", "Edit", "Delete"];
  if (!state.smSupplier) { target.innerHTML = table(headers, "", "Choose a supplier to see its material master."); return; }
  const query = new URLSearchParams({ supplier_id: state.smSupplier, search: $("#smSearch")?.value || "", limit: "1000" });
  const result = await api(`/api/planning/supplier-materials?${query}`);
  state.smItems = result.items;
  const rows = result.items.map((m) => `<tr><td>${esc(m.source === "Own" ? "Supplier's own" : m.source)}</td><td>${esc(m.code)}</td><td>${esc(m.name)}</td><td>${esc(m.uom)}</td><td>${m.unit_price === undefined ? "Restricted" : number(m.unit_price, 2)}</td><td>${esc(m.currency || "")}</td><td>${number(m.lead_time_days, 0)} days</td><td>${number(m.buffer_days, 0)} days</td><td><button class="text-button" data-sm-edit="${m.id}">Edit</button></td><td><button class="text-button button-danger" data-sm-delete="${m.id}">Delete</button></td></tr>`).join("");
  target.innerHTML = `${table(headers, rows, "No materials for this supplier yet.")}<div class="table-toolbar pagination-toolbar"><span>${result.total.toLocaleString()} material${result.total === 1 ? "" : "s"}${result.total > result.items.length ? ` (first ${result.items.length.toLocaleString()} shown: search to narrow)` : ""}</span></div>`;
}

function openSupplierMaterialEdit(material) {
  const own = material.source === "Own";
  $("#modalTitle").textContent = `Edit ${material.code}${own ? "" : ` (from ${material.source}: code and name follow your master)`}`;
  $("#modalBody").innerHTML = `<form id="smEditForm" class="form-grid">${own ? `${field("Code", "code", "text", { required: true, value: material.code })}${field("Name", "name", "text", { required: true, value: material.name })}${unitField("Unit", "uom", material.uom)}` : `<label class="field full"><span>Code · name</span><input class="form-control" value="${esc(material.code)} · ${esc(material.name)}" readonly /></label>`}${field("Supplier's lead time (days)", "lead_time_days", "number", { min: 0, value: material.lead_time_days ?? 0 })}${field("Supplier's buffer stock (days)", "buffer_days", "number", { min: 0, value: material.buffer_days ?? 0 })}${canSeePrices() ? `${field("Unit price", "unit_price", "number", { min: 0, step: "0.0001", value: material.unit_price ?? 0 })}<label class="field"><span>Currency</span><select class="form-control" name="currency">${["INR", "USD", "EUR", "GBP", "JPY"].map((c) => `<option ${c === material.currency ? "selected" : ""}>${c}</option>`).join("")}</select></label>` : ""}<div class="modal-footer field full"><button class="button button-primary">Save</button></div></form>`;
  modalLayer.hidden = false;
  $("#smEditForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const values = Object.fromEntries(new FormData(event.currentTarget).entries());
    try {
      await api(`/api/planning/supplier-materials/${material.id}`, { method: "PATCH", body: JSON.stringify({ supplier_id: material.supplier_id, code: values.code ?? material.code, name: values.name ?? material.name, uom: values.uom ?? material.uom, unit_price: Number(values.unit_price ?? material.unit_price ?? 0), currency: values.currency || material.currency || "INR", lead_time_days: Number(values.lead_time_days || 0), buffer_days: Number(values.buffer_days || 0) }) });
      modalLayer.hidden = true;
      toast("Supplier material saved");
      loadSupplierMaterials();
    } catch (error) { toast(error.message, true); }
  });
}

// ---------------- Supplier BOM tab (cost breakdown)
function initSupplierBomTab() {
  const choose = $("#sbSupplier");
  choose.addEventListener("change", () => { state.sbSupplier = choose.value; state.sbProduct = ""; loadSupplierBomProducts(); });
  $("#sbProduct").addEventListener("change", (event) => { state.sbProduct = event.target.value; loadSupplierBom(); });
  loadSupplierBomProducts();
}

async function loadSupplierBomProducts() {
  const select = $("#sbProduct");
  if (!state.sbSupplier) { select.innerHTML = '<option value="">Choose supplier first</option>'; return loadSupplierBom(); }
  const result = await api(`/api/planning/supplier-bom/products?supplier_id=${state.sbSupplier}`);
  if (!result.items.length) {
    select.innerHTML = '<option value="">No parts: add a T1 supplier/material relationship first</option>';
    state.sbProduct = "";
  } else {
    if (!result.items.some((item) => String(item.id) === String(state.sbProduct))) state.sbProduct = String(result.items[0].id);
    select.innerHTML = result.items.map((item) => `<option value="${item.id}" ${String(item.id) === String(state.sbProduct) ? "selected" : ""}>${esc(item.code)} · ${esc(item.name)}${item.has_bom ? "" : " (no BOM yet)"}</option>`).join("");
  }
  loadSupplierBom();
}

async function loadSupplierBom() {
  const tableTarget = $("#sbTable");
  const summary = $("#sbSummary");
  const actions = $("#sbActions");
  if (!tableTarget) return;
  const headers = ["", "L1 code", "L1 name", "L2 code", "L2 name", "L3 code", "L3 name", "Source", "Quantity", "Unit", "Unit price", "Line cost", "Price from", "Sources", "Add line below", "Edit", "Delete"];
  if (!state.sbSupplier || !state.sbProduct) {
    actions.innerHTML = "";
    summary.innerHTML = "";
    tableTarget.innerHTML = table(headers, "", "Choose a supplier and a part.");
    return;
  }
  const view = await api(`/api/planning/supplier-bom?supplier_id=${state.sbSupplier}&material_id=${state.sbProduct}`);
  state.sbView = view;
  const query = `supplier_id=${state.sbSupplier}&material_id=${state.sbProduct}`;
  actions.innerHTML = `<button class="button button-small button-primary" data-sb-add="">Add L1 line</button><form id="sbUploadForm" class="bom-upload-inline"><input class="form-control bom-file-input" type="file" name="workbook" accept=".xlsx" required aria-label="Supplier BOM workbook" /><button class="button button-small" type="submit" title="Columns: Parent code (blank = L1), Code, Quantity, Unit price (optional), Source (optional: our / supplier). Replaces this BOM.">Upload BOM</button></form><a class="button button-small" href="/api/planning/supplier-bom/xlsx?${query}">Download Excel</a><a class="button button-small" href="/api/planning/supplier-bom/pdf?${query}">Download PDF</a>`;
  if (view.should_cost !== undefined) {
    const quoted = view.quoted_price;
    const gap = view.gap;
    summary.innerHTML = `<div class="sb-summary"><div><span>Should-cost</span><strong>${money(view.should_cost, view.currency)}</strong></div><div><span>Quoted price</span><strong>${quoted === null ? "No quote" : money(quoted, view.currency)}</strong></div><div class="${gap === null ? "" : gap >= 0 ? "sb-gap-plus" : "sb-gap-minus"}"><span>Quoted − should-cost</span><strong>${gap === null ? "–" : `${money(gap, view.currency)}${view.should_cost ? ` (${number((gap / view.should_cost) * 100, 1)}%)` : ""}`}</strong></div>${view.excluded_lines.length ? `<p class="sb-warning">${view.excluded_lines.length} line(s) priced in another currency are not added to the should-cost (shown in red).</p>` : ""}</div>`;
  } else summary.innerHTML = "";
  const paths = {};
  const rows = view.lines.map((line) => {
    const mine = [...(paths[line.parent_line_id] || []), line.component.code, line.component.name];
    paths[line.id] = mine;
    const cells = [...mine, "", "", "", "", "", ""].slice(0, 6).map((text, index) => `<td class="${index < mine.length - 2 ? "tree-inherited" : ""}">${esc(text)}</td>`).join("");
    const priced = line.unit_price !== undefined;
    const sources = line.sources.length ? line.sources.map((src) => `${esc(src.sub_supplier || "(no supplier)")}${src.alternate ? ` · ${esc(src.alternate.code)}` : ""} ${number(src.sob_percent, 0)}%`).join("<br>") : "";
    return `<tr class="${line.has_children ? "tree-has-children" : ""} ${line.currency_mismatch ? "sb-mismatch" : ""}" data-sb-line="${line.id}" data-tree-depth="${line.depth - 1}"><td class="tree-toggle-cell">${line.has_children ? `<button type="button" class="tree-toggle" data-sb-toggle aria-expanded="true">−</button>` : ""}</td>${cells}<td>${esc(line.component.origin === "Own" ? "Supplier's own" : `${line.component.origin}`)}</td><td>${number(line.quantity, 4)}</td><td>${esc(line.component.uom)}</td><td>${priced ? money(line.unit_price, line.currency) : "Restricted"}</td><td>${priced ? money(line.line_cost, line.currency) : "Restricted"}</td><td>${esc(line.price_source)}</td><td>${line.has_children ? "" : `<button class="text-button" data-sb-sources="${line.id}">${line.sources.length ? "Edit" : "Add"}</button>${sources ? `<div class="sb-sources">${sources}</div>` : ""}`}</td><td>${line.depth < 3 ? `<button class="text-button" data-sb-add="${line.id}">+ L${line.depth + 1}</button>` : ""}</td><td><button class="text-button" data-sb-edit="${line.id}">Edit</button></td><td><button class="text-button button-danger" data-sb-delete="${line.id}">Delete</button></td></tr>`;
  }).join("");
  tableTarget.innerHTML = table(headers, rows, "No BOM lines yet: use Add L1 line or Upload BOM.");
  $("#sbUploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    if (state.sbView.lines.length && !window.confirm("Replace this supplier BOM with the uploaded file?")) return;
    try {
      const result = await uploadExcel(`/api/planning/supplier-bom/upload?${query}`, new FormData(form).get("workbook"));
      toast(`${result.rows} supplier BOM line(s) uploaded`);
      loadSupplierBom();
    } catch (error) { toast(error.message, true); }
  });
}

function componentPicker() {
  return `<label class="field"><span>Search the supplier's material master</span><input class="form-control" id="sbPartSearch" type="search" placeholder="Type code or name" /></label>
    <label class="field full"><span>Part</span><select class="form-control" id="sbPart" required></select></label>`;
}

async function fillComponentOptions(selectNode, kind, search, includeBlank = false) {
  const query = new URLSearchParams({ search, limit: "100" });
  const result = kind === "our"
    ? await api(`/api/planning/materials?${query}`)
    : await api(`/api/planning/supplier-materials?supplier_id=${state.sbSupplier}&${query}`);
  const previous = selectNode.value;
  selectNode.innerHTML = `${includeBlank ? '<option value="">Same code</option>' : ""}${result.items.map((item) => `<option value="${kind}:${item.id}">${esc(item.code)} · ${esc(item.name)}</option>`).join("")}`;
  if ([...selectNode.options].some((option) => option.value === previous)) selectNode.value = previous;  // keep the user's pick
}

function openSupplierBomLine(parentLineId, line = null) {
  const editing = Boolean(line);
  $("#modalTitle").textContent = editing ? `Edit ${line.component.code}` : (parentLineId ? "Add line below" : "Add L1 line");
  $("#modalBody").innerHTML = `<form id="sbLineForm" class="form-grid">${editing ? `<label class="field full"><span>Part</span><input class="form-control" value="${esc(line.component.code)} · ${esc(line.component.name)}" readonly /></label>` : componentPicker()}
    ${field("Quantity per parent", "quantity", "number", { required: true, step: "any", value: line ? line.quantity : 1 })}
    ${canSeePrices() ? `${field("Unit price (blank = from sources / material list)", "unit_price", "number", { min: 0, step: "0.0001", value: line?.entered_price ?? "" })}<label class="field"><span>Currency</span><select class="form-control" name="currency">${["INR", "USD", "EUR", "GBP", "JPY"].map((c) => `<option ${c === (line?.entered_currency || state.sbView.currency) ? "selected" : ""}>${c}</option>`).join("")}</select></label>` : ""}
    <div class="modal-footer field full"><button class="button button-primary">Save line</button></div></form>`;
  modalLayer.hidden = false;
  if (!editing) {
    const refresh = () => fillComponentOptions($("#sbPart"), "supplier", $("#sbPartSearch").value).catch((error) => toast(error.message, true));
    $("#sbPartSearch").addEventListener("input", () => { clearTimeout(state.searchTimer); state.searchTimer = setTimeout(refresh, 250); });
    refresh();
  }
  $("#sbLineForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const quantity = Number(data.get("quantity"));
    const priceText = data.get("unit_price");
    const unitPrice = priceText === null || priceText === "" ? null : Number(priceText);
    const currency = unitPrice === null ? null : data.get("currency");
    try {
      if (editing) {
        await api(`/api/planning/supplier-bom/lines/${line.id}`, { method: "PATCH", body: JSON.stringify({ quantity, unit_price: unitPrice, currency }) });
      } else {
        const [kind, id] = ($("#sbPart").value || ":").split(":");
        if (!id) return toast("Choose a part.", true);
        await api("/api/planning/supplier-bom/lines", { method: "POST", body: JSON.stringify({
          supplier_id: Number(state.sbSupplier), material_id: Number(state.sbProduct), parent_line_id: parentLineId ? Number(parentLineId) : null,
          item_id: kind === "our" ? Number(id) : null, supplier_material_id: kind === "supplier" ? Number(id) : null, quantity, unit_price: unitPrice, currency }) });
      }
      modalLayer.hidden = true;
      toast("Supplier BOM line saved");
      loadSupplierBom();
    } catch (error) { toast(error.message, true); }
  });
}

function sourceRow(source = {}) {
  const suppliers = state.suppliers.map((s) => `<option value="${s.id}" ${s.id === source.sub_supplier_id ? "selected" : ""}>${esc(s.code)} · ${esc(s.name)}</option>`).join("");
  const alternate = source.alternate ? `<option value="${source.alternate.kind}:${source.alternate.id}" selected>${esc(source.alternate.code)} · ${esc(source.alternate.name)}</option>` : "";
  return `<tr data-sb-source><td><input class="form-control supplier-search" type="search" data-supplier-search="row" placeholder="Search" /><select class="form-control" data-src-supplier><option value="">(none)</option>${suppliers}</select></td>
    <td><input class="form-control" type="search" data-src-search placeholder="Alternate code from supplier's master (optional)" /><select class="form-control" data-src-alt><option value="">Same code</option>${alternate}</select></td>
    <td><input class="form-control" type="number" min="0" step="0.0001" data-src-price value="${source.unit_price ?? 0}" /></td>
    <td><select class="form-control" data-src-currency>${["INR", "USD", "EUR", "GBP", "JPY"].map((c) => `<option ${c === (source.currency || state.sbView.currency) ? "selected" : ""}>${c}</option>`).join("")}</select></td>
    <td><input class="form-control" type="number" min="0.01" max="100" step="0.01" data-src-sob value="${source.sob_percent ?? 100}" /></td>
    <td><button type="button" class="text-button button-danger" data-src-remove>Remove</button></td></tr>`;
}

function openSupplierBomSources(line) {
  $("#modalTitle").textContent = `Sources of ${line.component.code}: sub-suppliers / alternate codes`;
  $("#modalBody").innerHTML = `<div class="sb-sources-editor"><p class="subtle">Unit price of this line = Σ (price × SOB) ÷ Σ SOB. SOB in total at most 100%. Sub-suppliers come from your Supplier register; an alternate code comes from the supplier's material master.</p><div class="new-bom-actions"><button type="button" class="button button-small" id="sbAddSource">+ Add source</button><button type="button" class="button button-small button-primary" id="sbSaveSources">Save sources</button></div><div class="table-wrap"><table><thead><tr><th>Sub-supplier</th><th>Alternate code</th><th>Unit price</th><th>Currency</th><th>SOB %</th><th></th></tr></thead><tbody id="sbSourceRows">${(line.sources.length ? line.sources : [{}]).map(sourceRow).join("")}</tbody></table></div></div>`;
  modalLayer.hidden = false;
  const body = $("#sbSourceRows");
  const wire = () => {
    bindSupplierSearch(body);
    body.querySelectorAll("[data-src-search]:not([data-wired])").forEach((input) => {
      input.dataset.wired = "true";
      const row = input.closest("tr");
      const refresh = () => fillComponentOptions(row.querySelector("[data-src-alt]"), "supplier", input.value, true).catch((error) => toast(error.message, true));
      input.addEventListener("input", () => { clearTimeout(row.timer); row.timer = setTimeout(refresh, 250); });
    });
  };
  wire();
  $("#sbAddSource").addEventListener("click", () => { body.insertAdjacentHTML("beforeend", sourceRow()); wire(); });
  body.addEventListener("click", (event) => { if (event.target.closest("[data-src-remove]")) event.target.closest("tr").remove(); });
  $("#sbSaveSources").addEventListener("click", async () => {
    const payload = [...body.rows].map((row) => {
      const [kind, id] = (row.querySelector("[data-src-alt]").value || ":").split(":");
      return {
        sub_supplier_id: Number(row.querySelector("[data-src-supplier]").value) || null,
        item_id: kind === "our" && id ? Number(id) : null, supplier_material_id: kind === "supplier" && id ? Number(id) : null,
        unit_price: Number(row.querySelector("[data-src-price]").value || 0), currency: row.querySelector("[data-src-currency]").value,
        sob_percent: Number(row.querySelector("[data-src-sob]").value),
      };
    });
    try {
      await api(`/api/planning/supplier-bom/lines/${line.id}/sources`, { method: "PUT", body: JSON.stringify(payload) });
      modalLayer.hidden = true;
      toast("Sources saved");
      loadSupplierBom();
    } catch (error) { toast(error.message, true); }
  });
}

async function renderSupplierControl() {
  if (!SUPPLIER_TABS.includes(state.supplierTab)) state.supplierTab = SUPPLIER_TABS[0];
  const tab = state.supplierTab;
  state.suppliers = await api("/api/suppliers");
  const tabs = `<div class="tab-strip">${SUPPLIER_TABS.map((name) => `<button class="tab-button ${tab === name ? "active" : ""}" data-supplier-tab="${esc(name)}">${esc(name)}</button>`).join("")}</div>`;
  let body = "";
  if (tab === "Supplier/material relationship" && state.coverageLinkId) {
    body = `<nav class="crumbs" id="covCrumbs"></nav><section class="panel table-panel">${panelHeader("Supplier/material coverage", "The part as this supplier makes it: its BOM (L1 materials, their L2 BOM and L3 BOM, each with quantity per unit of the line above) and the suppliers of every material with SOB (T2 for L1, T3 for L2, T4 for L3). A part may have only the T1 supplier, or go down to L3: add what you know. Type the requirement of the part to see the quantity of every material and supplier: e.g. 60,000 beams × Savera 60% × 2 kg steel = 72,000 kg; JGI 30% = 21,600 kg, TATA 70% = 50,400 kg; × 1.7 kg ore per kg steel = 122,400 kg ore. + / − opens and closes a level. Excel: Download template gives every part of this supplier with its BOM and suppliers (Level L1/L2/L3 = material with BOM qty; T2/T3/T4 = supplier of the line above, with SOB and terms); edit it and Upload coverage. Uploading replaces the BOM of each part in the file.")}
      <form id="coverageUploadForm" class="panel-body inline-upload"><label class="field"><span>Supplier-wise coverage workbook (all parts of this supplier)</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label><button class="button button-primary" type="submit">Upload coverage</button><a class="button" id="coverageTemplate" href="#">Download template</a></form>
      <div class="table-toolbar"><div class="toolbar-group"><label class="field"><span>Requirement of the part (all suppliers)</span><input class="form-control" type="number" min="0" step="any" id="covRequirement" placeholder="e.g. 60000" /></label><span class="subtle" id="covShare"></span></div><span class="toolbar-add"><button type="button" class="button button-small button-primary" data-cov-add-line="">+ L1 material</button></span></div>
      <div id="coverageTable" class="coverage-tree"><div class="tree-loading">Loading…</div></div>
    </section>`;
  } else if (tab === "Supplier/material relationship") {
    const chosen = state.suppliers.find((supplier) => supplier.id === state.relSupplierId) || null;
    if (!chosen) state.relSupplierId = null;
    body = `${chosen ? `<nav class="crumbs"><strong>${esc(chosen.code)} ${esc(chosen.name)}</strong> <a href="#" class="crumb-action" data-rel-all>Show all suppliers</a></nav>` : ""}
      <section class="panel table-panel">${panelHeader(`Supplier/material relationship${chosen ? ` · ${chosen.code} ${chosen.name}` : ""}`, `The parts ${chosen ? esc(chosen.name) : "each supplier"} supplies to you (T1): SOB, lead time, MOQ, buffer, price. One supplier can supply many parts; one part can have several suppliers (SOB up to 100% per part). Click the supplier or part (first 5 columns) to open its Supplier/material coverage: the part's BOM and the T2 / T3 suppliers of its materials. Type new rows below (+ Add row, Save all), or upload Excel with the same columns: Supplier Code, Supplier Name, T1 Part Code, T1 Part Name, SOB, Lead time, MOQ, Buffer days, Unit price, Currency (the template${chosen ? " has this supplier's rows: edit and upload back" : " has the columns"}). Needed: Supplier Code (or name) and T1 Part Code; a row already there is updated; the whole file is rejected if a row is wrong.`)}
        <form id="supplierPartUploadForm" class="panel-body inline-upload">
          <label class="field"><span>Relationship workbook</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label>
          <button class="button button-primary" type="submit">Upload relationships</button><a class="button" id="relTemplate" href="/api/planning/supplier-relationships/template.xlsx${chosen ? `?supplier_id=${chosen.id}` : ""}">Download template${chosen ? ` (${esc(chosen.code)}'s rows)` : ""}</a>
        </form>
        <div class="table-toolbar"><label class="search-field"><input class="form-control" id="relSearch" type="search" placeholder="Search supplier or part" value="${esc(state.relSearch || "")}" /></label><a class="button button-small" id="relExcel" href="#">Download Excel</a><a class="button button-small" id="relPdf" href="#">Download PDF</a><span class="toolbar-add"><button type="button" class="button button-small" id="relAddRow">+ Add row</button><button type="button" class="button button-small button-primary" id="relSaveAll">Save all</button></span></div>
        ${entryTable("relEntryTable", ["Supplier Code", "Supplier Name", "T1 Part Code", "T1 Part Name", "SOB (%)", "Lead time (days)", "MOQ", "Buffer days", "Unit price", "Currency"])}
        <div id="relListTable"><div class="tree-loading">Loading…</div></div>
      </section>`;
  } else if (tab === "Upload supplier stock") {
    body = `<section class="panel table-panel">${panelHeader("Upload supplier stock", "Upload Excel: Download template gives the supply chain tree of the month with Stock quantity, Available quantity and Arrival month columns; type the stock on the T1 part rows and on the T2 / T3 supplier rows (stock at that supplier for that part) and upload it; material rows are for reading. Then the supply chain tree: every T1 supplier, the parts it supplies you, each part's BOM materials and their T2 / T3 suppliers (+ / − or Open all / Close all), with the stock entered for the month on each part. Then upload a workbook, or type each supplier's stock for the month in the table below (+ Add row, Save all; no Excel needed). Workbook columns: supplier_code, material_code, tier, parent_supplier_code, quantity, available_quantity, arrival_month; tier and parent identify the exact relationship when the same supplier/material exists at more than one level. One row per supplier and material; a row already reported for the month is updated.")}
      <div class="table-toolbar"><div class="toolbar-group"><label class="field"><span>Report month</span><input class="form-control" type="month" id="stockMonth" required value="${monthNow()}" /></label><label class="field"><span>Workspace</span>${workspaceSelect()}</label></div></div>
      <form id="supplierStockForm" class="panel-body inline-upload">
        <label class="field"><span>Supplier stock workbook (the template below, or your sheet with supplier_code, material_code, quantity)</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label>
        <button class="button button-primary" type="submit">Upload Excel</button><a class="button" id="stockTemplate" href="#">Download template</a><a class="button button-small" href="/api/planning/templates/supplier-stock.xlsx">Simple template (one row per supplier)</a>
      </form>
      <h3 class="subtable-heading">Supply chain tree: T1 suppliers → parts → BOM materials → T2 suppliers → L2 materials → T3 suppliers</h3>
      <div class="table-toolbar"><label class="search-field"><input class="form-control" id="supplyTreeSearch" type="search" placeholder="Search supplier or material" /></label><span class="toolbar-add"><a class="button button-small" id="supplyTreeExcel" href="/api/planning/supply-tree.xlsx">Download Excel</a><a class="button button-small" id="supplyTreePdf" href="/api/planning/supply-tree.pdf">Download PDF</a><button type="button" class="button button-small" id="supplyTreeOpen">Open all</button><button type="button" class="button button-small" id="supplyTreeClose">Close all</button></span></div>
      <div id="supplyTreeTable" class="supply-tree"><div class="tree-loading">Loading…</div></div>
      <h3 class="subtable-heading">Enter stock</h3>
      <div class="new-bom-actions entry-actions"><button type="button" class="button button-small" id="stockAddRow">+ Add row</button><button type="button" class="button button-small button-primary" id="stockSaveAll">Save all</button></div>
      ${entryTable("stockEntryTable", ["Supplier code", "Supplier name", "Supplier's material (relationship)", "Stock quantity", "Available quantity", "Arrival month (if not yet there)"])}
      <h3 class="subtable-heading">Stock entered for the month</h3>
      <div id="stockListTable"></div>
    </section>`;
  } else if (tab === "Supplier register") {
    // one list for T1, T2 and T3 suppliers; the Tiers column (and its filter) says which
    const tierCounts = SUPPLIER_TIER_TABS.map(([tier]) => `${tier} ${state.suppliers.filter((supplier) => (supplier.tiers || ["T1"]).includes(tier)).length}`).join(" · ");
    const supplierRows = state.suppliers.map((supplier) => `<tr data-search="${esc(`${supplier.code} ${supplier.name} ${supplier.category}`.toLowerCase())}"><td class="cell-link" data-open-supplier-rel="${supplier.id}" title="Open Supplier/material relationship">${esc(supplier.code)}</td><td class="cell-link" data-open-supplier-rel="${supplier.id}" title="Open Supplier/material relationship">${esc(supplier.name)}</td><td>${esc((supplier.tiers || ["T1"]).join(", "))}</td><td>${esc(supplier.category)}</td><td>${esc(supplier.country || "")}</td><td>${esc(supplier.contact_name || "")}</td><td>${esc(supplier.email || "")}</td><td>${esc(supplier.phone || "")}</td><td>${esc(supplier.status)}</td><td>${supplier.status === "Pending review" ? `<button class="text-button" data-approve-supplier="${supplier.id}">Approve</button>` : ""}<button class="text-button" data-edit-supplier="${supplier.id}">Edit</button><button class="text-button button-danger" data-delete-supplier="${supplier.id}">Delete</button></td></tr>`).join("");
    body = `<section class="panel table-panel">${panelHeader(`Supplier Master (${tierCounts})`, `Click a supplier's Code or name to open its Supplier/material relationship (the parts it supplies). T1 = supplies you; T2 = supplies a T1 supplier; T3 = supplies a T2 supplier. One supplier can be in more than one (e.g. TATA Steel: T1 for you, T2 for Supangita); a relationship saved at a tier adds that tier to the supplier; filter the Tiers column to see one tier. Add one supplier with Add supplier, or many with an Excel upload (your own headings are fine; suppliers are T1 unless the sheet has a Supplier tier column). Needed: Supplier name. Optional: Supplier code (blank = next free SUP- code), Supplier tier, Category, Country, Contact person, Email, Phone, Approval. A row whose code (or, without a code, name) is already here updates that supplier; the whole file is rejected if any row is wrong.`)}
      <form id="supplierUploadForm" class="panel-body inline-upload"><label class="field"><span>Supplier list workbook</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label><button class="button button-primary" type="submit">Upload suppliers</button><a class="button" href="/api/planning/templates/suppliers.xlsx">Download template</a></form>
      <div class="table-toolbar"><span class="toolbar-add"><button class="button button-small button-primary" data-action="new-supplier">Add supplier</button></span><label class="search-field"><input class="form-control" id="supplierSearch" type="search" placeholder="Search supplier" /></label><div id="supplierFilters"></div>${tableExportLinks("suppliers")}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div><div id="supplierRegisterTable">${table(["Code", "Supplier", "Tiers", "Category", "Country", "Contact", "Email", "Phone", "Approval", "Action"], supplierRows, "No suppliers yet.")}</div></section>`;
  } else if (tab === "Supplier-to-supplier material transfer") {
    body = `<section class="panel table-panel">${panelHeader("Supplier-to-supplier material transfer", "Record transferable supplier-held surplus, one row per transfer; + Add row for more, then Save all. The receiving supplier is credited in projected coverage in the transfer month. Saved transfers are listed in Transfer schedule.")}
      <div class="table-toolbar"><div class="toolbar-group"><label class="field"><span>Workspace</span>${workspaceSelect()}</label></div></div>
      <div class="new-bom-actions entry-actions"><button type="button" class="button button-small" id="transferAddRow">+ Add row</button><button type="button" class="button button-small button-primary" id="transferSaveAll">Save all</button></div>
      ${entryTable("transferEntryTable", ["Part code (purchased part)", "Part name", "From supplier code", "From supplier name", "To supplier code", "To supplier name", "Transfer quantity", "Transfer / arrival month", "Transfer note"])}
    </section>`;
  } else if (tab === "Supplier materials") {
    body = `<section class="panel table-panel">${panelHeader("Supplier materials", "Each supplier's own material codes (e.g. solder wire, paste, connector). They are used only in that supplier's BOMs for cost breakdown.")}
      <div class="table-toolbar"><div class="toolbar-group">${supplierChooser("smSupplier", "Supplier")}<label class="search-field"><input class="form-control" id="smSearch" type="search" placeholder="Search code or name" /></label></div></div>
      <form id="smAddForm" class="panel-body form-grid supplier-inline-form">
        <label class="field"><span>Code from</span><select class="form-control" id="smFrom"><option value="T1">T1 Material Master</option><option value="T2">T2 Material Master</option><option value="T3">T3 Material Master</option><option value="Own">Supplier's own code</option></select></label>
        <label class="field sm-pick"><span>Search code / name</span><input class="form-control" id="smPickSearch" type="search" placeholder="Type to search" /></label>
        <label class="field sm-pick"><span>Code · name (from your master)</span><select class="form-control" id="smPick"></select></label>
        <span class="sm-own" hidden>${field("Supplier's code", "code", "text", { maxlength: 48 })}</span><span class="sm-own" hidden>${field("Name", "name", "text", { maxlength: 180 })}</span><span class="sm-own" hidden>${unitField("Unit", "uom", "ea")}</span>
        ${field("Supplier's unit price", "unit_price", "number", { min: 0, step: "0.0001", value: 0 })}
        <label class="field"><span>Currency</span><select class="form-control" name="currency"><option>INR</option><option>USD</option><option>EUR</option><option>GBP</option><option>JPY</option></select></label>
        ${field("Supplier's lead time (days)", "lead_time_days", "number", { min: 0, value: 0 })}${field("Supplier's buffer stock (days)", "buffer_days", "number", { min: 0, value: 0 })}
        <button class="button button-primary" type="submit">Add to supplier's master</button>
      </form>
      <form id="smUploadForm" class="panel-body inline-upload">
        <label class="field"><span>Upload this supplier's material master (Code, Name, Unit, Unit price / Unit cost, Currency, Lead time, Buffer days; your own headings are fine). Codes in your T1/T2/T3 are linked automatically; other codes are the supplier's own (name needed).</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label>
        <button class="button" type="submit">Upload list</button>
      </form>
      <div id="smTable"></div>
    </section>`;
  } else if (tab === "Supplier BOM") {
    body = `<section class="panel table-panel">${panelHeader("Supplier BOM: cost breakdown", "The supplier's own BOM of a part you buy, down to L3. Line cost = quantity × unit price; a line with lines below costs their sum. Unit price comes from the line's sources (sub-suppliers / alternate codes, SOB-weighted), else the price entered on the line, else the material list. Cost view only: planning is not affected.")}
      <div class="table-toolbar"><div class="toolbar-group">${supplierChooser("sbSupplier", "Supplier")}<label class="field"><span>Part bought from this supplier</span><select class="form-control" id="sbProduct"><option value="">Choose supplier first</option></select></label></div>
        <div class="toolbar-group" id="sbActions"></div></div>
      <div id="sbSummary"></div>
      <div id="sbTable"></div>
    </section>`;
  } else {
    body = `<section id="transferTable" class="panel table-panel">${panelHeader("Transfer schedule")}<div class="tree-loading">Loading…</div></section>`;
  }
  content.innerHTML = `${heading("Supplier control", "Manage supplier tier, share of business, lead time through T2/Tn, MOQ, buffer, dynamic current price and stock availability.", "")}${tabs}${body}`;
  bindSupplierSearch(content);

  if (tab === "Supplier/material relationship" && state.coverageLinkId) {
    try { await loadCoverage(); }
    catch (error) { state.coverageLinkId = null; toast(error.message, true); return renderSupplierControl(); }
    const { t1 } = state.coverage;
    $("#coverageTemplate").href = `/api/planning/coverage-workbook/template.xlsx?supplier_id=${t1.supplier_id}`;
    $("#coverageUploadForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      if (!window.confirm(`Replace the BOM and suppliers of every ${t1.supplier_name} part in the file?`)) return;
      try {
        const result = await uploadExcel(`/api/planning/coverage-workbook/upload?supplier_id=${t1.supplier_id}`, new FormData(form).get("workbook"));
        toast(`${result.parts} part${result.parts === 1 ? "" : "s"}: ${result.lines} material line${result.lines === 1 ? "" : "s"}, ${result.suppliers} supplier row${result.suppliers === 1 ? "" : "s"} saved`);
        form.reset();
        state.suppliers = await api("/api/suppliers");
        loadCoverage();
      } catch (error) { toast(error.message, true); }
    });
    $("#covCrumbs").innerHTML = `<a href="#" data-open-supplier-rel="${t1.supplier_id}">${esc(t1.supplier_code)} ${esc(t1.supplier_name)}</a> › <strong>${esc(t1.material_code)} ${esc(t1.material_name)}</strong>`;
    state.covRequirement = state.covRequirement || {};
    const box = $("#covRequirement");
    if (state.covRequirement[t1.id] !== undefined) box.value = state.covRequirement[t1.id];
    renderCoverageTable();
    box.addEventListener("input", () => { state.covRequirement[t1.id] = box.value; renderCoverageTable(); });
  } else if (tab === "Supplier/material relationship") {
    const chosen = state.suppliers.find((supplier) => supplier.id === state.relSupplierId) || null;
    const entries = bindEntryTable("relEntryTable", t1EntryRow);
    const prefill = () => { if (chosen) entries.rows().forEach((row) => { if (!cellValue(row, '[data-row-supplier="supplier"]')) setRowPick(row, "supplier", "supplier", chosen); }); };
    prefill();
    $("#relAddRow").addEventListener("click", () => { entries.add(); prefill(); });
    $("#relSaveAll").addEventListener("click", async () => {
      const rows = [];
      for (const [index, row] of entries.rows().entries()) {
        const typed = (key) => cellValue(row, `[data-pick-key="${key}"][data-pick-part="code"]`).trim();
        const supplier = Number(cellValue(row, '[data-row-supplier="supplier"]'));
        const material = Number(cellValue(row, '[data-row-material="material"]'));
        const value = (key) => cellValue(row, `[data-f="${key}"]`);
        if (!typed("material") && !material && !value("sob_percent")) continue;  // empty row (supplier may be prefilled)
        if (!supplier) return toast(typed("supplier") ? `Row ${index + 1}: supplier ${typed("supplier")} is not in the Supplier register; pick it from the hint list.` : `Row ${index + 1}: type the supplier.`, true);
        if (!material) return toast(typed("material") ? `Row ${index + 1}: part ${typed("material")} is not a purchased T1 code; pick it from the hint list.` : `Row ${index + 1}: type the T1 part code.`, true);
        const out = { supplier_id: supplier, material_id: material, tier: "T1", sob_percent: value("sob_percent") === "" ? 100 : Number(value("sob_percent")),
          lead_time_days: Number(value("lead_time_days") || 0), moq: Number(value("moq") || 0), buffer_days: Number(value("buffer_days") || 0), currency: value("currency") };
        if (row.querySelector('[data-f="unit_price"]')) out.unit_price = Number(value("unit_price") || 0);
        rows.push(out);
      }
      if (!rows.length) return toast("Fill in at least one row.", true);
      try {
        const result = await api("/api/planning/supplier-parts/rows", { method: "POST", body: JSON.stringify({ rows }) });
        toast(`${result.rows} supplier/material relationship${result.rows === 1 ? "" : "s"} saved`);
        entries.reset();
        prefill();
        state.suppliers = await api("/api/suppliers");
        loadRelList();
      } catch (error) { toast(error.message, true); }
    });
    $("#relSearch").addEventListener("input", (event) => {
      clearTimeout(state.searchTimer);
      state.searchTimer = setTimeout(() => { state.relSearch = event.target.value; loadRelList(); }, 250);
    });
    loadRelList();
  }
  if (tab === "Supplier/material relationship") {
    $("#supplierPartUploadForm")?.addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      try {
        const result = await uploadExcel("/api/planning/supplier-parts/upload", new FormData(form).get("workbook"));
        toast(`${result.rows} row(s): ${result.added} added, ${result.updated} updated, ${result.suppliers_created} new supplier(s)`);
        form.reset();
        state.suppliers = await api("/api/suppliers");
        loadRelList();
      } catch (error) { toast(error.message, true); }
    });
  }
  if (tab === "Upload supplier stock") {
    const workspaceOf = () => content.querySelector('select[name="workspace"]')?.value || state.workspace;
    const loadStockList = async () => {
      try {
        const result = await api(`/api/planning/supplier-stock/${$("#stockMonth").value}?workspace=${workspaceOf()}`);
        $("#stockListTable").innerHTML = table(["Supplier code", "Supplier name", "Material code", "Material name", "Tier", "Stock quantity", "Available quantity", "Unit", "Arrival month"],
          result.items.map((r) => `<tr><td>${esc(r.supplier_code)}</td><td>${esc(r.supplier_name)}</td><td>${esc(r.material_code)}</td><td>${esc(r.material_name)}</td><td>${esc(r.tier)}</td><td>${number(r.quantity)}</td><td>${number(r.available_quantity)}</td><td>${esc(r.uom)}</td><td>${r.arrival_month ? esc(monthLabel(r.arrival_month)) : "—"}</td></tr>`).join(""),
          `No supplier stock entered for ${monthLabel($("#stockMonth").value)}.`);
      } catch (error) { toast(error.message, true); }
    };
    const entries = bindEntryTable("stockEntryTable", stockEntryRow, async (row, target) => {
      if (!target.matches('[data-row-supplier="supplier"]')) return;
      const linkSelect = row.querySelector("[data-row-link]");
      if (!target.value) { linkSelect.innerHTML = '<option value="">Choose supplier first</option>'; return; }
      try {
        const result = await api(`/api/planning/supplier-parts?${new URLSearchParams({ supplier_id: target.value, limit: "1000" })}`);
        linkSelect.innerHTML = result.items.length
          ? `<option value="">Select material</option>${result.items.map((link) => `<option value="${link.id}">${esc(link.material_code)} · ${esc(link.material_name)} · ${esc(link.tier)}${link.parent_supplier_name ? ` via ${esc(link.parent_supplier_name)}` : ""}</option>`).join("")}`
          : '<option value="">No relationships for this supplier</option>';
        if (result.items.length === 1) linkSelect.value = String(result.items[0].id);
      } catch (error) { toast(error.message, true); }
    });
    $("#stockAddRow").addEventListener("click", () => entries.add());
    const loadTree = async () => {
      const query = new URLSearchParams({ report_month: $("#stockMonth").value, workspace: workspaceOf() });
      $("#supplyTreeExcel").href = `/api/planning/supply-tree.xlsx?${query}`;
      $("#supplyTreePdf").href = `/api/planning/supply-tree.pdf?${query}`;
      $("#stockTemplate").href = `/api/planning/supply-tree.xlsx?${query}&template=true`;
      try {
        state.supplyTree = (await api(`/api/planning/supply-tree?${new URLSearchParams({ report_month: $("#stockMonth").value, workspace: workspaceOf() })}`)).items;
        renderSupplyTree();
      } catch (error) { toast(error.message, true); }
    };
    loadTree();
    $("#stockMonth").addEventListener("change", loadTree);
    content.querySelector('select[name="workspace"]')?.addEventListener("change", loadTree);
    $("#supplyTreeOpen").addEventListener("click", () => renderSupplyTree(true));
    $("#supplyTreeClose").addEventListener("click", () => renderSupplyTree(false));
    $("#supplyTreeSearch").addEventListener("input", () => { clearTimeout(state.searchTimer); state.searchTimer = setTimeout(() => renderSupplyTree(!!$("#supplyTreeSearch").value.trim()), 200); });
    $("#stockMonth").addEventListener("change", loadStockList);
    content.querySelector('select[name="workspace"]')?.addEventListener("change", loadStockList);
    $("#stockSaveAll").addEventListener("click", async () => {
      const rows = [];
      for (const [index, row] of entries.rows().entries()) {
        const link = Number(cellValue(row, "[data-row-link]"));
        const quantity = cellValue(row, '[data-f="quantity"]');
        if (!link && !quantity) continue;
        if (!link) return toast(`Row ${index + 1}: choose the supplier and its material.`, true);
        if (quantity === "") return toast(`Row ${index + 1}: enter the stock quantity.`, true);
        const available = cellValue(row, '[data-f="available_quantity"]');
        rows.push({ supplier_part_id: link, quantity: Number(quantity), available_quantity: available === "" ? null : Number(available), arrival_month: cellValue(row, '[data-f="arrival_month"]') || null });
      }
      if (!rows.length) return toast("Fill in at least one row.", true);
      state.workspace = workspaceOf();
      try {
        const result = await api(`/api/planning/supplier-stock/${$("#stockMonth").value}?workspace=${state.workspace}`, { method: "POST", body: JSON.stringify({ rows }) });
        toast(`${result.rows} supplier stock row${result.rows === 1 ? "" : "s"} saved for ${monthLabel(result.report_month)}`);
        entries.reset();
        loadStockList(); loadTree();
      } catch (error) { toast(error.message, true); }
    });
    $("#supplierStockForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      state.workspace = workspaceOf();
      try {
        const result = await uploadExcel(`/api/planning/supplier-stock/${$("#stockMonth").value}/upload?workspace=${state.workspace}`, new FormData(form).get("workbook"));
        toast(`${result.rows} supplier stock row(s) uploaded`);
        form.reset();
        loadStockList(); loadTree();
      } catch (error) { toast(error.message, true); }
    });
    loadStockList();
  }
  if (tab === "Supplier register") {
    $("#supplierUploadForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const form = event.currentTarget;
      try {
        const result = await uploadExcel("/api/suppliers/upload?tier=T1", new FormData(form).get("workbook"));
        toast(`${result.rows} row(s): ${result.added} supplier(s) added, ${result.updated} updated`);
        form.reset();
        await reloadCategories();
        renderSupplierControl();
      } catch (error) { toast(error.message, true); }
    });
    setupTableFilters("#supplierRegisterTable", "#supplierFilters", $("#supplierSearch"), [{ title: "Tiers", index: 2 }, { title: "Category", index: 3 }, { title: "Approval", index: 8 }]);
    $("#supplierSearch").addEventListener("input", applyTableSearch);
  }
  if (tab === "Supplier-to-supplier material transfer") {
    const entries = bindEntryTable("transferEntryTable", transferEntryRow);
    $("#transferAddRow").addEventListener("click", () => entries.add());
    $("#transferSaveAll").addEventListener("click", async () => {
      const workspace = content.querySelector('select[name="workspace"]')?.value || state.workspace;
      const filled = entries.rows().filter((row) => cellValue(row, "[data-row-material]") || cellValue(row, '[data-f="quantity"]'));
      if (!filled.length) return toast("Fill in at least one row.", true);
      for (const [index, row] of entries.rows().entries()) {
        if (!filled.includes(row)) continue;
        const values = { material_id: Number(cellValue(row, "[data-row-material]")), from_supplier_id: Number(cellValue(row, '[data-row-supplier="from"]')),
          to_supplier_id: Number(cellValue(row, '[data-row-supplier="to"]')), quantity: Number(cellValue(row, '[data-f="quantity"]')),
          transfer_month: cellValue(row, '[data-f="transfer_month"]'), note: cellValue(row, '[data-f="note"]'), workspace };
        if (!values.material_id || !values.from_supplier_id || !values.to_supplier_id || !values.quantity || !values.transfer_month) return toast(`Row ${index + 1}: fill in material, both suppliers, quantity and month.`, true);
      }
      state.workspace = workspace;
      let saved = 0;
      for (const row of filled) {
        const values = { material_id: Number(cellValue(row, "[data-row-material]")), from_supplier_id: Number(cellValue(row, '[data-row-supplier="from"]')),
          to_supplier_id: Number(cellValue(row, '[data-row-supplier="to"]')), quantity: Number(cellValue(row, '[data-f="quantity"]')),
          transfer_month: cellValue(row, '[data-f="transfer_month"]'), note: cellValue(row, '[data-f="note"]'), workspace };
        try {
          await api("/api/planning/transfers", { method: "POST", body: JSON.stringify(values) });
          saved += 1;
          row.remove();  // saved rows leave the table; a failing row stays to be fixed
        } catch (error) {
          toast(`${saved} saved; row with error: ${error.message}`, true);
          return;
        }
      }
      toast(`${saved} supplier-to-supplier transfer${saved === 1 ? "" : "s"} saved`);
      state.supplierTab = "Transfer schedule";
      renderSupplierControl();
    });
  }
  if (tab === "Supplier materials") initSupplierMaterialsTab();
  if (tab === "Supplier BOM") initSupplierBomTab();
  if (tab === "Transfer schedule") {
    const transfers = await api(`/api/planning/transfers?workspace=${state.workspace}`);
    const transferRows = transfers.map((transfer) => `<tr><td>${esc(transfer.material_code)} · ${esc(transfer.material_name)}</td><td>${esc(transfer.from_supplier)}</td><td>${esc(transfer.to_supplier)}</td><td>${number(transfer.quantity)}</td><td>${monthLabel(transfer.transfer_month)}</td><td>${esc(transfer.status)}</td></tr>`).join("");
    $("#transferTable").innerHTML = `${panelHeader("Transfer schedule")}<div class="table-toolbar"><label class="search-field"><input class="form-control" id="transferSearch" type="search" placeholder="Search material or supplier" /></label><div id="transferFilters"></div>${tableExportLinks("transfers", { workspace: state.workspace })}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div><div id="transferRowsTable">${table(["Material", "From supplier", "To supplier", "Quantity", "Transfer month", "Status"], transferRows, "No transfers recorded.")}</div>`;
    setupTableFilters("#transferRowsTable", "#transferFilters", $("#transferSearch"), [{ title: "From supplier", index: 1 }, { title: "To supplier", index: 2 }, { title: "Status", index: 5 }]);
    $("#transferSearch").addEventListener("input", applyTableSearch);
  }
}

function setupTableFilters(tableRootSelector, filterRootSelector, searchInput, specifications) {
  const tableRoot = $(tableRootSelector);
  const filterRoot = $(filterRootSelector);
  if (!tableRoot || !filterRoot) return;
  const tableNode = tableRoot.querySelector("table");
  if (!tableNode) return;
  const rows = [...tableNode.querySelectorAll("tbody tr")].filter((row) => !row.querySelector(".table-empty"));
  filterRoot.replaceChildren();
  if (searchInput) searchInput.dataset.tableFilterRoot = tableRootSelector;
  tableRoot.applyTableFilters = apply;
  function apply() {
    const selected = new Map();
    tableNode.querySelectorAll("thead [data-column-value]:checked").forEach((input) => {
      const column = input.closest("th").cellIndex;
      const key = `column-${column}`;
      if (!selected.has(key)) selected.set(key, { values: [], column });
      selected.get(key).values.push(input.value);
    });
    const search = searchInput?.value.trim().toLowerCase() || "";
    rows.forEach((row) => {
      const matchesSearch = !search || row.textContent.toLowerCase().includes(search);
      const matchesFilters = [...selected.values()].every((selection) => selection.values.includes(row.cells[selection.column]?.textContent.trim() || ""));
      row.hidden = !(matchesSearch && matchesFilters);
    });
  }

}

function runTableFilters(tableNode) {
  if (!tableNode) return;
  let filterRoot = tableNode.parentElement;
  while (filterRoot && !filterRoot.applyTableFilters) filterRoot = filterRoot.parentElement;
  if (filterRoot) {
    filterRoot.applyTableFilters();
    return;
  }
  let tableContainer = tableNode.parentElement;
  while (tableContainer && !tableContainer.id) tableContainer = tableContainer.parentElement;
  const searchInput = tableContainer?.id
    ? document.querySelector(`[data-table-filter-root="#${tableContainer.id}"]`)
    : null;
  const needle = searchInput?.value.trim().toLowerCase() || "";
  const selected = new Map();
  tableNode.querySelectorAll("thead [data-column-value]:checked").forEach((input) => {
    const column = input.closest("th").cellIndex;
    if (!selected.has(column)) selected.set(column, []);
    selected.get(column).push(input.value);
  });
  tableNode.querySelectorAll("tbody tr").forEach((row) => {
    if (row.querySelector(".table-empty")) return;
    const matchesSearch = !needle || row.textContent.toLowerCase().includes(needle);
    row.hidden = !matchesSearch || [...selected].some(([column, values]) => !values.includes(row.cells[column]?.textContent.trim() || ""));
  });
}

let activeColumnResize = null;
document.addEventListener("pointerdown", (event) => {
  const handle = event.target.closest(".column-resizer");
  if (!handle) return;
  const header = handle.closest("th");
  const tableNode = header?.closest("table");
  if (!header || !tableNode) return;
  event.preventDefault();
  event.stopPropagation();
  if (header.parentElement !== headerRow(tableNode)) return;  // group headings are not resized
  const widths = [...header.parentElement.cells].map((cell) => cell.getBoundingClientRect().width);
  const tableWidth = widths.reduce((total, width) => total + width, 0);
  tableNode.style.tableLayout = "fixed";
  tableNode.style.width = `${tableWidth}px`;
  [...tableNode.rows].filter((row) => !row.classList.contains("group-head")).forEach((row) => [...row.cells].forEach((cell, index) => {
    cell.style.width = `${widths[index]}px`;
    cell.style.minWidth = `${widths[index]}px`;
  }));
  widths.forEach((width, index) => setColWidth(tableNode, index, width));
  activeColumnResize = { tableNode, column: header.cellIndex, startX: event.clientX, startWidth: widths[header.cellIndex] };
});

document.addEventListener("pointermove", (event) => {
  if (!activeColumnResize) return;
  const { tableNode, column, startX, startWidth } = activeColumnResize;
  const width = Math.max(70, startWidth + event.clientX - startX);
  setColWidth(tableNode, column, width);
  [...tableNode.rows].filter((row) => !row.classList.contains("group-head")).forEach((row) => {
    const cell = row.cells[column];
    if (cell) {
      cell.style.width = `${width}px`;
      cell.style.minWidth = `${width}px`;
    }
  });
});

document.addEventListener("pointerup", () => {
  if (activeColumnResize) saveColumnWidths(activeColumnResize.tableNode);
  activeColumnResize = null;
});

// Master-data tabs remember column widths (per tab and table) in this browser, so a
// redraw (tab change, filter, reload) keeps them. Double-click a column edge to reset.
// Tables with a group heading row (T1 / T2 / T3 Supplier) get <col> widths: a fixed-layout table
// otherwise takes its column widths from the first (group) row.
function setColWidth(tableNode, index, width) {
  if (!tableNode.querySelector("thead tr.group-head")) return;
  let group = tableNode.querySelector(":scope > colgroup");
  if (!group) { group = document.createElement("colgroup"); tableNode.prepend(group); }
  while (group.children.length < headerRow(tableNode).cells.length) group.append(document.createElement("col"));
  group.children[index].style.width = `${width}px`;
}
// the heading row with one cell per column (the last one; a group row may sit above it)
function headerRow(tableNode) {
  const rows = tableNode.tHead?.rows;
  return rows ? rows[rows.length - 1] : null;
}
function columnWidthKey(tableNode) {
  const view = content.dataset.view;
  if (view !== "master" && view !== "suppliers") return "";
  const owner = tableNode.closest("[id]")?.id || "";
  const tab = view === "master" ? state.bomTab : state.supplierTab;
  return `dtcColumnWidths:${tab}:${owner}:${headerRow(tableNode)?.cells.length || 0}`;
}

function saveColumnWidths(tableNode) {
  const key = columnWidthKey(tableNode);
  if (!key) return;
  const widths = [...headerRow(tableNode).cells].map((cell) => Math.round(cell.getBoundingClientRect().width));
  try { localStorage.setItem(key, JSON.stringify(widths)); } catch { /* storage unavailable: widths last until redraw */ }
}

function applySavedColumnWidths(tableNode) {
  const key = columnWidthKey(tableNode);
  if (!key || !tableNode.tHead) return;
  let widths = null;
  try { widths = JSON.parse(localStorage.getItem(key) || "null"); } catch { widths = null; }
  const headers = [...headerRow(tableNode).cells];
  if (!Array.isArray(widths) || widths.length !== headers.length) return;
  tableNode.style.tableLayout = "fixed";
  tableNode.style.width = `${widths.reduce((total, width) => total + width, 0)}px`;
  headers.forEach((cell, index) => { cell.style.width = `${widths[index]}px`; cell.style.minWidth = `${widths[index]}px`; setColWidth(tableNode, index, widths[index]); });
}

function resetColumnWidths(tableNode) {
  const key = columnWidthKey(tableNode);
  try { if (key) localStorage.removeItem(key); } catch { /* ignore */ }
  tableNode.style.tableLayout = "";
  tableNode.style.width = "";
  [...tableNode.rows].forEach((row) => [...row.cells].forEach((cell) => { cell.style.width = ""; cell.style.minWidth = ""; }));
  tableNode.querySelector(":scope > colgroup")?.remove();
}

document.addEventListener("dblclick", (event) => {
  const handle = event.target.closest(".column-resizer");
  const tableNode = handle?.closest("table");
  if (tableNode) resetColumnWidths(tableNode);
});

new MutationObserver(() => {
  content.querySelectorAll("table:not([data-widths-checked])").forEach((tableNode) => {
    tableNode.dataset.widthsChecked = "true";
    applySavedColumnWidths(tableNode);
  });
}).observe(content, { childList: true, subtree: true });

function reportRows(rows = state.report) {
  const checked = {};
  document.querySelectorAll("[data-filter-key]:checked").forEach((input) => {
    checked[input.dataset.filterKey] ||= [];
    checked[input.dataset.filterKey].push(input.value);
  });
  const visible = rows.filter((row) => Object.entries(checked).every(([key, values]) => values.includes(String(row[key] ?? ""))));
  return visible.map((row) => `<tr>
    <td>${esc(monthLabel(row.month))}</td><td><strong>${esc(row.material_code)}</strong><div class="secondary-cell">${esc(row.material_name)}</div></td>
    <td>${esc(row.supplier_name)}</td><td>${esc(row.parent_supplier_name || "—")}</td><td>${esc(row.tier)}</td><td>${number(row.sob_percent, 1)}%</td>
    <td>${number(row.plan_demand)} ${esc(row.uom)}</td><td>${row.actual_demand == null ? "—" : `${number(row.actual_demand)} ${esc(row.uom)}`}</td>
    <td>${number(row.own_stock)} ${esc(row.uom)}</td><td>${number(row.own_stock_allocated)} ${esc(row.uom)}</td>
    <td>${number(row.supplier_stock)} ${esc(row.uom)}</td><td>${number(row.projected_balance)} ${esc(row.uom)}</td>
    <td>${number(row.coverage_days, 0)} days</td>
    <td class="${row.shortage_qty > 0 ? "shortage-cell" : ""}">${number(row.shortage_qty)} ${esc(row.uom)}${row.shortage_value === undefined ? "" : `<small>${money(row.shortage_value, row.currency)}</small>`}</td>
    <td>${number(row.recommended_po_qty)} ${esc(row.uom)}</td>
    <td>${number(row.excess_qty)} ${esc(row.uom)}${row.excess_value === undefined ? "" : `<small>${money(row.excess_value, row.currency)}</small>`}</td>
    <td>${number(row.lead_time_days, 0)} days</td><td><span class="criticality criticality-${row.criticality.toLowerCase().replace(" ", "-")}">${esc(row.criticality)}</span></td>
    <td>${esc(row.plan_revision)}</td>
    <td>${row.sob_change_from ? `from ${esc(monthLabel(row.sob_change_from))}<small class="${row.sob_demand_change < 0 ? "sob-down" : "sob-up"}">${row.sob_demand_change > 0 ? "+" : ""}${number(row.sob_demand_change)} ${esc(row.uom)} demand</small>` : "—"}</td>
  </tr>`).join("");
}

async function renderReports() {
  const end = monthNow();
  const start = monthOffset(end, -11);
  content.innerHTML = `${heading("Shortage & excess analysis", "One-click, month-by-month shortage report by material and supplier. Planning demand and actual demand are shown separately.", "")}
    <section class="panel">${panelHeader("Report controls", "Select any historical period. Supplier buffer norms and long lead times determine criticality and target coverage.")}
      <form id="reportForm" class="table-toolbar report-controls">
        ${field("Start month", "start_month", "month", { required: true, value: start })}
        ${field("End month", "end_month", "month", { required: true, value: end })}
        <label class="field"><span>Workspace</span>${workspaceSelect()}</label>
        <button class="button button-primary" type="submit">Generate report</button>
        <a id="reportExcel" class="button" href="#">Download Excel</a><a id="reportPdf" class="button" href="#">Download PDF</a>
      </form>
    </section>
    <section class="panel table-panel">${panelHeader("Supplier shortage & excess table", "Critical shortages are listed first; prices and values are visible only to users with explicit access.")}
      <div class="table-toolbar"><div class="toolbar-group" id="reportFilters"></div><button class="button button-small" data-action="clear-report-filters">Clear all filters</button></div>
      <div id="reportTable">${table(["Month", "Material", "Supplier", "Upstream parent", "Tier", "SOB", "Plan demand", "Actual demand", "Own stock", "Own stock allocated", "Supplier stock", "Projected balance", "Coverage", "Shortage", "Recommended PO", "Excess", "Lead time", "Criticality", "Plan revision", "SOB change"], "", "Generate a report to see shortages and excess.")}</div>
    </section>`;
  $("#reportForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    await generateReport();
  });
  if (state.report.length) renderReportTable();
  $("#reportForm").requestSubmit();
}

async function generateReport() {
  const data = new FormData($("#reportForm"));
  state.workspace = data.get("workspace");
  state.reportQuery = { start_month: data.get("start_month"), end_month: data.get("end_month"), workspace: state.workspace };
  try {
    await fetchReportPage(0);
    const suffix = `start_month=${data.get("start_month")}&end_month=${data.get("end_month")}&workspace=${state.workspace}`;
    $("#reportExcel").href = `/api/planning/report.xlsx?${suffix}`;
    $("#reportPdf").href = `/api/planning/report.pdf?${suffix}`;
    renderReportTable();
    toast(`${state.reportTotal.toLocaleString()} report rows generated`);
  } catch (error) { toast(error.message, true); }
}

async function fetchReportPage(offset) {
  const query = new URLSearchParams({ ...state.reportQuery, offset: String(offset), limit: "1000" });
  const result = await api(`/api/planning/report?${query}`);
  state.report = result.items;
  state.reportTotal = result.total;
  state.reportOffset = result.offset;
  renderReportTable();
}

function renderReportTable() {
  const filters = $("#reportFilters");
  if (!filters) return;
  filters.innerHTML = filterSelector("Material", "material_code", state.report.map((row) => row.material_code))
    + filterSelector("Supplier", "supplier_name", state.report.map((row) => row.supplier_name))
    + filterSelector("Tier", "tier", state.report.map((row) => row.tier))
    + filterSelector("Criticality", "criticality", state.report.map((row) => row.criticality));
  const headers = ["Month", "Material", "Supplier", "Upstream parent", "Tier", "SOB", "Plan demand", "Actual demand", "Own stock", "Own stock allocated", "Supplier stock", "Projected balance", "Coverage", "Shortage", "Recommended PO", "Excess", "Lead time", "Criticality", "Plan revision", "SOB change"];
  $("#reportTable").innerHTML = `${table(headers, reportRows(), "No rows match the active filters.")}<div class="table-toolbar pagination-toolbar"><span>${state.reportTotal ? `${(state.reportOffset + 1).toLocaleString()}–${Math.min(state.reportOffset + state.report.length, state.reportTotal).toLocaleString()} of ${state.reportTotal.toLocaleString()}` : "No report rows"}</span><div><button class="button button-small" data-report-page="${Math.max(0, state.reportOffset - 1000)}" ${state.reportOffset === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" data-report-page="${state.reportOffset + 1000}" ${state.reportOffset + state.report.length >= state.reportTotal ? "disabled" : ""}>Next</button></div></div>`;
  filters.querySelectorAll("[data-filter-key]").forEach((input) => input.addEventListener("change", () => {
    $("#reportTable").innerHTML = `${table(headers, reportRows(), "No rows match the active filters.")}<div class="table-toolbar pagination-toolbar"><span>${state.reportTotal ? `${(state.reportOffset + 1).toLocaleString()}–${Math.min(state.reportOffset + state.report.length, state.reportTotal).toLocaleString()} of ${state.reportTotal.toLocaleString()}` : "No report rows"}</span><div><button class="button button-small" data-report-page="${Math.max(0, state.reportOffset - 1000)}" ${state.reportOffset === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" data-report-page="${state.reportOffset + 1000}" ${state.reportOffset + state.report.length >= state.reportTotal ? "disabled" : ""}>Next</button></div></div>`;
  }));
  filters.querySelectorAll(".multi-search").forEach((input) => input.addEventListener("input", () => {
    const needle = input.value.toLowerCase();
    input.parentElement.querySelectorAll(".multi-options label").forEach((label) => { label.hidden = !label.textContent.toLowerCase().includes(needle); });
  }));
}

async function renderPurchaseOrders() {
  const end = monthNow();
  const start = monthOffset(end, -11);
  const orders = await api(`/api/planning/purchase-orders?start_month=${start}&end_month=${end}&workspace=${state.workspace}`);
  const links = (await api("/api/planning/supplier-parts?offset=0&limit=500")).items;
  state.supplierParts = links;
  content.innerHTML = `${heading("PO schedule", "Issue direct-price purchase orders, preserve NCNR commitments, and request delivery preponement or postponement.", `<a class="button" href="/api/planning/templates/supplier-stock.xlsx">Stock template</a>`)}
    <section class="panel">${panelHeader("Issue a planning purchase order", "The current supplier/material price is copied to the PO and will not change when the master price is updated.")}
      <form id="poForm" class="panel-body form-grid">
        <label class="field"><span>Search supplier/material</span><input class="form-control" id="poPartSearch" type="search" placeholder="Type part, supplier, or tier" /></label>
        <label class="field"><span>Supplier and material</span><select class="form-control" name="supplier_part_id" id="poSupplierPartSelect" required>${links.filter((link) => link.material_id).map((link) => `<option value="${link.id}">${esc(link.material_code)} · ${esc(link.supplier_name)} · ${esc(link.tier)}${link.unit_price === undefined ? "" : ` · ${money(link.unit_price, link.currency)}`}</option>`).join("")}</select></label>
        ${field("Quantity", "quantity_ordered", "number", { required: true, min: 0.001, step: "any" })}
        ${field("Expected arrival month", "expected_month", "month", { required: true, value: monthOffset(monthNow(), 1) })}
        <label class="field"><span>Workspace</span>${workspaceSelect()}</label>
        <label class="check-option full"><input name="ncnr" type="checkbox" /> Non-cancellable / non-returnable (NCNR)</label>
        <button class="button button-primary" type="submit">Issue PO</button>
      </form></section>
    <section class="panel table-panel">${panelHeader("Purchase order register", "Change an expected date to pre-pone or postpone delivery, with a reason retained on the record.")}
      <div class="table-toolbar"><div class="toolbar-group">${field("Start month", "po_start", "month", { value: start })}${field("End month", "po_end", "month", { value: end })}<label class="search-field"><input class="form-control" id="poSearch" type="search" placeholder="Search PO or supplier" /></label><div id="poFilters"></div></div><div class="toolbar-group">${tableExportLinks("purchase-orders", { start_month: start, end_month: end, workspace: state.workspace })}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div></div>
      <div id="poTable">${table(["PO", "Material", "Supplier / tier", "Quantity", "Expected arrival", "Value", "Commitment", "Status", "Actions"], orders.map((po) => `<tr><td>${esc(po.po_number)}</td><td>${esc(po.material_code)} · ${esc(po.material_name)}</td><td>${esc(po.supplier_name)} · ${esc(po.tier)}</td><td>${number(po.quantity_ordered)}</td><td>${monthLabel(po.expected_month)}</td><td>${po.unit_price === undefined ? "Restricted" : money(po.total, po.currency)}</td><td>${po.ncnr ? "NCNR" : "Cancelable"}</td><td>${esc(po.status)}</td><td><button class="text-button" data-reschedule-po="${po.id}" data-current-month="${esc(po.expected_month)}">Advance / postpone</button>${po.status !== "Received" ? `<button class="text-button" data-receive-po="${po.id}" data-remaining="${po.quantity_ordered - po.quantity_received}">Record receipt</button>` : ""}</td></tr>`).join(""), "No planning POs in this period.")}</div></section>`;
  setupTableFilters("#poTable", "#poFilters", $("#poSearch"), [{ title: "Supplier / tier", index: 2 }, { title: "Status", index: 7 }, { title: "Commitment", index: 6 }]);
  $("#poSearch").addEventListener("input", applyTableSearch);
  $("#poPartSearch").addEventListener("input", (event) => {
    clearTimeout(state.searchTimer);
    state.searchTimer = setTimeout(async () => {
      try {
        const query = new URLSearchParams({ offset: "0", limit: "100", search: event.target.value });
        const result = await api(`/api/planning/supplier-parts?${query}`);
        state.supplierParts = result.items;
        $("#poSupplierPartSelect").innerHTML = result.items.map((link) => `<option value="${link.id}">${esc(link.material_code)} · ${esc(link.supplier_name)} · ${esc(link.tier)}${link.unit_price === undefined ? "" : ` · ${money(link.unit_price, link.currency)}`}</option>`).join("");
      } catch (error) { toast(error.message, true); }
    }, 250);
  });
  $("#poForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const values = Object.fromEntries(new FormData(event.currentTarget).entries());
    state.workspace = values.workspace;
    values.supplier_part_id = Number(values.supplier_part_id);
    values.quantity_ordered = Number(values.quantity_ordered);
    values.ncnr = Boolean(event.currentTarget.elements.ncnr.checked);
    try {
      await api("/api/planning/purchase-orders", { method: "POST", body: JSON.stringify(values) });
      toast("Planning PO issued");
      renderPurchaseOrders();
    } catch (error) { toast(error.message, true); }
  });
}

async function renderSupplierBusiness() {
  const end = monthNow();
  const start = monthOffset(end, -11);
  content.innerHTML = `${heading("Supplier business", "PO value issued during the selected period, split by direct/T1/T2/Tn tier.", "")}
    <section class="panel">${panelHeader("Period controls")}
      <form id="businessForm" class="table-toolbar report-controls">
        ${field("Start month", "start_month", "month", { required: true, value: start })}
        ${field("End month", "end_month", "month", { required: true, value: end })}
        <label class="field"><span>Workspace</span>${workspaceSelect()}</label>
        <button class="button button-primary" type="submit">Generate</button>
      </form>
    </section><section id="businessResults" class="panel table-panel">${panelHeader("Issued PO value")}</section>`;
  $("#businessForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const values = new FormData(event.currentTarget);
    state.workspace = values.get("workspace");
    const result = await api(`/api/planning/supplier-business?start_month=${values.get("start_month")}&end_month=${values.get("end_month")}&workspace=${state.workspace}`);
    const rows = result.map((row) => `<tr><td>${esc(row.supplier_code)} · ${esc(row.supplier_name)}</td><td>${esc(row.tier)}</td><td>${esc(row.currency)}</td><td>${number(row.po_count, 0)}</td><td>${row.po_value === undefined ? "Restricted" : money(row.po_value, row.currency)}</td></tr>`).join("");
    const suffix = `start_month=${values.get("start_month")}&end_month=${values.get("end_month")}&workspace=${state.workspace}`;
    $("#businessResults").innerHTML = `${panelHeader("Issued PO value")}<div class="table-toolbar"><div class="toolbar-group"><label class="search-field"><input class="form-control" type="search" placeholder="Search supplier or tier" id="businessSearch" /></label><div id="businessFilters"></div></div><button class="button button-small" data-action="clear-table-filters">Clear all filters</button><a class="button button-small" href="/api/planning/supplier-business.xlsx?${suffix}">Download Excel</a><a class="button button-small" href="/api/planning/supplier-business.pdf?${suffix}">Download PDF</a></div><div id="businessTable">${table(["Supplier", "Tier", "Currency", "PO count", "Issued PO value"], rows, "No POs in this period.")}</div>`;
    setupTableFilters("#businessTable", "#businessFilters", $("#businessSearch"), [{ title: "Supplier", index: 0 }, { title: "Tier", index: 1 }, { title: "Currency", index: 2 }]);
    $("#businessSearch").addEventListener("input", applyTableSearch);
  });
  $("#businessForm").requestSubmit();
}

async function renderUsers() {
  if (state.user.role !== "Admin") throw new Error("Admin access is required.");
  const users = await api("/api/users");
  const rows = users.map((user) => `<tr><td>${esc(user.full_name)}</td><td>${esc(user.username)}</td><td>${esc(user.role)}</td><td>${user.can_view_prices ? "Allowed" : "No price access"}</td><td>${user.enabled ? "Enabled" : "Disabled"}</td><td><button class="text-button" data-user-price="${user.id}" data-enabled="${!user.can_view_prices}">${user.can_view_prices ? "Revoke price view" : "Grant price view"}</button></td></tr>`).join("");
  content.innerHTML = `${heading("User access", "Create multiple Admin accounts and grant explicit per-user access to part prices.")}
    <section class="panel">${panelHeader("Add user")}
      <form id="userForm" class="panel-body form-grid">
        ${field("Username", "username", "text", { required: true })}
        ${field("Full name", "full_name", "text", { required: true })}
        ${field("Temporary password (12+ characters)", "password", "password", { required: true })}
        <label class="field"><span>Role</span><select class="form-control" name="role"><option>User</option><option>HOD</option><option>Sourcing</option><option>Supply Chain</option><option>Admin</option></select></label>
        <label class="check-option full"><input name="can_view_prices" type="checkbox" /> Grant part-price access</label>
        <button class="button button-primary" type="submit">Create user</button>
      </form></section>
    <section class="panel table-panel">${panelHeader("User accounts")}<div class="table-toolbar"><label class="search-field"><input class="form-control" id="userSearch" type="search" placeholder="Search name or username" /></label><div id="userFilters"></div>${tableExportLinks("users")}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div><div id="userTable">${table(["Name", "Username", "Role", "Part prices", "Account", "Action"], rows, "No users.")}</div></section>
    <section class="panel table-panel">${panelHeader("Trial data reset", "This irreversible action deletes trial plans, actuals, stock snapshots, POs and transfers only.")}
      <div class="panel-body"><button class="button button-danger" data-action="reset-trial">Clear all trial data</button></div></section>`;
  setupTableFilters("#userTable", "#userFilters", $("#userSearch"), [{ title: "Role", index: 2 }, { title: "Part prices", index: 3 }, { title: "Account", index: 4 }]);
  $("#userSearch").addEventListener("input", applyTableSearch);
  $("#userForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const values = Object.fromEntries(new FormData(event.currentTarget).entries());
    values.can_view_prices = event.currentTarget.elements.can_view_prices.checked;
    try {
      await api("/api/users", { method: "POST", body: JSON.stringify(values) });
      toast("User account created");
      renderUsers();
    } catch (error) { toast(error.message, true); }
  });
  document.querySelectorAll("[data-user-price]").forEach((button) => button.addEventListener("click", async () => {
    await api(`/api/users/${button.dataset.userPrice}`, { method: "PATCH", body: JSON.stringify({ can_view_prices: button.dataset.enabled === "true" }) });
    toast("Price permission updated");
    renderUsers();
  }));
}

function applyTableSearch(event) {
  const tableRoot = event.target.dataset.tableFilterRoot ? $(event.target.dataset.tableFilterRoot) : event.target.closest(".panel") || content;
  if (tableRoot.applyTableFilters) return tableRoot.applyTableFilters();
  const tableNode = tableRoot.querySelector("table");
  if (tableNode) runTableFilters(tableNode);
}

function openModal(type, record = {}) {
  let title = "";
  let form = "";
  if (type === "product") {
    title = record.id ? "Edit product" : "Add product";
    const models = state.models.map((model) => `<option value="${model.id}" ${(record.model_ids || []).includes(model.id) ? "selected" : ""}>${esc(model.code)} · ${esc(model.name)}</option>`).join("");
    const rowModel = state.models.find((model) => model.id === Number(record.editModelId));
    const modelAssignment = record.id && record.editModelId
      ? `<label class="field"><span>Model code</span><input class="form-control" value="${esc(rowModel?.code || "")}" readonly /></label><label class="field"><span>Model name</span><input class="form-control" value="${esc(rowModel?.name || "")}" readonly /></label>${(record.model_ids || []).map((id) => `<input type="hidden" name="model_ids" value="${id}" />`).join("")}`
      : `<label class="field full"><span>Models from Material Master (select multiple)</span><select class="form-control" name="model_ids" multiple size="7">${models}</select></label>`;
    form = `<form id="modalForm" class="form-grid">${field("Product code", "code", "text", { required: true, value: record.code })}${field("Product name", "name", "text", { required: true, value: record.name })}${modelAssignment}<div class="modal-footer field full"><button class="button button-primary">Save product</button></div></form>`;
  } else if (type === "model") {
    title = record.id ? "Edit model" : "Add model";
    const products = state.products.map((product) => `<option value="${product.id}" ${(record.product_ids || []).includes(product.id) ? "selected" : ""}>${esc(product.code)} · ${esc(product.name)}</option>`).join("");
    form = `<form id="modalForm" class="form-grid">${field("Model code", "code", "text", { required: true, value: record.code })}${field("Model name", "name", "text", { required: true, value: record.name })}<label class="field full"><span>Assigned products (optional; select multiple)</span><select class="form-control" name="product_ids" multiple size="6">${products}</select></label><div class="modal-footer field full"><button class="button button-primary">Save model</button></div></form>`;
  } else if (type === "material") {
    title = record.id ? "Edit material" : "Add material";
    form = `<form id="modalForm" class="form-grid">${field("Material code", "code", "text", { required: true, value: record.code })}${field("Material name", "name", "text", { required: true, value: record.name })}${categoryField(record.category || "Raw material")}${unitField("Unit of measure", "uom", record.uom || "ea")}${field("Opening stock", "stock_on_hand", "number", { required: true, min: 0, step: "any", value: record.stock_on_hand ?? 0 })}${field("Own-location buffer (days)", "buffer_days", "number", { required: true, min: 0, value: record.buffer_days ?? 0 })}${state.user.can_view_prices || state.user.role === "Admin" ? `${field("Current unit cost", "unit_cost", "number", { required: true, min: 0, step: "any", value: record.unit_cost ?? 0 })}<label class="field"><span>Currency</span><select class="form-control" name="currency">${["USD", "INR", "EUR", "GBP"].map((currency) => `<option ${currency === (record.currency || "USD") ? "selected" : ""}>${currency}</option>`).join("")}</select></label>` : '<input type="hidden" name="unit_cost" value="0" /><input type="hidden" name="currency" value="USD" />'}<div class="modal-footer field full"><button class="button button-primary">Save material</button></div></form>`;
  } else if (type === "bom") {
    title = record.root_bom_id ? `Edit ${state.bomTab} line` : `Add ${state.bomTab} line`;
    const materialUnit = record.uom || state.materials.find((m) => String(m.id) === String(record.material_id))?.uom || "ea";
    const modelId = Number(record.model_id || state.selectedModel);
    const model = state.models.find((item) => item.id === modelId);
    const material = state.materials.find((item) => String(item.id) === String(record.material_id));
    const modelField = record.root_bom_id
      ? `<input type="hidden" name="model_id" value="${modelId}" />`
      : `<label class="field full"><span>Model</span><select class="form-control" name="model_id" id="modalModelSelect" required>${state.models.map((item) => `<option value="${item.id}" ${item.id === modelId ? "selected" : ""}>${esc(item.code)} · ${esc(item.name)}</option>`).join("")}</select></label>`;
    const materialFields = record.root_bom_id
      ? `<input type="hidden" name="material_id" value="${record.material_id}" /><label class="field"><span>Part code</span><input class="form-control" value="${esc(material?.code || "")}" readonly /></label><label class="field"><span>Part name</span><input class="form-control" value="${esc(material?.name || "")}" readonly /></label>`
      : `<label class="field full"><span>Search materials</span><input class="form-control" id="modalMaterialSearch" type="search" placeholder="Type a material code or name" /></label><label class="field"><span>Part code</span><select class="form-control" name="material_id" id="modalMaterialSelect" required>${state.materials.map((item) => `<option value="${item.id}" ${String(item.id) === String(record.material_id) ? "selected" : ""}>${esc(item.code)}</option>`).join("")}</select></label><label class="field"><span>Part name</span><input class="form-control" id="modalMaterialName" value="${esc(material?.name || "")}" readonly /></label>`;
    form = `<form id="modalForm" class="form-grid">${modelField}<label class="field"><span>Model code</span><input class="form-control" id="modalModelCode" value="${esc(model?.code || "")}" readonly /></label><label class="field"><span>Model name</span><input class="form-control" id="modalModelName" value="${esc(model?.name || "")}" readonly /></label>${materialFields}${field("Quantity", "quantity", "number", { required: true, min: 0.000001, step: "any", value: record.quantity })}<label class="field"><span>Unit</span><input class="form-control" name="uom" value="${esc(materialUnit)}" readonly required /></label><div class="modal-footer field full"><button class="button button-primary">Save BOM line</button></div></form>`;
  } else if (type === "component") {
    title = record.component_id ? `Edit ${state.bomTab} line` : `Add ${state.bomTab} line`;
    const parents = state.materials.map((m) => `<option value="${m.id}" ${String(m.id) === String(record.parent_item_id) ? "selected" : ""}>${esc(m.code)} · ${esc(m.name)}</option>`).join("");
    const selectedMaterial = state.materials.find((item) => String(item.id) === String(record.material_id));
    const parent = state.materials.find((item) => item.id === Number(record.parent_item_id));
    form = `<form id="modalForm" class="form-grid"><label class="field full"><span>Parent part</span><select class="form-control" name="parent_item_id" required ${record.parent_item_id ? "disabled" : ""}>${parents}</select>${record.parent_item_id ? `<input type="hidden" name="parent_item_id" value="${record.parent_item_id}" />` : ""}</label><label class="field full"><span>Search Material Master by code or name</span><input class="form-control" id="modalMaterialSearch" type="search" placeholder="Type a material code or name" /></label><label class="field"><span>Part code from Material Master</span><select class="form-control" name="child_item_id" id="modalMaterialSelect" required>${state.materials.map((item) => `<option value="${item.id}" data-code="${esc(item.code)}" data-name="${esc(item.name)}" data-uom="${esc(item.uom)}" ${String(item.id) === String(record.material_id) ? "selected" : ""}>${esc(item.code)}</option>`).join("")}</select></label><label class="field"><span>Part name from Material Master</span><select class="form-control" id="modalMaterialNameSelect" required>${state.materials.map((item) => `<option value="${item.id}" data-code="${esc(item.code)}" data-name="${esc(item.name)}" data-uom="${esc(item.uom)}" ${String(item.id) === String(record.material_id) ? "selected" : ""}>${esc(item.name)}</option>`).join("")}</select></label><label class="field"><span>Unit from Material Master</span><input class="form-control" id="modalMaterialUnit" value="${esc(selectedMaterial?.uom || "")}" readonly /></label>${field("Quantity per parent", "quantity", "number", { required: true, min: -1000000000, max: 1000000000, step: "any", value: record.quantity ?? 1 })}<div class="score-explainer field full">${parent ? `Adding parts to ${esc(parent.code)} · ${esc(parent.name)}. Choose each purchased or in-house-made child part from Material Master; repeat at the next level for further operations.` : "Choose the child part from Material Master. Repeat at the next level for further in-house operations."} Use a negative quantity for scrap or a by-product. Quantity cannot be zero.</div><div class="modal-footer field full"><button class="button button-primary">Save BOM line</button></div></form>`;
  } else if (type === "supplier") {
    title = "Add supplier";
    form = `<form id="modalForm" class="form-grid">${field("Supplier code (blank = next free SUP- code)", "code", "text", { maxlength: 24, placeholder: "e.g. ILN-01" })}${field("Supplier name", "name", "text", { required: true })}${supplierTierField(["T1"])}${categoryField("")}${field("Country", "country", "text", { required: true })}${field("Contact person", "contact_name")}${field("Email", "email", "email")}${field("Phone", "phone")}<div class="modal-footer field full"><button class="button button-primary">Add supplier</button></div></form>`;
  }
  // Alternate code SOB on this BOM line: shown here, changed only in Alternate codes → Manage SOB
  const lineAlternates = (type === "bom" || type === "component") && (record.root_bom_id || record.component_id) ? record.alternates || [] : [];
  if (lineAlternates.length) {
    form = form.replace('<div class="modal-footer field full">', `<fieldset class="field full line-alternates"><legend>Alternate code SOB on this BOM line (this month)</legend><div class="field full">${lineAlternates.map((alt) => `<strong>${esc(alt.code)}</strong> ${number(alt.sob, 2)}%`).join(" · ")} · <strong>main</strong> ${number(100 - lineAlternates.reduce((sum, alt) => sum + Number(alt.sob || 0), 0), 2)}%</div><p class="subtle field full">To change the SOB, use Product Master → Alternate codes → Manage SOB: it updates every use of this code at once and keeps the SOB change log.</p><div class="field full" id="lineSobHistory"><span class="subtle">Loading SOB history…</span></div><button type="button" class="button button-small" data-alt-manage="${record.material_id}">Open Manage SOB</button></fieldset><div class="modal-footer field full">`);
  }
  $("#modalTitle").textContent = title;
  $("#modalBody").innerHTML = form;
  modalLayer.hidden = false;
  if (lineAlternates.length) {
    const lineQuery = new URLSearchParams(record.root_bom_id ? { model_bom_id: record.root_bom_id } : { component_id: record.component_id });
    api(`/api/planning/bom-line-alternates?${lineQuery}`).then(({ history }) => {
      const holder = $("#lineSobHistory");
      if (!holder) return;
      holder.innerHTML = history.length ? `<table class="sob-history"><thead><tr><th>From</th>${lineAlternates.map((alt) => `<th>${esc(alt.code)}</th>`).join("")}<th>Main</th></tr></thead><tbody>${history.map((change) => `<tr><td>${change.from_month ? esc(monthLabel(change.from_month)) : "Start"}</td>${lineAlternates.map((alt) => `<td>${number(change.alternates.find((a) => a.alternate_item_id === alt.id)?.sob_percent || 0, 2)}%</td>`).join("")}<td>${number(change.main_sob_percent, 2)}%</td></tr>`).join("")}</tbody></table>` : '<span class="subtle">No SOB set on this line yet (main code 100%).</span>';
    }).catch((error) => { const holder = $("#lineSobHistory"); if (holder) holder.textContent = error.message; });
  }
  bindCategoryField($("#modalBody"));
  if (type === "component") bindMaterialSearch("modalMaterialSearch", "modalMaterialSelect", "modalMaterialNameSelect");
  if (type === "bom" && !record.root_bom_id) bindMaterialSearch("modalMaterialSearch", "modalMaterialSelect", "modalMaterialName");
  if (type === "bom") {
    const materialSelect = $("#modalMaterialSelect");
    if (materialSelect) materialSelect.addEventListener("change", (event) => {
        const material = state.materials.find((item) => String(item.id) === event.target.value);
        if (material) {
          $("#modalForm").elements.uom.value = material.uom;
          $("#modalMaterialName").value = material.name;
        }
      });
    const modelSelect = $("#modalModelSelect");
    if (modelSelect) modelSelect.addEventListener("change", (event) => {
      const model = state.models.find((item) => String(item.id) === event.target.value);
      if (model) {
        $("#modalModelCode").value = model.code;
        $("#modalModelName").value = model.name;
      }
    });
  }
  if (type === "component") {
    const syncMaterialFields = (event) => {
      const option = event.target.selectedOptions[0];
      const materialSelect = $("#modalMaterialSelect");
      const nameSelect = $("#modalMaterialNameSelect");
      if (materialSelect.value !== event.target.value) materialSelect.value = event.target.value;
      if (nameSelect.value !== event.target.value) nameSelect.value = event.target.value;
      $("#modalMaterialUnit").value = option?.dataset.uom || "";
    };
    $("#modalMaterialSelect").addEventListener("change", syncMaterialFields);
    $("#modalMaterialNameSelect").addEventListener("change", syncMaterialFields);
    $("#modalMaterialSelect").dispatchEvent(new Event("change"));
  }
  $("#modalForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const formEl = event.currentTarget;  // event.currentTarget is empty after the first await
    const values = Object.fromEntries(new FormData(formEl).entries());
    try {
      let result;
      if (type === "product") {
        values.model_ids = [...new FormData(formEl).getAll("model_ids")].map(Number);
        result = await api(record.id ? `/api/planning/products/${record.id}` : "/api/planning/products", { method: record.id ? "PATCH" : "POST", body: JSON.stringify(values) });
      }
      if (type === "model") {
        values.product_ids = [...new FormData(formEl).getAll("product_ids")].map(Number);
        result = await api(record.id ? `/api/planning/models/${record.id}` : "/api/planning/models", { method: record.id ? "PATCH" : "POST", body: JSON.stringify(values) });
      }
      if (type === "material") {
        takeCategory(values);
        for (const key of ["stock_on_hand", "buffer_days", "unit_cost"]) values[key] = Number(values[key] || 0);
        result = await api(record.id ? `/api/planning/materials/${record.id}` : "/api/planning/materials", { method: record.id ? "PATCH" : "POST", body: JSON.stringify(values) });
      }
      if (type === "bom") {
        values.model_id = Number(values.model_id);
        values.material_id = Number(values.material_id);
        values.quantity = Number(values.quantity);
        values.uom = formEl.elements.uom.value;
        result = await api(record.root_bom_id ? `/api/planning/model-bom/${values.model_id}/${record.root_bom_id}` : `/api/planning/model-bom/${values.model_id}`, { method: record.root_bom_id ? "PATCH" : "POST", body: JSON.stringify(values) });
        state.selectedModel = String(values.model_id);
      }
      if (type === "component") {
        values.parent_item_id = Number(values.parent_item_id);
        values.child_item_id = Number(values.child_item_id);
        values.quantity = Number(values.quantity);
        const path = `/api/items/${values.parent_item_id}/bom${record.component_id ? `/${record.component_id}` : ""}`;
        result = await api(path, { method: record.component_id ? "PATCH" : "POST", body: JSON.stringify(values) });
      }
      if (type === "supplier") {
        values.tiers = new FormData(formEl).getAll("tiers");
        if (!values.tiers.length) throw new Error("Tick at least one tier: T1, T2 or T3.");
        result = await api("/api/suppliers", { method: "POST", body: JSON.stringify(takeCategory(values)) });
      }
      if (type === "supplier" || type === "material") await reloadCategories();
      if (type === "product" || type === "model" || type === "material" || type === "bom" || type === "component") await loadCoreData();
      modalLayer.hidden = true;
      toast(`${title} saved`);
      await renderPage();
    } catch (error) { toast(error.message, true); }
  });
}

// Sort a table by a column (from the column's filter pop-up). Numbers (also "₹1,800", "40%",
// "14 days") sort as numbers, text alphabetically. Tree tables keep each parent's lower-level
// rows directly below it: groups of rows move together.
function sortValue(cell) {
  const text = cellText(cell);
  const match = text.replace(/^[₹$€£¥]\s*/, "").replace(/,/g, "").match(/^-?\d+(\.\d+)?/);
  return { text, number: match ? Number(match[0]) : null };
}
function sortTableByColumn(tableNode, column, direction) {
  const body = tableNode.tBodies[0];
  if (!body) return;
  const rows = [...body.rows].filter((row) => !row.querySelector(".table-empty"));
  const depth = (row) => Number(row.dataset.treeDepth || 0);
  const groups = [];
  rows.forEach((row) => {
    if (!groups.length || depth(row) <= depth(groups[groups.length - 1][0])) groups.push([row]);
    else groups[groups.length - 1].push(row);
  });
  const sign = direction === "desc" ? -1 : 1;
  groups.sort((a, b) => {
    const x = sortValue(a[0].cells[column]);
    const y = sortValue(b[0].cells[column]);
    if (x.number !== null && y.number !== null && x.number !== y.number) return sign * (x.number - y.number);
    if (!x.text !== !y.text) return x.text ? -1 : 1;  // blanks last
    return sign * x.text.localeCompare(y.text, undefined, { numeric: true, sensitivity: "base" });
  });
  groups.flat().forEach((row) => body.appendChild(row));
  tableNode.querySelectorAll("thead th").forEach((th) => th.classList.remove("sorted-asc", "sorted-desc"));
  headerRow(tableNode).cells[column]?.classList.add(direction === "desc" ? "sorted-desc" : "sorted-asc");
}

document.addEventListener("click", async (event) => {
  const sortButton = event.target.closest("[data-column-sort]");
  if (sortButton) {
    const header = sortButton.closest("th");
    sortTableByColumn(header.closest("table"), header.cellIndex, sortButton.dataset.columnSort);
    sortButton.closest("details")?.removeAttribute("open");
    return;
  }
  const columnAction = event.target.closest("[data-column-select-all], [data-column-clear], [data-column-apply]");
  if (columnAction) {
    const filter = columnAction.closest(".table-column-filter");
    if (columnAction.hasAttribute("data-column-select-all")) {
      filter.querySelectorAll("[data-column-value]").forEach((input) => { input.checked = true; });
      return;
    }
    if (columnAction.hasAttribute("data-column-clear")) {
      filter.querySelectorAll("[data-column-value]").forEach((input) => { input.checked = false; });
      runTableFilters(filter.closest("table"));
      filter.open = false;
      return;
    }
    runTableFilters(filter.closest("table"));
    filter.open = false;
    return;
  }
  const modelSelectionAction = event.target.closest("[data-model-apply], [data-model-clear], [data-model-select-all]");
  if (modelSelectionAction) {
    const picker = modelSelectionAction.closest(".model-picker") || $(".model-picker");
    if (modelSelectionAction.hasAttribute("data-model-clear")) {
      picker.querySelectorAll("[data-model-option]").forEach((input) => { input.checked = false; });
      picker.open = false;
      try { return await applyModelSelection(); }
      catch (error) { toast(error.message, true); return; }
    }
    if (modelSelectionAction.hasAttribute("data-model-select-all")) {
      picker.querySelectorAll("[data-model-option]").forEach((input) => { input.checked = true; });
      return;
    }
    picker.open = false;
    try { return await applyModelSelection(); }
    catch (error) { toast(error.message, true); return; }
  }
  const parentSelectionAction = event.target.closest("[data-parent-apply], [data-parent-clear], [data-parent-select-all]");
  if (parentSelectionAction) {
    const picker = parentSelectionAction.closest(".bom-parent-picker");
    if (parentSelectionAction.hasAttribute("data-parent-select-all")) {
      picker.querySelectorAll("[data-parent-option]").forEach((input) => { input.checked = true; });
      return;
    }
    if (parentSelectionAction.hasAttribute("data-parent-clear")) {
      picker.querySelectorAll("[data-parent-option]").forEach((input) => { input.checked = false; });
    }
    try { return await applyParentSelection(); }
    catch (error) { toast(error.message, true); return; }
  }
  const treeExpand = event.target.closest("[data-tree-expand]");
  if (treeExpand) {
    try { return await expandBomTree(treeExpand); }
    catch (error) { toast(error.message, true); return; }
  }
  // Edit / Delete columns (last two cells) are not part of the row's click area:
  // only the buttons themselves act there.
  const clickedCell = event.target.closest("td");
  const clickedRow = clickedCell?.parentElement;
  if (clickedRow?.matches(ROW_CLICK_SELECTOR) && clickedCell.cellIndex >= clickedRow.cells.length - 2 && !event.target.closest("button, a")) return;
  const treeGroup = event.target.closest("[data-tree-group]");
  if (treeGroup && !event.target.closest("button, a, input, select")) {
    treeGroup.querySelector("[data-tree-expand]")?.click();
    return;
  }
  if (event.target.closest('[data-action="new-part-bom"]')) return openNewBomEditor();
  const smEdit = event.target.closest("[data-sm-edit]");
  if (smEdit) {
    const material = (state.smItems || []).find((item) => String(item.id) === smEdit.dataset.smEdit);
    if (material) return openSupplierMaterialEdit(material);
  }
  const smDelete = event.target.closest("[data-sm-delete]");
  if (smDelete) {
    if (!window.confirm("Delete this supplier material?")) return;
    try { await api(`/api/planning/supplier-materials/${smDelete.dataset.smDelete}`, { method: "DELETE" }); toast("Supplier material deleted"); return loadSupplierMaterials(); }
    catch (error) { toast(error.message, true); return; }
  }
  const sbLine = (attribute) => { const node = event.target.closest(`[${attribute}]`); return node ? (state.sbView?.lines || []).find((line) => String(line.id) === node.getAttribute(attribute)) : null; };
  const sbAdd = event.target.closest("[data-sb-add]");
  if (sbAdd) return openSupplierBomLine(sbAdd.dataset.sbAdd || null);
  const sbEditLine = sbLine("data-sb-edit");
  if (sbEditLine) return openSupplierBomLine(sbEditLine.parent_line_id, sbEditLine);
  const sbSourcesLine = sbLine("data-sb-sources");
  if (sbSourcesLine) return openSupplierBomSources(sbSourcesLine);
  const sbDelete = event.target.closest("[data-sb-delete]");
  if (sbDelete) {
    if (!window.confirm("Delete this line and every line below it?")) return;
    try { await api(`/api/planning/supplier-bom/lines/${sbDelete.dataset.sbDelete}`, { method: "DELETE" }); toast("Supplier BOM line deleted"); return loadSupplierBom(); }
    catch (error) { toast(error.message, true); return; }
  }
  const sbToggle = event.target.closest("[data-sb-toggle]");
  if (sbToggle) {
    const row = sbToggle.closest("tr");
    const depth = Number(row.dataset.treeDepth);
    const open = sbToggle.getAttribute("aria-expanded") === "true";
    let next = row.nextElementSibling;
    while (next && Number(next.dataset.treeDepth) > depth) {
      next.hidden = open;
      const inner = next.querySelector("[data-sb-toggle]");
      if (inner && !open) { inner.textContent = "−"; inner.setAttribute("aria-expanded", "true"); }
      next = next.nextElementSibling;
    }
    sbToggle.textContent = open ? "+" : "−";
    sbToggle.setAttribute("aria-expanded", String(!open));
    return;
  }
  const supplierTab = event.target.closest("[data-supplier-tab]");
  if (supplierTab) {
    event.preventDefault();
    state.supplierTab = supplierTab.dataset.supplierTab;
    state.coverageLinkId = null;
    if (supplierTab.classList.contains("tab-button")) state.relSupplierId = null;
    return renderSupplierControl();
  }
  const treeExport = event.target.closest("[data-bom-tree-export]");
  if (treeExport) {
    event.preventDefault();
    window.location.href = bomTreeExportUrl(treeExport.dataset.bomTreeExport);
    return;
  }
  if (event.target.closest("[data-execute-master-search]")) {
    if (state.bomTab === "Material Master") return loadMaterialMasterPage(0);  // server search over all codes
    if (TIER_TABS[state.bomTab]) return loadTierMaterialPage(0);
    if (state.bomTab === "Alternate codes") return loadAlternateCodes(0);
    if (state.bomTab === "SOB change log") return loadSobLog(0);
    return applyTableSearch({ target: $("#masterSearch") });
  }
  if (event.target.closest('[data-action="new-tier-material"]')) return openTierMaterial();
  const tierEdit = event.target.closest("[data-tier-edit]");
  if (tierEdit) return openTierMaterial((state.tierItems || []).find((m) => String(m.id) === tierEdit.dataset.tierEdit));
  const tierDelete = event.target.closest("[data-tier-delete]");
  if (tierDelete) {
    if (!window.confirm("Delete this code?")) return;
    try { await api(`/api/planning/tier-materials/${tierDelete.dataset.tierDelete}`, { method: "DELETE" }); toast("Code deleted"); return loadTierMaterialPage(0); }
    catch (error) { toast(error.message, true); return; }
  }
  const altDelete = event.target.closest("[data-alt-delete]");
  if (altDelete) {
    if (!window.confirm("Delete this alternate code? Its SOB on every BOM line goes back to the main code.")) return;
    try { await api(`/api/planning/alternate-codes/${altDelete.dataset.altDelete}`, { method: "DELETE" }); toast("Alternate code deleted"); return loadAlternateCodes(0); }
    catch (error) { toast(error.message, true); return; }
  }
  const categoryRename = event.target.closest("[data-category-rename]");
  if (categoryRename) {
    const category = (state.categoryItems || []).find((c) => String(c.id) === categoryRename.dataset.categoryRename);
    const name = window.prompt(`Rename "${category.name}" to (every part and supplier using it follows):`, category.name);
    if (!name || name.trim() === category.name) return;
    try { const result = await api(`/api/planning/categories/${category.id}`, { method: "PATCH", body: JSON.stringify({ name }) }); toast(`Renamed to ${result.name} (${result.records} record${result.records === 1 ? "" : "s"} updated)`); await reloadCategories(); return loadCategoryTab(); }
    catch (error) { toast(error.message, true); return; }
  }
  const categoryDelete = event.target.closest("[data-category-delete]");
  if (categoryDelete) {
    if (!window.confirm("Delete this category from the list?")) return;
    try { await api(`/api/planning/categories/${categoryDelete.dataset.categoryDelete}`, { method: "DELETE" }); toast("Category deleted"); await reloadCategories(); return loadCategoryTab(); }
    catch (error) { toast(error.message, true); return; }
  }
  const categoryReplace = event.target.closest("[data-category-replace]");
  if (categoryReplace) {
    const row = categoryReplace.closest("tr");
    const target = row.querySelector("[data-replace-with]").value;
    if (!target) return toast("Choose the category to replace it with.", true);
    try { const result = await api("/api/planning/categories/replace", { method: "POST", body: JSON.stringify({ old: row.dataset.unlisted, new: target }) }); toast(`${result.old} replaced by ${result.new} (${result.records} record${result.records === 1 ? "" : "s"})`); await reloadCategories(); return loadCategoryTab(); }
    catch (error) { toast(error.message, true); return; }
  }
  const categoryAdopt = event.target.closest("[data-category-adopt]");
  if (categoryAdopt) {
    try { const saved = await api("/api/planning/categories", { method: "POST", body: JSON.stringify({ name: categoryAdopt.closest("tr").dataset.unlisted }) }); toast(`${saved.name} added to your list`); await reloadCategories(); return loadCategoryTab(); }
    catch (error) { toast(error.message, true); return; }
  }
  const altManage = event.target.closest("[data-alt-manage]");
  if (altManage) return openManageSob(Number(altManage.dataset.altManage));
  const sobLogPage = event.target.closest("[data-sob-log-page]");
  if (sobLogPage) return loadSobLog(Number(sobLogPage.dataset.sobLogPage));
  const altPage = event.target.closest("[data-alt-page]");
  if (altPage) return loadAlternateCodes(Number(altPage.dataset.altPage));
  const tierPage = event.target.closest("[data-tier-page]");
  if (tierPage) return loadTierMaterialPage(Number(tierPage.dataset.tierPage));
  const moveMaterial = event.target.closest("[data-move-material]");
  if (moveMaterial) return openMoveMaterial((state.materialPageItems || []).find((item) => String(item.id) === moveMaterial.dataset.moveMaterial));
  const materialPage = event.target.closest("[data-material-page]");
  if (materialPage) return loadMaterialMasterPage(Number(materialPage.dataset.materialPage));
  const l1Page = event.target.closest("[data-l1-page]");
  if (l1Page) { state.l1Page = Number(l1Page.dataset.l1Page); return loadBomTab(); }
  if (event.target.closest("[data-execute-model-search]")) return applyTableSearch({ target: $("#modelMasterSearch") });
  const isDeleteAction = event.target.closest("[data-delete-product], [data-delete-model], [data-delete-material], [data-delete-bom-row], [data-delete-component-row]");
  const productModelRow = isDeleteAction ? null : event.target.closest("[data-product-model-row]");
  if (productModelRow && productModelRow.dataset.hasBom === "true" && !event.target.closest("button, a, input, select")) {
    state.page = "master";
    state.bomTab = "Model BOM Level 1";
    state.selectedModel = String(productModelRow.dataset.productModelRow);
    state.selectedModelIds = [Number(productModelRow.dataset.productModelRow)];
    state.selectedBomParent = "";
    state.selectedBomParentIds = [];
    return renderMaster();
  }
  const unassignedProductCell = isDeleteAction ? null : event.target.closest("[data-product-id] td");
  const unassignedProductRow = isDeleteAction ? null : event.target.closest("[data-product-id]");
  const unassignedProductColumn = unassignedProductCell?.cellIndex ?? (event.target === unassignedProductRow ? 1 : -1);
  if (unassignedProductRow && unassignedProductColumn >= 1 && unassignedProductColumn < 5 && !event.target.closest("button, a, input, select")) {
    const record = state.products.find((item) => item.id === Number(unassignedProductRow.dataset.productId));
    if (record) return openModal("product", record);
  }
  const level1Row = isDeleteAction ? null : event.target.closest("[data-bom-level1-row]");
  if (level1Row && level1Row.dataset.hasBom === "true" && !event.target.closest("button, a, input, select")) {
    state.bomTab = "Level 2 Part BOM";
    state.selectedModel = String(level1Row.dataset.modelId);
    state.selectedModelIds = [Number(level1Row.dataset.modelId)];
    state.selectedBomParent = String(level1Row.dataset.materialId);
    state.selectedBomParentIds = [Number(level1Row.dataset.materialId)];
    return renderMaster();
  }
  const level2Row = isDeleteAction ? null : event.target.closest("[data-bom-level2-row]");
  if (level2Row && state.bomTab === "Level 2 Part BOM" && level2Row.dataset.hasBom === "true" && !event.target.closest("button, a, input, select")) {
    state.bomTab = "Level 3 Part BOM";
    state.selectedBomParent = String(level2Row.dataset.materialId);
    state.selectedBomParentIds = [Number(level2Row.dataset.materialId)];
    return renderMaster();
  }
  const productEdit = isDeleteAction ? null : event.target.closest("[data-edit-product]");
  if (productEdit) {
    const record = state.products.find((item) => item.id === Number(productEdit.dataset.editProduct));
    if (record) return openModal("product", { ...record, editModelId: Number(productEdit.dataset.editModelId) || null });
  }
  const modelEdit = isDeleteAction ? null : event.target.closest("[data-edit-model]");
  if (modelEdit) {
    const record = state.models.find((item) => item.id === Number(modelEdit.dataset.editModel));
    if (record) return openModal("model", record);
  }
  const materialEdit = isDeleteAction ? null : event.target.closest("[data-edit-material]");
  if (materialEdit) {
    const id = Number(materialEdit.dataset.editMaterial);
    const record = (state.materialPageItems || []).find((item) => item.id === id) || state.materials.find((item) => item.id === id);
    if (record) return openModal("material", record);
  }
  const bomEdit = isDeleteAction ? null : event.target.closest("[data-edit-bom-row], [data-edit-component-row]");
  if (bomEdit) {
    const record = {
      root_bom_id: Number(bomEdit.dataset.editBomRow) || null,
      component_id: Number(bomEdit.dataset.editComponentRow) || null,
      parent_item_id: Number(bomEdit.dataset.parentItem) || null,
      material_id: Number(bomEdit.dataset.materialId),
      quantity: Number(bomEdit.dataset.quantity),
      uom: bomEdit.dataset.uom,
      model_id: Number(bomEdit.dataset.modelId || state.selectedModel),
      alternates: JSON.parse(bomEdit.closest("tr")?.dataset.alternates || "[]"),
    };
    const model = state.models.find((item) => item.id === record.model_id);
    if (model) {
      record.model_code = model.code;
      record.model_name = model.name;
    }
    if (record.root_bom_id) return openModal("bom", record);
    if (record.component_id) return openModal("component", record);
  }
  const productDelete = event.target.closest("[data-delete-product]");
  const modelDelete = event.target.closest("[data-delete-model]");
  const materialDelete = event.target.closest("[data-delete-material]");
  const bomDeleteCandidate = event.target.closest("[data-delete-bom-row], [data-delete-component-row]");
  const bomDelete = bomDeleteCandidate && (Number(bomDeleteCandidate.dataset.deleteBomRow) || Number(bomDeleteCandidate.dataset.deleteComponentRow)) ? bomDeleteCandidate : null;
  if (productDelete || modelDelete || materialDelete || bomDelete) {
    const isRootBOM = bomDelete && Number(bomDelete.dataset.deleteBomRow);
    const message = productDelete ? "Delete this product and its model links? This action cannot be undone."
      : modelDelete ? "Delete this model? This action cannot be undone."
      : materialDelete ? "Delete this material? This action cannot be undone."
      : isRootBOM ? "Delete this model BOM line?"
      : "Delete this shared sub-BOM link? This affects every model using the parent part.";
    if (!window.confirm(message)) return;
    try {
      let path;
      if (productDelete) path = `/api/planning/products/${productDelete.dataset.deleteProduct}`;
      else if (modelDelete) path = `/api/planning/models/${modelDelete.dataset.deleteModel}`;
      else if (materialDelete) path = `/api/planning/materials/${materialDelete.dataset.deleteMaterial}`;
      else if (isRootBOM) path = `/api/planning/model-bom/${bomDelete.dataset.modelId || state.selectedModel}/${bomDelete.dataset.deleteBomRow}`;
      else path = `/api/items/${bomDelete.dataset.parentItem}/bom/${bomDelete.dataset.deleteComponentRow}`;
      await api(path, { method: "DELETE" });
      toast("Line deleted");
      if (productDelete || modelDelete || materialDelete) await loadCoreData();
      return await renderPage();
    } catch (error) { toast(error.message, true); return; }
  }
  const nav = event.target.closest("[data-page]");
  if (nav) {
    event.preventDefault();
    return renderPage(nav.dataset.page);
  }
  if (event.target.closest("[data-close-modal]")) {
    modalLayer.hidden = true;
    return;
  }
  const planRow = event.target.closest("[data-plan-id]");
  if (planRow) {
    try { return await openPlanEditor(Number(planRow.dataset.planId)); }
    catch (error) { toast(error.message, true); }
  }
  const planPage = event.target.closest("[data-plan-editor-page]");
  if (planPage) {
    try { return await loadPlanEditorPage(Number(planPage.dataset.planEditorPage)); }
    catch (error) { toast(error.message, true); }
  }
  if (event.target.closest("[data-plan-editor-close]")) {
    state.planEditor = null;
    $("#planRevisionEditor").hidden = true;
    return;
  }
  const tab = event.target.closest("[data-bom-tab]");
  if (tab) {
    state.bomTab = tab.dataset.bomTab;
    state.bomOffset = 0;
    state.selectedBomParent = "";
    state.selectedBomParentIds = [];
    if (state.bomTab === "Model BOM Level 1") {
      state.selectedModelIds = [];
      state.selectedModel = "";
    }
    return renderMaster();
  }
  if (event.target.closest('[data-action="refresh"]')) return renderPage();
  if (event.target.closest('[data-action="clear-report-filters"]')) return renderReportTable();
  const reportPage = event.target.closest("[data-report-page]");
  if (reportPage && state.reportQuery) {
    try { return await fetchReportPage(Number(reportPage.dataset.reportPage)); }
    catch (error) { toast(error.message, true); }
  }
  const bomPage = event.target.closest("[data-bom-page]");
  if (bomPage) {
    try { return await loadBomTab(Number(bomPage.dataset.bomPage)); }
    catch (error) { toast(error.message, true); }
  }
  const supplierPage = event.target.closest("[data-supplier-page]");
  if (supplierPage) {
    try { return await loadSupplierPartPage(Number(supplierPage.dataset.supplierPage)); }
    catch (error) { toast(error.message, true); }
  }
  if (event.target.closest('[data-action="clear-table-filters"]')) {
    const panel = event.target.closest(".panel");
    panel.querySelectorAll('input[type="search"], input[type="month"], select:not([data-preserve-clear])').forEach((input) => { input.value = ""; });
    panel.querySelectorAll(".table-column-filter .multi-search").forEach((input) => {
      input.value = "";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    panel.querySelectorAll("[data-model-search]").forEach((input) => {
      input.value = "";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    panel.querySelectorAll("[data-parent-search]").forEach((input) => {
      input.value = "";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    });
    panel.querySelectorAll("tbody tr").forEach((row) => { row.hidden = false; });
    panel.querySelectorAll("[data-filter-key], [data-column-value]").forEach((input) => { input.checked = false; });
    panel.querySelectorAll("[data-model-option]").forEach((input) => { input.checked = false; });
    panel.querySelectorAll("[data-parent-option]").forEach((input) => { input.checked = false; });
    if (panel.querySelector("#modelPickerSummary")) {
      state.selectedModelIds = state.models.map((model) => model.id);
      state.selectedModel = state.selectedModelIds.length === 1 ? String(state.selectedModelIds[0]) : "";
      panel.querySelectorAll("[data-model-option]").forEach((input) => { input.checked = true; });
      const summary = $("#modelPickerSummary");
      if (summary) summary.textContent = "All models";
      try { await loadBomTab(0); }
      catch (error) { toast(error.message, true); }
    }
    if (panel.querySelector("#bomParentPicker")) {
      state.selectedBomParentIds = [];
      state.selectedBomParent = "";
      try { return await renderMaster(); }
      catch (error) { toast(error.message, true); return; }
    }
    panel.querySelectorAll("table").forEach((tableNode) => {
      runTableFilters(tableNode);
    });
    if (panel.querySelector("#supplierPartSearch")) {
      try { await loadSupplierPartPage(0, ""); }
      catch (error) { toast(error.message, true); }
    }
    return;
  }
  if (event.target.closest('[data-action="new-product"]')) return openModal("product");
  if (event.target.closest('[data-action="new-model"]')) return openModal("model");
  if (event.target.closest('[data-action="new-material"]')) return openModal("material");
  if (event.target.closest('[data-action="new-bom"]')) {
    if (state.bomTab === "Model BOM Level 1") return openModal("bom");
    if (!state.selectedBomParent) return toast("Select one parent part first.", true);
    return openModal("component", { parent_item_id: Number(state.selectedBomParent) });
  }
  if (event.target.closest('[data-action="new-supplier"]')) return openModal("supplier");
  if (event.target.closest('[data-action="reset-trial"]')) {
    if (!window.confirm("Delete all trial plans, actuals, stock snapshots, POs and transfers? Production data and master data will be preserved.")) return;
    try {
      await api("/api/planning/trial-data", { method: "DELETE" });
      toast("Trial data cleared");
      return renderUsers();
    } catch (error) { toast(error.message, true); }
  }
  if (event.target.closest("[data-approve-supplier]")) {
    const approval = event.target.closest("[data-approve-supplier]");
    try {
      await api(`/api/suppliers/${approval.dataset.approveSupplier}`, { method: "PATCH", body: JSON.stringify({ status: "Approved" }) });
      toast("Supplier approved");
      return renderSupplierControl();
    } catch (error) { toast(error.message, true); }
  }
  const reschedule = event.target.closest("[data-reschedule-po]");
  if (reschedule) {
    const newMonth = window.prompt("Enter the revised expected arrival month (YYYY-MM):", reschedule.dataset.currentMonth);
    if (!newMonth) return;
    const note = window.prompt("Reason for advancing or postponing the delivery:");
    if (!note) return;
    try {
      await api(`/api/planning/purchase-orders/${reschedule.dataset.reschedulePo}/schedule`, { method: "PATCH", body: JSON.stringify({ expected_month: newMonth, change_note: note }) });
      toast("PO schedule updated");
      return renderPurchaseOrders();
    } catch (error) { toast(error.message, true); }
  }
  const receipt = event.target.closest("[data-receive-po]");
  if (receipt) {
    const received = Number(window.prompt(`Quantity received (remaining ${receipt.dataset.remaining}):`, receipt.dataset.remaining));
    if (!Number.isFinite(received) || received <= 0) return;
    const receivedMonth = window.prompt("Receipt month (YYYY-MM):", monthNow());
    if (!receivedMonth) return;
    try {
      await api(`/api/planning/purchase-orders/${receipt.dataset.receivePo}/receipts`, { method: "POST", body: JSON.stringify({ quantity_received: received, received_month: receivedMonth }) });
      toast("PO receipt recorded and company stock updated");
      return renderPurchaseOrders();
    } catch (error) { toast(error.message, true); }
  }
  const tierTabButton = event.target.closest("[data-supplier-tier-tab]");
  if (tierTabButton) { state.supplierTierTab = tierTabButton.dataset.supplierTierTab; return renderSupplierControl(); }
  const supplierEdit = event.target.closest("[data-edit-supplier]");
  if (supplierEdit) return openSupplierEdit(state.suppliers.find((supplier) => String(supplier.id) === supplierEdit.dataset.editSupplier));
  const supplierDelete = event.target.closest("[data-delete-supplier]");
  if (supplierDelete) {
    const supplier = state.suppliers.find((item) => String(item.id) === supplierDelete.dataset.deleteSupplier);
    if (!window.confirm(`Delete supplier ${supplier?.name || ""}? This cannot be undone.`)) return;
    try { await api(`/api/suppliers/${supplierDelete.dataset.deleteSupplier}`, { method: "DELETE" }); toast("Supplier deleted"); return renderSupplierControl(); }
    catch (error) { toast(error.message, true); return; }
  }
  const openRel = event.target.closest("[data-open-supplier-rel]");
  if (openRel) { event.preventDefault(); state.supplierTab = "Supplier/material relationship"; state.relSupplierId = Number(openRel.dataset.openSupplierRel) || null; state.coverageLinkId = null; return renderSupplierControl(); }
  const relAll = event.target.closest("[data-rel-all]");
  if (relAll) { event.preventDefault(); state.relSupplierId = null; state.coverageLinkId = null; return renderSupplierControl(); }
  const openCoverage = event.target.closest("[data-open-coverage]");
  if (openCoverage) { state.supplierTab = "Supplier/material relationship"; state.coverageLinkId = Number(openCoverage.dataset.openCoverage); return renderSupplierControl(); }
  const covAddLine = event.target.closest("[data-cov-add-line]");
  if (covAddLine) return openCoverageAddLine(covAddLine.dataset.covAddLine || null);
  const covAddSource = event.target.closest("[data-cov-add-source]");
  if (covAddSource) return openCoverageSource(coverageLine(covAddSource.dataset.covAddSource));
  const covEditLine = event.target.closest("[data-cov-edit-line]");
  if (covEditLine) return openCoverageEditLine(coverageLine(covEditLine.dataset.covEditLine));
  const covDeleteLine = event.target.closest("[data-cov-delete-line]");
  if (covDeleteLine) {
    const line = coverageLine(covDeleteLine.dataset.covDeleteLine);
    if (!window.confirm(`Delete ${line.component.code}, its suppliers and every line below it?`)) return;
    try { await api(`/api/planning/supplier-bom/lines/${line.id}`, { method: "DELETE" }); toast("Line deleted"); return loadCoverage(); }
    catch (error) { toast(error.message, true); return; }
  }
  const covSource = event.target.closest("[data-cov-edit-source], [data-cov-delete-source]");
  if (covSource) {
    const [lineId, index] = (covSource.dataset.covEditSource || covSource.dataset.covDeleteSource).split(":");
    const line = coverageLine(lineId);
    if (covSource.dataset.covEditSource) return openCoverageSource(line, Number(index));
    if (!canSeePrices()) return toast("Price access is needed to change the suppliers of a line.", true);
    if (!window.confirm(`Remove ${line.sources[Number(index)].sub_supplier} as a supplier of ${line.component.code}?`)) return;
    return saveSources(line, line.sources.filter((_, i) => i !== Number(index)).map(sourcePayload), "Supplier removed");
  }
  const partEdit = event.target.closest("[data-edit-supplier-part]");
  if (partEdit) return openSupplierPartEdit((state.supplierParts || []).find((link) => String(link.id) === partEdit.dataset.editSupplierPart));
  const partDelete = event.target.closest("[data-delete-supplier-part]");
  if (partDelete) {
    if (!window.confirm("Delete this supplier/material relationship?")) return;
    try { await api(`/api/planning/supplier-parts/${partDelete.dataset.deleteSupplierPart}`, { method: "DELETE" }); modalLayer.hidden = true; toast("Relationship deleted"); return loadSupplierPartPage(state.supplierOffset || 0); }
    catch (error) { toast(error.message, true); return; }
  }
  const priceUpdate = event.target.closest("[data-update-price]");
  if (priceUpdate) {
    const updatedPrice = Number(window.prompt("Enter the new current unit price:", priceUpdate.dataset.currentPrice));
    if (!Number.isFinite(updatedPrice) || updatedPrice < 0) return;
    try {
      await api(`/api/planning/supplier-parts/${priceUpdate.dataset.updatePrice}`, { method: "PATCH", body: JSON.stringify({ unit_price: updatedPrice }) });
      toast("Current supplier price updated");
      return renderSupplierControl();
    } catch (error) { toast(error.message, true); }
  }
});

document.addEventListener("click", (event) => {
  const search = event.target.closest(".multi-filter summary");
  if (!search) return;
  event.stopPropagation();
});

// Filter lists and pickers close when you click anywhere outside them (or press Esc), even with nothing chosen
const DROPDOWN_SELECTOR = "details.multi-filter[open], details.model-picker[open]";
document.addEventListener("click", (event) => {
  document.querySelectorAll(DROPDOWN_SELECTOR).forEach((dropdown) => { if (!dropdown.contains(event.target)) dropdown.open = false; });
}, true);
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    const open = [...document.querySelectorAll(DROPDOWN_SELECTOR)];
    if (open.length) { open.forEach((dropdown) => { dropdown.open = false; }); return; }
  }
  if (event.key === "Escape" && !modalLayer.hidden) modalLayer.hidden = true;
  if ((event.key === "Enter" || event.key === " ") && event.target.matches("[data-plan-id]")) {
    event.preventDefault();
    event.target.click();
  }
  if ((event.key === "Enter" || event.key === " ") && event.target.matches("[data-product-model-row], [data-product-id]")) {
    event.preventDefault();
    event.target.click();
  }
  if ((event.key === "Enter" || event.key === " ") && event.target.matches("[data-bom-level1-row], [data-bom-level2-row]")) {
    if (event.target.closest("button")) return;
    event.preventDefault();
    event.target.click();
  }
  if (event.key === "Enter" && event.target.matches("#masterSearch")) {
    event.preventDefault();
    $("[data-execute-master-search]")?.click();
  }
  if (event.key === "Enter" && event.target.matches("#modelMasterSearch")) {
    event.preventDefault();
    $("[data-execute-model-search]")?.click();
  }
});

document.addEventListener("input", (event) => {
  if (event.target.matches(".table-column-filter .multi-search")) {
    const needle = event.target.value.toLowerCase();
    event.target.closest(".table-column-filter").querySelectorAll(".multi-options label").forEach((label) => {
      label.hidden = !label.textContent.toLowerCase().includes(needle);
    });
    return;
  }
  if (!event.target.matches("[data-plan-quantity]") || !state.planEditor) return;
  const input = event.target;
  const key = `${input.dataset.modelCode}\u0000${input.dataset.month}`;
  state.planEditor.edits.set(key, {
    model_code: input.dataset.modelCode, target_month: input.dataset.month, quantity: input.value,
  });
  $("#savePlanRevision").disabled = !state.planEditor.latest || state.planEditor.edits.size === 0;
});

$("#refreshButton").addEventListener("click", () => renderPage());
$("#logoutButton").addEventListener("click", async () => {
  try {
    await api("/api/auth/logout", { method: "POST", body: "{}" });
  } catch (error) {
    toast(error.message, true);
    return;
  }
  state.user = null;
  authScreen(false);
});

async function boot() {
  try {
    const response = await fetch("/api/auth/me");
    if (!response.ok) throw new Error("not signed in");
    const data = await response.json();
    state.user = data.user;
    $("#appShell").hidden = false;
    $("#authView").hidden = true;
    initUserShell();
    await loadCoreData();
    await renderPage("dashboard");
  } catch {
    try {
      const setup = await api("/api/auth/setup/status");
      authScreen(setup.setup_required);
    } catch (error) {
      $("#authView").hidden = false;
      $("#authView").innerHTML = `<div class="auth-card"><h1>Workspace unavailable</h1><p>${esc(error.message)}</p></div>`;
    }
  }
}

boot();
