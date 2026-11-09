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
    return `<th><div class="table-heading"><details class="multi-filter table-column-filter"><summary aria-label="Filter ${esc(header)}">${esc(header)}</summary><div class="multi-popover"><input class="form-control multi-search" type="search" placeholder="Search ${esc(header.toLowerCase())}" /><div class="column-filter-actions"><button type="button" class="text-button" data-column-select-all>Select all</button><button type="button" class="text-button" data-column-clear>Clear all</button></div><div class="multi-options">${options || '<span class="subtle">No values</span>'}</div><button type="button" class="button button-small button-primary column-filter-apply" data-column-apply>Apply</button></div></details><span class="column-resizer" aria-hidden="true"></span></div></th>`;
  }).join("");
  return `<div class="table-wrap"><table><thead><tr>${headings}</tr></thead><tbody>${rows || `<tr><td class="table-empty" colspan="${headers.length}">${esc(empty)}</td></tr>`}</tbody></table></div>`;
}

function tableExportLinks(kind, parameters = {}) {
  const query = new URLSearchParams(Object.entries(parameters).filter(([, value]) => value !== undefined && value !== null && value !== ""));
  const suffix = query.size ? `?${query}` : "";
  return `<a class="button button-small" href="/api/planning/tables/${encodeURIComponent(kind)}/xlsx${suffix}">Download Excel</a><a class="button button-small" href="/api/planning/tables/${encodeURIComponent(kind)}/pdf${suffix}">Download PDF</a>`;
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
  if (selected && !names.some((name) => name.toLowerCase() === selected.toLowerCase())) names.unshift(selected);
  return `<label class="field"><span>Category</span><select class="form-control" name="category_choice" id="categorySelect" required>${names.map((name) => `<option value="${esc(name)}" ${name.toLowerCase() === String(selected).toLowerCase() ? "selected" : ""}>${esc(name)}</option>`).join("")}<option value="__new__">+ Add new category</option></select><input class="form-control category-new" id="categoryNew" name="category_new" placeholder="Type new category name" minlength="2" maxlength="80" hidden /></label>`;
}

function unitField(label = "Unit", name = "uom", selected = "ea") {
  const units = [...new Set([...state.materials.map((material) => material.uom), "ea", "PCS", "Nos", "kg", "g", "grm", "m", "mtr", "mm", "cm", "L", "ml", "set", "pair", "roll", "box", "pack", "lot"].filter(Boolean))];
  return `<label class="field"><span>${esc(label)}</span><select class="form-control" name="${esc(name)}" required>${units.map((unit) => `<option value="${esc(unit)}" ${unit === selected ? "selected" : ""}>${esc(unit)}</option>`).join("")}</select></label>`;
}

function workspaceSelect() {
  return `<select class="form-control" name="workspace"><option value="trial" ${state.workspace === "trial" ? "selected" : ""}>Trial</option><option value="production" ${state.workspace === "production" ? "selected" : ""}>Production</option></select>`;
}

function bindMaterialSearch(searchId, selectId, nameId = null) {
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
  search.addEventListener("input", () => {
    clearTimeout(state.searchTimer);
    state.searchTimer = setTimeout(async () => {
      try {
        const query = new URLSearchParams({ offset: "0", limit: "100", search: search.value });
        const result = await api(`/api/planning/materials?${query}`);
        renderMaterialOptions(result.items);
        select.dispatchEvent(new Event("change"));
      } catch (error) { toast(error.message, true); }
    }, 250);
  });
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
  const planRows = state.plans.map((plan) => `<tr class="clickable-row" data-plan-id="${plan.id}" tabindex="0" title="Open ${esc(plan.revision)} to review or edit monthly quantities"><td>${monthLabel(plan.issue_month)}</td><td>${esc(plan.revision)}</td><td>${number(plan.row_count, 0)}</td><td>${new Date(plan.created_at).toLocaleString()}</td><td>${esc(plan.workspace)}</td></tr>`).join("");
  content.innerHTML = `${heading("Monthly plans & actuals", "Import a monthly 24-month plan snapshot. Revisions are immutable; missing model/month rows are treated as zero demand.", `<a class="button" id="planTemplateLink" href="/api/planning/templates/plan.xlsx?issue_month=${month}">Plan template</a><a class="button" href="/api/planning/templates/actuals.xlsx">Actuals template</a>`)}
    <div class="planning-grid">
      <section class="panel">${panelHeader("Upload plan revision", "Use one model per row and one forecast month per column from the Plan template. Long-form model_code, target_month, quantity workbooks are also accepted. Target months must fit within 24 months from the issue month.")}
        <form id="planUploadForm" class="panel-body form-grid">
          ${field("Plan issue month", "issue_month", "month", { required: true, value: month })}
          <label class="field"><span>Workspace</span>${workspaceSelect()}</label>
          <label class="field full"><span>24-month plan workbook (.xlsx)</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label>
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
      <div id="planHistory">${table(["Issue month", "Revision", "Rows", "Revision date", "Workspace"], planRows, "No monthly plans uploaded yet.")}</div>
    </section>`;
  $("#planUploadForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const data = new FormData(form);
    const uploadMonth = data.get("issue_month");
    state.workspace = data.get("workspace");
    try {
      const result = await uploadExcel(`/api/planning/plans/${uploadMonth}/upload?workspace=${state.workspace}`, data.get("workbook"));
      toast(`Plan saved as ${result.revision} for ${result.issue_month}`);
      await renderPlans();
    } catch (error) { toast(error.message, true); }
  });
  $('#planUploadForm [name="issue_month"]').addEventListener("change", (event) => {
    $("#planTemplateLink").href = `/api/planning/templates/plan.xlsx?issue_month=${encodeURIComponent(event.target.value)}`;
  });
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
  setupTableFilters("#planHistory", "#planFilters", null, [{ title: "Issue month", index: 0 }, { title: "Revision", index: 1 }, { title: "Workspace", index: 4 }]);
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
  const materialRows = state.materials.map((item) => `<tr data-edit-material="${item.id}" tabindex="0"><td>${esc(item.code)}</td><td>${esc(item.name)}</td><td>${esc(item.category)}</td><td>${esc(item.uom)}</td><td>${number(item.stock_on_hand)}</td><td>${number(item.buffer_days, 0)} days</td><td>${state.user.can_view_prices || state.user.role === "Admin" ? money(item.unit_cost, item.currency) : "Restricted"}</td><td><button class="text-button" data-edit-material="${item.id}">Edit</button></td><td><button class="text-button button-danger" data-delete-material="${item.id}">Delete</button></td></tr>`).join("");
  const tabs = ["Product Master", "Model BOM Level 1", "Level 2 Part BOM", "Level 3 Part BOM", "Material Master"];
  const bomTemplateKind = {
    "Model BOM Level 1": "bom-oem", "Level 2 Part BOM": "bom-level-2", "Level 3 Part BOM": "bom-level-3",
  }[state.bomTab];
  const actions = state.bomTab === "Product Master"
    ? '<button class="button button-primary" data-action="new-product">Add line</button>'
    : state.bomTab === "Material Master"
      ? '<button class="button" data-action="new-model">Add model</button><button class="button button-primary" data-action="new-material">Add material</button>'
      : "";
  content.innerHTML = `${heading("Supply Chain Management", "Build products from model codes in Material Master, define units once, and manage each model’s BOM by level.", actions)}
    <div class="tab-strip">${tabs.map((tab) => `<button class="tab-button ${state.bomTab === tab ? "active" : ""}" data-bom-tab="${esc(tab)}">${esc(tab)}</button>`).join("")}</div>
    ${state.bomTab === "Product Master" ? `<section class="panel table-panel">${panelHeader("Product Master")}
      <div class="table-toolbar"><label class="search-field"><input class="form-control" id="masterSearch" type="search" placeholder="Search product or model" /></label><button class="button button-small" type="button" data-execute-master-search>Execute</button>${tableExportLinks("products")}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <div id="productMasterTable">${table(treeHeaders("Product Master"), productRows, "No products yet.")}</div></section>`
    : state.bomTab === "Material Master" ? `<section class="panel table-panel">${panelHeader("Material Master", "Define each material code once. The unit selected here is used throughout its BOM and supply chain records.")}
      <div class="master-summaries"><article class="mini-stat"><span>Materials</span><strong>${state.materials.length}</strong></article><article class="mini-stat"><span>Models</span><strong>${state.models.length}</strong></article><article class="mini-stat"><span>Products</span><strong>${state.products.length}</strong></article></div>
      <div class="table-toolbar"><label class="search-field"><input class="form-control" id="masterSearch" type="search" placeholder="Search material code or name" /></label><button class="button button-small" type="button" data-execute-master-search>Execute</button><a class="button button-small" href="/api/planning/templates/own-stock.xlsx">Stock template</a>${tableExportLinks("materials")}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <div id="materialMasterTable">${table(["Material code", "Material name", "Category", "Unit", "On hand", "Buffer norm", "Part price", "Edit", "Delete"], materialRows, "No materials yet.")}</div>
      <h3 class="subtable-heading">Model Master <button class="button button-small button-primary" data-action="new-model">Add line</button>${tableExportLinks("models")}</h3>
      <div class="table-toolbar"><label class="search-field"><input class="form-control" id="modelMasterSearch" type="search" placeholder="Search model code or name" /></label><button class="button button-small" type="button" data-execute-model-search>Execute</button><button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div>
      <div id="modelMasterTable">${table(["Model code", "Model name", "Assigned products", "Action"], modelRows, "No models yet.")}</div></section>`
    : `<section class="panel table-panel ${state.bomTab === "Model BOM Level 1" ? "" : "part-bom-page"}">${panelHeader(state.bomTab)}
      <div class="table-toolbar bom-controls ${state.bomTab === "Model BOM Level 1" ? "" : "bom-controls-part"}"><div class="toolbar-group">
        ${state.bomTab === "Model BOM Level 1"
          ? modelPickerMarkup()
          : bomParentPickerMarkup()}
      </div>
      ${bomTemplateKind ? `<form id="bomUploadForm" class="bom-upload-inline"><input class="form-control bom-file-input" type="file" name="workbook" accept=".xlsx" aria-label="Select ${esc(state.bomTab)} Excel workbook" required /><button class="button button-small button-primary" id="bomUploadButton" type="submit" ${state.bomTab === "Model BOM Level 1" ? (state.selectedModelIds.length === 1 ? "" : "disabled") : (state.selectedBomParent ? "" : "disabled")}>Upload BOM</button></form><a class="button button-small" href="/api/planning/templates/${bomTemplateKind}.xlsx">Level ${state.bomTab === "Model BOM Level 1" ? "1" : state.bomTab === "Level 2 Part BOM" ? "2" : "3"} template</a>` : ""}
      ${state.bomTab === "Model BOM Level 1" && state.selectedModel ? tableExportLinks("bom", { model_id: state.selectedModel, tab: state.bomTab }) : ""}
      ${state.bomTab !== "Model BOM Level 1" && state.selectedBomParent ? tableExportLinks("bom", { parent_item_id: state.selectedBomParent, tab: state.bomTab }) : ""}
      <button class="button button-small" data-action="clear-table-filters">Clear all filters</button>
      ${state.bomTab === "Model BOM Level 1" ? '<button class="button button-small button-primary" data-action="new-bom">Add line</button>' : '<button class="button button-small button-primary" id="newPartBomButton" data-action="new-bom" disabled>Select one parent part</button><button class="button button-small button-primary" data-action="new-part-bom">New BOM</button>'}
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
  } else {
    $("#masterTable").dataset.filtertable = "true";
  }
  if (state.bomTab === "Model BOM Level 1") await loadBomTab(0);
  else if ($("#bomParentPicker")) updatePartBomActions();
  if ($("#bomUploadForm")) {
    $("#bomUploadForm").addEventListener("submit", async (event) => {
      event.preventDefault();
      const file = new FormData(event.currentTarget).get("workbook");
      try {
        const query = new URLSearchParams({ tab: state.bomTab });
        const uploadPath = state.bomTab === "Model BOM Level 1"
          ? `/api/planning/model-bom/${state.selectedModel}/upload?${query}`
          : `/api/planning/part-bom/${state.selectedBomParent}/upload?${query}`;
        const result = await uploadExcel(uploadPath, file);
        toast(`${result.rows} ${state.bomTab} row(s) uploaded`);
        if (state.bomTab === "Model BOM Level 1") await loadBomTab(0);
        else await loadPartBomTab();
        event.currentTarget.reset();
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
  const levelCells = TREE_LEVELS[tab].map((level) => row.cells[level]
    ? `<td>${esc(row.cells[level][0])}</td><td>${esc(row.cells[level][1])}</td>`
    : "<td></td><td></td>").join("");
  return `<tr class="${row.className || ""}" ${row.attributes || ""} data-tree-depth="${row.depth}" tabindex="0"><td class="tree-toggle-cell">${row.toggle || ""}</td>${levelCells}<td>${row.quantity ?? ""}</td><td>${esc(row.uom || "")}</td><td>${row.price ?? ""}</td><td>${row.edit || ""}</td><td>${row.remove || ""}</td></tr>`;
}

function treePrice(item) {
  return item.unit_price === undefined ? "Restricted" : money(item.unit_price, item.currency);
}

function level1RowHtml(tab, item, modelId, depth, showModel) {
  const data = `data-model-id="${modelId}" data-material-id="${item.material_id}" data-quantity="${item.edit_quantity}" data-uom="${esc(item.uom)}"`;
  return treeRowHtml(tab, {
    depth, className: "clickable-row",
    attributes: `data-bom-level1-row data-has-bom="${item.has_bom}" data-edit-bom-row="${item.root_bom_id || ""}" ${data}`,
    cells: { ...(showModel ? { model: [item.model_code, item.model_name] } : {}), l1: [item.code, item.name] },
    toggle: item.has_bom ? treeToggle(item.code, `data-tree-expand="part" data-tree-parent-id="${item.material_id}" data-tree-level="Level 2 Part BOM"`) : "",
    quantity: number(item.quantity), uom: item.uom, price: treePrice(item),
    edit: `<button class="text-button" data-edit-bom-row="${item.root_bom_id || ""}" ${data}>Edit</button>`,
    remove: `<button class="text-button button-danger" data-delete-bom-row="${item.root_bom_id || ""}" data-model-id="${modelId}">Delete</button>`,
  });
}

function partRowHtml(tab, item, parent, lineLevel, depth) {
  const level = lineLevel === "Level 2 Part BOM" ? "l2" : "l3";
  const expandable = level === "l2" && item.has_bom;
  const data = `data-parent-item="${parent.id}" data-material-id="${item.material_id}" data-quantity="${item.quantity}" data-uom="${esc(item.uom)}"`;
  return treeRowHtml(tab, {
    depth, className: "clickable-row",
    attributes: `data-bom-level2-row data-has-bom="${expandable}" data-edit-component-row="${item.component_id}" ${data}`,
    cells: { [level]: [item.code, item.name] },
    toggle: expandable ? treeToggle(item.code, `data-tree-expand="part" data-tree-parent-id="${item.material_id}" data-tree-level="Level 3 Part BOM"`) : "",
    quantity: number(item.quantity), uom: item.uom, price: treePrice(item),
    edit: `<button class="text-button" data-edit-component-row="${item.component_id}" ${data}>Edit</button>`,
    remove: `<button class="text-button button-danger" data-delete-component-row="${item.component_id}" data-parent-item="${parent.id}">Delete</button>`,
  });
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
  try {
    let rows;
    if (button.dataset.treeExpand === "model") {
      const modelId = Number(button.dataset.treeModelId);
      const query = new URLSearchParams({ tab: "Model BOM Level 1", offset: "0", limit: "500" });
      const result = await api(`/api/planning/model-bom/${modelId}?${query}`);
      rows = result.items.map((item) => level1RowHtml(state.bomTab, item, modelId, depth + 1, false));
    } else {
      const level = button.dataset.treeLevel;
      const query = new URLSearchParams({ tab: level, offset: "0", limit: "500" });
      const result = await api(`/api/planning/part-bom/${Number(button.dataset.treeParentId)}?${query}`);
      rows = result.items.map((item) => partRowHtml(state.bomTab, item, result.parent, level, depth + 1));
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
  holder.innerHTML = `<div class="new-bom-editor"><div class="new-bom-head"><label class="field"><span>Parent part (${levelName} part without BOM)</span><select class="form-control" id="newBomParent"><option value="">Select parent part</option>${candidates.map((item) => `<option value="${item.id}">${esc(item.code)} · ${esc(item.name)}</option>`).join("")}</select></label><p class="subtle">Parts come from Material Master only. Use a negative quantity for scrap or a by-product; zero is not allowed.</p></div><div class="table-wrap"><table class="new-bom-lines"><thead><tr><th>Search</th><th>Part code</th><th>Part name</th><th>Quantity per parent</th><th>Unit</th><th></th></tr></thead><tbody id="newBomLines">${newBomLineRow()}</tbody></table></div><div class="new-bom-actions"><button type="button" class="button button-small" data-new-bom-add>+ Add line</button><button type="button" class="button button-small button-primary" data-new-bom-save>Save BOM</button><button type="button" class="button button-small" data-new-bom-cancel>Cancel</button></div></div>`;
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

async function loadBomTab() {
  const modelIds = state.selectedModelIds.length ? state.selectedModelIds : state.models.map((model) => model.id);
  if (!modelIds.length) {
    $("#masterTable").innerHTML = table(treeHeaders("Model BOM Level 1"), "", "Select one or more models to view their BOM.");
    return;
  }
  const results = await Promise.all(modelIds.map(async (modelId) => {
    const items = [];
    let offset = 0;
    let metadata;
    do {
      const query = new URLSearchParams({ tab: state.bomTab, offset: String(offset), limit: "500" });
      metadata = await api(`/api/planning/model-bom/${modelId}?${query}`);
      items.push(...metadata.items);
      offset += metadata.items.length;
    } while (items.length < metadata.total && metadata.items.length);
    return items.map((row) => ({ ...row, model_id: modelId, model_code: metadata.model_code, model_name: metadata.model_name }));
  }));
  const items = results.flat();
  $("#masterTable").innerHTML = table(treeHeaders("Model BOM Level 1"), items.map((item) => level1RowHtml("Model BOM Level 1", item, item.model_id, 0, true)).join(""), "No Model BOM Level 1 lines found.");
}

function renderSupplierPartsTable(result) {
  const canEditPrices = state.user.can_view_prices || state.user.role === "Admin";
  const rows = result.items.map((link) => `<tr><td>${esc(link.material_code)} · ${esc(link.material_name)}</td><td>${esc(link.supplier_code)} · ${esc(link.supplier_name)}</td><td>${esc(link.tier)}${link.parent_supplier_name ? ` via ${esc(link.parent_supplier_name)}` : ""}</td><td>${number(link.sob_percent, 1)}%</td><td>${number(link.lead_time_days, 0)} days</td><td>${number(link.moq)}</td><td>${number(link.buffer_days, 0)} days</td><td>${link.unit_price === undefined ? "Restricted" : `${money(link.unit_price, link.currency)} · ${esc(link.price_mode)}`}</td><td>${esc(link.price_status)}</td><td>${canEditPrices && link.unit_price !== undefined ? `<button class="text-button" data-update-price="${link.id}" data-current-price="${link.unit_price}">Update price</button>` : ""}</td></tr>`).join("");
  const pageSize = result.limit;
  const markup = `${table(["Material", "Supplier", "Tier / route", "SOB", "Lead time", "MOQ", "Buffer", "Price", "Price status", "Action"], rows, "No supplier/material relationships configured.")}<div class="table-toolbar pagination-toolbar"><span>${result.total ? `${(result.offset + 1).toLocaleString()}–${Math.min(result.offset + result.items.length, result.total).toLocaleString()} of ${result.total.toLocaleString()}` : "No supplier/material relationships"}</span><div><button class="button button-small" data-supplier-page="${Math.max(0, result.offset - pageSize)}" ${result.offset === 0 ? "disabled" : ""}>Previous</button><button class="button button-small" data-supplier-page="${result.offset + pageSize}" ${result.offset + result.items.length >= result.total ? "disabled" : ""}>Next</button></div></div>`;
  const target = $("#supplierPartTable");
  if (target) {
    target.innerHTML = markup;
    setupTableFilters("#supplierPartTable", "#supplierPartFilters", null, [{ title: "Supplier", index: 1 }, { title: "Tier", index: 2 }, { title: "Price status", index: 8 }]);
  }
  return markup;
}

async function loadSupplierPartPage(offset = 0, search = state.supplierSearch) {
  const query = new URLSearchParams({ offset: String(offset), limit: "500", search });
  const result = await api(`/api/planning/supplier-parts?${query}`);
  state.supplierParts = result.items;
  state.supplierOffset = result.offset;
  state.supplierTotal = result.total;
  state.supplierSearch = search;
  if ($("#supplierPartTable")) renderSupplierPartsTable(result);
}

async function renderSupplierControl() {
  await loadCoreData();
  const [suppliers, links] = await Promise.all([api("/api/suppliers"), api("/api/planning/supplier-parts?offset=0&limit=500")]);
  state.suppliers = suppliers;
  state.supplierParts = links.items;
  state.supplierOffset = 0;
  state.supplierTotal = links.total;
  state.supplierSearch = "";
  const supplierRows = suppliers.map((supplier) => `<tr data-search="${esc(`${supplier.code} ${supplier.name} ${supplier.category}`.toLowerCase())}"><td>${esc(supplier.code)}</td><td>${esc(supplier.name)}</td><td>${esc(supplier.category)}</td><td>${esc(supplier.status)}</td><td>${supplier.status === "Pending review" ? `<button class="text-button" data-approve-supplier="${supplier.id}">Approve</button>` : ""}</td></tr>`).join("");
  const supplierOptions = suppliers.map((s) => `<option value="${s.id}">${esc(s.code)} · ${esc(s.name)}</option>`).join("");
  const materialOptions = state.materials.map((m) => `<option value="${m.id}">${esc(m.code)} · ${esc(m.name)}</option>`).join("");
  content.innerHTML = `${heading("Supplier control", "Manage supplier tier, share of business, lead time through T2/Tn, MOQ, buffer, dynamic current price and stock availability.", `<button class="button" data-action="new-supplier">Add supplier</button>`)}
    <section class="panel">${panelHeader("Supplier/material relationship", "SOB applies within a material, tier and parent-supplier route; each group is limited to 100%.")}
      <form id="supplierPartForm" class="panel-body form-grid">
        <label class="field"><span>Search all materials</span><input class="form-control" id="supplierMaterialSearch" type="search" placeholder="Type a material code or name" /></label>
        <label class="field"><span>Material</span><select class="form-control" name="material_id" id="supplierMaterialSelect" required>${materialOptions}</select></label>
        <label class="field"><span>Supplier</span><select class="form-control" name="supplier_id" required>${supplierOptions}</select></label>
        ${field("Tier (Direct, T1, T2…)", "tier", "text", { required: true, value: "Direct" })}
        <label class="field"><span>Parent customer supplier (for T1+)</span><select class="form-control" name="parent_supplier_id"><option value="">None / Direct</option>${supplierOptions}</select></label>
        ${field("Supplier SOB (%)", "sob_percent", "number", { required: true, min: 0, step: "0.01", value: 100 })}
        ${field("End-to-end lead time (days)", "lead_time_days", "number", { required: true, min: 0, value: 14 })}
        ${field("MOQ", "moq", "number", { required: true, min: 0, step: "any", value: 0 })}
        ${field("Supplier buffer (days)", "buffer_days", "number", { required: true, min: 0, value: 0 })}
        ${field("Current unit price", "unit_price", "number", { required: true, min: 0, step: "0.0001", value: 0 })}
        <label class="field"><span>Currency</span><select class="form-control" name="currency"><option>USD</option><option>INR</option><option>EUR</option><option>GBP</option><option>JPY</option></select></label>
        <label class="field"><span>Price agreement</span><select class="form-control" name="price_mode"><option value="spot">Spot</option><option value="fixed_monthly">Fixed monthly</option><option value="every_n_months">Changes every N months</option></select></label>
        ${field("Price interval (months)", "price_interval_months", "number", { required: true, min: 1, value: 1 })}
        <button class="button button-primary" type="submit">Save supplier/material</button>
      </form>
    </section>
    <section class="panel table-panel">${panelHeader("Upload supplier stock", "Columns: supplier_code, material_code, tier, parent_supplier_code, quantity, available_quantity, arrival_month. Tier and parent identify the exact relationship when the same supplier/material exists at multiple levels.")}
      <form id="supplierStockForm" class="panel-body inline-upload">
        <label class="field"><span>Report month</span><input class="form-control" type="month" name="report_month" required value="${monthNow()}" /></label>
        <label class="field"><span>Workspace</span>${workspaceSelect()}</label>
        <label class="field"><span>Supplier stock workbook</span><input class="form-control" type="file" name="workbook" accept=".xlsx" required /></label>
        <button class="button button-primary" type="submit">Upload supplier stock</button><a class="button" href="/api/planning/templates/supplier-stock.xlsx">Download template</a>
      </form>
    </section>
    <section class="panel table-panel">${panelHeader("Supplier/material coverage", "Search filters below apply without loading the complete supplier universe into the browser.")}
      <div class="table-toolbar"><div class="toolbar-group"><label class="search-field"><input class="form-control" id="supplierPartSearch" type="search" placeholder="Search part, supplier, tier" /></label></div><div class="toolbar-group" id="supplierPartFilters"></div><div class="toolbar-group">${tableExportLinks("supplier-parts")}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div></div>
      <div id="supplierPartTable"></div>
    </section>
    <section class="panel table-panel">${panelHeader("Supplier register")}<div class="table-toolbar"><label class="search-field"><input class="form-control" id="supplierSearch" type="search" placeholder="Search supplier" /></label><div id="supplierFilters"></div>${tableExportLinks("suppliers")}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div><div id="supplierRegisterTable">${table(["Code", "Supplier", "Category", "Approval", "Action"], supplierRows, "No suppliers entered.")}</div></section>
    <section class="panel table-panel">${panelHeader("Supplier-to-supplier material transfer", "Record transferable supplier-held surplus. The receiving supplier is credited in projected coverage in the transfer month.")}
      <form id="transferForm" class="panel-body form-grid">
        <label class="field"><span>Search all materials</span><input class="form-control" id="transferMaterialSearch" type="search" placeholder="Type a material code or name" /></label>
        <label class="field"><span>Material</span><select class="form-control" name="material_id" id="transferMaterialSelect" required>${materialOptions}</select></label>
        <label class="field"><span>From supplier</span><select class="form-control" name="from_supplier_id" required>${supplierOptions}</select></label>
        <label class="field"><span>To supplier</span><select class="form-control" name="to_supplier_id" required>${supplierOptions}</select></label>
        ${field("Transfer quantity", "quantity", "number", { required: true, min: 0.001, step: "any" })}
        ${field("Transfer/arrival month", "transfer_month", "month", { required: true, value: monthOffset(monthNow(), 1) })}
        <label class="field"><span>Workspace</span>${workspaceSelect()}</label>
        ${field("Transfer note", "note", "text")}
        <button class="button button-primary" type="submit">Record transfer</button>
      </form>
    </section>
    <section id="transferTable" class="panel table-panel">${panelHeader("Transfer schedule")}<div class="table-toolbar"><label class="search-field"><input class="form-control" id="transferSearch" type="search" placeholder="Search material or supplier" /></label><div id="transferFilters"></div>${tableExportLinks("transfers", { workspace: state.workspace })}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div></section>`;
  $("#supplierPartForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    try {
      const values = Object.fromEntries(new FormData(event.currentTarget).entries());
      for (const key of ["material_id", "supplier_id", "sob_percent", "lead_time_days", "moq", "buffer_days", "unit_price", "price_interval_months"]) values[key] = Number(values[key]);
      values.parent_supplier_id = values.parent_supplier_id ? Number(values.parent_supplier_id) : null;
      const result = await api("/api/planning/supplier-parts", { method: "POST", body: JSON.stringify(values) });
      toast(`Supplier link saved for ${result.tier}`);
      await renderSupplierControl();
    } catch (error) { toast(error.message, true); }
  });
  bindMaterialSearch("supplierMaterialSearch", "supplierMaterialSelect");
  bindMaterialSearch("transferMaterialSearch", "transferMaterialSelect");
  $("#supplierStockForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    state.workspace = data.get("workspace");
    try {
      const result = await uploadExcel(`/api/planning/supplier-stock/${data.get("report_month")}/upload?workspace=${state.workspace}`, data.get("workbook"));
      toast(`${result.rows} supplier stock row(s) uploaded`);
    } catch (error) { toast(error.message, true); }
  });
  renderSupplierPartsTable(links);
  $("#supplierPartSearch").addEventListener("input", (event) => {
    clearTimeout(state.searchTimer);
    state.searchTimer = setTimeout(() => loadSupplierPartPage(0, event.target.value), 250);
  });
  $("#transferForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const values = Object.fromEntries(new FormData(event.currentTarget).entries());
    for (const key of ["material_id", "from_supplier_id", "to_supplier_id"]) values[key] = Number(values[key]);
    values.quantity = Number(values.quantity);
    state.workspace = values.workspace;
    try {
      await api("/api/planning/transfers", { method: "POST", body: JSON.stringify(values) });
      toast("Supplier-to-supplier transfer saved");
      renderSupplierControl();
    } catch (error) { toast(error.message, true); }
  });
  const transfers = await api(`/api/planning/transfers?workspace=${state.workspace}`);
  const transferRows = transfers.map((transfer) => `<tr><td>${esc(transfer.material_code)} · ${esc(transfer.material_name)}</td><td>${esc(transfer.from_supplier)}</td><td>${esc(transfer.to_supplier)}</td><td>${number(transfer.quantity)}</td><td>${monthLabel(transfer.transfer_month)}</td><td>${esc(transfer.status)}</td></tr>`).join("");
  $("#transferTable").innerHTML = `${panelHeader("Transfer schedule")}<div class="table-toolbar"><label class="search-field"><input class="form-control" id="transferSearch" type="search" placeholder="Search material or supplier" /></label><div id="transferFilters"></div>${tableExportLinks("transfers", { workspace: state.workspace })}<button class="button button-small" data-action="clear-table-filters">Clear all filters</button></div><div id="transferRowsTable">${table(["Material", "From supplier", "To supplier", "Quantity", "Transfer month", "Status"], transferRows, "No transfers recorded.")}</div>`;
  setupTableFilters("#supplierRegisterTable", "#supplierFilters", $("#supplierSearch"), [{ title: "Category", index: 2 }, { title: "Approval", index: 3 }]);
  setupTableFilters("#transferRowsTable", "#transferFilters", $("#transferSearch"), [{ title: "From supplier", index: 1 }, { title: "To supplier", index: 2 }, { title: "Status", index: 5 }]);
  $("#supplierSearch").addEventListener("input", applyTableSearch);
  $("#transferSearch").addEventListener("input", applyTableSearch);
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
  const widths = [...header.parentElement.cells].map((cell) => cell.getBoundingClientRect().width);
  const tableWidth = widths.reduce((total, width) => total + width, 0);
  tableNode.style.tableLayout = "fixed";
  tableNode.style.width = `${tableWidth}px`;
  [...tableNode.rows].forEach((row) => [...row.cells].forEach((cell, index) => {
    cell.style.width = `${widths[index]}px`;
    cell.style.minWidth = `${widths[index]}px`;
  }));
  activeColumnResize = { tableNode, column: header.cellIndex, startX: event.clientX, startWidth: widths[header.cellIndex] };
});

document.addEventListener("pointermove", (event) => {
  if (!activeColumnResize) return;
  const { tableNode, column, startX, startWidth } = activeColumnResize;
  const width = Math.max(70, startWidth + event.clientX - startX);
  [...tableNode.rows].forEach((row) => {
    const cell = row.cells[column];
    if (cell) {
      cell.style.width = `${width}px`;
      cell.style.minWidth = `${width}px`;
    }
  });
});

document.addEventListener("pointerup", () => { activeColumnResize = null; });

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
      <div id="reportTable">${table(["Month", "Material", "Supplier", "Upstream parent", "Tier", "SOB", "Plan demand", "Actual demand", "Own stock", "Own stock allocated", "Supplier stock", "Projected balance", "Coverage", "Shortage", "Recommended PO", "Excess", "Lead time", "Criticality", "Plan revision"], "", "Generate a report to see shortages and excess.")}</div>
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
  const headers = ["Month", "Material", "Supplier", "Upstream parent", "Tier", "SOB", "Plan demand", "Actual demand", "Own stock", "Own stock allocated", "Supplier stock", "Projected balance", "Coverage", "Shortage", "Recommended PO", "Excess", "Lead time", "Criticality", "Plan revision"];
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
        <label class="field"><span>Supplier and material</span><select class="form-control" name="supplier_part_id" id="poSupplierPartSelect" required>${links.map((link) => `<option value="${link.id}">${esc(link.material_code)} · ${esc(link.supplier_name)} · ${esc(link.tier)}${link.unit_price === undefined ? "" : ` · ${money(link.unit_price, link.currency)}`}</option>`).join("")}</select></label>
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
    form = `<form id="modalForm" class="form-grid">${field("Supplier name", "name", "text", { required: true })}${field("Category", "category", "text", { required: true })}${field("Country", "country", "text", { required: true })}${field("Contact person", "contact_name")}${field("Email", "email", "email")}${field("Phone", "phone")}<div class="modal-footer field full"><button class="button button-primary">Add supplier</button></div></form>`;
  }
  $("#modalTitle").textContent = title;
  $("#modalBody").innerHTML = form;
  modalLayer.hidden = false;
  if (type === "material") {
    const choice = $("#categorySelect");
    const fresh = $("#categoryNew");
    const sync = () => {
      const adding = choice.value === "__new__";
      fresh.hidden = !adding;
      fresh.required = adding;
      if (adding) fresh.focus();
    };
    choice.addEventListener("change", sync);
    sync();
  }
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
    const values = Object.fromEntries(new FormData(event.currentTarget).entries());
    try {
      let result;
      if (type === "product") {
        values.model_ids = [...new FormData(event.currentTarget).getAll("model_ids")].map(Number);
        result = await api(record.id ? `/api/planning/products/${record.id}` : "/api/planning/products", { method: record.id ? "PATCH" : "POST", body: JSON.stringify(values) });
      }
      if (type === "model") {
        values.product_ids = [...new FormData(event.currentTarget).getAll("product_ids")].map(Number);
        result = await api(record.id ? `/api/planning/models/${record.id}` : "/api/planning/models", { method: record.id ? "PATCH" : "POST", body: JSON.stringify(values) });
      }
      if (type === "material") {
        values.category = values.category_choice === "__new__" ? (values.category_new || "").trim() : values.category_choice;
        const known = state.materialCategories.find((name) => name.toLowerCase() === values.category.toLowerCase());
        if (known) values.category = known;
        delete values.category_choice;
        delete values.category_new;
        for (const key of ["stock_on_hand", "buffer_days", "unit_cost"]) values[key] = Number(values[key] || 0);
        result = await api(record.id ? `/api/planning/materials/${record.id}` : "/api/planning/materials", { method: record.id ? "PATCH" : "POST", body: JSON.stringify(values) });
      }
      if (type === "bom") {
        values.model_id = Number(values.model_id);
        values.material_id = Number(values.material_id);
        values.quantity = Number(values.quantity);
        values.uom = event.currentTarget.elements.uom.value;
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
      if (type === "supplier") result = await api("/api/suppliers", { method: "POST", body: JSON.stringify(values) });
      if (type === "product" || type === "model" || type === "material" || type === "bom" || type === "component") await loadCoreData();
      modalLayer.hidden = true;
      toast(`${title} saved`);
      await renderPage();
    } catch (error) { toast(error.message, true); }
  });
}

document.addEventListener("click", async (event) => {
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
  if (event.target.closest("[data-execute-master-search]")) return applyTableSearch({ target: $("#masterSearch") });
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
    const record = state.materials.find((item) => item.id === Number(materialEdit.dataset.editMaterial));
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

document.addEventListener("keydown", (event) => {
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
