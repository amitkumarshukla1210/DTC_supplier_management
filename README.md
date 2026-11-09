# Design to Cost — Supply Chain Management

A local-first monthly demand and supply planning workspace for OEMs and their supplier networks. The application is named **Supply Chain Management** and carries the **Design to Cost** brand. It manages model forecasts, actual production, nested BOMs, supplier stock and share of business (SOB), lead times through T1/T2/Tn, purchase-order schedules, transfers, and shortage/excess analysis.

The application uses **Python, FastAPI, SQLite, SQLAlchemy, and a plain HTML/CSS/JavaScript frontend**. The DTC logo is served locally from `static/logo.png`.

## Run locally

Python 3.10+ is recommended.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app:app --reload
```

Open [http://127.0.0.1:8000](http://127.0.0.1:8000). On first launch, create the initial Admin account from the sign-in page. The Admin can add additional Admins and assign roles and price-view permissions. The default database is `data/suppliers.db`; set `SUPPLIER_DATABASE_URL` to use another SQLAlchemy database URL.

On Windows, you can instead double-click `start_app.bat` in the project folder. It creates the virtual environment if needed, installs the requirements, starts the server in a separate window, and opens the app in your browser. Keep the server window open while working; close it to stop the server.

If port 8000 is already occupied when using `start_app.bat`, the launcher starts this project on port 8001 so an older running copy is not mistaken for the updated one.

## Monthly planning

- Maintain masters in workflow order: Product Master first, Model BOM Level 1 and the Level 2/3 Part BOM tabs next, and Material Master last. Product Master displays each assigned model on its own row. Create models once in Material Master, define each material code and its unit there, then assign one or more catalog models to a product. Model codes and material codes are checked case-insensitively for duplicates. Product, model, material, and BOM rows can be edited; deletes are blocked where planning history or active references depend on the record. Edit and Delete actions have separate columns, and deletes ask for confirmation. Clicking any non-action cell of a Product Master row with a BOM opens that model's Level 1 BOM; clicking a Level 1 or Level 2 BOM row with child BOM lines opens the next level with that model/parent selected and its BOM loaded. Editing a model-specific product row preserves its other model links. Use the plus controls to expand Product Master into Level 1, Level 1 into Level 2, and Level 2 into Level 3 only when that row has child BOM lines; clicking a row navigates only when the next BOM level exists. BOM tabs use one table per page with a code and name column pair for each level (Product, Model, L1, L2, L3), so expanded rows appear directly under their parent and every item of a level lines up in the same columns. Row highlight and row click stop before the Edit and Delete columns. On these tabs, drag a column header's edge to change its width; widths are remembered in this browser per tab, and double-clicking the edge resets them. Level 1 uses a searchable multiple-model selector; Level 2 and Level 3 select their parent by part (Level 1 parent part for Level 2, Level 2 parent part for Level 3), without a model selector. With no model or parent filter selected, all BOMs are shown by default. Level 2/3 parent parts can be searched and multi-selected. Level 1 is linked to a model; Level 2 is linked to a Level 1 parent part, and Level 3 to a Level 2 parent part. Add each purchased or in-house-made child part from Material Master; Level 2/3 BOM entry provides searchable part-code and part-name dropdowns that stay synchronized. Then build deeper operations by adding that part's own BOM at the next level. Level 2/3 quantities may be positive for consumed components or negative for scrap/by-products; zero is not allowed. Part BOM links are shared across models using the same parent part.
- Every table header opens its searchable multi-select filter; choose values and Apply, or clear that column's choices. Table columns can be resized by dragging the header edge, and table-level clear-filter controls remove active column and text filters. The model picker closes after Apply. BOM tabs provide matching Excel templates and uploads, and quantities are rolled up per parent to material demand.
- BOM controls include model/parent selection, workbook upload, level template, exports, filters, and add-line actions. Level 2/3 controls wrap and their tables fit the available width without horizontal scrolling; Level 1 keeps its single-line toolbar. The workbook field uses a compact accessible label rather than a separate heading.
- Upload a 24-month plan for an issue month using month columns, or upload the long-form model/month/quantity layout. Month headers accept Excel dates and common month labels such as `Apr-26`. Each subsequent upload or edit creates a preserved `Rev.00`, `Rev.01`, and later revision with its saved date; click a revision row to edit monthly quantities and save them as a new revision.
- Upload actual production and OEM stock by month, and upload company stock and supplier stock separately. Supplier stock workbooks support supplier/material/tier/parent identification, available quantity, and arrival month.
- Configure supplier SOB, tier/route, lead time, MOQ, buffer days, price, currency, and spot/monthly/N-month price terms.
- Review monthly material and supplier demand, stock, projected balance, coverage, shortage, MOQ-sized PO recommendation, excess, lead-time criticality, and authorized price values.
- Issue planning POs with current price captured on the order; record partial receipts and advance/postpone the expected arrival month with a reason. Plan material transfers between suppliers.
- Run historical supplier-business reports over a selected period, split by supplier tier. Excel and PDF reports carry Design to Cost and Supply Chain Management branding.
- Use the **Trial** workspace for imports and analysis before entering **Production** data. Admins can clear trial plans, actuals, snapshots, POs, and transfers without deleting production or master data.

Planning workbooks can be downloaded from the relevant page. API documentation is available at [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs).

## Access and deployment

Admin users manage accounts and have full access. HOD, Sourcing, and Supply Chain roles can manage supplier controls and planning reports. Users can be individually granted part-price visibility; report exports also enforce that permission.

This is a local-first MVP, not a hardened production deployment. Before exposing it to a network, configure secure database credentials, backups, HTTPS, operational monitoring, and a production migration strategy. Large masters and reports use API pagination, but report calculation and exports should be load-tested against the organization's expected production data volume.

## Verify

```powershell
python -m unittest discover -s tests -v
node --check static/app.js
python -m py_compile app.py planning.py
```
