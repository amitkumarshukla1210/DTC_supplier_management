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

- Maintain masters in workflow order: Product Master first, Model BOM Level 1 and the Level 2/3 Part BOM tabs next, and Material Master last. Product Master displays each assigned model on its own row. Create models once in Material Master, define each material code and its unit there, then assign one or more catalog models to a product. Model codes and material codes are checked case-insensitively for duplicates. Product, model, material, and BOM rows can be edited; deletes are blocked where planning history or active references depend on the record. Edit and Delete actions have separate columns, and deletes ask for confirmation. Clicking any non-action cell of a Product Master row with a BOM opens that model's Level 1 BOM; clicking a Level 1 or Level 2 BOM row with child BOM lines opens the next level with that model/parent selected and its BOM loaded. Editing a model-specific product row preserves its other model links. Use the plus controls to expand Product Master into Level 1, Level 1 into Level 2, and Level 2 into Level 3 only when that row has child BOM lines; clicking a row navigates only when the next BOM level exists. BOM tabs use one table per page with a code and name column pair for each level (Product, Model, L1, L2, L3), so expanded rows appear directly under their parent and every item of a level lines up in the same columns. Each row repeats its parents' codes and names; rows that have lines under them are bold. Download Excel on these tabs gives the full tree breakdown (honouring the selected models / parent parts) with Excel outline +/- buttons. The Excel file opens collapsed (top rows only, each with +) and the +/- buttons work straight away. Download PDF gives the same rows with parent rows in bold, every row on one line (column widths fit the text; the font shrinks or the page widens for very wide tables). Row highlight and row click stop before the Edit and Delete columns. On these tabs, drag a column header's edge to change its width; widths are remembered in this browser per tab, and double-clicking the edge resets them. Level 1 uses a searchable multiple-model selector; Level 2 and Level 3 select their parent by part (Level 1 parent part for Level 2, Level 2 parent part for Level 3), without a model selector. With no model or parent filter selected, all BOMs are shown by default. Level 2/3 parent parts can be searched and multi-selected. Level 1 is linked to a model; Level 2 is linked to a Level 1 parent part, and Level 3 to a Level 2 parent part. Add each purchased or in-house-made child part from Material Master; Level 2/3 BOM entry provides searchable part-code and part-name dropdowns that stay synchronized. Then build deeper operations by adding that part's own BOM at the next level. Level 2/3 quantities may be positive for consumed components or negative for scrap/by-products; zero is not allowed. Part BOM links are shared across models using the same parent part.
- Every table header opens its searchable multi-select filter; choose values and Apply, or clear that column's choices. Table columns can be resized by dragging the header edge, and table-level clear-filter controls remove active column and text filters. The model picker closes after Apply. BOM tabs provide matching Excel templates and uploads, and quantities are rolled up per parent to material demand.
- BOM controls include model/parent selection, workbook upload, level template, exports, filters, and add-line actions. Level 2/3 controls wrap and their tables fit the available width without horizontal scrolling; Level 1 keeps its single-line toolbar. The workbook field uses a compact accessible label rather than a separate heading.
- Upload a 24-month plan for an issue month using month columns, or upload the long-form model/month/quantity layout. Month headers accept Excel dates and common month labels such as `Apr-26`. Each subsequent upload or edit creates a preserved `Rev.00`, `Rev.01`, and later revision with its saved date; click a revision row to edit monthly quantities and save them as a new revision.
- Upload actual production and OEM stock by month, and upload company stock and supplier stock separately. Supplier stock workbooks support supplier/material/tier/parent identification, available quantity, and arrival month.
- Configure supplier SOB, tier/route, lead time, MOQ, buffer days, price, currency, and spot/monthly/N-month price terms. The Suppliers page has one tab per area (relationship, stock upload, coverage, register, transfer, transfer schedule) with adjustable, remembered column widths. Only purchased parts can get a supplier: parts with an in-house BOM (made in your factory) or used with a negative BOM quantity (generated in your factory) are excluded from the material pickers and refused by the server. Material pickers search the server, so they work with 10,000+ codes. Relationships can be uploaded from Excel, either the template or your own sheet as it is (heading row found automatically; columns such as Material code, Supplier Name, Unit cost, Currency, Buffer days recognised; SOB defaults to 100%; other columns ignored): one row per material and supplier; new supplier names are added to the register, existing relationships are updated, and the whole file is rejected if any row is wrong. The BOM Part price is the SOB-weighted price of the code's Direct suppliers (sum of price x SOB / sum of SOB); it falls back to the Material Master price when no Direct supplier has a price or when suppliers quote in different currencies. Hover the price to see its source.
- Review monthly material and supplier demand, stock, projected balance, coverage, shortage, MOQ-sized PO recommendation, excess, lead-time criticality, and authorized price values.
- Issue planning POs with current price captured on the order; record partial receipts and advance/postpone the expected arrival month with a reason. Plan material transfers between suppliers.
- Run historical supplier-business reports over a selected period, split by supplier tier. Excel and PDF reports carry Design to Cost and Supply Chain Management branding.
- Use the **Trial** workspace for imports and analysis before entering **Production** data. Admins can clear trial plans, actuals, snapshots, POs, and transfers without deleting production or master data.

Planning workbooks can be downloaded from the relevant page. API documentation is available at [http://127.0.0.1:8000/docs](http://127.0.0.1:8000/docs).

## T1 / T2 / T3 Material Masters

- **T1 Material Master** (Product Master tab): codes your company uses directly or indirectly and buys from T1 suppliers. Only T1 codes are used in your BOMs, supplier relationships and planning.
- **T2 / T3 Material Master** tabs: codes used further down the supply chain (by your suppliers' suppliers): code, name, category, unit, reference price. Form, edit/delete, search, Excel upload/download.
- A code can be in only one of T1, T2 or T3 (checked on every save and upload, case-insensitive).

## Alternate codes

**Product Master → Alternate codes** holds your own alternate code for a part code you buy (e.g. `FL001MICROCHIPNEO` for `FL001MICROCHIP`), with the alternate's SOB %. It applies everywhere the main code is used:

- **BOM price:** main share × main price + each alternate's SOB × its own price (each code's own price is its SOB-weighted Direct supplier price, or the Material Master price). If the codes are priced in different currencies, the main code's price is shown.
- **BOM tree (screen, Excel, PDF):** the main code row shows its own SOB share of the quantity ("main 40% of 4" → 1.6), and each alternate follows it with its share and its own price; main + alternates = the BOM quantity. Edit still changes the full BOM quantity.
- **Planning:** Shortage & excess and the PO plan split the main code's demand: the alternate takes its SOB %, the main code keeps the rest (100 − all its alternates' SOB).

Rules: both codes are in T1 Material Master, are bought parts (no in-house BOM, no negative BOM quantity) and use the same unit; an alternate belongs to one main code and has no alternates of its own; the alternates of a code add up to at most 100%. A code in Alternate codes cannot be deleted from T1 Material Master. Excel upload accepts your own headings (Main code / Part code, Alternate code, SOB %); all rows or none.

## Supplier materials and supplier BOM (cost breakdown)

- **Supplier materials** tab: each supplier's material master. An entry is either picked from your T1, T2 or T3 Material Master (code and name follow your master) or the supplier's own code. Each entry carries the supplier's own price (may differ from yours), lead time and buffer stock. Excel upload links codes found in T1/T2/T3 automatically; other codes become the supplier's own. Supplier BOM lines and alternate codes use only codes from that supplier's material master.
- **Supplier BOM** tab: for a part you buy from a supplier, the supplier's BOM down to L3. Lines use your Material Master codes or that supplier's own codes. Line cost = quantity × unit price; a line with lines below costs their sum. A line's unit price is the SOB-weighted price of its sources (sub-suppliers from the Supplier register, optionally under an alternate code), else the price entered on the line, else the material list price. The top shows should-cost, the supplier's quoted price and the gap. Lines priced in another currency are listed and not added. Excel upload (Parent code, Code, Quantity, Unit price, Source) and Excel/PDF download. Cost view only; planning is not affected.

## Scale

Tested with 1,000 suppliers, 10,000 materials, 1,000 models (20,000 Model BOM lines) and 9,400 supplier links: every screen opens in about a second. Model BOM Level 1 shows 50 models a page (one request), Material Master shows 500 materials a page with server search, and material pickers search the server. The load-test kit is in `tests/load/`.

## Access and deployment

Admin users manage accounts and have full access. HOD, Sourcing, and Supply Chain roles can manage supplier controls and planning reports. Users can be individually granted part-price visibility; report exports also enforce that permission.

This is a local-first MVP, not a hardened production deployment. Before exposing it to a network, configure secure database credentials, backups, HTTPS, operational monitoring, and a production migration strategy. Large masters and reports use API pagination, but report calculation and exports should be load-tested against the organization's expected production data volume.

## Verify

```powershell
python -m unittest discover -s tests -v
node --check static/app.js
python -m py_compile app.py planning.py
```
