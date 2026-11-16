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

## Tables

Every table's column filter has **Sort A → Z / Z → A** (numbers such as "₹1,800", "40%" or "14 days" sort as numbers; the sorted column is marked ▲ / ▼; in BOM tree tables each parent keeps its lower-level rows under it). Paged tables sort the rows shown on the page. **Add** buttons (Add supplier, Add line, Add material, Add T2/T3 material, + Add row / Save all) sit right above their table, at the right-hand end. Upload pages use one panel: heading, upload row (workbook, Upload, Download template), then + Add row / Save all and the typing table.

## T1 Material Master: Used in

The **Used in** column (right after Material name, also in the Excel / PDF download) shows the BOM levels a code is a line of: **L1** (Model BOM Level 1), **L2** (a Level 2 Part BOM, under an L1 part), **L3** (a Level 3 Part BOM, under an L2 part), or combinations such as **L1 & L2**, **L2 & L3**, **L1 & L3**, **L1 & L2 & L3**. An **alternate code** counts as used at the level of each BOM line where it has an SOB above 0% (in any month) and shows e.g. **L3 (alternate)**. A code used in no BOM (and not an alternate above 0% anywhere) shows **Not used** and its row is red.

## T1 Material Master: delete and Move to T2/T3

- **Delete** says exactly what still uses a code (e.g. "used in 1 T1 supplier/material relationship, 2 part BOM lines"). Part BOMs not linked to any model also count. Material requests of the retired stock screens (install sample) no longer block a delete.
- **Move to T2/T3** (per row, shown only for codes in no BOM (L1/L2/L3) and not in Alternate codes): a code put in T1 by mistake moves to T2 or T3 Material Master with its name, category, unit and price. Supplier material-master entries and T2/T3 relationships switch to the moved code. Anything that needs a T1 code (your BOM lines, its own part BOM, T1 relationships, supplier BOMs, Alternate codes, own stock, transfers) blocks the move and is listed.

## Category list

**Product Master → Category** holds the one category list used for parts (T1, T2, T3) and suppliers; every category field is a dropdown from it with **+ Add new category**. Add, rename (every part and supplier using it follows) or delete (refused while used). A new category typed in a dropdown or found in an Excel upload is added to the list automatically, written as in the list when only the capitals differ ("steel raw MATERIAL" → "Steel Raw material").

Starting list: Raw material, Finished assembly, Electrical components, Electronic PCB Assy, Sheet Metal, Wire Harness, Plastic Raw material, Steel Raw material, Electronic Child Part, Casting, Machining, Packing Material, Fasteners, Subassembly. When the list is first set up, old names are replaced once: CNC machining → Machining, Metal fabrication → Sheet Metal, Plastics → Plastic Raw material. Any other category still used by records but not in the list is shown under the table, to **Replace** with a listed category or **Add to my list**.

## Alternate codes

**Product Master → Alternate codes** lists which code is an alternate of which (e.g. `FL101STEEL0.8` for `FL001STEEL0.8`): no SOB there. Excel upload accepts your own headings (Main code / Part code, Alternate code); all rows or none.

**SOB is managed only in Alternate codes → Manage SOB** (it opens by itself after adding a pair). It shows the main code's uses in three panels: **L1** (direct use in each model's Model BOM Level 1), **L2** and **L3** (part BOMs), each use with its BOM quantity, SOB this month, next change and history. Enter the alternates' SOB and the **SOB valid from month** (empty = all months) once; it is saved on all ticked uses (each panel has a tick-all box; untick a use or a whole level to keep its own SOB). **Remove this month's change from ticked uses** undoes a change. Every line starts at 0% (main code 100%). The BOM line window shows the line's SOB and history read-only, with a button to Manage SOB.

- **BOM tree (screen, Excel, PDF):** the main code row shows its own share of the quantity ("main 20% of 2" → 0.4) and each alternate follows it with its share and its own price; main + alternates = the BOM quantity. Edit still changes the full BOM quantity.
- **BOM price of the line:** main share × main price + each alternate's SOB × its own price.
- **Planning:** Shortage & excess and the PO plan split each line's demand by that line's SOB.

**SOB changes by month (history):** every change keeps its from month; earlier months are planned on the SOB valid then, later months on the new one. A **code replacement** is an SOB change: add the new code in Alternate codes, then give it 100% from the month (main code 0%). The BOM tree shows this month's SOB and the next change ("main 40% → 0% from Jan 2027").

**SOB change log** (Product Master tab, also **Download SOB history** on Alternate codes): every SOB set or removed, one row per use: changed on, changed by, action, main code, level, used in, from month and the split. Search, Excel and PDF download.

**Shortage & excess** gets an **SOB change** column (also in the Excel report): from the month of the latest SOB change that moved a code's demand, it shows that month and the demand change against the SOB before it (e.g. "from Nov 2026, −100 PCS" on the main code, "+100 PCS" on the alternate), so you see which code goes into shortage and which into excess because of the change.

Rules: both codes are in T1 Material Master, are bought parts and use the same unit; an alternate belongs to one main code and has no alternates of its own; a line's alternates add up to at most 100%. Deleting a pair or a BOM line removes its SOB. A code in Alternate codes cannot be deleted from T1 Material Master.

## Suppliers page

Tabs, in order: **Supplier register** (first), Supplier/material relationship, Supplier/material coverage, Upload supplier stock, Supplier-to-supplier material transfer, Transfer schedule, Supplier materials, Supplier BOM. Every table's column widths can be dragged and are remembered per tab.

- **Supplier register:** code, name, category, country, contact, email, phone, approval; Action = Approve (pending), **Edit** (including the supplier code, which must be unique), **Delete**. **Upload suppliers** in bulk from Excel (template available; your own headings are fine): needed Supplier name; optional Supplier code (blank = next free SUP- code), Category (new ones join the list), Country, Contact person, Email, Phone, Approval (new suppliers default to Approved). A row whose code (or, without a code, name) is already in the register updates that supplier; blank cells keep the old values; all rows or none. Delete is refused while the supplier is used (supplier/material relationships, supplier materials / BOM, transfers) and the message says where; set it On hold instead. Old records of the retired RFQ / purchase-order screens (the install sample) do not block a delete and go with the supplier. Sample suppliers and materials are only added to a brand-new database, so deleted ones never come back at start-up.
- **Supplier/material relationship** is one page: the Excel upload (+ template) at the top, then **+ Add row / Save all** and an Excel-like entry table (one row per part + supplier, every field a column; all rows or none, errors name the row; empty rows are skipped). Supplier code and name come first, then part code and name (parent supplier also has a code and a name column): type in either box and pick from the hint list (only purchased parts; capitals ignored, the master's spelling is shown); the other box fills in. Column headings have the usual filter (sort, search, Select all, Clear all) working on the typed values.
- **Supplier/material coverage:** one value per column (material code / name, supplier code / name, tier, parent, SOB, lead time, MOQ, buffer, unit price, currency, price agreement, price status); Action = **Edit** (SOB, lead time, MOQ, buffer, price, currency, agreement) and **Delete** (refused while the relationship has POs or supplier stock).
- **Upload supplier stock:** type stock in an entry table (supplier → its material relationships, stock quantity, available quantity, arrival month; Save all) or upload a workbook; the month's saved stock is listed below.
- **Supplier-to-supplier material transfer:** entry table, one row per transfer, Save all.

## Supplier tiers (T1 / T2 / T3)

- **T1** = supplies you (formerly "Direct"); **T2** = supplies a T1 supplier; **T3** = supplies a T2 supplier. On first start of this version, saved relationships are renamed once (Direct → T1, T1 → T2, T2 → T3) after a copy of the database file is kept (`suppliers.db.before-tier-update-<date>.bak`). "Direct" and "Tier 2" are still accepted in uploads.
- **Supplier register** has three tabs: **T1 / T2 / T3 Supplier Master**. A supplier can be in more than one (e.g. TATA Steel: T1 for you, T2 for Supangita); tick the tiers in Add / Edit supplier. A relationship saved at a tier adds that tier to the supplier (and the tier above to its parent supplier); tiers are never removed automatically. Supplier list uploads join the open tab unless the sheet has a Supplier tier column.
- **Relationships**: tier is a drop-down (T1, T2, T3); T2/T3 rows need the parent supplier and may use a **T1, T2 or T3 Material Master code** (hints are labelled); T1 rows use T1 codes you buy. Your BOM price uses T1 rows only; POs are issued for T1 codes only. Coverage shows **Code from** (T1/T2/T3) and **Edit** opens the whole row (supplier, part, tier, parent, SOB, lead time, MOQ, buffer, price).

## Supplier/material coverage (tree) and typing chains

- **Coverage** is a tree panel for each part you buy: the part (with its number of T1 suppliers and total SOB, red if not 100%), then each T1 supplier, then the T2 suppliers of that T1 with their part (a T1, T2 or T3 Material Master code), then T3. **+ / −** opens and closes a level. Columns are in blocks: **T1** (your part code and name, then supplier code and name), **T2** and **T3** (supplier code and name, then part code and name). Faded cells repeat the level above, so column filters and sort keep the context. Every level has its own **SOB, lead time, MOQ, buffer days, price mode, price status, unit price and currency** (columns at the end), and **Edit** (last column) opens that relationship, with **Delete** inside the Edit window. A T1 relationship with T2 suppliers under it cannot be deleted.
- **+ Add T2** on a T1 row (and **+ Add T3** on a T2 row) adds suppliers under that relationship: the window lists the suppliers already there and the SOB left; type supplier, part and terms (one or more rows). Too much SOB is refused with the % left; lower another supplier's SOB with its Edit first.
- One supplier can supply many parts; one part can have several suppliers, with SOB adding up to 100% per part (and, for T2/T3, per parent supplier's part).
- **Download Excel / PDF** gives the same tree (Level column, T1 / T2 / T3 blocks, then each row's terms; Excel rows are grouped so + / − works there too).
- **Supplier/material relationship** page: type chains in one table row. Your part, its T1 supplier, then (optional) the T2 supplier and its part, then T3; each block has its own SOB, lead time, MOQ, buffer, price mode, interval, unit price and currency. Leave the T1 block blank on the next row to add another T2 under the same T1 (e.g. JGI 30% and TATA 70% steel for Supangita's Base Front Beam). **Save all** creates or updates the relationships (all rows or none; an error names the row).

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
