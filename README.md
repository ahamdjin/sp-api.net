# SP-API.NET

VB.NET Windows Forms Amazon SP-API workbench for Visual Studio 2019 and .NET Framework 4.7.2.

## Run
1. Open `SP-API.NET.sln` in Visual Studio 2019.
2. Install **.NET desktop development** and the **.NET Framework 4.7.2 targeting pack**, if not installed.
3. Press **F5**.

No Node.js, npm, web server, or third-party NuGet packages.

## Source files and where to make changes

- `MainForm.vb` — desktop UI/UX and readable **View** for all operations, including product details, record tables, fees, inventory, orders, documents, report/feed/inbound results; plus Summary and Raw response tabs.
- `ApiRequests.vb` — login/LWA, SP-API HTTP requests, Sandbox/Production behavior, validations, pagination, retries, reports/feeds/inbound and documents.
- `Database.vb` — isolated SQL Server catalogue-to-`InventoryDimensions` mapping, read-only connection checks, previews, and confirmed insert/update.
- `ProductionApi.vb` — live endpoint and Production-only request rules.
- `SandboxApi.vb` — test endpoint and Sandbox-only request rules. Shared requests are not duplicated.

In Visual Studio, press **Ctrl+M, Ctrl+O** to collapse code sections. Expand the named section you need. `MainForm.vb` keeps offline tests at the bottom, away from startup and ordinary editing.

| What you want to change | Where to look |
|---|---|
| Credentials or SQL connection | Project Properties > Settings |
| Countries or operation names | MainForm.vb, section 01 |
| Default request values | MainForm.vb, section 04 |
| Layout, buttons, images | MainForm.vb, section 05 |
| Fields for a particular request | MainForm.vb, section 06, BuildOperationFields |
| Product/result presentation | MainForm.vb, sections 09–11 |
| Database connection and save experience | MainForm.vb, section 12 |
| Amazon operation URL/body/parameters | ApiRequests.vb, sections 02–06 |
| Production or Sandbox rules | ProductionApi.vb or SandboxApi.vb |
| Which Amazon values go into SQL columns | Database.vb, MapCatalogItem |
| Save transaction and verification | Database.vb, section 03 |

Database changes follow `MapCatalogItem` → `Preview` → `Save` → read-back verification. When adding a stored field, update the supported column lists, mapping, expected schema, and tests together. Keep SQL parameterized and keep verification inside the transaction. Weight/dimensions retain Amazon's supplied units; changing stored units requires an explicit conversion.

The .sln and .vbproj files are Visual Studio scaffolding, not extra code files.

## Everyday use

Project Properties > Settings includes `FTSSQL`, `ClientId`, `ClientSecret`, and `RefreshToken`. The app loads these defaults at startup. Catalogue item, United States, and Production are the startup defaults. Amazon credential defaults are blank. Settings are plain text: keep real credentials local and never commit them to GitHub. Editing fields in the running app changes only that session. Deploy `SP-API.NET.exe.config` alongside the executable to retain configured settings.

Enter Client ID, Client Secret and Refresh Token; select Sandbox or Production and marketplace; click **Test connection**. Choose an operation and run it. See the readable **View** tab by default. Summary and Raw response remain available if needed. Sandbox examples are populated only by explicitly clicking **Load Sandbox example**. Live write operations require confirmation.

Complete credentials connect automatically after typing stops and when the app opens with saved settings. A green Connected status unlocks requests. Empty or changed credentials immediately lock requests and database saves until verification succeeds again. Test connection remains available for a manual retry. Unavailable legacy utilities are hidden from navigation. Database activity appears in the Database tab without connection-success popups.

## SQL Server inventory updates

Joel has confirmed that the existing `InventoryDimensions` table is stored in **SQL Server**, not directly in Access. The original VB.NET method used `SqlConnection` and the application setting `FTSSQL`. The SQL Server connection for both reads and writes now lives entirely in `Database.vb`.

For a **Production** Amazon catalogue result:
1. Set `FTSSQL` in Project Properties > Settings, or enter a session override in the **SQL Server (FTSSQL)** field in the View footer. The default points to the local `SPAPI_LocalTest` database using Windows Authentication. **Never share database passwords or connection strings publicly.**
2. Click **Test database**. This checks the `InventoryDimensions` columns and read access without changing any records.
3. Select one returned ASIN and click **Preview / Save selected ASIN**.
4. Review whether this will INSERT or UPDATE the ASIN, inspect the values/units, and confirm the write.
5. The application reads the saved row back and shows its stored values.

Expected columns (spelling maintained from the original function): `Asin`, `Weight`, `Height`, `Length`, `Width`, `LastUpdated`, `ProductCatagory`, `ImageURL`. The application does not create or modify the database schema. Existing rows are updated using ASIN, otherwise a new row is inserted. Missing Amazon values do not overwrite existing values. Parameterized SQL and a transaction protect each update.

**Back up the SQL Server database before the first write.** A valid connection string and sufficient permissions are still necessary. Category identifiers from Amazon Catalog Items 2022-04-01 may not equal the older SDK's ProductCategoryId, and the units must match Joel's intended stored units. Validate with real, noncritical data before deployment.

Sandbox database writes are disabled for all databases, including the local test database. Saving requires a Production catalogue result and Production selected; switching environments does not make an old Sandbox result eligible. Connection checks are read-only and remain available in either environment. The Database tab shows connection details, the save preview, and the verified save result, with one save confirmation and no success popups. The older unrelated **Convert Amazon_US** and **SKU / FC Bulk Update** utilities remain disconnected.

## Verification

GitHub Actions validates the five application source files, builds Debug and Release on Windows, and runs the built-in offline WinForms self-test (including API request validation and Sandbox/Production endpoint checks).

It also launches an isolated SQL Server LocalDB instance, creates a **temporary test database**, and tests actual `InventoryDimensions` inserts, updates, read-back, partial-field preservation, apostrophes in string data, transaction rollback on SQL errors, invalid ASIN rejection, numeric column types, and missing-column detection. The temporary database is deleted at the end of the test. No access to Joel's server is used.

**Not yet verified against Joel's environment:** his actual SQL Server connection and permissions; the real `InventoryDimensions` schema and constraints; the desired dimension measurement units and legacy category ID mapping; and live Amazon seller authorization. Use a backup or test database for the first end-to-end integration check.
