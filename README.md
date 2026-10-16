# SP-API.NET

VB.NET Windows Forms Amazon SP-API workbench for Visual Studio 2019 and .NET Framework 4.7.2.

## Run
1. Open `SP-API.NET.sln` in Visual Studio 2019.
2. Install **.NET desktop development** and the **.NET Framework 4.7.2 targeting pack**, if not installed.
3. Press **F5**.

No Node.js, npm, web server, or third-party NuGet packages.

## Three source files

- `MainForm.vb` — desktop UI/UX and readable **View** for all operations, including product details, record tables, fees, inventory, orders, documents, report/feed/inbound results; plus Summary and Raw response tabs.
- `ApiRequests.vb` — login/LWA, SP-API HTTP requests, Sandbox/Production behavior, validations, pagination, retries, reports/feeds/inbound and documents.
- `Database.vb` — isolated SQL Server catalogue-to-`InventoryDimensions` mapping, read-only connection checks, previews, and confirmed insert/update.

The .sln and .vbproj files are Visual Studio scaffolding, not extra code files.

## Everyday use

Enter Client ID, Client Secret and Refresh Token; select Sandbox or Production and marketplace; click **Test connection**. Choose an operation and run it. See the readable **View** tab by default. Summary and Raw response remain available if needed. Sandbox examples are populated only by explicitly clicking **Load Sandbox example**. Live write operations require confirmation.

## SQL Server inventory updates

Joel has confirmed that the existing `InventoryDimensions` table is stored in **SQL Server**, not directly in Access. The original VB.NET method used `SqlConnection` and the application setting `FTSSQL`. The SQL Server connection for both reads and writes now lives entirely in `Database.vb`.

For a **Production** Amazon catalogue result:
1. Enter the existing SQL Server connection setting (such as `FTSSQL`) in the **SQL Server (FTSSQL)** field in the View footer. The field is masked and the app does not save the setting. **Never share database passwords or connection strings publicly.**
2. Click **Test database**. This checks the `InventoryDimensions` columns and read access without changing any records.
3. Select one returned ASIN and click **Preview / Save selected ASIN**.
4. Review whether this will INSERT or UPDATE the ASIN, inspect the values/units, and confirm the write.
5. The application reads the saved row back and shows its stored values.

Expected columns (spelling maintained from the original function): `Asin`, `Weight`, `Height`, `Length`, `Width`, `LastUpdated`, `ProductCatagory`, `ImageURL`. The application does not create or modify the database schema. Existing rows are updated using ASIN, otherwise a new row is inserted. Missing Amazon values do not overwrite existing values. Parameterized SQL and a transaction protect each update.

**Back up the SQL Server database before the first write.** A valid connection string and sufficient permissions are still necessary. Category identifiers from Amazon Catalog Items 2022-04-01 may not equal the older SDK's ProductCategoryId, and the units must match Joel's intended stored units. Validate with real, noncritical data before deployment.

Sandbox catalogue data cannot be saved, even after switching to Production. The older unrelated **Convert Amazon_US** and **SKU / FC Bulk Update** utilities remain disconnected.

## Verification

GitHub Actions validates the three-file layout, builds Debug and Release on Windows, and runs the built-in offline WinForms self-test (including API request validation and Sandbox/Production endpoint checks).

It also launches an isolated SQL Server LocalDB instance, creates a **temporary test database**, and tests actual `InventoryDimensions` inserts, updates, read-back, partial-field preservation, apostrophes in string data, transaction rollback on SQL errors, invalid ASIN rejection, numeric column types, and missing-column detection. The temporary database is deleted at the end of the test. No access to Joel's server is used.

**Not yet verified against Joel's environment:** his actual SQL Server connection and permissions; the real `InventoryDimensions` schema and constraints; the desired dimension measurement units and legacy category ID mapping; and live Amazon seller authorization. Use a backup or test database for the first end-to-end integration check.
