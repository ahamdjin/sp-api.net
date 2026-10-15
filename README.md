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
- `Database.vb` — isolated Microsoft Access catalogue-to-`InventoryDimensions` mapping, read-only connection checks, previews, and confirmed insert/update.

The .sln and .vbproj files are Visual Studio scaffolding, not extra code files.

## Everyday use

Enter Client ID, Client Secret and Refresh Token; select Sandbox or Production and marketplace; click **Test connection**. Choose an operation and run it. See the readable **View** tab by default. Summary and Raw response remain available if needed. Sandbox examples are populated only by explicitly clicking **Load Sandbox example**. Live write operations require confirmation.

## Microsoft Access

Microsoft Access `(.accdb/.mdb)` is the current working assumption. The old `SqlConnection` code may indicate that some Access tables are linked to SQL Server, so the actual storage and permissions must be checked on Joel's installation.

For a **Production** catalogue result:
1. In the View footer, choose the existing Access file.
2. Click **Test Access**. This performs only read and schema checks on `InventoryDimensions`.
3. Select a returned ASIN. Click **Preview / Save selected ASIN**.
4. Review the proposed INSERT or UPDATE, including existing values and units. Confirm explicitly to write a single record.

The table must already exist with the eight fields `Asin`, `Weight`, `Height`, `Length`, `Width`, `LastUpdated`, `ProductCatagory`, and `ImageURL`. The app does not create or alter tables, and missing Amazon fields do not overwrite existing values. Writes use parameterized queries within a transaction. Sandbox catalogue results cannot be written, including after changing the environment dropdown. Database file paths and Amazon credentials are not persisted by the app.

**Before the first real write, make a backup of the database.** Ensure a compatible ACE OLE DB provider is installed for the app's 32-/64-bit process. A linked SQL Server table may require its existing driver, connection configuration, and permissions. The field units are shown, not converted; the newer classification ID may differ from the older SDK's product category ID. Verify both with Joel before production updates.

The two older legacy database utilities remain disconnected; they are separate from the new `InventoryDimensions` save. The Amazon SP-API itself does **not** require Access.

## Verification

GitHub Actions builds Debug and Release on Windows and runs the built-in offline WinForms self-test, including all result-view operations and Sandbox/Production endpoint checks. Live Amazon seller authorization and actual reads/writes against a copy of Joel's Microsoft Access database must be tested separately. CI does not contain a real Access database or a live Amazon account.
