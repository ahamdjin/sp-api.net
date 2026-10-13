# SP-API.NET

VB.NET Windows Forms Amazon SP-API workbench for Visual Studio 2019 and .NET Framework 4.7.2.

## Run
1. Open `SP-API.NET.sln` in Visual Studio 2019.
2. Install **.NET desktop development** and the **.NET Framework 4.7.2 targeting pack**, if not installed.
3. Press **F5**.

No Node.js, npm, web server, or third-party NuGet packages.

## Two source files

- `MainForm.vb` — desktop UI/UX and readable **View** for all operations, including product details, record tables, fees, inventory, orders, documents, report/feed/inbound results; plus Summary and Raw response tabs.
- `ApiRequests.vb` — login/LWA, SP-API HTTP requests, Sandbox/Production behavior, validations, pagination, retries, reports/feeds/inbound and documents.

The .sln and .vbproj files are Visual Studio scaffolding, not extra code files.

## Everyday use

Enter Client ID, Client Secret and Refresh Token; select Sandbox or Production and marketplace; click **Test connection**. Choose an operation and run it. See the readable **View** tab by default. Summary and Raw response remain available if needed. Sandbox examples are populated only by explicitly clicking **Load Sandbox example**. Live write operations require confirmation.

## Microsoft Access

Joel's legacy database is **Microsoft Access (.accdb/.mdb)**, not SQL Server. The two legacy database utilities remain disconnected; connecting them safely requires his Access file (or sanitized copy), table/column structure, and original queries/business rules. Back up the database before making any live changes. The Amazon SP-API itself does **not** depend on Access.

## Verification

GitHub Actions builds Debug and Release on Windows and runs the built-in offline WinForms self-test, including all result-view operations and Sandbox/Production endpoint checks. Live Amazon seller authorization and Joel's real Microsoft Access database must be tested separately.
