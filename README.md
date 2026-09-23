# SP-API.NET

VB.NET WinForms version of the Amazon SP-API workbench.

## Run

1. Open `SP-API.NET.sln` in Visual Studio 2019.
2. Make sure the **.NET desktop development** workload and **.NET Framework 4.8 targeting pack** are installed.
3. Press **F5**.

There is no Node.js, npm, browser server, or NuGet package setup. The app uses only .NET Framework libraries.

## Included SP-API workflows

Catalog Items, Product Fees, FBA Inventory, Orders 2026, Reports, Feeds, Fulfillment Inbound plans/shipments/operations/prep, item labels, shipment labels, and bill of lading. Sandbox and Production use the same visible inputs. Production write operations require explicit confirmation.

Credentials stay in the running desktop process and are not written to disk by the app.
