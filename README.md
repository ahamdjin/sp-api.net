# SP-API.NET

Simple VB.NET WinForms version of the Amazon SP-API workbench for Visual Studio 2019.

## Run

1. Open `SP-API.NET.sln` in Visual Studio 2019.
2. Make sure the **.NET desktop development** workload and **.NET Framework 4.8 targeting pack** are installed.
3. Press **F5**.

There is no Node.js, npm, browser server, local web server, or NuGet package setup. The app uses the normal Windows/.NET networking stack, including the Windows certificate store and system proxy support.

## Using the app

Enter the LWA Client ID, Client Secret, and Refresh Token, choose Sandbox or Production and the marketplace, then click **Test connection**. If any of those connection values change, the app correctly returns to **Not tested**.

Choose an operation on the left. The form shows only the fields for that operation. Required fields have `*`. Sandbox shows the exact fixture guidance and provides an explicit **Load Sandbox example into the form** button for static examples; nothing is silently substituted. Production is clearly marked as live, and write operations still require confirmation.

The **Result** tab gives the readable outcome and next step. **Raw response** keeps the complete Amazon payload and request metadata. IDs returned by report/feed/inbound workflows are retained for their follow-up screens.

## Included SP-API workflows

Catalog Items, Product Fees, FBA Inventory, Orders 2026, Reports, Feeds, Fulfillment Inbound plans/shipments/operations/prep, item labels, shipment labels, and bill of lading. The old company-specific SQL utilities remain visible but intentionally disconnected because their private database/schema is not part of this portable app.

Credentials stay only in the running desktop process and are not written to disk by the app.
