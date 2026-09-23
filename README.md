# SP-API.NET

Simple VB.NET WinForms version of the Amazon SP-API workbench for Visual Studio 2019.

## Run

1. Open `SP-API.NET.sln` in Visual Studio 2019.
2. Make sure the **.NET desktop development** workload is installed.
3. Press **F5**.

If Visual Studio says the target framework is unavailable, open **Visual Studio Installer → Modify → Individual components** and add the **.NET Framework 4.7.2 targeting pack/development tools**, then reopen the solution. No Node.js setup is involved.

The project targets **.NET Framework 4.7.2** for broader Visual Studio 2019 compatibility. It also runs normally on machines with newer .NET Framework 4.x versions installed.

There is no Node.js, npm, browser server, local web server, or NuGet package setup. The app uses the normal Windows/.NET networking stack, Windows TLS policy, Windows certificate trust, and the system proxy.

## Using the app

Enter the LWA Client ID, Client Secret, and Refresh Token, choose Sandbox or Production and the marketplace, then click **Test connection**.

- Credentials are kept only in the running process and are not saved to disk.
- Secrets are masked by default and can be shown explicitly.
- Changing credentials, marketplace, or environment resets the connection state and clears the cached access token.
- Successful access tokens are reused in memory until close to expiry.
- Production is clearly marked as live.
- Production write operations require both the form confirmation and a final Yes/No confirmation.
- While a request is running, the connection/request inputs are locked so the returned response cannot be confused with edited values.

Choose an operation on the left. The form shows only the fields for that operation. Required fields have `*`.

Sandbox displays Amazon's fixture guidance and has an explicit **Load Sandbox example into the form** button for static examples. Nothing is silently substituted. Dynamic Sandbox operations remain manual.

The **Result** tab shows the readable outcome, returned IDs, documents, and the next useful action. Asynchronous report/feed/inbound flows provide a simple **Next step** button. Paginated operations can load the returned next-page token without copy/paste. List operations such as orders, reports, feeds, inbound plans, and shipments expose returned records through a simple **Open selected** control. If Amazon returns several label/document URLs, all of them remain selectable.

The **Raw response** tab keeps the complete Amazon payload plus request method/path, environment, marketplace, request ID, rate-limit metadata, duration, and retry count. Credentials and access tokens are never included. Result and Raw response can both be copied with one click.

## Included SP-API workflows

Catalog Items, Product Fees, FBA Inventory, Orders 2026, Reports, Feeds, Fulfillment Inbound plans/shipments/operations/prep, item labels, shipment labels, and bill of lading.

The old company-specific SQL utilities remain visible but intentionally disconnected because their private database/schema is not part of this portable app.

## Reliability

The repository CI verifies:

- every visible SP-API operation is wired into the execution router;
- the audited endpoint families and important request-contract guards remain present;
- Debug build, matching the normal Visual Studio F5 configuration;
- Release build;
- WinForms startup;
- every operation screen can be constructed;
- every static Sandbox example can be loaded without hidden substitution;
- write confirmation is invalidated when request values change;
- marketplace/region endpoint mapping;
- HTTPS document handling and multi-document extraction;
- key request-validation helpers;
- stale pagination-token prevention and one-click pagination follow-up;
- returned-record follow-up routing;
- TLS/proxy/document/write-safety invariants and banned insecure patterns.

Reads retry transient Amazon/network failures with bounded backoff. Writes are not blindly replayed after ambiguous network failures. Amazon throttling, Retry-After, request IDs, business-status failures, document URLs, and bounded report/feed previews are handled explicitly.

The SP-API request surface was rechecked against Amazon's official `selling-partner-api-models` `main` revision `713565ff394d136629a342120872e65cb073d162` (2026-09-22), including the Sellers, Catalog Items 2022-04-01, Product Fees v0, FBA Inventory, Orders 2026-01-01, Reports 2021-06-30, Feeds 2021-06-30, Fulfillment Inbound 2024-03-20, and legacy Fulfillment Inbound label/BOL models.
