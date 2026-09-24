# SP-API.NET

Simple VB.NET WinForms version of the Amazon SP-API workbench for Visual Studio 2019.

## Run

1. Open `SP-API.NET.sln` in Visual Studio 2019.
2. Make sure the **.NET desktop development** workload is installed.
3. Press **F5**.

If Visual Studio says the target framework is unavailable, open **Visual Studio Installer → Modify → Individual components** and add the **.NET Framework 4.7.2 targeting pack/development tools**, then reopen the solution. No Node.js setup is involved.

The project targets **.NET Framework 4.7.2** for broader Visual Studio 2019 compatibility. It also runs normally on machines with newer .NET Framework 4.x versions installed.

There is no Node.js, npm, browser server, local web server, or NuGet package setup. The app uses the normal Windows/.NET networking stack, Windows TLS policy, Windows certificate trust, and the system proxy.

## Code layout

There are only **two VB code files to read**:

- `MainForm.vb` — WinForms view/UX, fields, navigation, results, Sandbox examples, and workflow follow-up controls.
- `ApiRequests.vb` — LWA/SP-API requests, Amazon endpoints, request validation, retries, feeds/reports/inbound/document handling, and API error normalization.

They are two parts of the same VB `Partial Class`, so the split stays simple: no dependency-injection framework, service container, generated designer files, or extra abstraction layer.

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

The **View** tab is the normal result screen. It renders Amazon data as readable tables and field/value details instead of making the user read JSON:

- Catalog results show products with ASIN, title, brand, product type, and marketplace. Selecting a product shows all returned nested product fields underneath.
- Orders, inventory, reports, feeds, inbound plans, and shipments are shown as readable rows with their important columns.
- Fee estimates show fee rows and the total.
- Single records, statuses, documents, and other responses are shown as structured field/value details.
- Double-clicking/opening a returned workflow record can carry its ID into the appropriate follow-up screen.
- **Next step**, pagination, and document open/download controls stay in the View so normal work does not require switching to JSON.

The **Summary** tab keeps the concise text outcome and is useful for copying a short result.

The **Raw response** tab is the technical/debugging view. It keeps the complete Amazon payload plus request method/path, environment, marketplace, request ID, rate-limit metadata, duration, and retry count. Credentials and access tokens are never included.

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
- structured Catalog/Orders/error View rendering and selected-record details;
- TLS/proxy/document/write-safety invariants and banned insecure patterns.

Reads retry transient Amazon/network failures with bounded backoff. Writes are not blindly replayed after ambiguous network failures. Amazon throttling, Retry-After, request IDs, business-status failures, document URLs, and bounded report/feed previews are handled explicitly.

The SP-API request surface was rechecked against Amazon's official `selling-partner-api-models` `main` revision `713565ff394d136629a342120872e65cb073d162` (2026-09-22), including the Sellers, Catalog Items 2022-04-01, Product Fees v0, FBA Inventory, Orders 2026-01-01, Reports 2021-06-30, Feeds 2021-06-30, Fulfillment Inbound 2024-03-20, and legacy Fulfillment Inbound label/BOL models.
