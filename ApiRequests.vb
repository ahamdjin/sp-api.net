Option Explicit On
Option Strict On
Option Infer On

Imports System
Imports System.Collections
Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.Script.Serialization

' Amazon/LWA requests, request validation, retries, documents, and API error handling.
Public Partial Class MainForm

    ' -------------------- Shared API types and HTTP client --------------------

    Private Class ApiProblem
        Public Property Code As String = ""
        Public Property Message As String = ""
        Public Property Details As String = ""
        Public Property Action As String = ""
        Public Property Retryable As Boolean
    End Class

    Private Class ApiResult
        Public Property Ok As Boolean
        Public Property Status As Integer
        Public Property StatusText As String = ""
        Public Property RequestId As String = ""
        Public Property GatewayId As String = ""
        Public Property TraceId As String = ""
        Public Property RateLimit As String = ""
        Public Property DurationMs As Long
        Public Property Attempts As Integer
        Public Property RequestMethod As String = ""
        Public Property RequestPath As String = ""
        Public Property Data As Object
        Public Property Problem As ApiProblem
        Public Property ErrorMessage As String = ""
    End Class

    Private Class AppException
        Inherits Exception
        Public ReadOnly Property Status As Integer
        Public ReadOnly Property Code As String
        Public ReadOnly Property Details As String
        Public Sub New(message As String, Optional status As Integer = 400, Optional code As String = "INPUT_VALIDATION", Optional details As String = "")
            MyBase.New(message)
            Me.Status = status
            Me.Code = code
            Me.Details = details
        End Sub
    End Class

    Private Enum RetryMode
        ReadRequest
        WriteRequest
        SafePost
    End Enum

    Private Shared ReadOnly Serializer As New JavaScriptSerializer With {
        .MaxJsonLength = Integer.MaxValue,
        .RecursionLimit = 200
    }

    Private Shared ReadOnly Http As New HttpClient(New HttpClientHandler With {
        .AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate,
        .UseProxy = True,
        .DefaultProxyCredentials = CredentialCache.DefaultCredentials
    }) With {.Timeout = TimeSpan.FromSeconds(30)}

    Private CachedAccessToken As String = ""
    Private CachedAccessTokenExpiresUtc As DateTimeOffset = DateTimeOffset.MinValue

    Private Const CoreOrderData As String = "PROCEEDS,EXPENSE,PROMOTION,CANCELLATION,FULFILLMENT,PACKAGES,TAX,PAYMENT,FULFILLMENT_ORDERS"
    Private Const DefaultCatalogData As String = "attributes,classifications,dimensions,identifiers,images,productTypes,relationships,salesRanks,summaries,vendorDetails"
    Private Const PreviewLimit As Integer = 2 * 1024 * 1024

    Private Function Endpoint() As String
        Dim prefix = If(IsSandbox(), "https://sandbox.sellingpartnerapi-", "https://sellingpartnerapi-")
        Return prefix & SelectedMarketplace().Region & ".amazon.com"
    End Function

    Private Async Function TestConnectionRequestAsync() As Task(Of ApiResult)
        Dim token = Await GetAccessTokenAsync(True)
        Return Await CallSpApiAsync("/sellers/v1/marketplaceParticipations", HttpMethod.Get, Nothing, token.Item1)
    End Function

    ' -------------------- Operation router --------------------
    Private Async Function ExecuteOperationAsync(operation As String) As Task(Of ApiResult)
        Select Case operation
            Case "catalog" : Return Await CatalogAsync()
            Case "fees" : Return Await FeesAsync()
            Case "inventory" : Return Await InventoryAsync()
            Case "orders" : Return Await OrdersAsync()
            Case "order" : Return Await OrderAsync()
            Case "reports" : Return Await ReportsAsync()
            Case "createReport" : Return Await CreateReportAsync()
            Case "report" : Return ApplyBusinessOutcome(operation, Await CallSpApiAsync("/reports/2021-06-30/reports/" & Encode(Required("reportId"))))
            Case "reportDocument" : Return Await GetAndDownloadDocumentAsync("/reports/2021-06-30/documents/" & Encode(Required("reportDocumentId")) & If(IsSandbox(), "", "?enableContentEncodingUrlHeader=true"), "report document")
            Case "feeds" : Return Await FeedsAsync()
            Case "feed" : Return ApplyBusinessOutcome(operation, Await CallSpApiAsync("/feeds/2021-06-30/feeds/" & Encode(Required("feedId"))))
            Case "feedDocument" : Return Await GetAndDownloadDocumentAsync("/feeds/2021-06-30/documents/" & Encode(Required("feedDocumentId")) & If(IsSandbox(), "", "?enableContentEncodingUrlHeader=true"), "feed processing report")
            Case "submitFeed" : Return ApplyBusinessOutcome(operation, Await SubmitFeedAsync())
            Case "inboundPlans" : Return Await InboundPlansAsync()
            Case "inboundPlan" : Return Await CallSpApiAsync("/inbound/fba/2024-03-20/inboundPlans/" & Encode(Required("inboundPlanId")))
            Case "inboundShipment" : Return Await CallSpApiAsync("/inbound/fba/2024-03-20/inboundPlans/" & Encode(Required("inboundPlanId")) & "/shipments/" & Encode(Required("shipmentId")))
            Case "inboundOperationStatus" : Return ApplyBusinessOutcome(operation, Await CallSpApiAsync("/inbound/fba/2024-03-20/operations/" & Encode(Required("operationId"))))
            Case "prepDetails" : Return Await PrepDetailsAsync()
            Case "createInboundPlan" : Return ApplyBusinessOutcome(operation, Await CreateInboundPlanAsync())
            Case "itemLabels" : Return ApplyBusinessOutcome(operation, Await ItemLabelsAsync())
            Case "shipmentLabels" : Return ApplyBusinessOutcome(operation, Await ShipmentLabelsAsync())
            Case "billOfLading" : Return ApplyBusinessOutcome(operation, Await CallSpApiAsync("/fba/inbound/v0/shipments/" & Encode(Required("shipmentId")) & "/billOfLading"))
            Case Else : Throw New AppException("Unsupported operation", 400, "UNSUPPORTED_OPERATION")
        End Select
    End Function

    ' -------------------- Catalog --------------------
    Private Async Function CatalogAsync() As Task(Of ApiResult)
        Dim marketplace = SelectedMarketplace()
        Dim mode = Required("catalogMode")
        Dim identifierType = Required("identifierType")
        Dim query = Required("query")
        Dim included = NormalizeIncludedData(S("includedData"))
        Dim token = (Await GetAccessTokenAsync()).Item1
        Dim identifiers = SplitValues(query, 20).Select(Function(x) If(identifierType = "ASIN", x.ToUpperInvariant(), x)).ToList()

        If mode = "identifier" AndAlso identifierType = "SKU" AndAlso S("sellerId") = "" Then Throw New AppException("Seller ID is required for SKU searches")

        If mode = "identifier" AndAlso identifierType = "ASIN" AndAlso identifiers.Count = 1 AndAlso B("includeVariations") Then
            Return Await CompleteCatalogFamilyAsync(identifiers(0), included, token)
        End If

        If mode = "identifier" AndAlso identifierType = "ASIN" AndAlso identifiers.Count = 1 Then
            Dim q As New List(Of KeyValuePair(Of String, String)) From {
                QPair("marketplaceIds", marketplace.Id), QPair("includedData", included)
            }
            If Not IsSandbox() Then q.Add(QPair("locale", marketplace.Locale))
            Dim exact = Await CallSpApiAsync("/catalog/2022-04-01/items/" & Encode(identifiers(0)) & "?" & BuildQuery(q), HttpMethod.Get, Nothing, token)
            If exact.Ok Then
                exact.Data = New Dictionary(Of String, Object) From {{"numberOfResults", 1}, {"items", New Object() {exact.Data}}}
            End If
            Return exact
        End If

        Dim params As New List(Of KeyValuePair(Of String, String)) From {
            QPair("marketplaceIds", marketplace.Id), QPair("includedData", included)
        }
        If Not IsSandbox() Then
            params.Add(QPair("locale", marketplace.Locale))
            params.Add(QPair("pageSize", IntField("catalogPageSize", 1, 20, 20).ToString(CultureInfo.InvariantCulture)))
        End If
        If mode = "keywords" Then
            params.Add(QPair("keywords", query))
            If Not IsSandbox() Then params.Add(QPair("keywordsLocale", marketplace.Locale))
            AddCsvParam(params, "brandNames", S("brandNames"), Integer.MaxValue)
            AddCsvParam(params, "classificationIds", S("classificationIds"), Integer.MaxValue)
            AddOptional(params, "pageToken", S("pageToken"))
        Else
            params.Add(QPair("identifiers", String.Join(",", identifiers)))
            params.Add(QPair("identifiersType", identifierType))
            If identifierType = "SKU" Then params.Add(QPair("sellerId", Required("sellerId")))
        End If
        Return Await CallSpApiAsync("/catalog/2022-04-01/items?" & BuildQuery(params), HttpMethod.Get, Nothing, token)
    End Function

    Private Async Function CompleteCatalogFamilyAsync(requestedAsin As String, included As String, accessToken As String) As Task(Of ApiResult)
        Dim exact = Await GetCatalogItemAsync(requestedAsin, included, accessToken)
        If Not exact.Ok Then Return exact

        Dim items As New Dictionary(Of String, Dictionary(Of String, Object))(StringComparer.OrdinalIgnoreCase)
        Dim warnings As New List(Of Object)()
        Dim attempted As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {requestedAsin}
        Dim family As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {requestedAsin}
        AddCatalogItem(items, exact.Data)
        Dim initial = RelatedAsins(exact.Data)
        Dim parents = initial.Item1.Where(Function(x) Not x.Equals(requestedAsin, StringComparison.OrdinalIgnoreCase)).ToList()
        For Each a In initial.Item1.Concat(initial.Item2) : family.Add(a) : Next
        Dim totalDuration = exact.DurationMs

        Do
            Dim missing = family.Where(Function(a) Not items.ContainsKey(a) AndAlso Not attempted.Contains(a)).ToList()
            If missing.Count = 0 Then Exit Do
            For i As Integer = 0 To missing.Count - 1 Step 20
                Dim batch = missing.Skip(i).Take(20).ToList()
                For Each a In batch : attempted.Add(a) : Next
                Dim related = Await SearchCatalogAsinsAsync(batch, included, accessToken)
                totalDuration += related.DurationMs
                If related.Ok Then
                    Dim data = AsDict(related.Data)
                    For Each raw In ListValue(GetValue(data, "items"))
                        AddCatalogItem(items, raw)
                        Dim rels = RelatedAsins(raw)
                        For Each a In rels.Item1.Concat(rels.Item2) : family.Add(a) : Next
                    Next
                Else
                    warnings.Add(New Dictionary(Of String, Object) From {
                        {"asins", batch.ToArray()}, {"status", related.Status},
                        {"code", If(related.Problem IsNot Nothing, related.Problem.Code, "HTTP_" & related.Status.ToString())},
                        {"message", If(related.Problem IsNot Nothing, related.Problem.Message, "Amazon did not return this related item.")},
                        {"action", If(related.Problem IsNot Nothing, related.Problem.Action, "Retry the related-ASIN lookup after correcting the Amazon error.")},
                        {"requestId", related.RequestId}
                    })
                End If
            Next
        Loop

        Dim ordered = items.Values.OrderBy(Function(x)
                                                Dim asin = StringValue(GetValue(x, "asin"))
                                                If asin.Equals(requestedAsin, StringComparison.OrdinalIgnoreCase) Then Return "0" & asin
                                                If parents.Contains(asin) Then Return "1" & asin
                                                Return "2" & asin
                                            End Function).Cast(Of Object)().ToArray()
        exact.DurationMs = totalDuration
        exact.Data = New Dictionary(Of String, Object) From {
            {"numberOfResults", ordered.Length}, {"items", ordered},
            {"family", New Dictionary(Of String, Object) From {
                {"requestedAsin", requestedAsin}, {"parentAsins", parents.ToArray()},
                {"relatedAsins", family.Where(Function(a) Not a.Equals(requestedAsin, StringComparison.OrdinalIgnoreCase)).ToArray()},
                {"requestedCount", family.Count}, {"returnedCount", ordered.Length},
                {"complete", warnings.Count = 0 AndAlso ordered.Length = family.Count}, {"warnings", warnings.ToArray()}
            }}
        }
        Return exact
    End Function

    Private Async Function GetCatalogItemAsync(asin As String, included As String, accessToken As String) As Task(Of ApiResult)
        Dim q As New List(Of KeyValuePair(Of String, String)) From {QPair("marketplaceIds", SelectedMarketplace().Id), QPair("includedData", included)}
        If Not IsSandbox() Then q.Add(QPair("locale", SelectedMarketplace().Locale))
        Return Await CallSpApiAsync("/catalog/2022-04-01/items/" & Encode(asin) & "?" & BuildQuery(q), HttpMethod.Get, Nothing, accessToken)
    End Function

    Private Async Function SearchCatalogAsinsAsync(asins As List(Of String), included As String, accessToken As String) As Task(Of ApiResult)
        Dim q As New List(Of KeyValuePair(Of String, String)) From {
            QPair("identifiers", String.Join(",", asins)), QPair("identifiersType", "ASIN"), QPair("marketplaceIds", SelectedMarketplace().Id), QPair("includedData", included)
        }
        If Not IsSandbox() Then q.Add(QPair("locale", SelectedMarketplace().Locale)) : q.Add(QPair("pageSize", "20"))
        Return Await CallSpApiAsync("/catalog/2022-04-01/items?" & BuildQuery(q), HttpMethod.Get, Nothing, accessToken)
    End Function

    Private Sub AddCatalogItem(items As Dictionary(Of String, Dictionary(Of String, Object)), value As Object)
        Dim item = AsDict(value)
        Dim asin = StringValue(GetValue(item, "asin"))
        If asin <> "" Then items(asin) = item
    End Sub

    Private Function RelatedAsins(value As Object) As Tuple(Of List(Of String), List(Of String))
        Dim parents As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim children As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim item = AsDict(value)
        For Each groupObj In ListValue(GetValue(item, "relationships"))
            Dim group = AsDict(groupObj)
            For Each relObj In ListValue(GetValue(group, "relationships"))
                Dim rel = AsDict(relObj)
                For Each a In ListValue(GetValue(rel, "parentAsins"))
                    If TypeOf a Is String Then parents.Add(CStr(a))
                Next
                For Each a In ListValue(GetValue(rel, "childAsins"))
                    If TypeOf a Is String Then children.Add(CStr(a))
                Next
            Next
        Next
        Return Tuple.Create(parents.ToList(), children.ToList())
    End Function

    Private Function NormalizeIncludedData(value As String) As String
        Dim allowed As New HashSet(Of String)({"attributes", "classifications", "dimensions", "identifiers", "images", "productTypes", "relationships", "salesRanks", "summaries", "vendorDetails"}, StringComparer.Ordinal)
        Dim entries = SplitValues(If(value = "", DefaultCatalogData, value), 10)
        Dim invalid = entries.Where(Function(x) Not allowed.Contains(x)).ToList()
        If invalid.Count > 0 Then Throw New AppException("Unsupported Catalog includedData value: " & String.Join(", ", invalid), 400, "INVALID_CATALOG_INCLUDED_DATA")
        Return String.Join(",", entries)
    End Function

    ' -------------------- Fees, inventory, and orders --------------------
    Private Async Function FeesAsync() As Task(Of ApiResult)
        Dim idType = Required("feeIdType")
        Dim identifier = Required("feeIdentifier")
        Dim price = DecimalField("price", 0.01D, Decimal.MaxValue)
        Dim shipping = DecimalField("shipping", 0D, Decimal.MaxValue, 0D)
        Dim requestIdentifier = S("requestIdentifier")
        If requestIdentifier = "" Then requestIdentifier = "sp-api-workbench-" & DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString(CultureInfo.InvariantCulture)
        Dim currency = SelectedMarketplace().Currency.ToUpperInvariant()
        Dim priceToEstimate As New Dictionary(Of String, Object) From {
            {"ListingPrice", Money(currency, price)}, {"Shipping", Money(currency, shipping)}
        }
        If S("pointsNumber") <> "" OrElse S("pointsAmount") <> "" Then
            priceToEstimate("Points") = New Dictionary(Of String, Object) From {
                {"PointsNumber", IntField("pointsNumber", 0, Integer.MaxValue, 0)},
                {"PointsMonetaryValue", Money(currency, DecimalField("pointsAmount", 0D, Decimal.MaxValue, 0D))}
            }
        End If
        Dim body As New Dictionary(Of String, Object) From {
            {"FeesEstimateRequest", New Dictionary(Of String, Object) From {
                {"MarketplaceId", SelectedMarketplace().Id}, {"IsAmazonFulfilled", B("isAmazonFulfilled")},
                {"PriceToEstimateFees", priceToEstimate}, {"Identifier", requestIdentifier}
            }}
        }
        Dim resource = If(idType = "ASIN", "items", "listings")
        Dim result = Await CallSpApiAsync("/products/fees/v0/" & resource & "/" & Encode(identifier) & "/feesEstimate", HttpMethod.Post, body)
        If result.Ok Then
            Dim payload = AsDict(GetValue(AsDict(result.Data), "payload"))
            Dim fee = AsDict(GetValue(payload, "FeesEstimateResult"))
            Dim feeStatus = StringValue(GetValue(fee, "Status"))
            If feeStatus <> "" AndAlso feeStatus <> "Success" Then
                Dim err = AsDict(GetValue(fee, "Error"))
                result.Ok = False : result.Status = 422 : result.StatusText = "Fee estimate " & feeStatus
                result.Problem = New ApiProblem With {
                    .Code = If(StringValue(GetValue(err, "Code")) <> "", StringValue(GetValue(err, "Code")), "FEE_ESTIMATE_" & feeStatus.ToUpperInvariant()),
                    .Message = If(StringValue(GetValue(err, "Message")) <> "", StringValue(GetValue(err, "Message")), "Amazon returned a fee-estimate business error."),
                    .Details = Json(GetValue(err, "Detail")),
                    .Action = If(feeStatus = "ServiceError", "Amazon's fee service reported a service error. Retry with backoff.", "Check the ASIN/SKU, marketplace, price, shipping, currency, and fulfillment choice."),
                    .Retryable = (feeStatus = "ServiceError")
                }
            End If
        End If
        Return result
    End Function

    Private Async Function InventoryAsync() As Task(Of ApiResult)
        Dim q As New List(Of KeyValuePair(Of String, String)) From {
            QPair("details", If(B("details"), "true", "false")), QPair("granularityType", "Marketplace"),
            QPair("granularityId", SelectedMarketplace().Id), QPair("marketplaceIds", SelectedMarketplace().Id)
        }
        AddCsvParam(q, "sellerSkus", S("sellerSkus"), 50)
        Dim start = OptionalDate("startDateTime")
        If start <> "" AndAlso DateTimeOffset.Parse(start, CultureInfo.InvariantCulture).UtcDateTime < DateTime.UtcNow.AddMonths(-18) Then Throw New AppException("startDateTime cannot be earlier than 18 months before the request", 400, "INVENTORY_START_TOO_OLD")
        AddOptional(q, "startDateTime", start)
        AddOptional(q, "nextToken", S("inventoryNextToken"))
        Return Await CallSpApiAsync("/fba/inventory/v1/summaries?" & BuildQuery(q))
    End Function

    Private Async Function OrdersAsync() As Task(Of ApiResult)
        Dim createdAfter = RequiredDate("createdAfter")
        Dim createdBefore = OptionalDate("createdBefore")
        If createdBefore <> "" Then
            Dim a = DateTimeOffset.Parse(createdAfter, CultureInfo.InvariantCulture)
            Dim b = DateTimeOffset.Parse(createdBefore, CultureInfo.InvariantCulture)
            If b < a Then Throw New AppException("createdBefore must be equal to or after createdAfter", 400, "INVALID_ORDER_DATE_RANGE")
            If b > DateTimeOffset.UtcNow.AddMinutes(-2) Then Throw New AppException("createdBefore must be at least two minutes before the request time", 400, "ORDER_CREATED_BEFORE_TOO_RECENT")
        End If
        Dim q As New List(Of KeyValuePair(Of String, String)) From {
            QPair("marketplaceIds", SelectedMarketplace().Id), QPair("createdAfter", createdAfter), QPair("includedData", OrderIncludedData())
        }
        If S("pageSize") <> "" Then q.Add(QPair("maxResultsPerPage", IntField("pageSize", 1, 100).ToString(CultureInfo.InvariantCulture)))
        AddOptional(q, "createdBefore", createdBefore)
        Dim statuses = SplitValues(S("statuses"), 7)
        ValidateEnum(statuses, {"PENDING_AVAILABILITY", "PENDING", "UNSHIPPED", "PARTIALLY_SHIPPED", "SHIPPED", "CANCELLED", "UNFULFILLABLE"}, "fulfillmentStatuses")
        Dim fulfilled = SplitValues(S("fulfilledBy"), 2)
        ValidateEnum(fulfilled, {"MERCHANT", "AMAZON"}, "fulfilledBy")
        If statuses.Count > 0 Then q.Add(QPair("fulfillmentStatuses", String.Join(",", statuses)))
        If fulfilled.Count > 0 Then q.Add(QPair("fulfilledBy", String.Join(",", fulfilled)))
        AddOptional(q, "paginationToken", S("orderPaginationToken"))
        Return Await CallSpApiAsync("/orders/2026-01-01/orders?" & BuildQuery(q))
    End Function

    Private Async Function OrderAsync() As Task(Of ApiResult)
        Return Await CallSpApiAsync("/orders/2026-01-01/orders/" & Encode(Required("orderId")) & "?includedData=" & Encode(OrderIncludedData()))
    End Function

    Private Function OrderIncludedData() As String
        Dim overrideValue = S("orderIncludedData")
        Dim allowed = {"BUYER", "RECIPIENT", "PROCEEDS", "EXPENSE", "PROMOTION", "CANCELLATION", "FULFILLMENT", "PACKAGES", "TAX", "PAYMENT", "FULFILLMENT_ORDERS"}
        If overrideValue <> "" Then
            Dim values = SplitValues(overrideValue, 11)
            ValidateEnum(values, allowed, "includedData")
            Return String.Join(",", values)
        End If
        Return If(B("includeOrderPii"), "BUYER,RECIPIENT," & CoreOrderData, CoreOrderData)
    End Function

    ' -------------------- Reports and feeds --------------------
    Private Async Function ReportsAsync() As Task(Of ApiResult)
        Dim nextToken = S("reportNextToken")
        Dim q As New List(Of KeyValuePair(Of String, String))()
        If nextToken <> "" Then
            q.Add(QPair("nextToken", nextToken))
        Else
            AddCsvParam(q, "reportTypes", S("reportTypes"), 10)
            Dim statuses = SplitValues(S("processingStatuses"), 5)
            ValidateEnum(statuses, {"CANCELLED", "DONE", "FATAL", "IN_PROGRESS", "IN_QUEUE"}, "processingStatuses")
            If statuses.Count > 0 Then q.Add(QPair("processingStatuses", String.Join(",", statuses)))
            Dim mids = SplitValues(S("reportMarketplaceIds"), 10)
            ValidateMarketplaceRegions(mids, "reportMarketplaceIds")
            If mids.Count > 0 Then q.Add(QPair("marketplaceIds", String.Join(",", mids)))
            If S("pageSize") <> "" Then q.Add(QPair("pageSize", IntField("pageSize", 1, 100).ToString(CultureInfo.InvariantCulture)))
            Dim since = OptionalDate("createdSince")
            Dim until = OptionalDate("createdUntil")
            ValidateRange(since, until, "createdSince", "createdUntil")
            AddOptional(q, "createdSince", since) : AddOptional(q, "createdUntil", until)
        End If
        Return Await CallSpApiAsync("/reports/2021-06-30/reports?" & BuildQuery(q))
    End Function

    Private Async Function CreateReportAsync() As Task(Of ApiResult)
        Dim startTime = OptionalDate("dataStartTime")
        Dim endTime = OptionalDate("dataEndTime")
        ValidateRange(startTime, endTime, "dataStartTime", "dataEndTime")
        Dim mids = SplitValues(S("reportMarketplaceIds"), 25)
        If mids.Count = 0 Then mids.Add(SelectedMarketplace().Id)
        ValidateMarketplaceRegions(mids, "reportMarketplaceIds")
        Dim body As New Dictionary(Of String, Object) From {{"reportType", Required("reportType")}, {"marketplaceIds", mids.ToArray()}}
        If startTime <> "" Then body("dataStartTime") = startTime
        If endTime <> "" Then body("dataEndTime") = endTime
        Return ApplyBusinessOutcome("createReport", Await CallSpApiAsync("/reports/2021-06-30/reports", HttpMethod.Post, body))
    End Function

    Private Async Function FeedsAsync() As Task(Of ApiResult)
        Dim nextToken = S("feedNextToken")
        Dim q As New List(Of KeyValuePair(Of String, String))()
        If nextToken <> "" Then
            q.Add(QPair("nextToken", nextToken))
        Else
            AddCsvParam(q, "feedTypes", S("feedTypes"), 10)
            Dim statuses = SplitValues(S("processingStatuses"), 5)
            ValidateEnum(statuses, {"CANCELLED", "DONE", "FATAL", "IN_PROGRESS", "IN_QUEUE"}, "processingStatuses")
            If statuses.Count > 0 Then q.Add(QPair("processingStatuses", String.Join(",", statuses)))
            Dim mids = SplitValues(S("feedMarketplaceIds"), 10)
            ValidateMarketplaceRegions(mids, "feedMarketplaceIds")
            If mids.Count > 0 Then q.Add(QPair("marketplaceIds", String.Join(",", mids)))
            If S("pageSize") <> "" Then q.Add(QPair("pageSize", IntField("pageSize", 1, 100).ToString(CultureInfo.InvariantCulture)))
            Dim since = OptionalDate("createdSince")
            Dim until = OptionalDate("createdUntil")
            ValidateRange(since, until, "createdSince", "createdUntil")
            AddOptional(q, "createdSince", since) : AddOptional(q, "createdUntil", until)
        End If
        Return Await CallSpApiAsync("/feeds/2021-06-30/feeds?" & BuildQuery(q))
    End Function

    Private Async Function SubmitFeedAsync() As Task(Of ApiResult)
        Dim feedType = Required("feedType")
        Dim removed = New HashSet(Of String)({"POST_PRODUCT_DATA", "POST_INVENTORY_AVAILABILITY_DATA", "POST_PRODUCT_OVERRIDES_DATA", "POST_PRODUCT_PRICING_DATA", "POST_PRODUCT_IMAGE_DATA", "POST_PRODUCT_RELATIONSHIP_DATA", "POST_FLAT_FILE_INVLOADER_DATA", "POST_FLAT_FILE_BOOKLOADER_DATA", "POST_FLAT_FILE_CONVERGENCE_LISTINGS_DATA", "POST_FLAT_FILE_LISTINGS_DATA", "POST_FLAT_FILE_PRICEANDQUANTITYONLY_UPDATE_DATA", "POST_UIEE_BOOKLOADER_DATA"}, StringComparer.Ordinal)
        If Not IsSandbox() AndAlso removed.Contains(feedType) Then Throw New AppException("This legacy listings feed type was removed by Amazon on July 31, 2025. Use JSON_LISTINGS_FEED in Production.", 400, "REMOVED_LISTING_FEED_TYPE")
        Dim contentType = If(S("contentType") = "", "application/json; charset=UTF-8", S("contentType"))
        ValidateContentType(contentType)
        Dim content = Required("content")
        Dim mids = SplitValues(S("feedMarketplaceIds"), 25)
        If mids.Count = 0 Then mids.Add(SelectedMarketplace().Id)
        ValidateMarketplaceRegions(mids, "feedMarketplaceIds")
        If feedType = "JSON_LISTINGS_FEED" Then ValidateJsonListingsFeed(contentType, content)
        If Encoding.UTF8.GetByteCount(content) > 5 * 1024 * 1024 Then Throw New AppException("Feed content is limited to 5 MB in this workbench", 413, "FEED_TOO_LARGE")

        Dim document = Await CallSpApiAsync("/feeds/2021-06-30/documents", HttpMethod.Post, New Dictionary(Of String, Object) From {{"contentType", contentType}})
        If Not document.Ok Then Return document
        Dim docData = AsDict(document.Data)
        Dim url = StringValue(GetValue(docData, "url"))
        Dim docId = StringValue(GetValue(docData, "feedDocumentId"))
        If url = "" OrElse docId = "" Then Throw New AppException("Amazon did not return a feed upload URL and document ID", 502, "FEED_UPLOAD_URL_MISSING")
        If Not IsSandbox() Then
            ValidateAmazonDocumentUrl(url, "feed upload")
            Using req As New HttpRequestMessage(HttpMethod.Put, url)
                req.Content = New StringContent(content, Encoding.UTF8)
                req.Content.Headers.ContentType = MediaTypeHeaderValue.Parse(contentType)
                Using response = Await Http.SendAsync(req)
                    If Not response.IsSuccessStatusCode Then Throw New AppException("Amazon feed document upload failed with HTTP " & CInt(response.StatusCode).ToString(), CInt(response.StatusCode), "FEED_DOCUMENT_UPLOAD_FAILED")
                End Using
            End Using
        End If
        Dim created = Await CallSpApiAsync("/feeds/2021-06-30/feeds", HttpMethod.Post, New Dictionary(Of String, Object) From {{"feedType", feedType}, {"marketplaceIds", mids.ToArray()}, {"inputFeedDocumentId", docId}})
        If created.Ok Then
            Dim data = AsDict(created.Data)
            data("inputFeedDocumentId") = docId
            data("verification") = If(IsSandbox(), "Static Sandbox validates Amazon's predefined examples and does not persist the upload like Production.", "Poll Feed status until DONE or FATAL, then inspect resultFeedDocumentId.")
            created.Data = data
        End If
        Return created
    End Function

    ' -------------------- Fulfillment Inbound --------------------
    Private Async Function InboundPlansAsync() As Task(Of ApiResult)
        Dim q As New List(Of KeyValuePair(Of String, String))()
        If S("pageSize") <> "" Then q.Add(QPair("pageSize", IntField("pageSize", 1, 30).ToString(CultureInfo.InvariantCulture)))
        AddEnumParam(q, "sortBy", S("sortBy"), {"LAST_UPDATED_TIME", "CREATION_TIME"})
        AddEnumParam(q, "sortOrder", S("sortOrder"), {"ASC", "DESC"})
        AddEnumParam(q, "status", S("status"), {"ACTIVE", "VOIDED", "SHIPPED"})
        AddOptional(q, "paginationToken", S("inboundPaginationToken"))
        Return Await CallSpApiAsync("/inbound/fba/2024-03-20/inboundPlans?" & BuildQuery(q))
    End Function

    Private Async Function PrepDetailsAsync() As Task(Of ApiResult)
        Dim lines = Required("mskus").Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries).Select(Function(x) x.Trim()).Where(Function(x) x <> "").ToList()
        If lines.Count > 100 Then Throw New AppException("mskus accepts at most 100 values", 400, "TOO_MANY_VALUES")
        Dim parts As New List(Of String) From {"marketplaceId=" & Encode(SelectedMarketplace().Id)}
        For Each msku In lines
            Dim pre = msku.Replace("%", "%25").Replace("+", "%2B").Replace(",", "%2C")
            parts.Add("mskus=" & Encode(pre))
        Next
        Return Await CallSpApiAsync("/inbound/fba/2024-03-20/items/prepDetails?" & String.Join("&", parts))
    End Function

    Private Async Function CreateInboundPlanAsync() As Task(Of ApiResult)
        Dim mids = SplitValues(S("destinationMarketplaces"), 1)
        If mids.Count = 0 Then mids.Add(SelectedMarketplace().Id)
        ValidateMarketplaceRegions(mids, "destinationMarketplaces")
        Dim country = If(S("countryCode") = "", SelectedMarketplace().Locale.Substring(SelectedMarketplace().Locale.Length - 2), S("countryCode")).ToUpperInvariant()
        If Not Regex.IsMatch(country, "^[A-Z]{2}$") Then Throw New AppException("countryCode must be a two-letter ISO country code", 400, "INVALID_COUNTRY_CODE")
        Dim items = ParseItems(Required("items"), 2000, 500000)
        If mids(0) = "ATVPDKIKX0DER" Then
            For Each itemObj In items
                Dim item = AsDict(itemObj)
                If StringValue(GetValue(item, "labelOwner")) = "AMAZON" Then Throw New AppException("Amazon does not accept labelOwner=AMAZON for US inbound-plan items. Use SELLER or NONE.", 400, "INVALID_US_LABEL_OWNER")
            Next
        End If
        Dim address As New Dictionary(Of String, Object) From {
            {"name", LimitedRequired("contactName", 50)}, {"addressLine1", LimitedRequired("addressLine1", 180)},
            {"city", LimitedRequired("city", 30)}, {"postalCode", LimitedRequired("postalCode", 32)},
            {"countryCode", country}, {"phoneNumber", LimitedRequired("phoneNumber", 20)}
        }
        AddOptionalObject(address, "companyName", LimitedOptional("companyName", 50))
        AddOptionalObject(address, "addressLine2", LimitedOptional("addressLine2", 60))
        AddOptionalObject(address, "districtOrCounty", LimitedOptional("districtOrCounty", 50))
        AddOptionalObject(address, "stateOrProvinceCode", LimitedOptional("stateOrProvinceCode", 64))
        AddOptionalObject(address, "email", LimitedOptional("email", 1024))
        Dim body As New Dictionary(Of String, Object) From {{"destinationMarketplaces", mids.ToArray()}, {"sourceAddress", address}, {"items", items.ToArray()}}
        Dim planName = LimitedOptional("planName", 40)
        If planName <> "" Then body("name") = planName
        Return Await CallSpApiAsync("/inbound/fba/2024-03-20/inboundPlans", HttpMethod.Post, body)
    End Function

    Private Async Function ItemLabelsAsync() As Task(Of ApiResult)
        Dim labelType = If(S("labelType") = "", "STANDARD_FORMAT", S("labelType"))
        ValidateEnum(New List(Of String) From {labelType}, {"STANDARD_FORMAT", "THERMAL_PRINTING"}, "labelType")
        Dim parsed = ParseItems(Required("items"), 100, 10000)
        Dim quantities As New List(Of Object)()
        For Each raw In parsed
            Dim item = AsDict(raw)
            quantities.Add(New Dictionary(Of String, Object) From {{"msku", GetValue(item, "msku")}, {"quantity", GetValue(item, "quantity")}})
        Next
        Dim body As New Dictionary(Of String, Object) From {{"marketplaceId", SelectedMarketplace().Id}, {"labelType", labelType}, {"localeCode", SelectedMarketplace().Locale}, {"mskuQuantities", quantities.ToArray()}}
        If labelType = "THERMAL_PRINTING" Then
            body("height") = Decimal.ToDouble(DecimalField("labelHeight", 25D, 100D, 25D))
            body("width") = Decimal.ToDouble(DecimalField("labelWidth", 25D, 100D, 100D))
        Else
            Dim page = If(S("pageType") = "", "A4_21", S("pageType"))
            ValidateEnum(New List(Of String) From {page}, {"A4_21", "A4_24", "A4_24_64x33", "A4_24_66x35", "A4_24_70x36", "A4_24_70x37", "A4_24i", "A4_27", "A4_40_52x29", "A4_44_48x25", "Letter_30"}, "pageType")
            body("pageType") = page
        End If
        Return Await CallSpApiAsync("/inbound/fba/2024-03-20/items/labels", HttpMethod.Post, body)
    End Function

    Private Async Function ShipmentLabelsAsync() As Task(Of ApiResult)
        Dim labelType = If(S("shipmentLabelType") = "", "UNIQUE", S("shipmentLabelType"))
        ValidateEnum(New List(Of String) From {labelType}, {"BARCODE_2D", "UNIQUE", "PALLET"}, "shipmentLabelType")
        If labelType = "PALLET" AndAlso S("numberOfPallets") = "" Then Throw New AppException("numberOfPallets is required for PALLET labels", 400, "MISSING_NUMBER_OF_PALLETS")
        Dim pageType = If(S("shipmentPageType") = "", "PackageLabel_Thermal_NonPCP", S("shipmentPageType"))
        ValidateEnum(New List(Of String) From {pageType}, {"PackageLabel_Letter_2", "PackageLabel_Letter_4", "PackageLabel_Letter_6", "PackageLabel_Letter_6_CarrierLeft", "PackageLabel_A4_2", "PackageLabel_A4_4", "PackageLabel_Plain_Paper", "PackageLabel_Plain_Paper_CarrierBottom", "PackageLabel_Thermal", "PackageLabel_Thermal_Unified", "PackageLabel_Thermal_NonPCP", "PackageLabel_Thermal_No_Carrier_Rotation"}, "shipmentPageType")
        Dim q As New List(Of KeyValuePair(Of String, String)) From {QPair("PageType", pageType), QPair("LabelType", labelType)}
        AddOptionalInteger(q, "NumberOfPackages", "numberOfPackages", 1)
        AddOptionalInteger(q, "NumberOfPallets", "numberOfPallets", 1)
        AddOptionalInteger(q, "PageSize", "shipmentPageSize", 1, 1000)
        AddOptionalInteger(q, "PageStartIndex", "pageStartIndex", 0)
        AddCsvParam(q, "PackageLabelsToPrint", S("packageLabelsToPrint"), 1000)
        Return Await CallSpApiAsync("/fba/inbound/v0/shipments/" & Encode(Required("shipmentId")) & "/labels?" & BuildQuery(q))
    End Function

    ' -------------------- Documents and Amazon business outcomes --------------------
    Private Async Function GetAndDownloadDocumentAsync(path As String, label As String) As Task(Of ApiResult)
        Dim metadata = Await CallSpApiAsync(path)
        If Not metadata.Ok Then Return metadata
        Dim doc = AsDict(metadata.Data)
        Dim url = StringValue(GetValue(doc, "url"))
        If url = "" Then Throw New AppException("Amazon returned no URL for the " & label, 502, "DOCUMENT_URL_MISSING")
        ValidateAmazonDocumentUrl(url, label)
        Dim downloaded As New Dictionary(Of String, Object) From {{"contentType", Nothing}, {"bytesRead", 0}, {"truncated", False}, {"content", ""}, {"error", Nothing}}
        Try
            Using timeout As New CancellationTokenSource(TimeSpan.FromSeconds(30))
                Using response = Await Http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, timeout.Token)
                    If Not response.IsSuccessStatusCode Then
                        downloaded("error") = "Preview failed with HTTP " & CInt(response.StatusCode).ToString() & " " & response.ReasonPhrase & ". The original URL is still available."
                    Else
                        downloaded("contentType") = If(response.Content.Headers.ContentType Is Nothing, Nothing, response.Content.Headers.ContentType.ToString())
                        downloaded("contentDisposition") = If(response.Content.Headers.ContentDisposition Is Nothing, Nothing, response.Content.Headers.ContentDisposition.ToString())
                        Dim preview = Await ReadPreviewAsync(response, PreviewLimit, timeout.Token)
                        downloaded("bytesRead") = preview.Item2
                        downloaded("truncated") = preview.Item3
                        downloaded("content") = preview.Item1
                    End If
                End Using
            End Using
        Catch ex As OperationCanceledException
            downloaded("error") = "Preview timed out after 30 seconds. The original URL is still available."
        Catch ex As Exception
            downloaded("error") = "Preview failed: " & ex.Message & ". The original URL is still available."
        End Try
        doc("downloaded") = downloaded
        metadata.Data = doc
        Return metadata
    End Function

    Private Async Function ReadPreviewAsync(response As HttpResponseMessage, maxBytes As Integer, cancellationToken As CancellationToken) As Task(Of Tuple(Of String, Integer, Boolean))
        Using input = Await response.Content.ReadAsStreamAsync()
            Using ms As New MemoryStream()
                Dim buffer(8191) As Byte
                Dim total As Integer = 0
                Dim truncated As Boolean = False
                Do
                    Dim remaining = maxBytes - total
                    If remaining <= 0 Then truncated = True : Exit Do
                    Dim count = Await input.ReadAsync(buffer, 0, Math.Min(buffer.Length, remaining), cancellationToken)
                    If count = 0 Then Exit Do
                    ms.Write(buffer, 0, count)
                    total += count
                    If total >= maxBytes Then
                        Dim extra(0) As Byte
                        Dim extraCount = Await input.ReadAsync(extra, 0, 1, cancellationToken)
                        truncated = extraCount > 0
                        Exit Do
                    End If
                Loop
                Return Tuple.Create(Encoding.UTF8.GetString(ms.ToArray()), total, truncated)
            End Using
        End Using
    End Function

    Private Function ApplyBusinessOutcome(operation As String, result As ApiResult) As ApiResult
        If Not result.Ok Then Return result
        Dim data = AsDict(result.Data)
        If operation = "feed" Then
            Dim status = StringValue(GetValue(data, "processingStatus"))
            If status = "FATAL" OrElse status = "CANCELLED" Then
                result.Ok = False : result.Status = 422 : result.StatusText = "Feed processing " & status.ToLowerInvariant()
                result.Problem = New ApiProblem With {.Code = "FEED_" & status, .Message = If(status = "FATAL", "Amazon aborted the feed during processing. Some records may or may not have been applied.", "Amazon cancelled the feed before processing completed."), .Details = If(StringValue(GetValue(data, "resultFeedDocumentId")) = "", "", "resultFeedDocumentId: " & StringValue(GetValue(data, "resultFeedDocumentId"))), .Action = "Review the feed processing report before resubmitting.", .Retryable = False}
            ElseIf status = "DONE" Then
                data("nextStep") = "Use Feed processing report with resultFeedDocumentId before treating individual records as successful."
            ElseIf status = "IN_QUEUE" OrElse status = "IN_PROGRESS" Then
                data("nextStep") = "The feed is still processing. Poll Feed status again later."
            End If
        ElseIf operation = "report" Then
            Dim status = StringValue(GetValue(data, "processingStatus"))
            If status = "FATAL" OrElse status = "CANCELLED" Then
                result.Ok = False : result.Status = 422 : result.StatusText = "Report processing " & status.ToLowerInvariant()
                result.Problem = New ApiProblem With {.Code = "REPORT_" & status, .Message = If(status = "FATAL", "Amazon could not complete the report job.", "The report job was cancelled."), .Action = "Verify report type, date range, marketplace, and role before creating another report.", .Retryable = False}
            ElseIf status = "DONE" Then
                data("nextStep") = "Use Report document with reportDocumentId to download and inspect the generated report."
            ElseIf status = "IN_QUEUE" OrElse status = "IN_PROGRESS" Then
                data("nextStep") = "The report is still processing. Poll Report status again later."
            End If
        ElseIf operation = "inboundOperationStatus" Then
            Dim status = StringValue(GetValue(data, "operationStatus"))
            If status = "FAILED" Then
                Dim problems = ListValue(GetValue(data, "operationProblems"))
                Dim firstProblem = If(problems.Count > 0, AsDict(problems(0)), New Dictionary(Of String, Object)())
                Dim problemCode = StringValue(GetValue(firstProblem, "code"))
                Dim problemMessage = StringValue(GetValue(firstProblem, "message"))
                If problemCode = "" Then problemCode = "INBOUND_OPERATION_FAILED"
                If problemMessage = "" Then problemMessage = "The asynchronous Fulfillment Inbound operation failed."

                result.Ok = False : result.Status = 422 : result.StatusText = "Inbound operation failed"
                result.Problem = New ApiProblem With {
                    .Code = problemCode,
                    .Message = problemMessage,
                    .Details = Json(GetValue(data, "operationProblems")),
                    .Action = "Read every operationProblem, correct the inbound data it identifies, then start a new valid operation. Do not assume the original write was applied.",
                    .Retryable = False
                }
            ElseIf status = "IN_PROGRESS" Then
                data("nextStep") = "Poll Operation status again before continuing."
            ElseIf status = "SUCCESS" AndAlso ListValue(GetValue(data, "operationProblems")).Count > 0 Then
                data("businessWarnings") = GetValue(data, "operationProblems")
                data("nextStep") = "The operation succeeded, but Amazon returned warnings. Review operationProblems."
            End If
        ElseIf operation = "createInboundPlan" AndAlso StringValue(GetValue(data, "operationId")) <> "" Then
            data("nextStep") = "Use Operation status with operationId and wait for SUCCESS before continuing."
        ElseIf operation = "createReport" AndAlso StringValue(GetValue(data, "reportId")) <> "" Then
            data("nextStep") = "Poll Report status with reportId until DONE, FATAL, or CANCELLED."
        ElseIf operation = "submitFeed" AndAlso StringValue(GetValue(data, "feedId")) <> "" Then
            data("nextStep") = "Poll Feed status with feedId until DONE, FATAL, or CANCELLED, then inspect resultFeedDocumentId."
        ElseIf operation = "itemLabels" Then
            data("nextStep") = "Open or download each returned documentDownloads URL before it expires."
        ElseIf operation = "shipmentLabels" OrElse operation = "billOfLading" Then
            data("nextStep") = "Open or download Amazon's returned DownloadURL before it expires."
        End If
        result.Data = data
        Return result
    End Function

    ' -------------------- LWA authentication and SP-API transport --------------------
    Private Async Function GetAccessTokenAsync(Optional forceRefresh As Boolean = False) As Task(Of Tuple(Of String, Integer))
        If Not forceRefresh AndAlso CachedAccessToken <> "" AndAlso CachedAccessTokenExpiresUtc > DateTimeOffset.UtcNow.AddSeconds(60) Then
            Dim remaining = Math.Max(1, CInt((CachedAccessTokenExpiresUtc - DateTimeOffset.UtcNow).TotalSeconds))
            Return Tuple.Create(CachedAccessToken, remaining)
        End If

        Dim pairs As New Dictionary(Of String, String) From {
            {"grant_type", "refresh_token"}, {"refresh_token", txtRefreshToken.Text.Trim()}, {"client_id", txtClientId.Text.Trim()}, {"client_secret", txtClientSecret.Text.Trim()}
        }
        Dim outcome = Await SendWithRetryAsync(Function()
                                                   Dim req As New HttpRequestMessage(HttpMethod.Post, "https://api.amazon.com/auth/o2/token")
                                                   req.Content = New FormUrlEncodedContent(pairs)
                                                   Return req
                                               End Function, RetryMode.SafePost)
        Using response = outcome.Item1
            Dim data = ParseJson(Await response.Content.ReadAsStringAsync())
            If Not response.IsSuccessStatusCode Then
                CachedAccessToken = ""
                CachedAccessTokenExpiresUtc = DateTimeOffset.MinValue
                Throw New AppException(ExtractMessage(data, "Amazon rejected the supplied LWA credentials"), CInt(response.StatusCode), ExtractCode(data, "LWA_AUTH_FAILED"), Json(data))
            End If

            Dim dict = AsDict(data)
            Dim token = StringValue(GetValue(dict, "access_token"))
            If token = "" Then Throw New AppException("Amazon returned no access token", 502, "LWA_TOKEN_MISSING", Json(data))

            Dim expires As Integer = 3600
            Integer.TryParse(Convert.ToString(GetValue(dict, "expires_in"), CultureInfo.InvariantCulture), expires)
            If expires < 1 Then expires = 3600

            CachedAccessToken = token
            CachedAccessTokenExpiresUtc = DateTimeOffset.UtcNow.AddSeconds(expires)
            Return Tuple.Create(token, expires)
        End Using
    End Function

    Private Async Function CallSpApiAsync(path As String, Optional method As HttpMethod = Nothing, Optional body As Object = Nothing, Optional accessToken As String = "") As Task(Of ApiResult)
        If method Is Nothing Then method = HttpMethod.Get
        If accessToken = "" Then accessToken = (Await GetAccessTokenAsync()).Item1
        Dim url = Endpoint() & path
        Dim retryPolicy = If(method = HttpMethod.Get, RetryMode.ReadRequest, RetryMode.WriteRequest)
        Dim sw = Stopwatch.StartNew()
        Dim outcome As Tuple(Of HttpResponseMessage, Integer)
        Try
            outcome = Await SendWithRetryAsync(Function()
                                                   Dim req As New HttpRequestMessage(method, url)
                                                   req.Headers.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
                                                   req.Headers.TryAddWithoutValidation("x-amz-access-token", accessToken)
                                                   req.Headers.TryAddWithoutValidation("x-amz-date", DateTime.UtcNow.ToString("yyyyMMdd'T'HHmmss'Z'", CultureInfo.InvariantCulture))
                                                   req.Headers.TryAddWithoutValidation("User-Agent", "SP-API-Workbench/1.2 (Language=VB.NET; Platform=.NET Framework 4.8)")
                                                   If body IsNot Nothing Then req.Content = New StringContent(Serializer.Serialize(body), Encoding.UTF8, "application/json")
                                                   Return req
                                               End Function, retryPolicy)
        Catch ex As Exception
            sw.Stop()
            Throw
        End Try

        Using response = outcome.Item1
            sw.Stop()
            Dim raw = Await response.Content.ReadAsStringAsync()
            Dim data = ParseJson(raw)
            Dim result As New ApiResult With {
                .Ok = response.IsSuccessStatusCode, .Status = CInt(response.StatusCode), .StatusText = response.ReasonPhrase,
                .RequestId = Header(response, "x-amzn-requestid"), .GatewayId = Header(response, "x-amz-apigw-id"),
                .TraceId = Header(response, "x-amzn-trace-id"), .RateLimit = Header(response, "x-amzn-ratelimit-limit"),
                .RequestMethod = method.Method, .RequestPath = path,
                .Data = data, .DurationMs = sw.ElapsedMilliseconds, .Attempts = outcome.Item2
            }
            If Not result.Ok Then
                If result.Status = 401 Then
                    CachedAccessToken = ""
                    CachedAccessTokenExpiresUtc = DateTimeOffset.MinValue
                End If
                result.Problem = BuildProblem(result.Status, result.StatusText, data, method, Header(response, "x-amzn-errortype"))
            End If
            Return result
        End Using
    End Function

    Private Async Function SendWithRetryAsync(factory As Func(Of HttpRequestMessage), mode As RetryMode) As Task(Of Tuple(Of HttpResponseMessage, Integer))
        Const maxAttempts As Integer = 4
        For attempt As Integer = 1 To maxAttempts
            Dim response As HttpResponseMessage = Nothing
            Dim networkFailure As Exception = Nothing
            Dim timedOut As Boolean = False

            Try
                Using request = factory()
                    response = Await Http.SendAsync(request)
                End Using
            Catch ex As TaskCanceledException
                networkFailure = ex
                timedOut = True
            Catch ex As HttpRequestException
                networkFailure = ex
            End Try

            If networkFailure IsNot Nothing Then
                If mode <> RetryMode.WriteRequest AndAlso attempt < maxAttempts Then
                    Await Task.Delay(ClampDelay(750 * CInt(Math.Pow(2, attempt - 1))))
                    Continue For
                End If

                If mode = RetryMode.WriteRequest Then
                    Throw New AppException(
                        "The connection failed while sending an Amazon write request. Amazon may have received it even though no response reached the app.",
                        502,
                        "AMBIGUOUS_WRITE_RESULT",
                        networkFailure.Message)
                End If

                If timedOut Then Throw New AppException("The request to Amazon timed out.", 502, "AMAZON_TIMEOUT", networkFailure.Message)
                Throw New AppException("The app could not reach Amazon.", 502, "AMAZON_NETWORK_ERROR", networkFailure.Message)
            End If

            If response Is Nothing Then Continue For
            If Not ShouldRetry(CInt(response.StatusCode), mode) OrElse attempt = maxAttempts Then Return Tuple.Create(response, attempt)

            Dim delay = RetryDelay(response, attempt)
            response.Dispose()
            Await Task.Delay(delay)
        Next

        Throw New AppException("Amazon request did not produce a response", 502, "NO_RESPONSE")
    End Function

    Private Function ShouldRetry(status As Integer, mode As RetryMode) As Boolean
        If status = 429 Then Return True
        If status = 500 OrElse status = 502 OrElse status = 503 OrElse status = 504 Then Return mode <> RetryMode.WriteRequest
        Return False
    End Function

    Private Function RetryDelay(response As HttpResponseMessage, attempt As Integer) As Integer
        Dim retryAfter As IEnumerable(Of String) = Nothing
        If response.Headers.TryGetValues("Retry-After", retryAfter) Then
            Dim raw = retryAfter.FirstOrDefault()
            Dim seconds As Double
            If Double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, seconds) Then Return ClampDelay(CInt(seconds * 1000))

            Dim retryAt As DateTimeOffset
            If DateTimeOffset.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces Or DateTimeStyles.AssumeUniversal, retryAt) Then
                Return ClampDelay(CInt((retryAt.ToUniversalTime() - DateTimeOffset.UtcNow).TotalMilliseconds))
            End If
        End If
        If CInt(response.StatusCode) = 429 Then
            Dim rateRaw As IEnumerable(Of String) = Nothing
            If response.Headers.TryGetValues("x-amzn-ratelimit-limit", rateRaw) Then
                Dim rate As Double
                If Double.TryParse(rateRaw.FirstOrDefault(), NumberStyles.Any, CultureInfo.InvariantCulture, rate) AndAlso rate > 0 Then Return ClampDelay(CInt(1000 / rate))
            End If
        End If
        Return ClampDelay(750 * CInt(Math.Pow(2, attempt - 1)))
    End Function

    Private Function ClampDelay(value As Integer) As Integer
        Return Math.Min(Math.Max(value, 500), 15000)
    End Function

    Private Function BuildProblem(status As Integer, statusText As String, data As Object, method As HttpMethod, Optional headerCode As String = "") As ApiProblem
        Dim code = ExtractCode(data, "")
        If code = "" AndAlso headerCode <> "" Then code = headerCode.Split(":"c)(0)
        If code = "" Then code = "HTTP_" & status.ToString(CultureInfo.InvariantCulture)
        Dim message = ExtractMessage(data, If(statusText = "", "Amazon returned HTTP " & status.ToString(), statusText))
        Dim retryable = status = 429 OrElse (method = HttpMethod.Get AndAlso {500, 502, 503, 504}.Contains(status))
        Return New ApiProblem With {.Code = code, .Message = message, .Details = ExtractDetails(data), .Action = RecommendedAction(code, status, message, retryable, method), .Retryable = retryable}
    End Function

    Private Function RecommendedAction(code As String, status As Integer, message As String, retryable As Boolean, method As HttpMethod) As String
        Dim key = (code & " " & message).ToLowerInvariant()
        If key.Contains("invalid_grant") OrElse key.Contains("refresh token") Then Return "Reconnect/self-authorize the seller account to obtain a fresh refresh token."
        If key.Contains("invalid_client") Then Return "Verify the LWA client ID and client secret belong to the same SP-API app."
        If key.Contains("expiredtoken") OrElse key.Contains("expired token") Then Return "Request a fresh LWA access token. If a fresh token is also rejected, reauthorize the seller account."
        If status = 401 OrElse status = 403 OrElse key.Contains("unauthorized") OrElse key.Contains("accessdenied") OrElse key.Contains("access denied") Then Return "Verify seller authorization and the Amazon role required by this operation. Reauthorize after changing roles."
        If key.Contains("could not match input arguments") OrElse key.Contains("sandbox request") Then Return "Amazon's static Sandbox accepts predefined request examples. Use the exact Sandbox values shown in the app."
        If status = 429 OrElse key.Contains("throttl") OrElse key.Contains("quota") Then Return "Slow the request rate and honor Retry-After / x-amzn-RateLimit-Limit before retrying."
        If status = 400 OrElse key.Contains("badrequest") OrElse key.Contains("invalidinput") OrElse key.Contains("invalid input") Then Return "Check required fields, IDs, date ranges, enum values, marketplace, and URL encoding."
        If status = 404 Then Return "Verify the resource ID, marketplace, and environment. Sandbox IDs and Production IDs are not interchangeable."
        If status = 409 Then Return "Re-read the current Amazon resource state, then retry only if the requested transition is still valid."
        If status = 413 Then Return "Reduce the request or document size and submit it in smaller supported batches."
        If status = 415 Then Return "Use the Content-Type required by this operation and make sure the body format matches it."
        If status = 422 Then Return "The request is syntactically valid but violates an Amazon business rule. Read the error details, correct the data/state, and retry."
        If status >= 500 AndAlso method <> HttpMethod.Get Then Return "Amazon returned a server error for a write. Do not immediately resubmit it; verify the related Amazon resource/job first."
        If retryable OrElse status >= 500 Then Return "Retry with backoff. Keep the Amazon request ID if the error persists."
        Return "Read the Amazon error details and request ID, correct the condition, then retry."
    End Function

    Private Function ExtractCode(data As Object, fallback As String) As String
        Dim d = AsDict(data)
        Dim errors = ListValue(GetValue(d, "errors"))
        If errors.Count > 0 Then
            Dim first = AsDict(errors(0))
            Dim c = StringValue(GetValue(first, "code"))
            If c <> "" Then Return c
        End If
        For Each key In {"error", "code"}
            Dim c = StringValue(GetValue(d, key))
            If c <> "" Then Return c
        Next
        Return fallback
    End Function

    Private Function ExtractMessage(data As Object, fallback As String) As String
        Dim d = AsDict(data)
        Dim errors = ListValue(GetValue(d, "errors"))
        If errors.Count > 0 Then
            Dim first = AsDict(errors(0))
            Dim m = StringValue(GetValue(first, "message"))
            If m <> "" Then Return m
        End If
        For Each key In {"error_description", "message", "error"}
            Dim m = StringValue(GetValue(d, key))
            If m <> "" Then Return m
        Next
        Return fallback
    End Function

    Private Function ExtractDetails(data As Object) As String
        Dim d = AsDict(data)
        Dim errors = ListValue(GetValue(d, "errors"))
        If errors.Count > 0 Then
            Dim first = AsDict(errors(0))
            If first.ContainsKey("details") Then
                Dim detailValue = GetValue(first, "details")
                If TypeOf detailValue Is String Then Return CStr(detailValue)
                Return Json(detailValue)
            End If
        End If
        If d.ContainsKey("details") Then
            Dim detailValue = GetValue(d, "details")
            If TypeOf detailValue Is String Then Return CStr(detailValue)
            Return Json(detailValue)
        End If
        Return ""
    End Function

    ' -------------------- Request validation and helpers --------------------
    Private Sub ValidateAmazonDocumentUrl(value As String, label As String)
        Dim uri As Uri = Nothing
        If Not Uri.TryCreate(value, UriKind.Absolute, uri) Then Throw New AppException("Amazon returned an invalid URL for the " & label, 502, "INVALID_DOCUMENT_URL")
        Dim host = uri.Host.ToLowerInvariant()
        Dim allowed = host = "amazonaws.com" OrElse host.EndsWith(".amazonaws.com", StringComparison.Ordinal) OrElse host = "cloudfront.net" OrElse host.EndsWith(".cloudfront.net", StringComparison.Ordinal)
        If uri.Scheme <> Uri.UriSchemeHttps OrElse Not allowed Then Throw New AppException("Amazon returned an unexpected host for the " & label, 502, "UNEXPECTED_DOCUMENT_HOST", host)
    End Sub

    Private Sub ValidateMarketplaceRegions(ids As List(Of String), key As String)
        If IsSandbox() OrElse ids.Count = 0 Then Return
        Dim selected = SelectedMarketplace()
        Dim unknown = ids.Where(Function(id) Not Marketplaces.Any(Function(m) m.Id = id)).ToList()
        If unknown.Count > 0 Then Throw New AppException(key & " contains unsupported marketplace IDs: " & String.Join(",", unknown), 400, "UNSUPPORTED_MARKETPLACE")
        Dim wrong = ids.Where(Function(id) Marketplaces.First(Function(m) m.Id = id).Region <> selected.Region).ToList()
        If wrong.Count > 0 Then Throw New AppException(key & " must use marketplaces in the same SP-API selling region as the selected marketplace", 400, "MARKETPLACE_REGION_MISMATCH", String.Join(",", wrong))
    End Sub

    Private Sub ValidateContentType(contentType As String)
        Dim parsed As MediaTypeHeaderValue = Nothing
        If Not MediaTypeHeaderValue.TryParse(contentType, parsed) OrElse parsed Is Nothing OrElse String.IsNullOrWhiteSpace(parsed.MediaType) Then
            Throw New AppException("contentType is not a valid HTTP media type", 400, "INVALID_CONTENT_TYPE", contentType)
        End If
    End Sub

    Private Sub ValidateJsonListingsFeed(contentType As String, content As String)
        If Not Regex.IsMatch(contentType.Trim(), "^application/json\b", RegexOptions.IgnoreCase) Then Throw New AppException("JSON_LISTINGS_FEED requires an application/json content type", 400, "INVALID_JSON_LISTINGS_CONTENT_TYPE")
        Dim parsed As Object
        Try
            parsed = Serializer.DeserializeObject(content)
        Catch ex As Exception
            Throw New AppException("Feed content is not valid JSON", 400, "INVALID_FEED_JSON", ex.Message)
        End Try
        Dim feed = AsDict(parsed)
        Dim header = AsDict(GetValue(feed, "header"))
        Dim messages = ListValue(GetValue(feed, "messages"))
        If StringValue(GetValue(header, "sellerId")) = "" OrElse StringValue(GetValue(header, "version")) = "" OrElse messages.Count = 0 Then Throw New AppException("JSON_LISTINGS_FEED requires header.sellerId, header.version, and at least one message", 400, "INVALID_JSON_LISTINGS_STRUCTURE")
        For i As Integer = 0 To messages.Count - 1
            Dim msg = AsDict(messages(i))
            Dim id As Integer
            If Not Integer.TryParse(Convert.ToString(GetValue(msg, "messageId"), CultureInfo.InvariantCulture), id) OrElse id < 1 OrElse StringValue(GetValue(msg, "sku")) = "" OrElse StringValue(GetValue(msg, "operationType")) = "" Then Throw New AppException("Each JSON listings message requires a positive integer messageId, sku, and operationType", 400, "INVALID_JSON_LISTINGS_MESSAGE", "messageIndex=" & i.ToString())
        Next
    End Sub

    Private Function ParseItems(value As String, maxItems As Integer, maxQuantity As Integer) As List(Of Object)
        Dim output As New List(Of Object)()
        For Each rawLine In value.Split({ControlChars.Cr, ControlChars.Lf}, StringSplitOptions.RemoveEmptyEntries)
            Dim line = rawLine.Trim()
            If line = "" Then Continue For
            Dim parts = ParseCsvLine(line)
            If parts.Count > 6 Then Throw New AppException("Each item must use: MSKU, quantity, prep owner, label owner[, expiration, manufacturing lot code]", 400, "INVALID_ITEM_ROW", line)
            While parts.Count < 6 : parts.Add("") : End While
            Dim msku = parts(0)
            Dim quantity As Integer
            If msku = "" OrElse msku.Length > 255 OrElse Not Integer.TryParse(parts(1), quantity) OrElse quantity < 1 OrElse quantity > maxQuantity Then Throw New AppException("Each item needs an MSKU up to 255 characters and an integer quantity between 1 and " & maxQuantity.ToString(), 400, "INVALID_ITEM_ROW", line)
            Dim prepOwner = If(parts(2) = "", "SELLER", parts(2)).ToUpperInvariant()
            Dim labelOwner = If(parts(3) = "", "SELLER", parts(3)).ToUpperInvariant()
            ValidateEnum(New List(Of String) From {prepOwner}, {"AMAZON", "SELLER", "NONE"}, "prepOwner")
            ValidateEnum(New List(Of String) From {labelOwner}, {"AMAZON", "SELLER", "NONE"}, "labelOwner")
            Dim item As New Dictionary(Of String, Object) From {{"msku", msku}, {"quantity", quantity}, {"prepOwner", prepOwner}, {"labelOwner", labelOwner}}
            If parts(4) <> "" Then
                If Not Regex.IsMatch(parts(4), "^\d{4}-\d{2}-\d{2}$") Then Throw New AppException("Expiration must use a real YYYY-MM-DD date", 400, "INVALID_EXPIRATION_DATE", line)
                Dim dateCheck As DateTime
                If Not DateTime.TryParseExact(parts(4), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, dateCheck) Then Throw New AppException("Expiration must use a real YYYY-MM-DD date", 400, "INVALID_EXPIRATION_DATE", line)
                item("expiration") = parts(4)
            End If
            If parts(5).Length > 256 Then Throw New AppException("Manufacturing lot code must be at most 256 characters", 400, "LOT_CODE_TOO_LONG", line)
            If parts(5) <> "" Then item("manufacturingLotCode") = parts(5)
            output.Add(item)
        Next
        If output.Count = 0 Then Throw New AppException("Add at least one item", 400, "NO_ITEMS")
        If output.Count > maxItems Then Throw New AppException("This operation accepts at most " & maxItems.ToString() & " items", 400, "TOO_MANY_ITEMS")
        Return output
    End Function

    Private Function ParseCsvLine(line As String) As List(Of String)
        Dim fields As New List(Of String)()
        Dim current As New StringBuilder()
        Dim quoted As Boolean = False
        Dim i As Integer = 0
        While i < line.Length
            Dim ch = line(i)
            If ch = """"c Then
                If quoted AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then current.Append(""""c) : i += 2 : Continue While
                quoted = Not quoted
            ElseIf ch = ","c AndAlso Not quoted Then
                fields.Add(current.ToString().Trim()) : current.Clear()
            Else
                current.Append(ch)
            End If
            i += 1
        End While
        If quoted Then Throw New AppException("Unclosed quoted CSV field", 400, "INVALID_ITEM_ROW", line)
        fields.Add(current.ToString().Trim())
        Return fields
    End Function

    Private Function Required(key As String) As String
        Dim value = S(key)
        If value = "" Then Throw New AppException(key & " is required", 400, "MISSING_REQUIRED_FIELD")
        Return value
    End Function

    Private Function RequiredDate(key As String) As String
        Dim value = Required(key)
        If Not ValidIsoInstant(value) Then Throw New AppException(key & " must be an ISO 8601 timestamp with an explicit timezone, for example 2026-09-18T12:00:00Z", 400, "INVALID_DATE")
        Return value
    End Function

    Private Function OptionalDate(key As String) As String
        Dim value = S(key)
        If value = "" Then Return ""
        If Not ValidIsoInstant(value) Then Throw New AppException(key & " must be an ISO 8601 timestamp with an explicit timezone", 400, "INVALID_DATE")
        Return value
    End Function

    Private Function ValidIsoInstant(value As String) As Boolean
        If Not Regex.IsMatch(value, "^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,9})?)?(?:Z|[+-]\d{2}:\d{2})$") Then Return False
        Dim parsed As DateTimeOffset
        Return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, parsed)
    End Function

    Private Sub ValidateRange(startValue As String, endValue As String, startLabel As String, endLabel As String)
        If startValue = "" OrElse endValue = "" Then Return
        If DateTimeOffset.Parse(endValue, CultureInfo.InvariantCulture) < DateTimeOffset.Parse(startValue, CultureInfo.InvariantCulture) Then Throw New AppException(endLabel & " must be equal to or after " & startLabel, 400, "INVALID_DATE_RANGE")
    End Sub

    Private Function LimitedRequired(key As String, maxLength As Integer) As String
        Dim value = Required(key)
        If value.Length > maxLength Then Throw New AppException(key & " must be at most " & maxLength.ToString() & " characters", 400, "VALUE_TOO_LONG")
        Return value
    End Function

    Private Function LimitedOptional(key As String, maxLength As Integer) As String
        Dim value = S(key)
        If value.Length > maxLength Then Throw New AppException(key & " must be at most " & maxLength.ToString() & " characters", 400, "VALUE_TOO_LONG")
        Return value
    End Function

    Private Function IntField(key As String, min As Integer, max As Integer, Optional fallback As Integer = Integer.MinValue) As Integer
        Dim raw = S(key)
        If raw = "" AndAlso fallback <> Integer.MinValue Then Return fallback
        Dim value As Integer
        If Not Integer.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, value) OrElse value < min OrElse value > max Then Throw New AppException(key & " must be an integer between " & min.ToString() & " and " & max.ToString(), 400, "INVALID_NUMBER")
        Return value
    End Function

    Private Function DecimalField(key As String, min As Decimal, max As Decimal, Optional fallback As Decimal = Decimal.MinValue) As Decimal
        Dim raw = S(key)
        If raw = "" AndAlso fallback <> Decimal.MinValue Then Return fallback
        Dim value As Decimal
        If Not Decimal.TryParse(raw, NumberStyles.Number, CultureInfo.InvariantCulture, value) OrElse value < min OrElse value > max Then Throw New AppException(key & " must be a number between " & min.ToString(CultureInfo.InvariantCulture) & " and " & max.ToString(CultureInfo.InvariantCulture), 400, "INVALID_NUMBER")
        Return value
    End Function

    Private Function SplitValues(value As String, max As Integer) As List(Of String)
        If String.IsNullOrWhiteSpace(value) Then Return New List(Of String)()
        Dim values = Regex.Split(value, "[\r\n,]+").Select(Function(x) x.Trim()).Where(Function(x) x <> "").ToList()
        If values.Count > max Then Throw New AppException("At most " & max.ToString() & " values are allowed", 400, "TOO_MANY_VALUES")
        Return values
    End Function

    Private Sub ValidateEnum(values As List(Of String), allowedValues As IEnumerable(Of String), key As String)
        Dim allowed As New HashSet(Of String)(allowedValues, StringComparer.Ordinal)
        Dim invalid = values.Where(Function(x) Not allowed.Contains(x)).ToList()
        If invalid.Count > 0 Then Throw New AppException(key & " contains unsupported value(s): " & String.Join(",", invalid), 400, "INVALID_ENUM_VALUE")
    End Sub

    Private Sub AddEnumParam(q As List(Of KeyValuePair(Of String, String)), key As String, value As String, allowed As IEnumerable(Of String))
        If value = "" Then Return
        ValidateEnum(New List(Of String) From {value}, allowed, key)
        q.Add(QPair(key, value))
    End Sub

    Private Sub AddOptionalInteger(q As List(Of KeyValuePair(Of String, String)), queryKey As String, fieldKey As String, min As Integer, Optional max As Integer = Integer.MaxValue)
        If S(fieldKey) = "" Then Return
        q.Add(QPair(queryKey, IntField(fieldKey, min, max).ToString(CultureInfo.InvariantCulture)))
    End Sub

    Private Sub AddCsvParam(q As List(Of KeyValuePair(Of String, String)), key As String, value As String, max As Integer)
        Dim values = SplitValues(value, max)
        If values.Count > 0 Then q.Add(QPair(key, String.Join(",", values)))
    End Sub

    Private Sub AddOptional(q As List(Of KeyValuePair(Of String, String)), key As String, value As String)
        If value <> "" Then q.Add(QPair(key, value))
    End Sub

    Private Sub AddOptionalObject(dict As Dictionary(Of String, Object), key As String, value As String)
        If value <> "" Then dict(key) = value
    End Sub

    Private Function QPair(key As String, value As String) As KeyValuePair(Of String, String)
        Return New KeyValuePair(Of String, String)(key, value)
    End Function

    Private Function BuildQuery(items As IEnumerable(Of KeyValuePair(Of String, String))) As String
        Return String.Join("&", items.Select(Function(p) Encode(p.Key) & "=" & Encode(p.Value)))
    End Function

    Private Function Encode(value As String) As String
        Return Uri.EscapeDataString(value)
    End Function

    Private Function Money(currency As String, amount As Decimal) As Dictionary(Of String, Object)
        Return New Dictionary(Of String, Object) From {{"CurrencyCode", currency}, {"Amount", amount}}
    End Function

    Private Function Header(response As HttpResponseMessage, name As String) As String
        Dim values As IEnumerable(Of String) = Nothing
        If response.Headers.TryGetValues(name, values) Then Return values.FirstOrDefault()
        Return ""
    End Function

    Private Function ParseJson(raw As String) As Object
        If String.IsNullOrWhiteSpace(raw) Then Return Nothing
        Try
            Return Serializer.DeserializeObject(raw)
        Catch
            Return New Dictionary(Of String, Object) From {{"message", raw}}
        End Try
    End Function

    Private Function AsDict(value As Object) As Dictionary(Of String, Object)
        Dim d = TryCast(value, Dictionary(Of String, Object))
        If d IsNot Nothing Then Return d
        Return New Dictionary(Of String, Object)()
    End Function

    Private Function ListValue(value As Object) As List(Of Object)
        If value Is Nothing Then Return New List(Of Object)()
        Dim arr = TryCast(value, Object())
        If arr IsNot Nothing Then Return arr.ToList()
        Dim list = TryCast(value, ArrayList)
        If list IsNot Nothing Then Return list.Cast(Of Object)().ToList()
        Dim generic = TryCast(value, IEnumerable(Of Object))
        If generic IsNot Nothing Then Return generic.ToList()
        Return New List(Of Object)()
    End Function

    Private Function GetValue(dict As Dictionary(Of String, Object), key As String) As Object
        Dim value As Object = Nothing
        If dict.TryGetValue(key, value) Then Return value
        Return Nothing
    End Function

    Private Function StringValue(value As Object) As String
        If TypeOf value Is String Then Return CStr(value)
        Return ""
    End Function

    Private Function Json(value As Object) As String
        If value Is Nothing Then Return ""
        Try
            Return Serializer.Serialize(value)
        Catch
            Return Convert.ToString(value, CultureInfo.InvariantCulture)
        End Try
    End Function

    ' -------------------- Error normalization for the view --------------------
    Private Function LocalFailure(ex As AppException) As ApiResult
        Dim action = "Correct the request values and run the operation again."
        Dim retryable = False
        Dim key = (ex.Code & " " & ex.Message).ToLowerInvariant()

        If key.Contains("ambiguous_write_result") Then
            action = "Do not submit the write again yet. First check the related Amazon job/resource to see whether the original request was applied."
        ElseIf key.Contains("amazon_timeout") Then
            action = "Retry the read/connection test. If it repeats, check the Windows network/proxy path to Amazon."
            retryable = True
        ElseIf key.Contains("amazon_network_error") OrElse key.Contains("connection") OrElse key.Contains("ssl") OrElse key.Contains("certificate") Then
            action = "Check Windows internet access, system proxy, and trusted certificate chain. The app uses Windows certificate trust and does not disable TLS validation."
            retryable = True
        ElseIf key.Contains("invalid_grant") OrElse key.Contains("refresh token") Then
            action = "Reauthorize the seller account and use the new refresh token."
        ElseIf key.Contains("invalid_client") OrElse key.Contains("client authentication") Then
            action = "Verify the LWA Client ID and Client Secret belong to the same SP-API application."
        ElseIf ex.Status = 401 OrElse ex.Status = 403 Then
            action = "Verify seller authorization and the Amazon role/permission required by this operation."
        ElseIf key.Contains("document") AndAlso key.Contains("url") Then
            action = "Do not open the returned URL. Keep the Amazon request ID and retry the document metadata request."
        ElseIf key.Contains("no_response") Then
            action = "Retry after checking Windows network/proxy connectivity to Amazon."
            retryable = True
        End If

        Dim statusText = If(ex.Status >= 500, "Local / network error", "Local validation error")
        If key.Contains("invalid_grant") OrElse key.Contains("invalid_client") OrElse key.Contains("lwa_") OrElse key.Contains("refresh token") Then
            statusText = "Amazon LWA authentication error"
        ElseIf ex.Code = "AMBIGUOUS_WRITE_RESULT" Then
            statusText = "Ambiguous Amazon write result"
        End If

        Return New ApiResult With {
            .Ok = False, .Status = ex.Status, .StatusText = statusText, .ErrorMessage = ex.Message,
            .Problem = New ApiProblem With {.Code = ex.Code, .Message = ex.Message, .Details = ex.Details, .Action = action, .Retryable = retryable}
        }
    End Function

End Class
