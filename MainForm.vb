Option Explicit On
Option Strict On
Option Infer On

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms

Module Program
    <STAThread>
    Public Sub Main()
        ServicePointManager.SecurityProtocol = SecurityProtocolType.SystemDefault
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        If String.Equals(Environment.GetEnvironmentVariable("SP_API_CI_SMOKE"), "1", StringComparison.Ordinal) Then
            Try
                Using form As New MainForm()
                    Dim handle = form.Handle
                    form.RunCiSelfTest()
                End Using
            Catch ex As Exception
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "smoke-error.txt"), ex.ToString())
                Environment.ExitCode = 1
            End Try
            Return
        End If

        Application.Run(New MainForm())
    End Sub
End Module

Public Partial Class MainForm
    Inherits Form

    ' View / UX only. Amazon request construction and transport live in ApiRequests.vb.

    Private Class Marketplace
        Public Property Id As String
        Public Property Name As String
        Public Property Region As String
        Public Property Currency As String
        Public Property Locale As String
        Public Overrides Function ToString() As String
            Return Name & " (" & Id & ")"
        End Function
    End Class

    Private Class ReturnedRecordAction
        Public Property Label As String = ""
        Public Property OperationId As String = ""
        Public Property Fields As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Public Overrides Function ToString() As String
            Return Label
        End Function
    End Class

    Private Class OperationInfo
        Public Property Id As String = ""
        Public Property Label As String = ""
        Public Property Kind As String = "read"
        Public Property Group As String = ""
        Public Overrides Function ToString() As String
            Return Label
        End Function
    End Class

    Private ReadOnly Marketplaces As New List(Of Marketplace) From {
        New Marketplace With {.Id = "ATVPDKIKX0DER", .Name = "United States", .Region = "na", .Currency = "USD", .Locale = "en_US"},
        New Marketplace With {.Id = "A2EUQ1WTGCTBG2", .Name = "Canada", .Region = "na", .Currency = "CAD", .Locale = "en_CA"},
        New Marketplace With {.Id = "A1AM78C64UM0Y8", .Name = "Mexico", .Region = "na", .Currency = "MXN", .Locale = "es_MX"},
        New Marketplace With {.Id = "A2Q3Y263D00KWC", .Name = "Brazil", .Region = "na", .Currency = "BRL", .Locale = "pt_BR"},
        New Marketplace With {.Id = "A28R8C7NBKEWEA", .Name = "Ireland", .Region = "eu", .Currency = "EUR", .Locale = "en_IE"},
        New Marketplace With {.Id = "A1F83G8C2ARO7P", .Name = "United Kingdom", .Region = "eu", .Currency = "GBP", .Locale = "en_GB"},
        New Marketplace With {.Id = "A1PA6795UKMFR9", .Name = "Germany", .Region = "eu", .Currency = "EUR", .Locale = "de_DE"},
        New Marketplace With {.Id = "A13V1IB3VIYZZH", .Name = "France", .Region = "eu", .Currency = "EUR", .Locale = "fr_FR"},
        New Marketplace With {.Id = "AMEN7PMS3EDWL", .Name = "Belgium", .Region = "eu", .Currency = "EUR", .Locale = "fr_BE"},
        New Marketplace With {.Id = "APJ6JRA9NG5V4", .Name = "Italy", .Region = "eu", .Currency = "EUR", .Locale = "it_IT"},
        New Marketplace With {.Id = "A1RKKUPIHCS9HS", .Name = "Spain", .Region = "eu", .Currency = "EUR", .Locale = "es_ES"},
        New Marketplace With {.Id = "A1805IZSGTT6HS", .Name = "Netherlands", .Region = "eu", .Currency = "EUR", .Locale = "nl_NL"},
        New Marketplace With {.Id = "A2NODRKZP88ZB9", .Name = "Sweden", .Region = "eu", .Currency = "SEK", .Locale = "sv_SE"},
        New Marketplace With {.Id = "AE08WJ6YKNBMC", .Name = "South Africa", .Region = "eu", .Currency = "ZAR", .Locale = "en_ZA"},
        New Marketplace With {.Id = "A1C3SOZRARQ6R3", .Name = "Poland", .Region = "eu", .Currency = "PLN", .Locale = "pl_PL"},
        New Marketplace With {.Id = "ARBP9OOSHTCHU", .Name = "Egypt", .Region = "eu", .Currency = "EGP", .Locale = "ar_EG"},
        New Marketplace With {.Id = "A33AVAJ2PDY3EV", .Name = "Turkey", .Region = "eu", .Currency = "TRY", .Locale = "tr_TR"},
        New Marketplace With {.Id = "A21TJRUUN4KGV", .Name = "India", .Region = "eu", .Currency = "INR", .Locale = "en_IN"},
        New Marketplace With {.Id = "A2VIGQ35RCS4UG", .Name = "United Arab Emirates", .Region = "eu", .Currency = "AED", .Locale = "en_AE"},
        New Marketplace With {.Id = "A17E79C6D8DWNP", .Name = "Saudi Arabia", .Region = "eu", .Currency = "SAR", .Locale = "ar_SA"},
        New Marketplace With {.Id = "A1VC38T7YXB528", .Name = "Japan", .Region = "fe", .Currency = "JPY", .Locale = "ja_JP"},
        New Marketplace With {.Id = "A39IBJ37TRP1C6", .Name = "Australia", .Region = "fe", .Currency = "AUD", .Locale = "en_AU"},
        New Marketplace With {.Id = "A19VAU5U5O7RUS", .Name = "Singapore", .Region = "fe", .Currency = "SGD", .Locale = "en_SG"}
    }

    Private ReadOnly Operations As New List(Of OperationInfo) From {
        New OperationInfo With {.Group = "Products", .Id = "catalog", .Label = "Catalogue item"},
        New OperationInfo With {.Group = "Products", .Id = "fees", .Label = "Fee estimate"},
        New OperationInfo With {.Group = "Products", .Id = "inventory", .Label = "FBA inventory"},
        New OperationInfo With {.Group = "Orders", .Id = "orders", .Label = "Search orders"},
        New OperationInfo With {.Group = "Orders", .Id = "order", .Label = "Get order"},
        New OperationInfo With {.Group = "Reports", .Id = "reports", .Label = "List reports"},
        New OperationInfo With {.Group = "Reports", .Id = "createReport", .Label = "Request report", .Kind = "write"},
        New OperationInfo With {.Group = "Reports", .Id = "report", .Label = "Report status"},
        New OperationInfo With {.Group = "Reports", .Id = "reportDocument", .Label = "Report document"},
        New OperationInfo With {.Group = "Feeds", .Id = "feeds", .Label = "List feeds"},
        New OperationInfo With {.Group = "Feeds", .Id = "feed", .Label = "Feed status"},
        New OperationInfo With {.Group = "Feeds", .Id = "feedDocument", .Label = "Feed processing report"},
        New OperationInfo With {.Group = "Feeds", .Id = "submitFeed", .Label = "Submit feed", .Kind = "write"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "inboundPlans", .Label = "List plans"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "inboundPlan", .Label = "Get plan"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "inboundShipment", .Label = "Get shipment"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "inboundOperationStatus", .Label = "Operation status"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "prepDetails", .Label = "Prep details"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "createInboundPlan", .Label = "Create plan", .Kind = "write"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "itemLabels", .Label = "Item labels"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "shipmentLabels", .Label = "Shipment labels"},
        New OperationInfo With {.Group = "FBA inbound", .Id = "billOfLading", .Label = "Bill of lading"},
        New OperationInfo With {.Group = "Legacy database", .Id = "legacyConvert", .Label = "Convert Amazon_US", .Kind = "legacy"},
        New OperationInfo With {.Group = "Legacy database", .Id = "legacyFc", .Label = "SKU / FC bulk update", .Kind = "legacy"}
    }

    Private ReadOnly FieldValues As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly FieldControls As New Dictionary(Of String, Control)(StringComparer.OrdinalIgnoreCase)

    Private ReadOnly txtClientId As New TextBox()
    Private ReadOnly txtClientSecret As New TextBox()
    Private ReadOnly txtRefreshToken As New TextBox()
    Private ReadOnly cboEnvironment As New ComboBox()
    Private ReadOnly cboMarketplace As New ComboBox()
    Private ReadOnly btnTest As New Button()
    Private ReadOnly lblConnection As New Label()
    Private ReadOnly chkShowSecrets As New CheckBox()
    Private ReadOnly operationTree As New TreeView()
    Private ReadOnly workspaceSplit As New SplitContainer()
    Private ReadOnly requestPanel As New FlowLayoutPanel()
    Private ReadOnly lblOperation As New Label()
    Private ReadOnly lblSandbox As New Label()
    Private ReadOnly btnRun As New Button()
    Private ReadOnly lblMeta As New Label()
    Private ReadOnly lblViewTitle As New Label()
    Private ReadOnly lblViewSubtitle As New Label()
    Private ReadOnly lblViewDetails As New Label()
    Private ReadOnly viewGrid As New DataGridView()
    Private ReadOnly viewDetails As New DataGridView()
    Private ReadOnly viewSplit As New SplitContainer()
    Private ReadOnly viewDetailLayout As New TableLayoutPanel()
    Private ReadOnly catalogHero As New TableLayoutPanel()
    Private ReadOnly picProduct As New PictureBox()
    Private ReadOnly productThumbnailStrip As New FlowLayoutPanel()
    Private ReadOnly lblProductTitle As New Label()
    Private ReadOnly lblProductEyebrow As New Label()
    Private ReadOnly lblProductMeta As New Label()
    Private ReadOnly lblProductCoverage As New Label()
    Private ReadOnly txtProductDescription As New RichTextBox()
    Private ReadOnly lblProductImageStatus As New Label()
    Private ReadOnly btnPreviousImage As New Button()
    Private ReadOnly btnNextImage As New Button()
    Private ReadOnly catalogDetailTabs As New TabControl()
    Private ReadOnly catalogOverviewGrid As New DataGridView()
    Private ReadOnly catalogSpecificationsGrid As New DataGridView()
    Private ReadOnly catalogRelatedGrid As New DataGridView()
    Private ReadOnly catalogAllFieldsGrid As New DataGridView()
    Private ReadOnly catalogSectionsPanel As New FlowLayoutPanel()
    Private ReadOnly txtResult As New TextBox()
    Private ReadOnly txtRaw As New TextBox()
    Private ReadOnly btnOpenDocument As New Button()
    Private ReadOnly btnNextStep As New Button()
    Private ReadOnly cboDocuments As New ComboBox()
    Private ReadOnly cboReturnedRecords As New ComboBox()
    Private ReadOnly btnOpenReturnedRecord As New Button()
    Private ReadOnly btnEditRequest As New Button()
    Private ReadOnly tabs As New TabControl()

    Private CurrentOperation As String = "catalog"
    Private LastResult As ApiResult
    Private LastDocumentUrl As String = ""
    Private ReadOnly DocumentUrls As New List(Of KeyValuePair(Of String, String))()
    Private ReadOnly ReturnedRecordActions As New List(Of ReturnedRecordAction)()
    Private NextOperationId As String = ""
    Private NextFieldKey As String = ""
    Private NextFieldValue As String = ""
    Private ConnectionVerified As Boolean
    Private RenderingView As Boolean
    Private CurrentViewOperation As String = ""
    Private ReadOnly CatalogImageUrls As New List(Of String)()
    Private ReadOnly CatalogThumbnailCards As New List(Of Panel)()
    Private CatalogImageIndex As Integer = -1


    Public Sub New()
        Text = "Amazon SP-API Workbench - VB.NET"
        StartPosition = FormStartPosition.CenterScreen
        MinimumSize = New Size(1120, 720)
        Size = New Size(1450, 900)
        Font = New Font("Segoe UI", 9.0F)
        InitializeFieldValues()
        BuildUi()
        BuildOperationTree()
        SelectOperation("catalog")
        NavigateToOperation("catalog")
    End Sub

    Public Sub RunCiSelfTest()
        If CurrentOperation <> "catalog" Then Throw New InvalidOperationException("Catalogue must be the initial active workflow.")
        If operationTree.SelectedNode Is Nothing OrElse operationTree.SelectedNode.Tag Is Nothing OrElse CStr(operationTree.SelectedNode.Tag) <> "catalog" Then
            Throw New InvalidOperationException("Catalogue must be visibly selected on startup.")
        End If
        If Marketplaces.Count <> 23 Then Throw New InvalidOperationException("Marketplace list must contain 23 entries.")
        If Marketplaces.Select(Function(m) m.Id).Distinct(StringComparer.Ordinal).Count() <> Marketplaces.Count Then Throw New InvalidOperationException("Marketplace IDs must be unique.")
        If Operations.Select(Function(op) op.Id).Distinct(StringComparer.Ordinal).Count() <> Operations.Count Then Throw New InvalidOperationException("Operation IDs must be unique.")

        cboEnvironment.SelectedIndex = 0
        For Each operation In Operations
            SelectOperation(operation.Id)
            If requestPanel.Controls.Count = 0 Then Throw New InvalidOperationException("No UI controls were built for " & operation.Id & ".")

            If operation.Kind <> "legacy" AndAlso operation.Id <> "inventory" Then
                LoadSandboxExample()
                If B("confirmed") Then Throw New InvalidOperationException("Sandbox examples must not auto-confirm writes for " & operation.Id & ".")
            End If
        Next

        SelectOperation("catalog")
        FieldValues("catalogMode") = "identifier"
        BuildOperationFields()
        LoadSandboxExample()
        If S("query") <> "B07N4M94X4" Then Throw New InvalidOperationException("Catalog Sandbox example did not remain loaded.")
        If S("identifierType") <> "ASIN" Then Throw New InvalidOperationException("Catalog Sandbox identifier type is incorrect.")

        SelectOperation("createReport")
        FieldValues("confirmed") = True
        BuildOperationFields()
        Dim reportTypeControl As Control = Nothing
        If Not FieldControls.TryGetValue("reportType", reportTypeControl) Then Throw New InvalidOperationException("Create report field was not built.")
        DirectCast(reportTypeControl, TextBox).Text = "GET_MERCHANT_LISTINGS_ALL_DATA_TEST"
        If B("confirmed") Then Throw New InvalidOperationException("Changing a write input must invalidate confirmation.")

        FieldValues("confirmed") = True
        BuildOperationFields()
        Dim confirmationControl As Control = Nothing
        If Not FieldControls.TryGetValue("confirmed", confirmationControl) Then Throw New InvalidOperationException("Write confirmation checkbox was not built.")
        DirectCast(confirmationControl, CheckBox).Checked = True
        InvalidateConnectionState()
        If DirectCast(confirmationControl, CheckBox).Checked OrElse B("confirmed") Then Throw New InvalidOperationException("Credential/context changes must visibly clear write confirmation.")

        SelectMarketplaceById("ATVPDKIKX0DER")
        cboEnvironment.SelectedIndex = 0
        If Endpoint() <> "https://sandbox.sellingpartnerapi-na.amazon.com" Then Throw New InvalidOperationException("North America Sandbox endpoint is incorrect.")
        cboEnvironment.SelectedIndex = 1
        If Endpoint() <> "https://sellingpartnerapi-na.amazon.com" Then Throw New InvalidOperationException("North America Production endpoint is incorrect.")

        SelectMarketplaceById("A1F83G8C2ARO7P")
        If Endpoint() <> "https://sellingpartnerapi-eu.amazon.com" Then Throw New InvalidOperationException("Europe Production endpoint is incorrect.")

        SelectMarketplaceById("A1VC38T7YXB528")
        If Endpoint() <> "https://sellingpartnerapi-fe.amazon.com" Then Throw New InvalidOperationException("Far East Production endpoint is incorrect.")

        If SafeHttpsUrl("http://example.com/file") <> "" Then Throw New InvalidOperationException("HTTP document URLs must be rejected.")
        If SafeHttpsUrl("https://example.com/file") = "" Then Throw New InvalidOperationException("HTTPS document URLs should be accepted.")

        Dim testDocuments = New Dictionary(Of String, Object) From {
            {"documentDownloads", New Object() {
                New Dictionary(Of String, Object) From {{"uri", "https://example.com/one.pdf"}, {"downloadType", "PDF"}},
                New Dictionary(Of String, Object) From {{"uri", "https://example.com/two.pdf"}, {"downloadType", "ZPL"}}
            }}
        }
        If FindDocumentUrls(testDocuments).Count <> 2 Then Throw New InvalidOperationException("All returned document links must remain accessible.")

        If Not ValidIsoInstant("2026-09-23T10:00:00Z") Then Throw New InvalidOperationException("Valid ISO timestamp was rejected.")
        If ValidIsoInstant("2026-09-23 10:00:00") Then Throw New InvalidOperationException("Timestamp without explicit timezone must be rejected.")

        Dim csv = ParseCsvLine("""SKU,ONE"", 2, SELLER, SELLER")
        If csv.Count <> 4 OrElse csv(0) <> "SKU,ONE" OrElse csv(1) <> "2" Then Throw New InvalidOperationException("Quoted CSV item parsing is incorrect.")

        If NormalizeIncludedData(DefaultCatalogData) <> DefaultCatalogData Then Throw New InvalidOperationException("Default Catalog datasets changed unexpectedly.")
        Dim invalidDatasetRejected As Boolean = False
        Try
            NormalizeIncludedData("summaries,notARealDataset")
        Catch ex As AppException
            invalidDatasetRejected = (ex.Code = "INVALID_CATALOG_INCLUDED_DATA")
        End Try
        If Not invalidDatasetRejected Then Throw New InvalidOperationException("Invalid Catalog includedData must be rejected.")

        SelectOperation("catalog")
        FieldValues("pageToken") = "STALE"
        BuildOperationFields()
        SelectOperation("fees")
        If S("pageToken") <> "" Then Throw New InvalidOperationException("Changing operations must clear stale pagination tokens.")

        SelectOperation("catalog")
        Dim pagedResult As New ApiResult With {
            .Ok = True,
            .Status = 200,
            .Data = New Dictionary(Of String, Object) From {
                {"pagination", New Dictionary(Of String, Object) From {{"nextToken", "NEXT_TOKEN"}}}
            }
        }
        ConfigureNextStep("catalog", pagedResult)
        If NextOperationId <> "catalog" OrElse NextFieldKey <> "pageToken" OrElse NextFieldValue <> "NEXT_TOKEN" Then Throw New InvalidOperationException("Catalog next-page action was not prepared.")
        OpenNextStep(Nothing, EventArgs.Empty)
        If S("pageToken") <> "NEXT_TOKEN" Then Throw New InvalidOperationException("Catalog next-page token was not loaded into the form.")

        Dim failedFeed As New ApiResult With {
            .Ok = False,
            .Status = 422,
            .Data = New Dictionary(Of String, Object) From {{"resultFeedDocumentId", "result-doc"}}
        }
        ConfigureNextStep("feed", failedFeed)
        If NextOperationId <> "feedDocument" OrElse NextFieldValue <> "result-doc" Then Throw New InvalidOperationException("Failed feed processing report must remain reachable.")

        SelectOperation("orders")
        Dim orderListResult As New ApiResult With {
            .Ok = True,
            .Status = 200,
            .Data = New Dictionary(Of String, Object) From {
                {"orders", New Object() {
                    New Dictionary(Of String, Object) From {{"orderId", "ORDER-1"}},
                    New Dictionary(Of String, Object) From {{"orderId", "ORDER-2"}}
                }}
            }
        }
        ConfigureReturnedRecords("orders", orderListResult)
        If ReturnedRecordActions.Count <> 2 Then Throw New InvalidOperationException("Returned order selector did not expose every record.")
        cboReturnedRecords.SelectedIndex = 1
        OpenReturnedRecord(Nothing, EventArgs.Empty)
        If CurrentOperation <> "order" OrElse S("orderId") <> "ORDER-2" Then Throw New InvalidOperationException("Returned order selector did not carry the selected ID forward.")

        Dim tooManyValuesRejected As Boolean = False
        Try
            SplitValues(String.Join(",", Enumerable.Range(1, 26).Select(Function(i) "M" & i.ToString(CultureInfo.InvariantCulture))), 25)
        Catch ex As AppException
            tooManyValuesRejected = (ex.Code = "TOO_MANY_VALUES")
        End Try
        If Not tooManyValuesRejected Then Throw New InvalidOperationException("25-value marketplace limit guard is not working.")

        ValidateContentType("application/json; charset=UTF-8")
        Dim invalidContentTypeRejected As Boolean = False
        Try
            ValidateContentType("not a media type <>")
        Catch ex As AppException
            invalidContentTypeRejected = (ex.Code = "INVALID_CONTENT_TYPE")
        End Try
        If Not invalidContentTypeRejected Then Throw New InvalidOperationException("Invalid feed Content-Type must be rejected before an Amazon write.")

        Dim catalogViewResult As New ApiResult With {
            .Ok = True,
            .Status = 200,
            .Data = New Dictionary(Of String, Object) From {
                {"items", New Object() {
                    New Dictionary(Of String, Object) From {
                        {"asin", "B000TEST01"},
                        {"summaries", New Object() {
                            New Dictionary(Of String, Object) From {
                                {"itemName", "Test Product"},
                                {"brand", "Test Brand"},
                                {"marketplaceId", "ATVPDKIKX0DER"}
                            }
                        }},
                        {"productTypes", New Object() {
                            New Dictionary(Of String, Object) From {{"productType", "TEST_PRODUCT"}}
                        }}
                    }
                }}
            }
        }
        RenderResultView("catalog", catalogViewResult)
        If viewGrid.Rows.Count <> 1 Then Throw New InvalidOperationException("Catalog View did not render the returned product.")
        If viewGrid.Columns.Count < 4 Then Throw New InvalidOperationException("Catalog View is missing product columns.")
        If Convert.ToString(viewGrid.Rows(0).Cells(1).Value, CultureInfo.InvariantCulture) <> "Test Product" Then Throw New InvalidOperationException("Catalog View did not show the product title.")
        If catalogSectionsPanel.Controls.Count < 6 Then Throw New InvalidOperationException("Catalog View did not render the complete product page sections.")
        If Not viewSplit.Panel1Collapsed Then Throw New InvalidOperationException("A single catalogue record must use the full product-page width.")
        ShowResultWorkspace()
        If Not workspaceSplit.Panel1Collapsed Then Throw New InvalidOperationException("Result page did not expand over the request editor.")
        ShowRequestWorkspace()
        If workspaceSplit.Panel1Collapsed Then Throw New InvalidOperationException("Edit request did not restore the request editor.")

        Dim richCatalogItem As New Dictionary(Of String, Object) From {
            {"attributes", New Dictionary(Of String, Object) From {
                {"product_description", New Object() {New Dictionary(Of String, Object) From {{"value", "A useful <b>product</b>."}}}},
                {"bullet_point", New Object() {New Dictionary(Of String, Object) From {{"value", "First benefit"}}}}
            }},
            {"images", New Object() {
                New Dictionary(Of String, Object) From {
                    {"images", New Object() {New Dictionary(Of String, Object) From {{"variant", "MAIN"}, {"link", "https://m.media-amazon.com/images/I/test.jpg"}}}}
                }
            }}
        }
        If Not CatalogDescription(richCatalogItem).Contains("A useful product.") OrElse Not CatalogDescription(richCatalogItem).Contains("First benefit") Then Throw New InvalidOperationException("Catalog product page did not render description and bullet attributes.")
        If FindCatalogImageUrls(richCatalogItem).Count <> 1 Then Throw New InvalidOperationException("Catalog product page did not retain the returned HTTPS image.")
        RenderCatalogProductSections(richCatalogItem)
        If catalogSectionsPanel.Controls.Count < 6 Then Throw New InvalidOperationException("Catalog product sections did not render as a continuous page.")

        Dim ordersViewResult As New ApiResult With {
            .Ok = True,
            .Status = 200,
            .Data = New Dictionary(Of String, Object) From {
                {"orders", New Object() {
                    New Dictionary(Of String, Object) From {
                        {"orderId", "ORDER-VIEW-1"},
                        {"createdTime", "2026-09-24T10:00:00Z"},
                        {"salesChannel", "Amazon.com"},
                        {"fulfillment", New Dictionary(Of String, Object) From {
                            {"fulfillmentStatus", "SHIPPED"},
                            {"fulfilledBy", "AMAZON"}
                        }},
                        {"proceeds", New Dictionary(Of String, Object) From {
                            {"grandTotal", New Dictionary(Of String, Object) From {
                                {"currencyCode", "USD"},
                                {"amount", "25.50"}
                            }}
                        }}
                    },
                    New Dictionary(Of String, Object) From {
                        {"orderId", "ORDER-VIEW-2"},
                        {"createdTime", "2026-09-24T11:00:00Z"},
                        {"fulfillment", New Dictionary(Of String, Object) From {
                            {"fulfillmentStatus", "UNSHIPPED"},
                            {"fulfilledBy", "MERCHANT"}
                        }}
                    }
                }}
            }
        }
        RenderResultView("orders", ordersViewResult)
        If viewGrid.Rows.Count <> 2 Then Throw New InvalidOperationException("Orders View did not render every returned order.")
        If Convert.ToString(viewGrid.Rows(0).Cells(0).Value, CultureInfo.InvariantCulture) <> "ORDER-VIEW-1" Then Throw New InvalidOperationException("Orders View did not show the order ID.")
        If Convert.ToString(viewGrid.Rows(0).Cells(5).Value, CultureInfo.InvariantCulture) <> "USD 25.50" Then Throw New InvalidOperationException("Orders View did not show the Orders 2026 grand total.")

        Dim inventoryViewResult As New ApiResult With {
            .Ok = True,
            .Status = 200,
            .Data = New Dictionary(Of String, Object) From {
                {"payload", New Dictionary(Of String, Object) From {
                    {"inventorySummaries", New Object() {
                        New Dictionary(Of String, Object) From {
                            {"productName", "Inventory Product"},
                            {"sellerSku", "SKU-1"},
                            {"asin", "ASIN-1"},
                            {"totalQuantity", 10},
                            {"inventoryDetails", New Dictionary(Of String, Object) From {
                                {"fulfillableQuantity", 7},
                                {"reservedQuantity", New Dictionary(Of String, Object) From {{"totalReservedQuantity", 2}}},
                                {"unfulfillableQuantity", New Dictionary(Of String, Object) From {{"totalUnfulfillableQuantity", 1}}}
                            }}
                        }
                    }}
                }}
            }
        }
        RenderResultView("inventory", inventoryViewResult)
        If viewGrid.Rows.Count <> 1 OrElse Convert.ToString(viewGrid.Rows(0).Cells(3).Value, CultureInfo.InvariantCulture) <> "7" Then Throw New InvalidOperationException("Inventory View did not show fulfillable quantity.")

        Dim prepViewResult As New ApiResult With {
            .Ok = True,
            .Status = 200,
            .Data = New Dictionary(Of String, Object) From {
                {"mskuPrepDetails", New Object() {
                    New Dictionary(Of String, Object) From {
                        {"msku", "MSKU-1"},
                        {"prepCategory", "FRAGILE"},
                        {"prepTypes", New Object() {"ITEM_BUBBLEWRAP", "ITEM_LABELING"}}
                    }
                }}
            }
        }
        RenderResultView("prepDetails", prepViewResult)
        If viewGrid.Rows.Count <> 1 OrElse Not Convert.ToString(viewGrid.Rows(0).Cells(2).Value, CultureInfo.InvariantCulture).Contains("Item Bubblewrap") Then Throw New InvalidOperationException("Prep View did not render prep types readably.")

        Dim documentViewResult As New ApiResult With {
            .Ok = True,
            .Status = 200,
            .Data = New Dictionary(Of String, Object) From {
                {"payload", New Dictionary(Of String, Object) From {{"DownloadURL", "https://example.com/document.pdf"}}}
            }
        }
        RenderResultView("billOfLading", documentViewResult)
        If viewDetails.Rows.Count = 0 Then Throw New InvalidOperationException("Document View did not render payload download information.")

        Dim singleOrderViewResult As New ApiResult With {
            .Ok = True,
            .Status = 200,
            .Data = New Dictionary(Of String, Object) From {
                {"orderId", "ORDER-SINGLE"},
                {"createdTime", "2026-09-24T10:00:00Z"},
                {"fulfillment", New Dictionary(Of String, Object) From {{"fulfillmentStatus", "SHIPPED"}}},
                {"orderItems", New Object() {
                    New Dictionary(Of String, Object) From {
                        {"quantityOrdered", 2},
                        {"product", New Dictionary(Of String, Object) From {
                            {"asin", "ASIN-ITEM"},
                            {"sellerSku", "SKU-ITEM"},
                            {"title", "Order Item"},
                            {"price", New Dictionary(Of String, Object) From {{"currencyCode", "USD"}, {"amount", "9.99"}}}
                        }}
                    }
                }}
            }
        }
        RenderResultView("order", singleOrderViewResult)
        If viewGrid.Rows.Count <> 1 OrElse Convert.ToString(viewGrid.Rows(0).Cells(2).Value, CultureInfo.InvariantCulture) <> "Order Item" Then Throw New InvalidOperationException("Single Order View did not render order items.")

        Dim viewFailure As New ApiResult With {
            .Ok = False,
            .Status = 403,
            .StatusText = "Forbidden",
            .RequestId = "request-view-test",
            .Problem = New ApiProblem With {.Code = "Unauthorized", .Message = "Test failure", .Action = "Check permissions."}
        }
        RenderResultView("orders", viewFailure)
        If Not viewSplit.Panel1Collapsed Then Throw New InvalidOperationException("Failure View should show readable details without an empty record table.")
        If viewDetails.Rows.Count = 0 Then Throw New InvalidOperationException("Failure View did not show error details.")

        CachedAccessToken = "test-token"
        CachedAccessTokenExpiresUtc = DateTimeOffset.UtcNow.AddMinutes(10)
        InvalidateConnectionState()
        If CachedAccessToken <> "" Then Throw New InvalidOperationException("Credential/context changes must clear the cached access token.")

        cboEnvironment.SelectedIndex = 0
        SelectMarketplaceById("ATVPDKIKX0DER")
        SelectOperation("catalog")
    End Sub

    Private Sub InitializeFieldValues()
        FieldValues("catalogMode") = "identifier"
        FieldValues("identifierType") = "ASIN"
        FieldValues("query") = ""
        FieldValues("sellerId") = ""
        FieldValues("includeVariations") = True
        FieldValues("brandNames") = ""
        FieldValues("classificationIds") = ""
        FieldValues("catalogPageSize") = "20"
        FieldValues("pageToken") = ""
        FieldValues("includedData") = DefaultCatalogData
        FieldValues("feeIdType") = "ASIN"
        FieldValues("feeIdentifier") = ""
        FieldValues("price") = ""
        FieldValues("shipping") = "0"
        FieldValues("isAmazonFulfilled") = True
        FieldValues("requestIdentifier") = ""
        FieldValues("pointsNumber") = ""
        FieldValues("pointsAmount") = ""
        FieldValues("details") = True
        FieldValues("includeOrderPii") = False
        FieldValues("pageSize") = "20"
        FieldValues("createdAfter") = DateTimeOffset.UtcNow.AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture)
        FieldValues("createdBefore") = ""
        FieldValues("statuses") = ""
        FieldValues("fulfilledBy") = ""
        FieldValues("orderIncludedData") = ""
        FieldValues("reportTypes") = ""
        FieldValues("reportType") = "GET_MERCHANT_LISTINGS_ALL_DATA"
        FieldValues("processingStatuses") = ""
        FieldValues("feedTypes") = ""
        FieldValues("feedType") = "JSON_LISTINGS_FEED"
        FieldValues("contentType") = "application/json; charset=UTF-8"
        FieldValues("status") = ""
        FieldValues("sortBy") = "LAST_UPDATED_TIME"
        FieldValues("sortOrder") = "DESC"
        FieldValues("countryCode") = "US"
        FieldValues("labelType") = "STANDARD_FORMAT"
        FieldValues("pageType") = "A4_21"
        FieldValues("shipmentPageType") = "PackageLabel_Thermal_NonPCP"
        FieldValues("shipmentLabelType") = "UNIQUE"
        FieldValues("items") = "MY-SKU-001, 1, SELLER, SELLER"
        FieldValues("confirmed") = False
    End Sub

    ' -------------------- Window and controls --------------------
    Private Sub BuildUi()
        Dim root As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Padding = New Padding(10)}
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        Controls.Add(root)

        Dim credentials As New GroupBox With {.Text = "Amazon credentials and environment", .Dock = DockStyle.Top, .AutoSize = True, .Padding = New Padding(10)}
        Dim cGrid As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 6, .AutoSize = True}
        For i As Integer = 0 To 5
            cGrid.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, If(i Mod 2 = 0, 12.0F, 21.333F)))
        Next
        credentials.Controls.Add(cGrid)

        AddCredential(cGrid, 0, 0, "Client ID", txtClientId, False)
        AddCredential(cGrid, 2, 0, "Client secret", txtClientSecret, True)
        AddCredential(cGrid, 4, 0, "Refresh token", txtRefreshToken, True)

        cboEnvironment.DropDownStyle = ComboBoxStyle.DropDownList
        cboEnvironment.Items.AddRange(New Object() {"Sandbox", "Production"})
        cboEnvironment.SelectedIndex = 0
        AddCredential(cGrid, 0, 1, "Environment", cboEnvironment, False)

        cboMarketplace.DropDownStyle = ComboBoxStyle.DropDownList
        cboMarketplace.Items.AddRange(Marketplaces.Cast(Of Object)().ToArray())
        cboMarketplace.SelectedIndex = 0
        AddCredential(cGrid, 2, 1, "Marketplace", cboMarketplace, False)

        btnTest.Text = "Test connection"
        btnTest.AutoSize = True
        btnTest.Padding = New Padding(10, 4, 10, 4)
        AddHandler btnTest.Click, Async Sub(sender, e) Await TestConnectionAsync()
        cGrid.Controls.Add(btnTest, 4, 1)

        lblConnection.AutoSize = False
        lblConnection.Dock = DockStyle.Fill
        lblConnection.AutoEllipsis = True
        lblConnection.TextAlign = ContentAlignment.MiddleLeft
        lblConnection.Text = "Not tested"
        lblConnection.ForeColor = Color.DimGray
        cGrid.Controls.Add(lblConnection, 5, 1)

        chkShowSecrets.Text = "Show Client Secret and Refresh Token"
        chkShowSecrets.AutoSize = True
        chkShowSecrets.Margin = New Padding(3, 6, 3, 4)
        AddHandler chkShowSecrets.CheckedChanged, Sub(sender, e)
                                                      txtClientSecret.UseSystemPasswordChar = Not chkShowSecrets.Checked
                                                      txtRefreshToken.UseSystemPasswordChar = Not chkShowSecrets.Checked
                                                  End Sub
        cGrid.Controls.Add(chkShowSecrets, 0, 2)
        cGrid.SetColumnSpan(chkShowSecrets, 2)

        Dim credentialNote As New Label With {
            .Text = "Credentials stay in this running app only; they are not saved to disk.",
            .AutoSize = True,
            .ForeColor = Color.DimGray,
            .Margin = New Padding(3, 8, 3, 4)
        }
        cGrid.Controls.Add(credentialNote, 2, 2)
        cGrid.SetColumnSpan(credentialNote, 4)

        AddHandler txtClientId.TextChanged, Sub(sender, e) InvalidateConnectionState()
        AddHandler txtClientSecret.TextChanged, Sub(sender, e) InvalidateConnectionState()
        AddHandler txtRefreshToken.TextChanged, Sub(sender, e) InvalidateConnectionState()
        AddHandler cboEnvironment.SelectedIndexChanged, Sub(sender, e)
                                                            InvalidateConnectionState()
                                                            BuildOperationFields()
                                                        End Sub
        AddHandler cboMarketplace.SelectedIndexChanged, Sub(sender, e)
                                                            InvalidateConnectionState()
                                                            BuildOperationFields()
                                                        End Sub
        InvalidateConnectionState()
        root.Controls.Add(credentials, 0, 0)

        Dim mainSplit As New SplitContainer With {.Dock = DockStyle.Fill, .Orientation = Orientation.Vertical, .SplitterDistance = 245, .FixedPanel = FixedPanel.Panel1}
        root.Controls.Add(mainSplit, 0, 1)

        operationTree.Dock = DockStyle.Fill
        operationTree.HideSelection = False
        operationTree.FullRowSelect = True
        AddHandler operationTree.AfterSelect, AddressOf OperationSelected
        mainSplit.Panel1.Controls.Add(operationTree)

        workspaceSplit.Dock = DockStyle.Fill
        workspaceSplit.Orientation = Orientation.Horizontal
        workspaceSplit.SplitterDistance = 360
        workspaceSplit.SplitterWidth = 18
        workspaceSplit.BackColor = Color.FromArgb(37, 99, 235)
        workspaceSplit.Panel1.BackColor = SystemColors.Control
        workspaceSplit.Panel2.BackColor = Color.FromArgb(247, 245, 240)
        AddHandler workspaceSplit.Paint, AddressOf PaintWorkspaceSplitter
        AddHandler workspaceSplit.SplitterMoved, Sub(sender, e) workspaceSplit.Invalidate()
        AddHandler workspaceSplit.MouseDoubleClick, AddressOf WorkspaceSplitterDoubleClick
        mainSplit.Panel2.Controls.Add(workspaceSplit)

        Dim requestHost As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(10)}
        workspaceSplit.Panel1.Controls.Add(requestHost)
        lblOperation.Dock = DockStyle.Top
        lblOperation.Height = 34
        lblOperation.Font = New Font(Font, FontStyle.Bold)
        lblOperation.Font = New Font(lblOperation.Font.FontFamily, 14.0F, FontStyle.Bold)
        requestHost.Controls.Add(lblOperation)

        lblSandbox.Dock = DockStyle.Top
        lblSandbox.AutoSize = False
        lblSandbox.Height = 55
        lblSandbox.Padding = New Padding(8)
        lblSandbox.BackColor = Color.FromArgb(255, 248, 220)
        lblSandbox.ForeColor = Color.FromArgb(90, 70, 0)
        requestHost.Controls.Add(lblSandbox)
        lblSandbox.BringToFront()

        requestPanel.Dock = DockStyle.Fill
        requestPanel.FlowDirection = FlowDirection.TopDown
        requestPanel.WrapContents = False
        requestPanel.AutoScroll = True
        requestPanel.Padding = New Padding(3)
        AddHandler requestPanel.Resize, Sub(sender, e) ResizeRequestFields()
        requestHost.Controls.Add(requestPanel)
        requestPanel.BringToFront()

        btnRun.Dock = DockStyle.Bottom
        btnRun.Height = 40
        btnRun.Text = "Run request"
        AddHandler btnRun.Click, Async Sub(sender, e) Await RunCurrentAsync()
        requestHost.Controls.Add(btnRun)

        tabs.Dock = DockStyle.Fill
        Dim viewTab As New TabPage("View")
        Dim resultTab As New TabPage("Summary")
        Dim rawTab As New TabPage("Raw response")
        tabs.TabPages.Add(viewTab)
        tabs.TabPages.Add(resultTab)
        tabs.TabPages.Add(rawTab)
        workspaceSplit.Panel2.Controls.Add(tabs)

        BuildResultView(viewTab)

        Dim resultLayout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 3,
            .Padding = New Padding(0)
        }
        resultLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        resultLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        resultLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        resultLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        resultTab.Controls.Add(resultLayout)

        lblMeta.Dock = DockStyle.Fill
        lblMeta.Height = 28
        lblMeta.Padding = New Padding(6)
        resultLayout.Controls.Add(lblMeta, 0, 0)

        txtResult.Dock = DockStyle.Fill
        txtResult.Multiline = True
        txtResult.ScrollBars = ScrollBars.Both
        txtResult.ReadOnly = True
        txtResult.WordWrap = False
        txtResult.Font = New Font("Consolas", 9.0F)
        resultLayout.Controls.Add(txtResult, 0, 1)

        Dim resultActions As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .Padding = New Padding(3, 3, 3, 3)
        }

        Dim btnCopyResult As New Button With {
            .Text = "Copy result",
            .AutoSize = True,
            .Padding = New Padding(8, 2, 8, 2)
        }
        AddHandler btnCopyResult.Click, Sub(sender, e) CopyTextToClipboard(txtResult.Text, "Result")
        resultActions.Controls.Add(btnCopyResult)

        resultLayout.Controls.Add(resultActions, 0, 2)

        Dim rawLayout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2,
            .Padding = New Padding(0)
        }
        rawLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        rawLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        rawLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        rawTab.Controls.Add(rawLayout)

        txtRaw.Dock = DockStyle.Fill
        txtRaw.Multiline = True
        txtRaw.ScrollBars = ScrollBars.Both
        txtRaw.ReadOnly = True
        txtRaw.WordWrap = False
        txtRaw.Font = New Font("Consolas", 9.0F)
        rawLayout.Controls.Add(txtRaw, 0, 0)

        Dim btnCopyRaw As New Button With {
            .Text = "Copy raw response",
            .AutoSize = True,
            .Padding = New Padding(8, 2, 8, 2),
            .Anchor = AnchorStyles.Left
        }
        AddHandler btnCopyRaw.Click, Sub(sender, e) CopyTextToClipboard(txtRaw.Text, "Raw response")
        rawLayout.Controls.Add(btnCopyRaw, 0, 1)
    End Sub

    Private Sub BuildResultView(viewTab As TabPage)
        Dim canvas = Color.FromArgb(247, 245, 240)
        Dim ink = Color.FromArgb(30, 38, 48)
        Dim accent = Color.FromArgb(37, 99, 235)
        viewTab.BackColor = canvas
        Dim layout As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 4,
            .Padding = New Padding(14, 10, 14, 10),
            .BackColor = canvas
        }
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        layout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        layout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        viewTab.Controls.Add(layout)

        lblViewTitle.AutoSize = True
        lblViewTitle.Font = New Font("Bahnschrift SemiBold", 17.0F, FontStyle.Bold)
        lblViewTitle.ForeColor = ink
        lblViewTitle.Text = "No result yet"
        lblViewTitle.Margin = New Padding(3, 2, 3, 2)
        layout.Controls.Add(lblViewTitle, 0, 0)

        lblViewSubtitle.AutoSize = True
        lblViewSubtitle.ForeColor = Color.DimGray
        lblViewSubtitle.MaximumSize = New Size(1000, 0)
        lblViewSubtitle.Text = "Run an operation to see a readable view here."
        lblViewSubtitle.Margin = New Padding(3, 0, 3, 8)
        layout.Controls.Add(lblViewSubtitle, 0, 1)

        viewSplit.Dock = DockStyle.Fill
        viewSplit.Orientation = Orientation.Horizontal
        viewSplit.SplitterDistance = 210
        viewSplit.Panel1MinSize = 80
        viewSplit.Panel2MinSize = 80
        layout.Controls.Add(viewSplit, 0, 2)

        viewGrid.Dock = DockStyle.Fill
        viewGrid.ReadOnly = True
        viewGrid.AllowUserToAddRows = False
        viewGrid.AllowUserToDeleteRows = False
        viewGrid.AllowUserToResizeRows = False
        viewGrid.MultiSelect = False
        viewGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        viewGrid.RowHeadersVisible = False
        viewGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        viewGrid.BackgroundColor = Color.White
        viewGrid.BorderStyle = BorderStyle.None
        viewGrid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        viewGrid.GridColor = Color.FromArgb(226, 222, 214)
        viewGrid.EnableHeadersVisualStyles = False
        viewGrid.ColumnHeadersDefaultCellStyle.BackColor = ink
        viewGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        viewGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        viewGrid.ColumnHeadersHeight = 34
        viewGrid.RowTemplate.Height = 34
        viewGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(251, 249, 245)
        viewGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(255, 235, 205)
        viewGrid.DefaultCellStyle.SelectionForeColor = ink
        viewGrid.AutoGenerateColumns = False
        AddHandler viewGrid.SelectionChanged, AddressOf ViewGridSelectionChanged
        AddHandler viewGrid.CellDoubleClick, AddressOf ViewGridDoubleClick
        viewSplit.Panel1.Controls.Add(viewGrid)

        viewDetailLayout.Dock = DockStyle.Fill
        viewDetailLayout.ColumnCount = 1
        viewDetailLayout.RowCount = 3
        viewDetailLayout.Padding = New Padding(8, 0, 0, 0)
        viewDetailLayout.BackColor = canvas
        viewDetailLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        viewDetailLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 0.0F))
        viewDetailLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        viewSplit.Panel2.Controls.Add(viewDetailLayout)

        lblViewDetails.AutoSize = True
        lblViewDetails.Font = New Font(Font, FontStyle.Bold)
        lblViewDetails.Text = "Details"
        lblViewDetails.Margin = New Padding(3, 4, 3, 4)
        viewDetailLayout.Controls.Add(lblViewDetails, 0, 0)

        BuildCatalogHero(ink, accent, canvas)
        viewDetailLayout.Controls.Add(catalogHero, 0, 1)

        viewDetails.Dock = DockStyle.Fill
        viewDetails.ReadOnly = True
        viewDetails.AllowUserToAddRows = False
        viewDetails.AllowUserToDeleteRows = False
        viewDetails.AllowUserToResizeRows = False
        viewDetails.MultiSelect = False
        viewDetails.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        viewDetails.RowHeadersVisible = False
        viewDetails.AutoGenerateColumns = False
        viewDetails.BackgroundColor = Color.White
        viewDetails.BorderStyle = BorderStyle.None
        viewDetails.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        viewDetails.GridColor = Color.FromArgb(230, 226, 218)
        viewDetails.EnableHeadersVisualStyles = False
        viewDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(63, 72, 82)
        viewDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        viewDetails.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        viewDetails.ColumnHeadersHeight = 32
        viewDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(251, 249, 245)
        viewDetails.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "Field",
            .HeaderText = "Field",
            .Width = 235,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        })
        viewDetails.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "Value",
            .HeaderText = "Value",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        })
        viewDetailLayout.Controls.Add(viewDetails, 0, 2)

        viewDetails.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        viewDetails.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders

        BuildCatalogDetailTabs(ink, accent, canvas)
        viewDetailLayout.Controls.Add(catalogDetailTabs, 0, 2)
        catalogDetailTabs.Visible = False

        catalogSectionsPanel.AutoSize = True
        catalogSectionsPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink
        catalogSectionsPanel.FlowDirection = FlowDirection.TopDown
        catalogSectionsPanel.WrapContents = False
        catalogSectionsPanel.Margin = New Padding(0)
        catalogSectionsPanel.Padding = New Padding(12, 8, 18, 24)
        catalogSectionsPanel.BackColor = canvas
        catalogSectionsPanel.Visible = False
        viewDetailLayout.Controls.Add(catalogSectionsPanel, 0, 2)

        Dim viewActions As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .Padding = New Padding(0, 6, 0, 0)
        }

        btnEditRequest.Text = "← Edit request"
        btnEditRequest.AutoSize = True
        btnEditRequest.Padding = New Padding(12, 4, 12, 4)
        btnEditRequest.FlatStyle = FlatStyle.Flat
        btnEditRequest.BackColor = ink
        btnEditRequest.ForeColor = Color.White
        btnEditRequest.FlatAppearance.BorderColor = ink
        btnEditRequest.Visible = False
        AddHandler btnEditRequest.Click, Sub(sender, e) ShowRequestWorkspace()
        viewActions.Controls.Add(btnEditRequest)

        cboReturnedRecords.DropDownStyle = ComboBoxStyle.DropDownList
        cboReturnedRecords.Width = 300
        cboReturnedRecords.Visible = False
        viewActions.Controls.Add(cboReturnedRecords)

        btnOpenReturnedRecord.Text = "Open selected"
        btnOpenReturnedRecord.AutoSize = True
        btnOpenReturnedRecord.Padding = New Padding(8, 2, 8, 2)
        btnOpenReturnedRecord.Visible = False
        AddHandler btnOpenReturnedRecord.Click, AddressOf OpenReturnedRecord
        viewActions.Controls.Add(btnOpenReturnedRecord)

        btnNextStep.Text = "Next step"
        btnNextStep.AutoSize = True
        btnNextStep.Padding = New Padding(8, 2, 8, 2)
        btnNextStep.Visible = False
        AddHandler btnNextStep.Click, AddressOf OpenNextStep
        viewActions.Controls.Add(btnNextStep)

        cboDocuments.DropDownStyle = ComboBoxStyle.DropDownList
        cboDocuments.Width = 220
        cboDocuments.Visible = False
        viewActions.Controls.Add(cboDocuments)

        btnOpenDocument.Text = "Open / download document"
        btnOpenDocument.AutoSize = True
        btnOpenDocument.Padding = New Padding(8, 2, 8, 2)
        btnOpenDocument.Visible = False
        AddHandler btnOpenDocument.Click, AddressOf OpenDocument
        viewActions.Controls.Add(btnOpenDocument)

        layout.Controls.Add(viewActions, 0, 3)
    End Sub

    Private Sub BuildCatalogDetailTabs(ink As Color, accent As Color, canvas As Color)
        catalogDetailTabs.Dock = DockStyle.Fill
        catalogDetailTabs.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        catalogDetailTabs.Padding = New Point(16, 7)

        Dim tabsAndGrids = New Tuple(Of String, DataGridView)() {
            Tuple.Create("Overview", catalogOverviewGrid),
            Tuple.Create("Specifications", catalogSpecificationsGrid),
            Tuple.Create("Category && related", catalogRelatedGrid),
            Tuple.Create("All SP-API fields", catalogAllFieldsGrid)
        }
        For Each entry In tabsAndGrids
            Dim page As New TabPage(entry.Item1) With {.BackColor = canvas, .Padding = New Padding(8)}
            ConfigureCatalogDetailGrid(entry.Item2, ink, accent)
            page.Controls.Add(entry.Item2)
            catalogDetailTabs.TabPages.Add(page)
        Next
    End Sub

    Private Sub ConfigureCatalogDetailGrid(grid As DataGridView, ink As Color, accent As Color)
        grid.Dock = DockStyle.Fill
        grid.ReadOnly = True
        grid.AllowUserToAddRows = False
        grid.AllowUserToDeleteRows = False
        grid.AllowUserToResizeRows = False
        grid.MultiSelect = False
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        grid.RowHeadersVisible = False
        grid.AutoGenerateColumns = False
        grid.BackgroundColor = Color.White
        grid.BorderStyle = BorderStyle.None
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        grid.GridColor = Color.FromArgb(231, 228, 221)
        grid.EnableHeadersVisualStyles = False
        grid.ColumnHeadersDefaultCellStyle.BackColor = ink
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        grid.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        grid.ColumnHeadersHeight = 34
        grid.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(255, 235, 205)
        grid.DefaultCellStyle.SelectionForeColor = ink
        grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(252, 250, 246)
        grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        grid.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "Property",
            .HeaderText = "Property",
            .Width = 245,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        })
        grid.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "Value",
            .HeaderText = "Value",
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        })
    End Sub

    Private Sub BuildCatalogHero(ink As Color, accent As Color, canvas As Color)
        catalogHero.Dock = DockStyle.Fill
        catalogHero.ColumnCount = 2
        catalogHero.RowCount = 1
        catalogHero.Padding = New Padding(0, 2, 0, 10)
        catalogHero.BackColor = canvas
        catalogHero.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 224.0F))
        catalogHero.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        catalogHero.Visible = False

        Dim imageCard As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .Padding = New Padding(10),
            .Margin = New Padding(0, 0, 18, 0),
            .ColumnCount = 1,
            .RowCount = 3
        }
        imageCard.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        imageCard.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        imageCard.RowStyles.Add(New RowStyle(SizeType.Absolute, 98.0F))
        imageCard.RowStyles.Add(New RowStyle(SizeType.Absolute, 58.0F))
        picProduct.Dock = DockStyle.Fill
        picProduct.BackColor = Color.White
        picProduct.SizeMode = PictureBoxSizeMode.Zoom
        picProduct.Cursor = Cursors.Hand
        AddHandler picProduct.MouseWheel, Sub(sender, e) MoveCatalogImage(If(e.Delta < 0, 1, -1))
        AddHandler picProduct.DoubleClick, Sub(sender, e) OpenCurrentCatalogImage()
        AddHandler picProduct.LoadCompleted, Sub(sender, e)
                                                    If e.Error IsNot Nothing Then
                                                        lblProductImageStatus.Text = "Image unavailable"
                                                    ElseIf CatalogImageIndex >= 0 Then
                                                        lblProductImageStatus.Text = (CatalogImageIndex + 1).ToString(CultureInfo.InvariantCulture) & " / " & CatalogImageUrls.Count.ToString(CultureInfo.InvariantCulture) & " · wheel"
                                                    End If
                                                End Sub
        imageCard.Controls.Add(picProduct, 0, 0)

        Dim imageFooter As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 3,
            .RowCount = 1,
            .BackColor = Color.White,
            .Padding = New Padding(6, 7, 6, 5)
        }
        imageFooter.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 64.0F))
        imageFooter.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        imageFooter.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 64.0F))
        imageFooter.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        btnPreviousImage.Text = "<"
        btnPreviousImage.Dock = DockStyle.Fill
        btnPreviousImage.Margin = New Padding(0)
        btnPreviousImage.FlatStyle = FlatStyle.Flat
        btnPreviousImage.FlatAppearance.BorderColor = accent
        btnPreviousImage.BackColor = accent
        btnPreviousImage.ForeColor = Color.White
        btnPreviousImage.Font = New Font("Segoe UI Semibold", 12.0F, FontStyle.Bold)
        btnNextImage.Text = ">"
        btnNextImage.Dock = DockStyle.Fill
        btnNextImage.Margin = New Padding(0)
        btnNextImage.FlatStyle = FlatStyle.Flat
        btnNextImage.FlatAppearance.BorderColor = accent
        btnNextImage.BackColor = accent
        btnNextImage.ForeColor = Color.White
        btnNextImage.Font = New Font("Segoe UI Semibold", 12.0F, FontStyle.Bold)
        lblProductImageStatus.AutoSize = False
        lblProductImageStatus.Dock = DockStyle.Fill
        lblProductImageStatus.Margin = New Padding(8, 0, 8, 0)
        lblProductImageStatus.TextAlign = ContentAlignment.MiddleCenter
        lblProductImageStatus.ForeColor = ink
        lblProductImageStatus.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        AddHandler btnPreviousImage.Click, Sub(sender, e) MoveCatalogImage(-1)
        AddHandler btnNextImage.Click, Sub(sender, e) MoveCatalogImage(1)
        Dim imageTip As New ToolTip()
        imageTip.SetToolTip(picProduct, "Use the mouse wheel or arrows to browse images. Double-click to open the current image.")
        imageTip.SetToolTip(btnPreviousImage, "Previous product image")
        imageTip.SetToolTip(btnNextImage, "Next product image")
        imageFooter.Controls.Add(btnPreviousImage, 0, 0)
        imageFooter.Controls.Add(lblProductImageStatus, 1, 0)
        imageFooter.Controls.Add(btnNextImage, 2, 0)

        productThumbnailStrip.Dock = DockStyle.Fill
        productThumbnailStrip.FlowDirection = FlowDirection.LeftToRight
        productThumbnailStrip.WrapContents = False
        productThumbnailStrip.AutoScroll = True
        productThumbnailStrip.BackColor = Color.White
        productThumbnailStrip.Padding = New Padding(6, 9, 6, 7)
        productThumbnailStrip.Margin = New Padding(0)
        imageCard.Controls.Add(productThumbnailStrip, 0, 1)
        imageCard.Controls.Add(imageFooter, 0, 2)
        catalogHero.Controls.Add(imageCard, 0, 0)

        Dim copy As New TableLayoutPanel With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 5,
            .BackColor = canvas,
            .Padding = New Padding(0)
        }
        copy.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        copy.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        copy.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        copy.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        copy.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        lblProductEyebrow.AutoSize = True
        lblProductEyebrow.Font = New Font("Segoe UI Semibold", 8.5F, FontStyle.Bold)
        lblProductEyebrow.ForeColor = accent
        lblProductEyebrow.Margin = New Padding(0, 2, 0, 4)
        copy.Controls.Add(lblProductEyebrow, 0, 0)

        lblProductTitle.AutoSize = True
        lblProductTitle.Font = New Font("Bahnschrift SemiBold", 15.0F, FontStyle.Bold)
        lblProductTitle.ForeColor = ink
        lblProductTitle.MaximumSize = New Size(700, 58)
        lblProductTitle.Margin = New Padding(0, 0, 0, 5)
        copy.Controls.Add(lblProductTitle, 0, 1)

        lblProductMeta.AutoSize = True
        lblProductMeta.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular)
        lblProductMeta.ForeColor = Color.FromArgb(82, 88, 94)
        lblProductMeta.MaximumSize = New Size(760, 72)
        lblProductMeta.Margin = New Padding(0, 0, 0, 7)
        copy.Controls.Add(lblProductMeta, 0, 2)

        lblProductCoverage.AutoSize = True
        lblProductCoverage.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        lblProductCoverage.ForeColor = Color.FromArgb(30, 64, 175)
        lblProductCoverage.BackColor = Color.FromArgb(235, 241, 255)
        lblProductCoverage.Padding = New Padding(9, 6, 9, 6)
        lblProductCoverage.Margin = New Padding(0, 0, 0, 9)
        lblProductCoverage.MaximumSize = New Size(760, 58)
        copy.Controls.Add(lblProductCoverage, 0, 3)

        txtProductDescription.Dock = DockStyle.Fill
        txtProductDescription.ReadOnly = True
        txtProductDescription.BorderStyle = BorderStyle.None
        txtProductDescription.BackColor = canvas
        txtProductDescription.ForeColor = ink
        txtProductDescription.Font = New Font("Segoe UI", 9.0F)
        txtProductDescription.ScrollBars = RichTextBoxScrollBars.None
        txtProductDescription.TabStop = False
        copy.Controls.Add(txtProductDescription, 0, 4)
        catalogHero.Controls.Add(copy, 1, 0)
    End Sub

    Private Sub ShowResultWorkspace()
        If workspaceSplit Is Nothing Then Return
        workspaceSplit.Panel1Collapsed = True
        btnEditRequest.Visible = True
        tabs.SelectedIndex = 0
        PerformLayout()
    End Sub

    Private Sub ShowRequestWorkspace()
        If workspaceSplit Is Nothing Then Return
        workspaceSplit.Panel1Collapsed = False
        btnEditRequest.Visible = False
    End Sub

    Private Sub PaintWorkspaceSplitter(sender As Object, e As PaintEventArgs)
        If workspaceSplit.Panel1Collapsed OrElse workspaceSplit.Panel2Collapsed Then Return
        Dim bounds = workspaceSplit.SplitterRectangle
        If bounds.Width <= 0 OrElse bounds.Height <= 0 Then Return

        Using splitterBrush As New SolidBrush(Color.FromArgb(37, 99, 235))
            e.Graphics.FillRectangle(splitterBrush, bounds)
        End Using
        Using labelFont As New Font("Segoe UI Semibold", 7.5F, FontStyle.Bold),
              labelBrush As New SolidBrush(Color.White)
            Dim label = "DRAG TO RESIZE   |   DOUBLE-CLICK FOR FULL RESULTS"
            Dim measured = e.Graphics.MeasureString(label, labelFont)
            Dim x = bounds.Left + Math.Max(8.0F, (bounds.Width - measured.Width) / 2.0F)
            Dim y = bounds.Top + Math.Max(0.0F, (bounds.Height - measured.Height) / 2.0F)
            e.Graphics.DrawString(label, labelFont, labelBrush, x, y)
        End Using
    End Sub

    Private Sub WorkspaceSplitterDoubleClick(sender As Object, e As MouseEventArgs)
        If workspaceSplit.Panel1Collapsed Then Return
        Dim bounds = workspaceSplit.SplitterRectangle
        bounds.Inflate(0, 4)
        If bounds.Contains(e.Location) Then ShowResultWorkspace()
    End Sub

    Private Sub AddCredential(grid As TableLayoutPanel, col As Integer, row As Integer, labelText As String, control As Control, secret As Boolean)
        Dim label As New Label With {.Text = labelText, .AutoSize = True, .Anchor = AnchorStyles.Left, .Margin = New Padding(3, 8, 3, 3)}
        control.Dock = DockStyle.Fill
        control.Margin = New Padding(3, 4, 10, 4)
        Dim tb = TryCast(control, TextBox)
        If tb IsNot Nothing AndAlso secret Then tb.UseSystemPasswordChar = True
        grid.Controls.Add(label, col, row)
        grid.Controls.Add(control, col + 1, row)
    End Sub

    Private Sub BuildOperationTree()
        operationTree.BeginUpdate()
        operationTree.Nodes.Clear()
        For Each group In Operations.GroupBy(Function(x) x.Group)
            Dim groupNode As New TreeNode(group.Key)
            For Each op In group
                groupNode.Nodes.Add(New TreeNode(op.Label) With {.Tag = op.Id})
            Next
            operationTree.Nodes.Add(groupNode)
            groupNode.Expand()
        Next
        operationTree.EndUpdate()
    End Sub

    Private Sub OperationSelected(sender As Object, e As TreeViewEventArgs)
        If e.Node.Tag Is Nothing Then Return
        SelectOperation(CStr(e.Node.Tag))
    End Sub

    Private Sub ClearPaginationTokens()
        For Each key In {"pageToken", "inventoryNextToken", "orderPaginationToken", "reportNextToken", "feedNextToken", "inboundPaginationToken"}
            FieldValues(key) = ""
        Next
    End Sub

    Private Sub SelectOperation(id As String)
        ShowRequestWorkspace()
        SaveVisibleFieldValues()
        If Not String.Equals(id, CurrentOperation, StringComparison.Ordinal) Then ClearPaginationTokens()
        CurrentOperation = id
        FieldValues("confirmed") = False
        Dim operation = Operations.First(Function(x) x.Id = id)
        lblOperation.Text = operation.Label
        btnRun.Text = If(operation.Kind = "legacy", "Legacy utility unavailable", "Run " & operation.Label)
        BuildOperationFields()
        txtResult.Text = If(operation.Kind = "legacy", "This legacy SQL utility is not part of the portable SP-API connection.", "Run the selected request to see a readable result here.")
        txtRaw.Text = If(operation.Kind = "legacy", "", "The complete Amazon response will appear here.")
        ClearResultView(If(operation.Kind = "legacy", "Legacy utility", operation.Label), If(operation.Kind = "legacy", "This private SQL utility is intentionally disconnected.", "Run the request to see the returned data here."))
        lblMeta.Text = ""
        btnOpenDocument.Visible = False
        cboDocuments.Visible = False
        cboDocuments.Items.Clear()
        DocumentUrls.Clear()
        cboReturnedRecords.Visible = False
        cboReturnedRecords.Items.Clear()
        btnOpenReturnedRecord.Visible = False
        ReturnedRecordActions.Clear()
        btnNextStep.Visible = False
        NextOperationId = ""
        NextFieldKey = ""
        NextFieldValue = ""
        btnRun.Enabled = operation.Kind <> "legacy"
    End Sub

    ' -------------------- Request-form UX --------------------
    Private Sub BuildOperationFields()
        If requestPanel Is Nothing Then Return
        requestPanel.SuspendLayout()
        requestPanel.Controls.Clear()
        FieldControls.Clear()
        Dim guide = If(IsSandbox(), SandboxGuide(CurrentOperation), "")
        Dim operationKind = Operations.First(Function(x) x.Id = CurrentOperation).Kind
        If operationKind = "legacy" Then
            lblSandbox.Visible = False
        Else
            lblSandbox.Visible = True
            If IsSandbox() Then
                lblSandbox.Text = "SANDBOX - Amazon test endpoint. Nothing is changed in Production."
                lblSandbox.BackColor = Color.FromArgb(255, 248, 220)
                lblSandbox.ForeColor = Color.FromArgb(90, 70, 0)
            Else
                lblSandbox.Text = "PRODUCTION - Live seller account. Read requests use live data; confirmed write requests can change Amazon data."
                lblSandbox.BackColor = Color.FromArgb(255, 238, 238)
                lblSandbox.ForeColor = Color.FromArgb(120, 35, 35)
            End If
        End If

        AddNote(OperationHelp(CurrentOperation))
        If IsSandbox() AndAlso guide <> "" Then
            AddNote(guide)
            If CurrentOperation <> "inventory" Then AddSandboxExampleButton()
        End If

        Select Case CurrentOperation
            Case "catalog"
                AddChoice("catalogMode", "Search mode", {"identifier", "keywords"}, True, AddressOf RebuildOnChange)
                If S("catalogMode") = "identifier" Then AddChoice("identifierType", "Identifier type", {"ASIN", "UPC", "EAN", "GTIN", "ISBN", "SKU", "JAN", "MINSAN"}, True, AddressOf RebuildOnChange)
                AddText("query", If(S("catalogMode") = "keywords", "Search terms", "Product identifier(s)"), True, True)
                If S("catalogMode") = "identifier" AndAlso S("identifierType") = "SKU" Then AddText("sellerId", "Seller ID", True)
                AddText("includedData", "Included data", True, True)
                If S("catalogMode") = "identifier" AndAlso S("identifierType") = "ASIN" Then AddCheck("includeVariations", "Fetch every related variation and package ASIN")
                If S("catalogMode") = "keywords" Then
                    AddText("brandNames", "Brand names")
                    AddText("classificationIds", "Classification IDs")
                    AddText("catalogPageSize", "Results per page")
                    AddText("pageToken", "Next-page token")
                End If
            Case "fees"
                AddChoice("feeIdType", "Lookup by", {"ASIN", "SKU"}, True)
                AddText("feeIdentifier", S("feeIdType"), True)
                AddText("price", "Listing price (" & SelectedMarketplace().Currency & ")", True)
                AddText("shipping", "Shipping (" & SelectedMarketplace().Currency & ")")
                AddChoice("fulfillment", "Fulfilment", {"FBA", "Merchant"}, True)
                AddText("requestIdentifier", "Request identifier")
                AddText("pointsNumber", "Points number")
                AddText("pointsAmount", "Points monetary value")
            Case "inventory"
                AddText("sellerSkus", "Seller SKUs", False, True)
                AddText("startDateTime", "Changed since (ISO 8601 + timezone)")
                AddText("inventoryNextToken", "Next-page token")
                AddCheck("details", "Include quantity details")
            Case "orders"
                AddText("createdAfter", "Created after (ISO 8601 + timezone)", True)
                AddText("createdBefore", "Created before (ISO 8601 + timezone)")
                AddText("statuses", "Order statuses")
                AddText("fulfilledBy", "Fulfilled by")
                AddText("pageSize", "Results per page")
                AddText("orderPaginationToken", "Next-page token")
                AddText("orderIncludedData", "Included data override", False, True)
                AddCheck("includeOrderPii", "Include buyer + recipient PII when no override is supplied")
            Case "order"
                AddText("orderId", "Amazon order ID", True)
                AddText("orderIncludedData", "Included data override", False, True)
                AddCheck("includeOrderPii", "Include buyer + recipient PII when no override is supplied")
            Case "reports"
                AddText("reportTypes", "Report type(s)", False, True)
                AddText("processingStatuses", "Processing statuses")
                AddText("reportMarketplaceIds", "Marketplace IDs")
                AddText("createdSince", "Created since")
                AddText("createdUntil", "Created until")
                AddText("pageSize", "Results per page")
                AddText("reportNextToken", "Next-page token")
            Case "createReport"
                AddText("reportType", "Report type", True)
                AddText("reportMarketplaceIds", "Marketplace IDs")
                AddText("dataStartTime", "Data start")
                AddText("dataEndTime", "Data end")
                AddCheck("confirmed", "I understand this starts a report job using exactly these values.")
            Case "report"
                AddText("reportId", "Report ID", True)
            Case "reportDocument"
                AddText("reportDocumentId", "Report document ID", True)
            Case "feeds"
                AddText("feedTypes", "Feed type(s)", False, True)
                AddText("processingStatuses", "Processing statuses")
                AddText("feedMarketplaceIds", "Marketplace IDs")
                AddText("createdSince", "Created since")
                AddText("createdUntil", "Created until")
                AddText("pageSize", "Results per page")
                AddText("feedNextToken", "Next-page token")
            Case "feed"
                AddText("feedId", "Feed ID", True)
            Case "feedDocument"
                AddText("feedDocumentId", "Result feed document ID", True)
            Case "submitFeed"
                AddText("feedType", "Feed type", True)
                AddText("feedMarketplaceIds", "Marketplace IDs")
                AddText("contentType", "Content type", True)
                AddText("content", "Feed content", True, True, 130)
                AddCheck("confirmed", "I understand this submits the feed using exactly these values.")
            Case "inboundPlans"
                AddChoice("status", "Plan status", {"", "ACTIVE", "SHIPPED", "VOIDED"})
                AddChoice("sortBy", "Sort by", {"LAST_UPDATED_TIME", "CREATION_TIME"})
                AddChoice("sortOrder", "Sort order", {"DESC", "ASC"})
                AddText("pageSize", "Results per page")
                AddText("inboundPaginationToken", "Next-page token")
            Case "inboundPlan"
                AddText("inboundPlanId", "Inbound plan ID", True)
            Case "inboundShipment"
                AddText("inboundPlanId", "Inbound plan ID", True)
                AddText("shipmentId", "Shipment ID", True)
            Case "inboundOperationStatus"
                AddText("operationId", "Operation ID", True)
            Case "prepDetails"
                AddText("mskus", "Merchant SKUs (one per line)", True, True)
            Case "createInboundPlan"
                AddText("planName", "Plan name")
                AddText("destinationMarketplaces", "Destination marketplace ID")
                AddText("items", "Items: MSKU, quantity, prep owner, label owner[, expiration, lot code]", True, True, 90)
                AddText("contactName", "Contact name", True)
                AddText("companyName", "Company")
                AddText("addressLine1", "Address line 1", True)
                AddText("addressLine2", "Address line 2")
                AddText("city", "City", True)
                AddText("districtOrCounty", "District / county")
                AddText("stateOrProvinceCode", "State / province")
                AddText("postalCode", "Postal code", True)
                AddText("countryCode", "Country code")
                AddText("phoneNumber", "Phone number", True)
                AddText("email", "Email")
                AddCheck("confirmed", "I understand this creates an inbound plan using exactly these values.")
            Case "itemLabels"
                AddText("items", "Items: MSKU, quantity", True, True)
                AddChoice("labelType", "Label format", {"STANDARD_FORMAT", "THERMAL_PRINTING"}, True, AddressOf RebuildOnChange)
                If S("labelType") = "THERMAL_PRINTING" Then
                    AddText("labelHeight", "Height (25-100)", True)
                    AddText("labelWidth", "Width (25-100)", True)
                Else
                    AddChoice("pageType", "Page type", {"A4_21", "A4_24", "A4_24_64x33", "A4_24_66x35", "A4_24_70x36", "A4_24_70x37", "A4_24i", "A4_27", "A4_40_52x29", "A4_44_48x25", "Letter_30"})
                End If
            Case "shipmentLabels"
                AddText("shipmentId", "Shipment ID", True)
                AddChoice("shipmentLabelType", "Label type", {"UNIQUE", "BARCODE_2D", "PALLET"}, True, AddressOf RebuildOnChange)
                AddText("shipmentPageType", "Page type", True)
                AddText("numberOfPackages", "Number of packages")
                AddText("numberOfPallets", "Number of pallets", S("shipmentLabelType") = "PALLET")
                AddText("packageLabelsToPrint", "Package labels to print", False, True)
                AddText("shipmentPageSize", "Page size")
                AddText("pageStartIndex", "Page start index")
            Case "billOfLading"
                AddText("shipmentId", "Shipment ID", True)
            Case "legacyConvert", "legacyFc"
                AddNote("This is not an SP-API request. It depends on the old private SQL database/tables and business rules, so it remains intentionally disconnected just like the current workbench.")
        End Select
        requestPanel.ResumeLayout()
    End Sub

    Private Sub RebuildOnChange(sender As Object, e As EventArgs)
        SaveVisibleFieldValues()
        BuildOperationFields()
    End Sub

    Private Sub AddText(key As String, labelText As String, Optional required As Boolean = False, Optional multiline As Boolean = False, Optional height As Integer = 70)
        Dim host As New Panel With {.Width = Math.Max(620, requestPanel.ClientSize.Width - 35), .Height = If(multiline, height + 26, 52), .Margin = New Padding(3, 2, 3, 5)}
        Dim label As New Label With {.Text = labelText & If(required, " *", ""), .AutoSize = True, .Location = New Point(0, 3)}
        Dim box As New TextBox With {.Name = key, .Tag = key, .Text = S(key), .Location = New Point(0, 23), .Width = Math.Max(580, host.Width - 8)}
        If multiline Then
            box.Multiline = True
            box.ScrollBars = ScrollBars.Vertical
            box.Height = height
        End If
        AddHandler box.TextChanged, Sub(sender, e) SetUserFieldValue(key, box.Text)
        host.Controls.Add(label)
        host.Controls.Add(box)
        requestPanel.Controls.Add(host)
        FieldControls(key) = box
    End Sub

    Private Sub AddChoice(key As String, labelText As String, options As IEnumerable(Of String), Optional required As Boolean = False, Optional extraHandler As EventHandler = Nothing)
        Dim host As New Panel With {.Width = Math.Max(620, requestPanel.ClientSize.Width - 35), .Height = 52, .Margin = New Padding(3, 2, 3, 5)}
        Dim label As New Label With {.Text = labelText & If(required, " *", ""), .AutoSize = True, .Location = New Point(0, 3)}
        Dim combo As New ComboBox With {.Name = key, .Tag = key, .Location = New Point(0, 23), .Width = Math.Max(580, host.Width - 8), .DropDownStyle = ComboBoxStyle.DropDownList}
        combo.Items.AddRange(options.Cast(Of Object)().ToArray())
        Dim value = S(key)
        If key = "fulfillment" Then value = If(B("isAmazonFulfilled"), "FBA", "Merchant")
        If combo.Items.Contains(value) Then
            combo.SelectedItem = value
        ElseIf combo.Items.Count > 0 Then
            combo.SelectedIndex = 0
        End If
        AddHandler combo.SelectedIndexChanged, Sub(sender, e)
                                                   If key = "fulfillment" Then
                                                       SetUserFieldValue("isAmazonFulfilled", CStr(combo.SelectedItem) = "FBA")
                                                   Else
                                                       SetUserFieldValue(key, CStr(combo.SelectedItem))
                                                   End If
                                               End Sub
        If extraHandler IsNot Nothing Then AddHandler combo.SelectedIndexChanged, extraHandler
        host.Controls.Add(label)
        host.Controls.Add(combo)
        requestPanel.Controls.Add(host)
        FieldControls(key) = combo
    End Sub

    Private Sub ResizeRequestFields()
        If requestPanel Is Nothing Then Return
        Dim availableWidth = Math.Max(580, requestPanel.ClientSize.Width - 35)

        For Each control As Control In requestPanel.Controls
            Dim panel = TryCast(control, Panel)
            If panel IsNot Nothing Then
                panel.Width = availableWidth
                For Each child As Control In panel.Controls
                    If TypeOf child Is TextBox OrElse TypeOf child Is ComboBox Then
                        child.Width = Math.Max(540, availableWidth - 8)
                    End If
                Next
                Continue For
            End If

            Dim label = TryCast(control, Label)
            If label IsNot Nothing Then
                label.MaximumSize = New Size(availableWidth, 0)
                Continue For
            End If

            Dim check = TryCast(control, CheckBox)
            If check IsNot Nothing Then check.MaximumSize = New Size(availableWidth, 0)
        Next
    End Sub

    Private Sub AddSandboxExampleButton()
        Dim button As New Button With {
            .Text = "Load Sandbox example into the form",
            .AutoSize = True,
            .Padding = New Padding(8, 3, 8, 3),
            .Margin = New Padding(3, 2, 3, 8)
        }
        AddHandler button.Click, Sub(sender, e) LoadSandboxExample()
        requestPanel.Controls.Add(button)
    End Sub

    Private Sub AddNote(message As String)
        Dim note As New Label With {
            .Text = message,
            .AutoSize = True,
            .MaximumSize = New Size(Math.Max(620, requestPanel.ClientSize.Width - 35), 0),
            .Padding = New Padding(8),
            .BackColor = Color.FromArgb(245, 245, 245),
            .ForeColor = Color.DimGray,
            .Margin = New Padding(3, 8, 3, 8)
        }
        requestPanel.Controls.Add(note)
    End Sub

    Private Sub AddCheck(key As String, labelText As String)
        Dim check As New CheckBox With {.Name = key, .Tag = key, .Text = labelText, .Checked = B(key), .AutoSize = True, .MaximumSize = New Size(Math.Max(620, requestPanel.ClientSize.Width - 35), 0), .Margin = New Padding(3, 7, 3, 7)}
        AddHandler check.CheckedChanged, Sub(sender, e) SetUserFieldValue(key, check.Checked)
        requestPanel.Controls.Add(check)
        FieldControls(key) = check
    End Sub

    Private Sub SetUserFieldValue(key As String, value As Object)
        FieldValues(key) = value
        If key <> "confirmed" AndAlso IsWriteOperation(CurrentOperation) AndAlso B("confirmed") Then
            ClearWriteConfirmation()
        End If
    End Sub

    Private Sub SaveVisibleFieldValues()
        For Each pair In FieldControls
            Dim box = TryCast(pair.Value, TextBox)
            If box IsNot Nothing Then FieldValues(pair.Key) = box.Text : Continue For
            Dim combo = TryCast(pair.Value, ComboBox)
            If combo IsNot Nothing AndAlso combo.SelectedItem IsNot Nothing Then
                If pair.Key = "fulfillment" Then FieldValues("isAmazonFulfilled") = (CStr(combo.SelectedItem) = "FBA") Else FieldValues(pair.Key) = CStr(combo.SelectedItem)
                Continue For
            End If
            Dim check = TryCast(pair.Value, CheckBox)
            If check IsNot Nothing Then FieldValues(pair.Key) = check.Checked
        Next
    End Sub

    Private Function S(key As String) As String
        Dim value As Object = Nothing
        If Not FieldValues.TryGetValue(key, value) OrElse value Is Nothing Then Return ""
        Return Convert.ToString(value, CultureInfo.InvariantCulture).Trim()
    End Function

    Private Function B(key As String) As Boolean
        Dim value As Object = Nothing
        If Not FieldValues.TryGetValue(key, value) OrElse value Is Nothing Then Return False
        If TypeOf value Is Boolean Then Return DirectCast(value, Boolean)
        Dim result As Boolean
        Return Boolean.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), result) AndAlso result
    End Function

    Private Function SelectedMarketplace() As Marketplace
        Return DirectCast(cboMarketplace.SelectedItem, Marketplace)
    End Function

    Private Function IsSandbox() As Boolean
        Return cboEnvironment.SelectedIndex = 0
    End Function

    Private Function EnvironmentName() As String
        Return If(IsSandbox(), "sandbox", "production")
    End Function

    Private Function OperationHelp(operation As String) As String
        Select Case operation
            Case "catalog" : Return "Find a catalogue item by ASIN/SKU/other identifier, or search by keywords. Related ASIN fetching is optional."
            Case "fees" : Return "Estimate Amazon selling fees for one ASIN or seller SKU at the price and fulfilment method you enter."
            Case "inventory" : Return "Read FBA inventory summaries. Seller SKUs and changed-since are optional filters."
            Case "orders" : Return "Search Orders API 2026 by creation date. Buyer/recipient data is opt-in because it can contain PII."
            Case "order" : Return "Retrieve one Amazon order by order ID."
            Case "reports" : Return "List report jobs. A next-page token is used by itself, exactly as Amazon requires."
            Case "createReport" : Return "Create a report job. Amazon first returns a report ID; use Report status until the job is DONE."
            Case "report" : Return "Check a report job. When DONE, the returned report document ID is saved for the Report document screen."
            Case "reportDocument" : Return "Get report-document metadata and a bounded text preview when Amazon returns a readable document URL."
            Case "feeds" : Return "List feed jobs and processing states."
            Case "feed" : Return "Check one feed. When DONE, the result feed document ID is saved for the Feed processing report screen."
            Case "feedDocument" : Return "Get the feed processing report and a bounded text preview when available."
            Case "submitFeed" : Return "Create the upload document, upload the feed in Production, then create the Amazon feed job."
            Case "inboundPlans" : Return "List Fulfillment Inbound plans with optional status, sort, page size, and pagination token."
            Case "inboundPlan" : Return "Retrieve one inbound plan by ID."
            Case "inboundShipment" : Return "Retrieve one shipment belonging to an inbound plan."
            Case "inboundOperationStatus" : Return "Check an asynchronous inbound operation and review Amazon operation problems/warnings."
            Case "prepDetails" : Return "Get prep requirements for up to 100 merchant SKUs."
            Case "createInboundPlan" : Return "Create an inbound plan using the destination marketplace, ship-from address, and item rows shown below."
            Case "itemLabels" : Return "Request FBA item labels for the entered MSKUs and quantities."
            Case "shipmentLabels" : Return "Request shipment/carton/pallet labels for an existing FBA shipment."
            Case "billOfLading" : Return "Request the bill of lading document for an existing FBA shipment."
            Case "legacyConvert", "legacyFc" : Return "Legacy company-specific SQL utility. It is intentionally disconnected because the private database/schema was not supplied."
        End Select
        Return ""
    End Function

    Private Sub ClearWriteConfirmation()
        FieldValues("confirmed") = False
        Dim confirmation As Control = Nothing
        If FieldControls.TryGetValue("confirmed", confirmation) Then
            Dim check = TryCast(confirmation, CheckBox)
            If check IsNot Nothing AndAlso check.Checked Then check.Checked = False
        End If
    End Sub

    Private Sub InvalidateConnectionState()
        ConnectionVerified = False
        CachedAccessToken = ""
        CachedAccessTokenExpiresUtc = DateTimeOffset.MinValue
        ClearWriteConfirmation()
        lblConnection.Text = "Not tested"
        lblConnection.ForeColor = Color.DimGray
        If cboEnvironment.SelectedIndex >= 0 Then btnTest.Text = "Test " & EnvironmentName() & " connection"
    End Sub

    Private Sub SelectMarketplaceById(id As String)
        For i As Integer = 0 To cboMarketplace.Items.Count - 1
            Dim marketplace = TryCast(cboMarketplace.Items(i), Marketplace)
            If marketplace IsNot Nothing AndAlso marketplace.Id = id Then
                cboMarketplace.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub

    Private Sub LoadSandboxExample()
        If Not IsSandbox() Then Return
        FieldValues("confirmed") = False

        Select Case CurrentOperation
            Case "catalog"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("includedData") = "classifications,dimensions,identifiers,images,productTypes,relationships,salesRanks,summaries,vendorDetails"
                If S("catalogMode") = "keywords" Then
                    FieldValues("query") = "samsung,tv"
                    FieldValues("brandNames") = ""
                    FieldValues("classificationIds") = ""
                    FieldValues("catalogPageSize") = "20"
                    FieldValues("pageToken") = ""
                Else
                    FieldValues("catalogMode") = "identifier"
                    FieldValues("identifierType") = "ASIN"
                    FieldValues("query") = "B07N4M94X4"
                    FieldValues("includeVariations") = False
                    FieldValues("sellerId") = ""
                End If
            Case "fees"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("feeIdType") = "ASIN"
                FieldValues("feeIdentifier") = "B00V5DG6IQ"
                FieldValues("price") = "10"
                FieldValues("shipping") = "10"
                FieldValues("isAmazonFulfilled") = False
                FieldValues("requestIdentifier") = "UmaS1"
                FieldValues("pointsNumber") = "0"
                FieldValues("pointsAmount") = "0"
            Case "orders"
                SelectMarketplaceById("A1VC38T7YXB528")
                FieldValues("createdAfter") = "2024-12-25T00:00:00Z"
                FieldValues("createdBefore") = ""
                FieldValues("statuses") = ""
                FieldValues("fulfilledBy") = ""
                FieldValues("pageSize") = ""
                FieldValues("orderPaginationToken") = ""
                FieldValues("orderIncludedData") = "BUYER,RECIPIENT,PROCEEDS,EXPENSE,PROMOTION,CANCELLATION,FULFILLMENT,PACKAGES"
                FieldValues("includeOrderPii") = False
            Case "order"
                SelectMarketplaceById("A1VC38T7YXB528")
                FieldValues("orderId") = "171-9876543-2109876"
                FieldValues("orderIncludedData") = "BUYER,RECIPIENT,PROCEEDS,EXPENSE,PROMOTION,CANCELLATION,FULFILLMENT,PACKAGES"
                FieldValues("includeOrderPii") = False
            Case "reports"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("reportTypes") = "FEE_DISCOUNTS_REPORT,GET_AFN_INVENTORY_DATA"
                FieldValues("processingStatuses") = "IN_QUEUE,IN_PROGRESS"
                FieldValues("reportMarketplaceIds") = ""
                FieldValues("createdSince") = ""
                FieldValues("createdUntil") = ""
                FieldValues("pageSize") = ""
                FieldValues("reportNextToken") = ""
            Case "createReport"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("reportType") = "GET_MERCHANT_LISTINGS_ALL_DATA"
                FieldValues("dataStartTime") = "2024-03-10T20:11:24.000Z"
                FieldValues("dataEndTime") = ""
                FieldValues("reportMarketplaceIds") = "A1PA6795UKMFR9,ATVPDKIKX0DER"
            Case "report"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("reportId") = "ID323"
            Case "reportDocument"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("reportDocumentId") = "0356cf79-b8b0-4226-b4b9-0ee058ea5760"
            Case "feeds"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("feedTypes") = "POST_PRODUCT_DATA"
                FieldValues("processingStatuses") = "CANCELLED,DONE"
                FieldValues("feedMarketplaceIds") = ""
                FieldValues("createdSince") = ""
                FieldValues("createdUntil") = ""
                FieldValues("pageSize") = "10"
                FieldValues("feedNextToken") = ""
            Case "feed"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("feedId") = "feedId1"
            Case "feedDocument"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("feedDocumentId") = "0356cf79-b8b0-4226-b4b9-0ee058ea5760"
            Case "submitFeed"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("feedType") = "POST_PRODUCT_DATA"
                FieldValues("contentType") = "text/tab-separated-values; charset=UTF-8"
                FieldValues("feedMarketplaceIds") = "ATVPDKIKX0DER,A1F83G8C2ARO7P"
                FieldValues("content") = "Sandbox test content"
            Case "inboundPlans"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("status") = "ACTIVE"
                FieldValues("sortBy") = "LAST_UPDATED_TIME"
                FieldValues("sortOrder") = "ASC"
                FieldValues("pageSize") = "2"
                FieldValues("inboundPaginationToken") = "paginationToken"
            Case "inboundPlan"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("inboundPlanId") = "wf1234abcd-1234-abcd-5678-1234abcd5678"
            Case "inboundShipment"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("inboundPlanId") = "wf1234abcd-1234-abcd-5678-1234abcd5678"
                FieldValues("shipmentId") = "sh1234abcd-1234-abcd-5678-1234abcd5678"
            Case "inboundOperationStatus"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("operationId") = "1234abcd-1234-abcd-5678-1234abcd5678"
            Case "prepDetails"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("mskus") = "msku1" & Environment.NewLine & "msku2"
            Case "createInboundPlan"
                SelectMarketplaceById("A2EUQ1WTGCTBG2")
                FieldValues("destinationMarketplaces") = "A2EUQ1WTGCTBG2"
                FieldValues("planName") = "FBA (03/20/2024, 12:01 PM)"
                FieldValues("items") = "msku, 2, AMAZON, AMAZON, 2024-01-01, lotCode"
                FieldValues("contactName") = "name"
                FieldValues("companyName") = "Acme"
                FieldValues("addressLine1") = "123 example street"
                FieldValues("addressLine2") = "Unit 102"
                FieldValues("city") = "Toronto"
                FieldValues("districtOrCounty") = ""
                FieldValues("stateOrProvinceCode") = "ON"
                FieldValues("postalCode") = "M1M1M1"
                FieldValues("countryCode") = "CA"
                FieldValues("phoneNumber") = "1234567890"
                FieldValues("email") = "email@email.com"
            Case "itemLabels"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("items") = "msku1, 1" & Environment.NewLine & "msku2, 1"
                FieldValues("labelType") = "STANDARD_FORMAT"
                FieldValues("pageType") = "A4_21"
            Case "shipmentLabels"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("shipmentId") = "348975493"
                FieldValues("shipmentLabelType") = "BARCODE_2D"
                FieldValues("shipmentPageType") = "PackageLabel_Letter_2"
                FieldValues("numberOfPackages") = ""
                FieldValues("numberOfPallets") = ""
                FieldValues("packageLabelsToPrint") = ""
                FieldValues("shipmentPageSize") = ""
                FieldValues("pageStartIndex") = ""
            Case "billOfLading"
                SelectMarketplaceById("ATVPDKIKX0DER")
                FieldValues("shipmentId") = "shipmentId"
        End Select

        BuildOperationFields()
    End Sub

    Private Function SandboxGuide(operation As String) As String
        Select Case operation
            Case "catalog"
                If S("catalogMode") = "keywords" Then Return "Sandbox example: US; keywords samsung,tv; includedData classifications,dimensions,identifiers,images,productTypes,relationships,salesRanks,summaries,vendorDetails."
                Return "Sandbox example: US; ASIN B07N4M94X4; same includedData list; turn related fetching off for the exact single-item fixture."
            Case "fees" : Return "Sandbox example: US; ASIN B00V5DG6IQ; price 10; shipping 10; Merchant; identifier UmaS1; points 0 / 0."
            Case "inventory" : Return "Amazon uses dynamic Sandbox for FBA inventory. Enter normal Sandbox SKUs/filters; empty inventory can be valid."
            Case "orders" : Return "Sandbox example: Japan; createdAfter 2024-12-25T00:00:00Z; includedData BUYER,RECIPIENT,PROCEEDS,EXPENSE,PROMOTION,CANCELLATION,FULFILLMENT,PACKAGES."
            Case "order" : Return "Sandbox example: Japan; order 171-9876543-2109876; same includedData override as Search orders."
            Case "reports" : Return "Sandbox example: report types FEE_DISCOUNTS_REPORT,GET_AFN_INVENTORY_DATA; statuses IN_QUEUE,IN_PROGRESS."
            Case "createReport" : Return "Sandbox example: GET_MERCHANT_LISTINGS_ALL_DATA; start 2024-03-10T20:11:24.000Z; marketplaces A1PA6795UKMFR9,ATVPDKIKX0DER."
            Case "report" : Return "Sandbox example: Report ID ID323."
            Case "reportDocument", "feedDocument" : Return "Sandbox example document ID: 0356cf79-b8b0-4226-b4b9-0ee058ea5760."
            Case "feeds" : Return "Sandbox example: feed type POST_PRODUCT_DATA; statuses CANCELLED,DONE; page size 10."
            Case "feed" : Return "Sandbox example: Feed ID feedId1."
            Case "submitFeed" : Return "Sandbox example: POST_PRODUCT_DATA; text/tab-separated-values; charset=UTF-8; marketplaces ATVPDKIKX0DER,A1F83G8C2ARO7P; any non-empty test content."
            Case "inboundPlans" : Return "Sandbox example: ACTIVE; LAST_UPDATED_TIME; ASC; page size 2; paginationToken."
            Case "inboundPlan" : Return "Sandbox example plan: wf1234abcd-1234-abcd-5678-1234abcd5678."
            Case "inboundShipment" : Return "Sandbox example plan wf1234abcd-1234-abcd-5678-1234abcd5678; shipment sh1234abcd-1234-abcd-5678-1234abcd5678."
            Case "inboundOperationStatus" : Return "Sandbox example operation: 1234abcd-1234-abcd-5678-1234abcd5678."
            Case "prepDetails" : Return "Sandbox example: US; msku1 and msku2, one per line."
            Case "createInboundPlan" : Return "Sandbox example: Canada destination; plan FBA (03/20/2024, 12:01 PM); item msku, 2, AMAZON, AMAZON, 2024-01-01, lotCode; Toronto address."
            Case "itemLabels" : Return "Sandbox example: US; msku1,1 and msku2,1; STANDARD_FORMAT; A4_21."
            Case "shipmentLabels" : Return "Sandbox example: shipment 348975493; BARCODE_2D; PackageLabel_Letter_2."
            Case "billOfLading" : Return "Sandbox example: Shipment ID shipmentId."
        End Select
        Return ""
    End Function

    Private Function CredentialsReady() As Boolean
        Return txtClientId.Text.Trim().Length > 0 AndAlso txtClientSecret.Text.Trim().Length > 0 AndAlso txtRefreshToken.Text.Trim().Length > 0
    End Function

    ' -------------------- User actions and workflow orchestration --------------------
    Private Async Function TestConnectionAsync() As Task
        If Not CredentialsReady() Then
            MessageBox.Show("Client ID, client secret, and refresh token are required.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        ToggleBusy(True, "Testing connection...")
        Try
            Dim probe = Await TestConnectionRequestAsync()
            If Not probe.Ok Then
                ConnectionVerified = False
                lblConnection.Text = "Connection failed - see Result"
                lblConnection.ForeColor = Color.DarkRed
                ShowResult("connection", probe)
                Return
            End If

            ConnectionVerified = True
            lblConnection.Text = "Connected to " & EnvironmentName() & " - " & SelectedMarketplace().Name
            lblConnection.ForeColor = Color.DarkGreen
            ShowResult("connection", probe)
        Catch ex As AppException
            ConnectionVerified = False
            lblConnection.Text = "Connection failed - see Result"
            lblConnection.ForeColor = Color.DarkRed
            ShowResult("connection", LocalFailure(ex))
        Catch ex As Exception
            ConnectionVerified = False
            lblConnection.Text = "Connection failed - see Result"
            lblConnection.ForeColor = Color.DarkRed
            ShowResult("connection", LocalFailure(New AppException(ex.Message, 500, "CONNECTION_TEST_FAILED", ex.ToString())))
        Finally
            ToggleBusy(False, "")
        End Try
    End Function

    Private Async Function RunCurrentAsync() As Task
        SaveVisibleFieldValues()
        If CurrentOperation = "legacyConvert" OrElse CurrentOperation = "legacyFc" Then
            MessageBox.Show("This legacy SQL utility is intentionally not connected in the portable SP-API workbench.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        If Not CredentialsReady() Then
            MessageBox.Show("Client ID, client secret, and refresh token are required.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If IsWriteOperation(CurrentOperation) AndAlso Not B("confirmed") Then
            MessageBox.Show("Confirm this Amazon write operation before running it.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If IsWriteOperation(CurrentOperation) AndAlso Not IsSandbox() Then
            Dim operationLabel = Operations.First(Function(x) x.Id = CurrentOperation).Label
            Dim confirmation = MessageBox.Show(
                "This will send a live Production write to Amazon." & Environment.NewLine & Environment.NewLine &
                "Operation: " & operationLabel & Environment.NewLine &
                "Marketplace: " & SelectedMarketplace().Name & Environment.NewLine & Environment.NewLine &
                "Continue?",
                "Confirm Production write",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2)
            If confirmation <> DialogResult.Yes Then Return
        End If

        ToggleBusy(True, "Waiting for Amazon...")
        LastDocumentUrl = ""
        btnOpenDocument.Visible = False
        Try
            Dim result = Await ExecuteOperationAsync(CurrentOperation)
            LastResult = result
            ApplyReturnedIds(CurrentOperation, result.Data)
            ShowResult(CurrentOperation, result)
            UpdateConnectionFromResult(result)
        Catch ex As AppException
            Dim failure = LocalFailure(ex)
            ShowResult(CurrentOperation, failure)
            UpdateConnectionFromResult(failure)
        Catch ex As Exception
            Dim failure = LocalFailure(New AppException(ex.Message, 500, "CLIENT_INTERNAL_ERROR", ex.ToString()))
            ShowResult(CurrentOperation, failure)
            UpdateConnectionFromResult(failure)
        Finally
            ToggleBusy(False, "")
        End Try
    End Function

    Private Sub ToggleBusy(busy As Boolean, message As String)
        btnRun.Enabled = Not busy AndAlso Operations.First(Function(x) x.Id = CurrentOperation).Kind <> "legacy"
        btnTest.Enabled = Not busy
        operationTree.Enabled = Not busy
        requestPanel.Enabled = Not busy
        txtClientId.Enabled = Not busy
        txtClientSecret.Enabled = Not busy
        txtRefreshToken.Enabled = Not busy
        cboEnvironment.Enabled = Not busy
        cboMarketplace.Enabled = Not busy
        chkShowSecrets.Enabled = Not busy
        Cursor = If(busy, Cursors.WaitCursor, Cursors.Default)
        If busy Then lblMeta.Text = message
    End Sub

    Private Sub UpdateConnectionFromResult(result As ApiResult)
        Dim code = If(result.Problem Is Nothing, "", result.Problem.Code).ToLowerInvariant()

        If result.Status = 401 OrElse code.Contains("invalid_grant") OrElse code.Contains("invalid_client") OrElse code.Contains("lwa_") Then
            ConnectionVerified = False
            lblConnection.Text = "Authentication failed - see Result"
            lblConnection.ForeColor = Color.DarkRed
            Return
        End If

        If code.Contains("amazon_network_error") OrElse code.Contains("amazon_timeout") OrElse code.Contains("no_response") Then
            ConnectionVerified = False
            lblConnection.Text = "Connection problem - see Result"
            lblConnection.ForeColor = Color.DarkRed
            Return
        End If

        If result.Ok OrElse result.RequestId <> "" Then
            ConnectionVerified = True
            lblConnection.Text = "Connected to " & EnvironmentName() & " - " & SelectedMarketplace().Name
            lblConnection.ForeColor = Color.DarkGreen
        End If
    End Sub

    Private Function IsWriteOperation(operation As String) As Boolean
        Return operation = "createReport" OrElse operation = "submitFeed" OrElse operation = "createInboundPlan"
    End Function

    Private Function PrettyJson(value As Object, Optional indent As Integer = 0) As String
        Dim sb As New StringBuilder()
        WritePrettyJson(sb, value, indent)
        Return sb.ToString()
    End Function

    Private Sub WritePrettyJson(sb As StringBuilder, value As Object, indent As Integer)
        If value Is Nothing Then sb.Append("null") : Return
        Dim d = TryCast(value, Dictionary(Of String, Object))
        If d IsNot Nothing Then
            sb.AppendLine("{")
            Dim index = 0
            For Each pair In d
                sb.Append(New String(" "c, indent + 2)).Append(Serializer.Serialize(pair.Key)).Append(": ")
                WritePrettyJson(sb, pair.Value, indent + 2)
                index += 1
                If index < d.Count Then sb.Append(",")
                sb.AppendLine()
            Next
            sb.Append(New String(" "c, indent)).Append("}")
            Return
        End If
        Dim list = ListValue(value)
        If (TypeOf value Is Object() OrElse TypeOf value Is ArrayList OrElse TypeOf value Is IEnumerable(Of Object)) Then
            sb.AppendLine("[")
            For i As Integer = 0 To list.Count - 1
                sb.Append(New String(" "c, indent + 2))
                WritePrettyJson(sb, list(i), indent + 2)
                If i < list.Count - 1 Then sb.Append(",")
                sb.AppendLine()
            Next
            sb.Append(New String(" "c, indent)).Append("]")
            Return
        End If
        sb.Append(Serializer.Serialize(value))
    End Sub

    ' -------------------- Readable result view --------------------
    Private Sub ClearResultView(title As String, subtitle As String)
        RenderingView = True
        Try
            CurrentViewOperation = ""
            SetCatalogViewMode(False)
            lblViewTitle.Text = title
            lblViewSubtitle.Text = subtitle
            lblViewDetails.Text = "Details"
            viewGrid.Rows.Clear()
            viewGrid.Columns.Clear()
            viewDetails.Rows.Clear()
            viewSplit.Panel1Collapsed = True
        Finally
            RenderingView = False
        End Try
    End Sub

    Private Sub RenderResultView(operation As String, result As ApiResult)
        RenderingView = True
        Try
            CurrentViewOperation = operation
            SetCatalogViewMode(operation = "catalog" AndAlso result.Ok)
            viewGrid.Rows.Clear()
            viewGrid.Columns.Clear()
            viewDetails.Rows.Clear()
            lblViewDetails.Text = "Details"

            If Not result.Ok Then
                lblViewTitle.Text = "Request failed"
                lblViewSubtitle.Text = If(result.Status > 0, result.Status.ToString(CultureInfo.InvariantCulture) & " " & result.StatusText, result.StatusText)
                viewSplit.Panel1Collapsed = True
                AddViewDetail("Code", If(result.Problem Is Nothing, "", result.Problem.Code))
                AddViewDetail("Message", If(result.Problem Is Nothing, result.ErrorMessage, result.Problem.Message))
                If result.Problem IsNot Nothing Then
                    AddViewDetail("Details", result.Problem.Details)
                    AddViewDetail("What to do", result.Problem.Action)
                    AddViewDetail("Retryable", If(result.Problem.Retryable, "Yes", "No"))
                End If
                AddViewDetail("Amazon request ID", result.RequestId)
                Return
            End If

            Dim data = AsDict(result.Data)
            Select Case operation
                Case "connection"
                    lblViewTitle.Text = "Connection successful"
                    lblViewSubtitle.Text = SelectedMarketplace().Name & " · " & EnvironmentName()
                    viewSplit.Panel1Collapsed = True
                    AddViewDetail("Environment", EnvironmentName())
                    AddViewDetail("Marketplace", SelectedMarketplace().Name)
                    AddViewDetail("Marketplace ID", SelectedMarketplace().Id)
                    AddViewDetail("SP-API endpoint", Endpoint())
                    AddViewDetail("Amazon request ID", result.RequestId)

                Case "catalog"
                    RenderCatalogView(data)

                Case "fees"
                    RenderFeesView(data)

                Case "inventory"
                    RenderInventoryView(data)

                Case "orders"
                    RenderOrdersView(data)

                Case "reports"
                    RenderReportsView(data)

                Case "feeds"
                    RenderFeedsView(data)

                Case "inboundPlans"
                    RenderInboundPlansView(data)

                Case "inboundPlan"
                    RenderInboundPlanView(data)

                Case "prepDetails"
                    RenderPrepDetailsView(data)

                Case "reportDocument", "feedDocument"
                    RenderDocumentView(If(operation = "reportDocument", "Report document", "Feed processing report"), data)

                Case "itemLabels"
                    RenderLabelDocumentsView(data)

                Case "shipmentLabels"
                    RenderDocumentView("Shipment labels", data)

                Case "billOfLading"
                    RenderDocumentView("Bill of lading", data)

                Case "order"
                    RenderOrderView(data)

                Case "inboundShipment"
                    RenderInboundShipmentView(data)

                Case "report"
                    RenderReportStatusView("Report status", data)

                Case "feed"
                    RenderFeedStatusView("Feed status", data)

                Case "createReport"
                    RenderReportStatusView("Report requested", data)

                Case "submitFeed"
                    RenderFeedStatusView("Feed submitted", data)

                Case "createInboundPlan"
                    RenderInboundOperationView("Inbound plan requested", data)

                Case "inboundOperationStatus"
                    RenderInboundOperationView("Inbound operation status", data)

                Case Else
                    RenderObjectView(Operations.Where(Function(op) op.Id = operation).Select(Function(op) op.Label).FirstOrDefault(), "Returned Amazon data", data)
            End Select
        Finally
            RenderingView = False
        End Try

        If viewGrid.Rows.Count > 0 Then
            viewGrid.ClearSelection()
            viewGrid.Rows(0).Selected = True
            viewGrid.CurrentCell = viewGrid.Rows(0).Cells(0)
            ShowSelectedViewRecord()
        End If
    End Sub

    Private Sub RenderCatalogView(data As Dictionary(Of String, Object))
        Dim items = ListValue(GetValue(data, "items"))
        lblViewTitle.Text = If(items.Count = 1, "Product page", "Product catalogue")
        lblViewSubtitle.Text = items.Count.ToString(CultureInfo.InvariantCulture) & " catalogue record(s) returned"
        ConfigureCatalogPageLayout(items.Count)

        Dim family = AsDict(GetValue(data, "family"))
        If family.Count > 0 Then
            Dim returnedCount = Convert.ToString(GetValue(family, "returnedCount"), CultureInfo.InvariantCulture)
            Dim requestedCount = Convert.ToString(GetValue(family, "requestedCount"), CultureInfo.InvariantCulture)
            If returnedCount <> "" AndAlso requestedCount <> "" Then
                lblViewSubtitle.Text &= " · related family " & returnedCount & " of " & requestedCount
            End If
        End If

        If items.Count = 0 Then
            RenderCatalogProductSections(data)
            viewSplit.Panel1Collapsed = True
            Return
        End If

        ConfigureViewTable(
            Tuple.Create("ASIN", 18),
            Tuple.Create("Title", 42),
            Tuple.Create("Brand", 16),
            Tuple.Create("Product type", 16),
            Tuple.Create("Marketplace", 16))

        For Each raw In items
            Dim item = AsDict(raw)
            Dim summary = FirstDictionary(item, "summaries")
            Dim productType = FirstDictionary(item, "productTypes")
            AddViewRow(item,
                       StringValue(GetValue(item, "asin")),
                       FirstNonEmpty(CatalogTitle(item), StringValue(GetValue(summary, "itemName"))),
                       StringValue(GetValue(summary, "brand")),
                       FriendlyToken(StringValue(GetValue(productType, "productType"))),
                       StringValue(GetValue(summary, "marketplaceId")))
        Next
        viewSplit.Panel1Collapsed = (items.Count = 1)
    End Sub

    Private Sub ConfigureCatalogPageLayout(productCount As Integer)
        Dim singleProduct = productCount = 1
        If catalogHero.ColumnStyles.Count > 0 Then catalogHero.ColumnStyles(0).Width = If(singleProduct, 400.0F, 310.0F)
        If viewDetailLayout.RowStyles.Count >= 2 Then viewDetailLayout.RowStyles(1).Height = If(singleProduct, 500.0F, 405.0F)
        catalogHero.Padding = New Padding(If(singleProduct, 18, 4), 12, If(singleProduct, 24, 4), 24)
    End Sub

    Private Sub SetCatalogViewMode(enabled As Boolean)
        catalogHero.Visible = enabled
        catalogDetailTabs.Visible = False
        catalogSectionsPanel.Visible = enabled
        viewDetails.Visible = Not enabled
        If viewDetailLayout.RowStyles.Count >= 2 Then
            viewDetailLayout.RowStyles(1).SizeType = SizeType.Absolute
            viewDetailLayout.RowStyles(1).Height = If(enabled, 245.0F, 0.0F)
        End If
        If viewDetailLayout.RowStyles.Count >= 3 Then
            viewDetailLayout.RowStyles(2).SizeType = If(enabled, SizeType.AutoSize, SizeType.Percent)
            viewDetailLayout.RowStyles(2).Height = If(enabled, 0.0F, 100.0F)
        End If

        If enabled Then
            If viewSplit.Orientation <> Orientation.Vertical Then
                viewSplit.Panel1MinSize = 0
                viewSplit.Panel2MinSize = 0
                viewSplit.SplitterDistance = 1
                viewSplit.Orientation = Orientation.Vertical
            End If
            Dim available = Math.Max(0, viewSplit.ClientSize.Width - viewSplit.SplitterWidth)
            If available > 80 Then viewSplit.SplitterDistance = Math.Min(390, Math.Max(40, CInt(available * 0.34R)))
            viewSplit.Panel2.AutoScroll = True
            viewSplit.Panel2.AutoScrollPosition = Point.Empty
            viewDetailLayout.Dock = DockStyle.Top
            viewDetailLayout.AutoSize = True
            viewDetailLayout.AutoSizeMode = AutoSizeMode.GrowAndShrink
            lblViewDetails.Text = "Selected product"
        Else
            picProduct.CancelAsync()
            picProduct.Image = Nothing
            CatalogImageUrls.Clear()
            CatalogImageIndex = -1
            viewSplit.Panel2.AutoScroll = False
            viewDetailLayout.AutoSize = False
            viewDetailLayout.Dock = DockStyle.Fill
            If viewSplit.Orientation <> Orientation.Horizontal Then
                viewSplit.Panel1MinSize = 0
                viewSplit.Panel2MinSize = 0
                viewSplit.SplitterDistance = 1
                viewSplit.Orientation = Orientation.Horizontal
            End If
            Dim available = Math.Max(0, viewSplit.ClientSize.Height - viewSplit.SplitterWidth)
            If available > 80 Then viewSplit.SplitterDistance = Math.Min(210, Math.Max(40, CInt(available * 0.48R)))
        End If
    End Sub

    Private Sub RenderCatalogProductDetail(item As Dictionary(Of String, Object))
        Dim summary = FirstDictionary(item, "summaries")
        Dim productType = FirstDictionary(item, "productTypes")
        Dim asin = StringValue(GetValue(item, "asin"))
        Dim brand = FirstNonEmpty(StringValue(GetValue(summary, "brand")), FirstCatalogAttribute(item, "brand"))
        Dim title = FirstNonEmpty(CatalogTitle(item), FirstCatalogAttribute(item, "item_name"), "Untitled product")
        Dim typeName = FirstNonEmpty(FriendlyToken(StringValue(GetValue(productType, "productType"))), FriendlyToken(FirstCatalogAttribute(item, "item_type_name")))
        Dim model = FirstCatalogAttribute(item, "model_number")
        Dim manufacturer = FirstCatalogAttribute(item, "manufacturer")

        lblViewDetails.Text = "Selected product · every returned field is listed below"
        lblProductEyebrow.Text = String.Join("  ·  ", New String() {brand.ToUpperInvariant(), asin}.Where(Function(value) value <> ""))
        lblProductTitle.Text = title
        Dim identityParts As New List(Of String)(New String() {typeName, If(asin = "", "", "ASIN " & asin), If(model = "", "", "Model " & model), manufacturer}.Where(Function(value) value <> ""))
        Dim identifiers = CatalogIdentifierLabels(item)
        If identifiers.Count > 0 Then identityParts.Add(String.Join("   ", identifiers))
        lblProductMeta.Text = String.Join(Environment.NewLine, identityParts)
        lblProductCoverage.Text = CatalogCoverageText(item)
        txtProductDescription.Text = CatalogDescription(item)
        If txtProductDescription.Text = "" Then txtProductDescription.Text = "Amazon did not return a description or bullet points for this catalogue record. Every available attribute is still listed below."

        CatalogImageUrls.Clear()
        CatalogImageUrls.AddRange(FindCatalogImageUrls(item))
        CatalogImageIndex = If(CatalogImageUrls.Count > 0, 0, -1)
        BuildCatalogThumbnails()
        ShowCatalogImage()
        RenderCatalogProductSections(item)
    End Sub

    Private Function FindCatalogImageUrls(item As Dictionary(Of String, Object)) As List(Of String)
        Dim ranked As New List(Of Tuple(Of Integer, String))()
        For Each groupObject In ListValue(GetValue(item, "images"))
            Dim group = AsDict(groupObject)
            For Each imageObject In ListValue(GetValue(group, "images"))
                Dim imageData = AsDict(imageObject)
                Dim link = StringValue(GetValue(imageData, "link"))
                If SafeCatalogImageUrl(link) = "" Then Continue For
                Dim imageVariant = StringValue(GetValue(imageData, "variant")).ToUpperInvariant()
                ranked.Add(Tuple.Create(If(imageVariant = "MAIN", 0, 1), link))
            Next
        Next

        Return ranked.OrderBy(Function(entry) entry.Item1).
            Select(Function(entry) entry.Item2).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Private Function SafeCatalogImageUrl(value As String) As String
        Dim uri As Uri = Nothing
        If Not Uri.TryCreate(value, UriKind.Absolute, uri) OrElse uri.Scheme <> Uri.UriSchemeHttps Then Return ""
        Dim host = uri.Host.ToLowerInvariant()
        Dim allowed = host.EndsWith(".media-amazon.com", StringComparison.Ordinal) OrElse
                      host.EndsWith(".ssl-images-amazon.com", StringComparison.Ordinal) OrElse
                      host.EndsWith(".amazon.com", StringComparison.Ordinal) OrElse
                      host.EndsWith(".amazonaws.com", StringComparison.Ordinal) OrElse
                      host.EndsWith(".cloudfront.net", StringComparison.Ordinal)
        Return If(allowed, uri.AbsoluteUri, "")
    End Function

    Private Sub BuildCatalogThumbnails()
        For Each card In CatalogThumbnailCards
            card.Dispose()
        Next
        CatalogThumbnailCards.Clear()
        productThumbnailStrip.Controls.Clear()

        For i As Integer = 0 To Math.Min(CatalogImageUrls.Count, 20) - 1
            Dim thumbnailIndex = i
            Dim card As New Panel With {
                .Width = 86,
                .Height = 76,
                .Padding = New Padding(3),
                .Margin = New Padding(4, 2, 4, 2),
                .BackColor = Color.FromArgb(218, 223, 232),
                .Cursor = Cursors.Hand,
                .Tag = thumbnailIndex
            }
            Dim thumbnail As New PictureBox With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.White,
                .SizeMode = PictureBoxSizeMode.Zoom,
                .Cursor = Cursors.Hand,
                .Tag = thumbnailIndex
            }
            AddHandler card.Click, Sub(sender, e) SelectCatalogImage(thumbnailIndex)
            AddHandler thumbnail.Click, Sub(sender, e) SelectCatalogImage(thumbnailIndex)
            card.Controls.Add(thumbnail)
            productThumbnailStrip.Controls.Add(card)
            CatalogThumbnailCards.Add(card)
            Try
                thumbnail.LoadAsync(CatalogImageUrls(i))
            Catch
                ' The main image retains the actionable error state.
            End Try
        Next
        productThumbnailStrip.Visible = CatalogImageUrls.Count > 0
        UpdateCatalogThumbnailSelection()
    End Sub

    Private Sub SelectCatalogImage(index As Integer)
        If index < 0 OrElse index >= CatalogImageUrls.Count Then Return
        CatalogImageIndex = index
        ShowCatalogImage()
    End Sub

    Private Sub UpdateCatalogThumbnailSelection()
        For i As Integer = 0 To CatalogThumbnailCards.Count - 1
            CatalogThumbnailCards(i).BackColor = If(i = CatalogImageIndex, Color.FromArgb(37, 99, 235), Color.FromArgb(218, 223, 232))
            CatalogThumbnailCards(i).Padding = If(i = CatalogImageIndex, New Padding(4), New Padding(2))
        Next
    End Sub

    Private Sub MoveCatalogImage(offset As Integer)
        If CatalogImageUrls.Count < 2 Then Return
        CatalogImageIndex = (CatalogImageIndex + offset + CatalogImageUrls.Count) Mod CatalogImageUrls.Count
        ShowCatalogImage()
    End Sub

    Private Sub OpenCurrentCatalogImage()
        If CatalogImageIndex < 0 OrElse CatalogImageIndex >= CatalogImageUrls.Count Then Return
        Try
            Process.Start(CatalogImageUrls(CatalogImageIndex))
        Catch ex As Exception
            MessageBox.Show("Could not open the product image: " & ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ShowCatalogImage()
        picProduct.CancelAsync()
        picProduct.Image = Nothing
        Dim hasImage = CatalogImageIndex >= 0 AndAlso CatalogImageIndex < CatalogImageUrls.Count
        btnPreviousImage.Enabled = CatalogImageUrls.Count > 1
        btnNextImage.Enabled = CatalogImageUrls.Count > 1
        UpdateCatalogThumbnailSelection()
        If Not hasImage Then
            lblProductImageStatus.Text = "No image returned"
            Return
        End If

        lblProductImageStatus.Text = "Loading " & (CatalogImageIndex + 1).ToString(CultureInfo.InvariantCulture) & " / " & CatalogImageUrls.Count.ToString(CultureInfo.InvariantCulture)
        Try
            picProduct.LoadAsync(CatalogImageUrls(CatalogImageIndex))
        Catch ex As Exception
            lblProductImageStatus.Text = "Image unavailable"
        End Try
    End Sub

    Private Function FirstCatalogAttribute(item As Dictionary(Of String, Object), key As String) As String
        Dim attributes = AsDict(GetValue(item, "attributes"))
        For Each raw In ListValue(GetValue(attributes, key))
            Dim entry = AsDict(raw)
            Dim value = GetValue(entry, "value")
            If value Is Nothing Then value = GetValue(entry, "displayValue")
            Dim text = ViewValue(value)
            If text <> "" Then Return CleanCatalogText(text)
        Next
        Return ""
    End Function

    Private Function CatalogAttributeValues(item As Dictionary(Of String, Object), key As String) As List(Of String)
        Dim values As New List(Of String)()
        Dim attributes = AsDict(GetValue(item, "attributes"))
        For Each raw In ListValue(GetValue(attributes, key))
            Dim entry = AsDict(raw)
            Dim value = GetValue(entry, "value")
            If value Is Nothing Then value = GetValue(entry, "displayValue")
            Dim text = CleanCatalogText(ViewValue(value))
            If text <> "" Then values.Add(text)
        Next
        Return values.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    Private Function CatalogDescription(item As Dictionary(Of String, Object)) As String
        Dim sections As New List(Of String)()
        Dim descriptions = CatalogAttributeValues(item, "product_description")
        If descriptions.Count > 0 Then sections.Add(String.Join(Environment.NewLine & Environment.NewLine, descriptions))

        Dim bullets = CatalogAttributeValues(item, "bullet_point")
        If bullets.Count > 0 Then
            sections.Add("HIGHLIGHTS" & Environment.NewLine & String.Join(Environment.NewLine, bullets.Select(Function(value) "• " & value)))
        End If

        Dim features = CatalogAttributeValues(item, "special_feature")
        If features.Count = 0 Then features = CatalogAttributeValues(item, "special_features")
        If features.Count > 0 Then sections.Add("FEATURES" & Environment.NewLine & String.Join(Environment.NewLine, features.Select(Function(value) "• " & value)))
        Return String.Join(Environment.NewLine & Environment.NewLine, sections)
    End Function

    Private Function CleanCatalogText(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim decoded = WebUtility.HtmlDecode(value)
        decoded = Regex.Replace(decoded, "<\s*br\s*/?\s*>", Environment.NewLine, RegexOptions.IgnoreCase)
        decoded = Regex.Replace(decoded, "<\s*/?\s*(p|li|ul|ol)\b[^>]*>", Environment.NewLine, RegexOptions.IgnoreCase)
        decoded = Regex.Replace(decoded, "<[^>]+>", "")
        decoded = Regex.Replace(decoded, "[ \t]+", " ")
        decoded = Regex.Replace(decoded, "(\r?\n\s*){3,}", Environment.NewLine & Environment.NewLine)
        Return decoded.Trim()
    End Function

    Private Function CatalogIdentifierLabels(item As Dictionary(Of String, Object)) As List(Of String)
        Dim labels As New List(Of String)()
        For Each groupObject In ListValue(GetValue(item, "identifiers"))
            Dim group = AsDict(groupObject)
            For Each identifierObject In ListValue(GetValue(group, "identifiers"))
                Dim identifier = AsDict(identifierObject)
                Dim kind = StringValue(GetValue(identifier, "identifierType"))
                Dim value = StringValue(GetValue(identifier, "identifier"))
                If kind <> "" AndAlso value <> "" Then labels.Add(kind.ToUpperInvariant() & " · " & value)
            Next
        Next
        Return labels.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    Private Function CatalogCoverageText(item As Dictionary(Of String, Object)) As String
        Dim groups = New Tuple(Of String, String)() {
            Tuple.Create("Attributes", "attributes"),
            Tuple.Create("Classifications", "classifications"),
            Tuple.Create("Dimensions", "dimensions"),
            Tuple.Create("Identifiers", "identifiers"),
            Tuple.Create("Images", "images"),
            Tuple.Create("Product types", "productTypes"),
            Tuple.Create("Relationships", "relationships"),
            Tuple.Create("Sales ranks", "salesRanks"),
            Tuple.Create("Summaries", "summaries"),
            Tuple.Create("Vendor details", "vendorDetails")
        }
        Dim returned = groups.Where(Function(group) HasCatalogContent(GetValue(item, group.Item2))).ToList()
        Return returned.Count.ToString(CultureInfo.InvariantCulture) & "/" & groups.Length.ToString(CultureInfo.InvariantCulture) & " data groups returned" &
               If(returned.Count = 0, "", Environment.NewLine & String.Join(" · ", returned.Select(Function(group) group.Item1)))
    End Function

    Private Sub RenderCatalogProductSections(item As Dictionary(Of String, Object))
        catalogSectionsPanel.SuspendLayout()
        catalogSectionsPanel.Controls.Clear()
        Dim pageWidth = Math.Max(620, viewSplit.Panel2.ClientSize.Width - 54)
        catalogSectionsPanel.Width = pageWidth + catalogSectionsPanel.Padding.Horizontal

        Dim summary = FirstDictionary(item, "summaries")
        Dim productRows As New List(Of KeyValuePair(Of String, String)) From {
            New KeyValuePair(Of String, String)("Colour", StringValue(GetValue(summary, "color"))),
            New KeyValuePair(Of String, String)("Item classification", FriendlyToken(StringValue(GetValue(summary, "itemClassification")))),
            New KeyValuePair(Of String, String)("Manufacturer", StringValue(GetValue(summary, "manufacturer"))),
            New KeyValuePair(Of String, String)("Model number", StringValue(GetValue(summary, "modelNumber"))),
            New KeyValuePair(Of String, String)("Package quantity", ViewValue(GetValue(summary, "packageQuantity"))),
            New KeyValuePair(Of String, String)("Part number", StringValue(GetValue(summary, "partNumber"))),
            New KeyValuePair(Of String, String)("Size", StringValue(GetValue(summary, "size"))),
            New KeyValuePair(Of String, String)("Style", StringValue(GetValue(summary, "style"))),
            New KeyValuePair(Of String, String)("Display group code", StringValue(GetValue(summary, "websiteDisplayGroup"))),
            New KeyValuePair(Of String, String)("Display group", StringValue(GetValue(summary, "websiteDisplayGroupName")))
        }
        productRows = productRows.Where(Function(row) Not String.IsNullOrWhiteSpace(row.Value)).ToList()
        Dim productCard = CreateProductSection("Product details", productRows.Count, pageWidth)
        For Each row In productRows : AddProductPageRow(productCard, row.Key, row.Value) : Next
        AddProductSectionCard(productCard)

        Dim attributes = AsDict(GetValue(item, "attributes"))
        Dim specificationKeys = attributes.Keys.Where(Function(key) Not {"item_name", "brand", "manufacturer", "model_number", "part_number", "product_description", "bullet_point", "special_feature", "special_features"}.Contains(key, StringComparer.OrdinalIgnoreCase)).OrderBy(Function(key) key).ToList()
        Dim specificationCard = CreateProductSection("Specifications", specificationKeys.Count, pageWidth)
        If specificationKeys.Count = 0 Then
            AddProductPageNote(specificationCard, "Amazon did not return product attributes for this product.")
        Else
            For Each key In specificationKeys
                Dim values = CatalogAttributeValues(item, key)
                If values.Count > 0 Then AddProductPageRow(specificationCard, PrettyFieldPath(key), String.Join(Environment.NewLine, values))
            Next
        End If
        AddProductSectionCard(specificationCard)

        Dim dimensionGroups = ListValue(GetValue(item, "dimensions"))
        Dim dimensions = If(dimensionGroups.Count > 0, AsDict(dimensionGroups(0)), New Dictionary(Of String, Object)())
        Dim itemDimensions = AsDict(GetValue(dimensions, "item"))
        Dim packageDimensions = AsDict(GetValue(dimensions, "package"))
        Dim measurementCount = itemDimensions.Count + packageDimensions.Count
        Dim measurementCard = CreateProductSection("Measurements", measurementCount, pageWidth)
        If itemDimensions.Count > 0 Then
            AddProductPageSubheading(measurementCard, "Item")
            AddMeasurementRows(measurementCard, itemDimensions)
        End If
        If packageDimensions.Count > 0 Then
            AddProductPageSubheading(measurementCard, "Package")
            AddMeasurementRows(measurementCard, packageDimensions)
        End If
        If measurementCount = 0 Then AddProductPageNote(measurementCard, "Amazon did not return item or package measurements.")
        AddProductSectionCard(measurementCard)

        Dim categoryRows = CatalogCategoryRows(item)
        Dim categoryCard = CreateProductSection("Category and sales rank", categoryRows.Count, pageWidth)
        For Each row In categoryRows : AddProductPageRow(categoryCard, row.Key, row.Value) : Next
        If categoryRows.Count = 0 Then AddProductPageNote(categoryCard, "Amazon did not return classification or sales-rank data.")
        AddProductSectionCard(categoryCard)

        Dim relatedRows = CatalogRelationshipRows(item)
        Dim relatedCard = CreateProductSection("Related products", relatedRows.Count, pageWidth)
        For Each row In relatedRows : AddProductPageRow(relatedCard, row.Key, row.Value) : Next
        If relatedRows.Count = 0 Then AddProductPageNote(relatedCard, "Amazon did not return parent, child, or variation relationships.")
        AddProductSectionCard(relatedCard)

        Dim vendorRows As New List(Of KeyValuePair(Of String, String))()
        FlattenViewValue(GetValue(item, "vendorDetails"), "", vendorRows, 0)
        Dim vendorCard = CreateProductSection("Vendor details", vendorRows.Count, pageWidth)
        For Each row In vendorRows.Take(80)
            AddProductPageRow(vendorCard, PrettyFieldPath(row.Key), FriendlyDetailValue(row.Key, row.Value))
        Next
        If vendorRows.Count = 0 Then AddProductPageNote(vendorCard, "Amazon did not return vendor details.")
        AddProductSectionCard(vendorCard)

        Dim rawCard = CreateProductSection("Complete returned data", 0, pageWidth)
        AddProductPageNote(rawCard, "The complete, unmodified Amazon response remains available in the Raw response tab.")
        AddProductSectionCard(rawCard)

        catalogSectionsPanel.ResumeLayout(True)
    End Sub

    Private Function CreateProductSection(title As String, count As Integer, width As Integer) As TableLayoutPanel
        Dim card As New TableLayoutPanel With {
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .ColumnCount = 2,
            .RowCount = 1,
            .Width = width,
            .BackColor = Color.White,
            .Padding = New Padding(18, 14, 18, 16),
            .Margin = New Padding(0, 0, 0, 14),
            .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
        }
        card.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 250.0F))
        card.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        card.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Dim heading As New Label With {
            .Text = title & If(count > 0, "   " & count.ToString(CultureInfo.InvariantCulture), ""),
            .AutoSize = True,
            .Dock = DockStyle.Fill,
            .Font = New Font("Bahnschrift SemiBold", 12.0F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(30, 38, 48),
            .BackColor = Color.FromArgb(247, 249, 252),
            .Padding = New Padding(8, 8, 8, 8),
            .Margin = New Padding(0, 0, 0, 8)
        }
        card.Controls.Add(heading, 0, 0)
        card.SetColumnSpan(heading, 2)
        Return card
    End Function

    Private Sub AddProductSectionCard(card As TableLayoutPanel)
        catalogSectionsPanel.Controls.Add(card)
    End Sub

    Private Sub AddProductPageRow(card As TableLayoutPanel, field As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        Dim row = card.RowCount
        card.RowCount += 1
        card.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Dim fieldLabel As New Label With {
            .Text = field,
            .AutoSize = True,
            .Dock = DockStyle.Fill,
            .ForeColor = Color.FromArgb(91, 99, 110),
            .Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold),
            .Padding = New Padding(8, 9, 8, 9),
            .Margin = New Padding(0)
        }
        Dim valueLabel As New Label With {
            .Text = value,
            .AutoSize = True,
            .Dock = DockStyle.Fill,
            .ForeColor = Color.FromArgb(30, 38, 48),
            .Font = New Font("Segoe UI", 9.0F),
            .MaximumSize = New Size(Math.Max(300, card.Width - 310), 0),
            .Padding = New Padding(8, 9, 8, 9),
            .Margin = New Padding(0)
        }
        card.Controls.Add(fieldLabel, 0, row)
        card.Controls.Add(valueLabel, 1, row)
    End Sub

    Private Sub AddProductPageSubheading(card As TableLayoutPanel, title As String)
        Dim row = card.RowCount
        card.RowCount += 1
        card.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Dim label As New Label With {
            .Text = title.ToUpperInvariant(),
            .AutoSize = True,
            .Dock = DockStyle.Fill,
            .Font = New Font("Segoe UI Semibold", 8.5F, FontStyle.Bold),
            .ForeColor = Color.FromArgb(37, 99, 235),
            .BackColor = Color.FromArgb(239, 246, 255),
            .Padding = New Padding(8, 7, 8, 7),
            .Margin = New Padding(0)
        }
        card.Controls.Add(label, 0, row)
        card.SetColumnSpan(label, 2)
    End Sub

    Private Sub AddProductPageNote(card As TableLayoutPanel, text As String)
        Dim row = card.RowCount
        card.RowCount += 1
        card.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Dim label As New Label With {
            .Text = text,
            .AutoSize = True,
            .Dock = DockStyle.Fill,
            .ForeColor = Color.FromArgb(91, 99, 110),
            .Font = New Font("Segoe UI", 9.0F, FontStyle.Italic),
            .Padding = New Padding(8, 12, 8, 12),
            .Margin = New Padding(0)
        }
        card.Controls.Add(label, 0, row)
        card.SetColumnSpan(label, 2)
    End Sub

    Private Sub AddMeasurementRows(card As TableLayoutPanel, measurements As Dictionary(Of String, Object))
        For Each pair In measurements.OrderBy(Function(entry) entry.Key)
            Dim measurement = AsDict(pair.Value)
            Dim value = ViewValue(GetValue(measurement, "value"))
            Dim unit = FriendlyToken(StringValue(GetValue(measurement, "unit")))
            AddProductPageRow(card, PrettyFieldPath(pair.Key), String.Join(" ", New String() {value, unit}.Where(Function(part) part <> "")))
        Next
    End Sub

    Private Function CatalogCategoryRows(item As Dictionary(Of String, Object)) As List(Of KeyValuePair(Of String, String))
        Dim rows As New List(Of KeyValuePair(Of String, String))()
        Dim summary = FirstDictionary(item, "summaries")
        Dim browse = AsDict(GetValue(summary, "browseClassification"))
        Dim browseName = StringValue(GetValue(browse, "displayName"))
        Dim browseId = StringValue(GetValue(browse, "classificationId"))
        If browseName <> "" Then rows.Add(New KeyValuePair(Of String, String)("Browse path", browseName & If(browseId = "", "", " (" & browseId & ")")))

        For Each groupObject In ListValue(GetValue(item, "salesRanks"))
            Dim group = AsDict(groupObject)
            For Each rankObject In ListValue(GetValue(group, "classificationRanks"))
                Dim rank = AsDict(rankObject)
                rows.Add(New KeyValuePair(Of String, String)(FirstNonEmpty(StringValue(GetValue(rank, "title")), "Classification"), "#" & ViewValue(GetValue(rank, "rank"))))
            Next
            For Each rankObject In ListValue(GetValue(group, "displayGroupRanks"))
                Dim rank = AsDict(rankObject)
                rows.Add(New KeyValuePair(Of String, String)(FirstNonEmpty(StringValue(GetValue(rank, "title")), "Display group"), "#" & ViewValue(GetValue(rank, "rank"))))
            Next
        Next
        Return rows
    End Function

    Private Function CatalogRelationshipRows(item As Dictionary(Of String, Object)) As List(Of KeyValuePair(Of String, String))
        Dim rows As New List(Of KeyValuePair(Of String, String))()
        For Each groupObject In ListValue(GetValue(item, "relationships"))
            Dim group = AsDict(groupObject)
            For Each relationObject In ListValue(GetValue(group, "relationships"))
                Dim relation = AsDict(relationObject)
                Dim relationshipType = FriendlyToken(StringValue(GetValue(relation, "type")))
                Dim theme = AsDict(GetValue(relation, "variationTheme"))
                Dim themeText = FriendlyList(GetValue(theme, "attributes"))
                Dim parents = ListValue(GetValue(relation, "parentAsins")).Select(Function(value) ViewValue(value)).Where(Function(value) value <> "").ToList()
                Dim children = ListValue(GetValue(relation, "childAsins")).Select(Function(value) ViewValue(value)).Where(Function(value) value <> "").ToList()
                If parents.Count > 0 Then rows.Add(New KeyValuePair(Of String, String)(FirstNonEmpty(relationshipType, "Parent"), String.Join(", ", parents) & If(themeText = "", "", " · " & themeText)))
                If children.Count > 0 Then rows.Add(New KeyValuePair(Of String, String)(FirstNonEmpty(relationshipType, "Children"), String.Join(", ", children) & If(themeText = "", "", " · " & themeText)))
            Next
        Next
        Return rows
    End Function

    Private Sub RenderCatalogDetailSections(item As Dictionary(Of String, Object))
        For Each grid In New DataGridView() {catalogOverviewGrid, catalogSpecificationsGrid, catalogRelatedGrid, catalogAllFieldsGrid}
            grid.Rows.Clear()
        Next

        Dim summary = FirstDictionary(item, "summaries")
        Dim productType = FirstDictionary(item, "productTypes")
        AddCatalogSection(catalogOverviewGrid, "Product identity")
        AddCatalogValue(catalogOverviewGrid, "ASIN", StringValue(GetValue(item, "asin")))
        AddCatalogValue(catalogOverviewGrid, "Title", FirstNonEmpty(CatalogTitle(item), FirstCatalogAttribute(item, "item_name")))
        AddCatalogValue(catalogOverviewGrid, "Brand", FirstNonEmpty(StringValue(GetValue(summary, "brand")), FirstCatalogAttribute(item, "brand")))
        AddCatalogValue(catalogOverviewGrid, "Manufacturer", FirstNonEmpty(StringValue(GetValue(summary, "manufacturer")), FirstCatalogAttribute(item, "manufacturer")))
        AddCatalogValue(catalogOverviewGrid, "Model number", FirstNonEmpty(StringValue(GetValue(summary, "modelNumber")), FirstCatalogAttribute(item, "model_number")))
        AddCatalogValue(catalogOverviewGrid, "Part number", FirstNonEmpty(StringValue(GetValue(summary, "partNumber")), FirstCatalogAttribute(item, "part_number")))
        AddCatalogValue(catalogOverviewGrid, "Product type", FriendlyToken(StringValue(GetValue(productType, "productType"))))
        AddCatalogValue(catalogOverviewGrid, "Marketplace", StringValue(GetValue(summary, "marketplaceId")))
        AddCatalogValue(catalogOverviewGrid, "Package quantity", ViewValue(GetValue(summary, "packageQuantity")))

        Dim description = CatalogDescription(item)
        If description <> "" Then
            AddCatalogSection(catalogOverviewGrid, "Description and highlights")
            AddCatalogValue(catalogOverviewGrid, "Product description", description)
        End If
        If HasCatalogContent(GetValue(item, "identifiers")) Then
            AddCatalogSection(catalogOverviewGrid, "Identifiers")
            AddCatalogObjectRows(catalogOverviewGrid, GetValue(item, "identifiers"), "", 80)
        End If
        If CatalogImageUrls.Count > 0 Then
            AddCatalogSection(catalogOverviewGrid, "Media")
            AddCatalogValue(catalogOverviewGrid, "Product images", CatalogImageUrls.Count.ToString(CultureInfo.InvariantCulture) & " image(s) returned by Amazon")
        End If

        Dim attributes = AsDict(GetValue(item, "attributes"))
        Dim descriptiveKeys As New HashSet(Of String)({"item_name", "brand", "manufacturer", "model_number", "part_number", "product_description", "bullet_point", "special_feature", "special_features"}, StringComparer.OrdinalIgnoreCase)
        If attributes.Count > 0 Then
            AddCatalogSection(catalogSpecificationsGrid, "Product specifications")
            For Each pair In attributes.OrderBy(Function(entry) entry.Key)
                If descriptiveKeys.Contains(pair.Key) Then Continue For
                Dim values = CatalogAttributeValues(item, pair.Key)
                If values.Count > 0 Then
                    AddCatalogValue(catalogSpecificationsGrid, PrettyFieldPath(pair.Key), String.Join(Environment.NewLine, values))
                Else
                    AddCatalogObjectRows(catalogSpecificationsGrid, pair.Value, pair.Key, 40)
                End If
            Next
        End If
        If HasCatalogContent(GetValue(item, "dimensions")) Then
            AddCatalogSection(catalogSpecificationsGrid, "Measurements")
            AddCatalogObjectRows(catalogSpecificationsGrid, GetValue(item, "dimensions"), "", 100)
        End If
        If catalogSpecificationsGrid.Rows.Count = 0 Then AddCatalogValue(catalogSpecificationsGrid, "Specifications", "Amazon returned no specification attributes for this product.")

        AddCatalogSectionIfPresent(catalogRelatedGrid, "Classification", GetValue(item, "classifications"), 100)
        AddCatalogSectionIfPresent(catalogRelatedGrid, "Category and sales rank", GetValue(item, "salesRanks"), 120)
        AddCatalogSectionIfPresent(catalogRelatedGrid, "Related products and variations", GetValue(item, "relationships"), 120)
        AddCatalogSectionIfPresent(catalogRelatedGrid, "Vendor details", GetValue(item, "vendorDetails"), 120)
        If catalogRelatedGrid.Rows.Count = 0 Then AddCatalogValue(catalogRelatedGrid, "Related data", "Amazon returned no classification, sales-rank, relationship, or vendor records.")

        AddCatalogSection(catalogAllFieldsGrid, "Complete Amazon SP-API response")
        AddCatalogObjectRows(catalogAllFieldsGrid, item, "", 1000)
        For Each grid In New DataGridView() {catalogOverviewGrid, catalogSpecificationsGrid, catalogRelatedGrid, catalogAllFieldsGrid}
            If grid.Rows.Count > 0 Then grid.ClearSelection()
        Next
    End Sub

    Private Sub AddCatalogSectionIfPresent(grid As DataGridView, title As String, value As Object, maxRows As Integer)
        If Not HasCatalogContent(value) Then Return
        AddCatalogSection(grid, title)
        AddCatalogObjectRows(grid, value, "", maxRows)
    End Sub

    Private Function HasCatalogContent(value As Object) As Boolean
        If value Is Nothing Then Return False
        Dim dict = TryCast(value, Dictionary(Of String, Object))
        If dict IsNot Nothing Then Return dict.Count > 0
        If TypeOf value Is Object() OrElse TypeOf value Is ArrayList OrElse TypeOf value Is IEnumerable(Of Object) Then Return ListValue(value).Count > 0
        Return Not String.IsNullOrWhiteSpace(ViewValue(value))
    End Function

    Private Sub AddCatalogSection(grid As DataGridView, title As String)
        Dim index = grid.Rows.Add(title.ToUpperInvariant(), "")
        Dim row = grid.Rows(index)
        row.Height = 36
        row.ReadOnly = True
        row.DefaultCellStyle.BackColor = Color.FromArgb(235, 241, 255)
        row.DefaultCellStyle.ForeColor = Color.FromArgb(30, 64, 175)
        row.DefaultCellStyle.Font = New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold)
        row.DefaultCellStyle.SelectionBackColor = Color.FromArgb(235, 241, 255)
        row.DefaultCellStyle.SelectionForeColor = Color.FromArgb(30, 64, 175)
    End Sub

    Private Sub AddCatalogValue(grid As DataGridView, field As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        grid.Rows.Add(field, value)
    End Sub

    Private Sub AddCatalogObjectRows(grid As DataGridView, value As Object, prefix As String, maxRows As Integer)
        Dim rows As New List(Of KeyValuePair(Of String, String))()
        FlattenViewValue(value, prefix, rows, 0)
        For Each row In rows.Take(maxRows)
            AddCatalogValue(grid, PrettyFieldPath(row.Key), FriendlyDetailValue(row.Key, row.Value))
        Next
        If rows.Count > maxRows Then AddCatalogValue(grid, "More fields", (rows.Count - maxRows).ToString(CultureInfo.InvariantCulture) & " additional field(s) remain available in All SP-API fields and Raw response.")
    End Sub

    Private Sub RenderFeesView(data As Dictionary(Of String, Object))
        lblViewTitle.Text = "Fee estimate"
        Dim payload = AsDict(GetValue(data, "payload"))
        Dim result = AsDict(GetValue(payload, "FeesEstimateResult"))
        Dim estimate = AsDict(GetValue(result, "FeesEstimate"))
        Dim total = AsDict(GetValue(estimate, "TotalFeesEstimate"))
        Dim status = StringValue(GetValue(result, "Status"))
        Dim totalText = MoneyText(total)
        lblViewSubtitle.Text = FirstNonEmpty(FriendlyToken(status), "Amazon fee response") & If(totalText = "", "", " · total " & totalText)

        Dim fees = ListValue(GetValue(estimate, "FeeDetailList"))
        If fees.Count = 0 Then
            viewSplit.Panel1Collapsed = True
            RenderDetailsFromObject(data)
            Return
        End If

        ConfigureViewTable(
            Tuple.Create("Fee", 34),
            Tuple.Create("Amount", 24),
            Tuple.Create("Promotion", 24),
            Tuple.Create("Tax", 18))
        For Each raw In fees
            Dim fee = AsDict(raw)
            Dim amount = AsDict(GetValue(fee, "FinalFee"))
            If amount.Count = 0 Then amount = AsDict(GetValue(fee, "FeeAmount"))
            AddViewRow(fee,
                       FirstNonEmpty(StringValue(GetValue(fee, "FeeType")), StringValue(GetValue(fee, "FeeName"))),
                       MoneyText(amount),
                       MoneyText(AsDict(GetValue(fee, "FeePromotion"))),
                       MoneyText(AsDict(GetValue(fee, "TaxAmount"))))
        Next
        viewSplit.Panel1Collapsed = False
        AddViewDetail("Status", FriendlyToken(status))
        AddViewDetail("Total fees", totalText)
    End Sub

    Private Sub RenderInventoryView(data As Dictionary(Of String, Object))
        Dim payload = AsDict(GetValue(data, "payload"))
        Dim records = ListValue(GetValue(payload, "inventorySummaries"))
        lblViewTitle.Text = "FBA inventory"
        lblViewSubtitle.Text = records.Count.ToString(CultureInfo.InvariantCulture) & " inventory record(s)"

        If records.Count = 0 Then
            viewSplit.Panel1Collapsed = True
            RenderDetailsFromObject(data)
            Return
        End If

        ConfigureViewTable(
            Tuple.Create("Product", 28),
            Tuple.Create("Seller SKU", 18),
            Tuple.Create("ASIN", 15),
            Tuple.Create("Fulfillable", 12),
            Tuple.Create("Reserved", 12),
            Tuple.Create("Unfulfillable", 12),
            Tuple.Create("Total", 11))
        For Each raw In records
            Dim record = AsDict(raw)
            Dim details = AsDict(GetValue(record, "inventoryDetails"))
            Dim reserved = AsDict(GetValue(details, "reservedQuantity"))
            Dim unfulfillable = AsDict(GetValue(details, "unfulfillableQuantity"))
            AddViewRow(record,
                       FirstNonEmpty(StringValue(GetValue(record, "productName")), StringValue(GetValue(record, "sellerSku"))),
                       StringValue(GetValue(record, "sellerSku")),
                       StringValue(GetValue(record, "asin")),
                       ViewValue(GetValue(details, "fulfillableQuantity")),
                       ViewValue(GetValue(reserved, "totalReservedQuantity")),
                       ViewValue(GetValue(unfulfillable, "totalUnfulfillableQuantity")),
                       ViewValue(GetValue(record, "totalQuantity")))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderOrdersView(data As Dictionary(Of String, Object))
        Dim records = ListValue(GetValue(data, "orders"))
        lblViewTitle.Text = "Orders"
        lblViewSubtitle.Text = records.Count.ToString(CultureInfo.InvariantCulture) & " order(s)"

        If records.Count = 0 Then
            viewSplit.Panel1Collapsed = True
            RenderDetailsFromObject(data)
            Return
        End If

        ConfigureViewTable(
            Tuple.Create("Order ID", 24),
            Tuple.Create("Status", 18),
            Tuple.Create("Created", 20),
            Tuple.Create("Fulfilled by", 14),
            Tuple.Create("Sales channel", 16),
            Tuple.Create("Total", 14))
        For Each raw In records
            Dim record = AsDict(raw)
            Dim fulfillment = AsDict(GetValue(record, "fulfillment"))
            Dim proceeds = AsDict(GetValue(record, "proceeds"))
            Dim total = AsDict(GetValue(proceeds, "grandTotal"))
            Dim status = StringValue(GetValue(fulfillment, "fulfillmentStatus"))
            AddViewRow(record,
                       StringValue(GetValue(record, "orderId")),
                       FriendlyToken(status),
                       DisplayDate(GetValue(record, "createdTime")),
                       FriendlyToken(StringValue(GetValue(fulfillment, "fulfilledBy"))),
                       StringValue(GetValue(record, "salesChannel")),
                       MoneyText(total))
            ApplyStatusStyle(viewGrid.Rows(viewGrid.Rows.Count - 1), status)
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderReportsView(data As Dictionary(Of String, Object))
        Dim records = ListValue(GetValue(data, "reports"))
        lblViewTitle.Text = "Reports"
        lblViewSubtitle.Text = records.Count.ToString(CultureInfo.InvariantCulture) & " report job(s)"
        If records.Count = 0 Then RenderObjectView("Reports", "No report jobs returned", data) : Return

        ConfigureViewTable(
            Tuple.Create("Report ID", 20),
            Tuple.Create("Type", 32),
            Tuple.Create("Status", 16),
            Tuple.Create("Created", 18),
            Tuple.Create("Document ID", 22))
        For Each raw In records
            Dim record = AsDict(raw)
            AddViewRow(record,
                       StringValue(GetValue(record, "reportId")),
                       StringValue(GetValue(record, "reportType")),
                       FriendlyToken(StringValue(GetValue(record, "processingStatus"))),
                       DisplayDate(GetValue(record, "createdTime")),
                       StringValue(GetValue(record, "reportDocumentId")))
            ApplyStatusStyle(viewGrid.Rows(viewGrid.Rows.Count - 1), StringValue(GetValue(record, "processingStatus")))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderFeedsView(data As Dictionary(Of String, Object))
        Dim records = ListValue(GetValue(data, "feeds"))
        lblViewTitle.Text = "Feeds"
        lblViewSubtitle.Text = records.Count.ToString(CultureInfo.InvariantCulture) & " feed job(s)"
        If records.Count = 0 Then RenderObjectView("Feeds", "No feed jobs returned", data) : Return

        ConfigureViewTable(
            Tuple.Create("Feed ID", 20),
            Tuple.Create("Type", 30),
            Tuple.Create("Status", 16),
            Tuple.Create("Created", 18),
            Tuple.Create("Result document", 24))
        For Each raw In records
            Dim record = AsDict(raw)
            AddViewRow(record,
                       StringValue(GetValue(record, "feedId")),
                       StringValue(GetValue(record, "feedType")),
                       FriendlyToken(StringValue(GetValue(record, "processingStatus"))),
                       DisplayDate(GetValue(record, "createdTime")),
                       StringValue(GetValue(record, "resultFeedDocumentId")))
            ApplyStatusStyle(viewGrid.Rows(viewGrid.Rows.Count - 1), StringValue(GetValue(record, "processingStatus")))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderInboundPlansView(data As Dictionary(Of String, Object))
        Dim records = ListValue(GetValue(data, "inboundPlans"))
        lblViewTitle.Text = "Inbound plans"
        lblViewSubtitle.Text = records.Count.ToString(CultureInfo.InvariantCulture) & " plan(s)"
        If records.Count = 0 Then RenderObjectView("Inbound plans", "No inbound plans returned", data) : Return

        ConfigureViewTable(
            Tuple.Create("Plan ID", 28),
            Tuple.Create("Name", 26),
            Tuple.Create("Status", 16),
            Tuple.Create("Created", 18),
            Tuple.Create("Updated", 18))
        For Each raw In records
            Dim record = AsDict(raw)
            AddViewRow(record,
                       StringValue(GetValue(record, "inboundPlanId")),
                       StringValue(GetValue(record, "name")),
                       FriendlyToken(StringValue(GetValue(record, "status"))),
                       DisplayDate(FirstNonEmpty(StringValue(GetValue(record, "createdAt")), StringValue(GetValue(record, "createdTime")))),
                       DisplayDate(FirstNonEmpty(StringValue(GetValue(record, "lastUpdatedAt")), StringValue(GetValue(record, "lastUpdatedTime")))))
            ApplyStatusStyle(viewGrid.Rows(viewGrid.Rows.Count - 1), StringValue(GetValue(record, "status")))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderInboundPlanView(data As Dictionary(Of String, Object))
        Dim shipments = ListValue(GetValue(data, "shipments"))
        lblViewTitle.Text = FirstNonEmpty(StringValue(GetValue(data, "name")), "Inbound plan")
        lblViewSubtitle.Text = FirstNonEmpty(StringValue(GetValue(data, "inboundPlanId")), shipments.Count.ToString(CultureInfo.InvariantCulture) & " shipment(s)")

        If shipments.Count = 0 Then
            viewSplit.Panel1Collapsed = True
            RenderDetailsFromObject(data)
            Return
        End If

        ConfigureViewTable(
            Tuple.Create("Shipment ID", 30),
            Tuple.Create("Name", 25),
            Tuple.Create("Status", 18),
            Tuple.Create("Destination", 27))
        For Each raw In shipments
            Dim shipment = AsDict(raw)
            Dim destination = AsDict(GetValue(shipment, "destination"))
            AddViewRow(shipment,
                       StringValue(GetValue(shipment, "shipmentId")),
                       StringValue(GetValue(shipment, "name")),
                       FriendlyToken(StringValue(GetValue(shipment, "status"))),
                       FirstNonEmpty(
                           StringValue(GetValue(destination, "warehouseId")),
                           FriendlyToken(StringValue(GetValue(destination, "destinationType"))),
                           FriendlyToken(StringValue(GetValue(shipment, "destinationType")))))
            ApplyStatusStyle(viewGrid.Rows(viewGrid.Rows.Count - 1), StringValue(GetValue(shipment, "status")))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderOrderView(data As Dictionary(Of String, Object))
        Dim orderId = StringValue(GetValue(data, "orderId"))
        Dim fulfillment = AsDict(GetValue(data, "fulfillment"))
        Dim status = StringValue(GetValue(fulfillment, "fulfillmentStatus"))
        Dim proceeds = AsDict(GetValue(data, "proceeds"))
        Dim total = AsDict(GetValue(proceeds, "grandTotal"))
        Dim items = ListValue(GetValue(data, "orderItems"))

        lblViewTitle.Text = If(orderId = "", "Order", "Order " & orderId)
        lblViewSubtitle.Text = String.Join(" · ", New String() {
            FriendlyToken(status),
            DisplayDate(GetValue(data, "createdTime")),
            MoneyText(total)
        }.Where(Function(value) value <> ""))

        If items.Count = 0 Then
            viewSplit.Panel1Collapsed = True
            RenderDetailsFromObject(data)
            Return
        End If

        ConfigureViewTable(
            Tuple.Create("ASIN", 17),
            Tuple.Create("Seller SKU", 19),
            Tuple.Create("Product", 35),
            Tuple.Create("Qty", 9),
            Tuple.Create("Price", 14))
        For Each raw In items
            Dim item = AsDict(raw)
            Dim product = AsDict(GetValue(item, "product"))
            AddViewRow(item,
                       StringValue(GetValue(product, "asin")),
                       StringValue(GetValue(product, "sellerSku")),
                       StringValue(GetValue(product, "title")),
                       ViewValue(GetValue(item, "quantityOrdered")),
                       MoneyText(AsDict(GetValue(product, "price"))))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderPrepDetailsView(data As Dictionary(Of String, Object))
        Dim records = FirstList(data, "mskuPrepDetails", "prepDetails", "items")
        lblViewTitle.Text = "Prep details"
        lblViewSubtitle.Text = records.Count.ToString(CultureInfo.InvariantCulture) & " SKU preparation record(s)"

        If records.Count = 0 Then
            viewSplit.Panel1Collapsed = True
            RenderDetailsFromObject(data)
            Return
        End If

        ConfigureViewTable(
            Tuple.Create("MSKU", 22),
            Tuple.Create("Category", 24),
            Tuple.Create("Prep types", 36),
            Tuple.Create("Prep owner", 18),
            Tuple.Create("Label owner", 18))
        For Each raw In records
            Dim record = AsDict(raw)
            AddViewRow(record,
                       StringValue(GetValue(record, "msku")),
                       FriendlyToken(StringValue(GetValue(record, "prepCategory"))),
                       FriendlyList(GetValue(record, "prepTypes")),
                       FriendlyList(GetValue(record, "prepOwnerConstraint")),
                       FriendlyList(GetValue(record, "labelOwnerConstraint")))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderLabelDocumentsView(data As Dictionary(Of String, Object))
        Dim records = FirstList(data, "documentDownloads", "documents")
        lblViewTitle.Text = "Item labels"
        lblViewSubtitle.Text = If(records.Count = 0, "No label documents returned", records.Count.ToString(CultureInfo.InvariantCulture) & " label document(s) returned")

        If records.Count = 0 Then
            viewSplit.Panel1Collapsed = True
            RenderDetailsFromObject(data)
            Return
        End If

        ConfigureViewTable(
            Tuple.Create("Type", 24),
            Tuple.Create("Expires", 26),
            Tuple.Create("Download", 50))
        For Each raw In records
            Dim record = AsDict(raw)
            AddViewRow(record,
                       FriendlyToken(StringValue(GetValue(record, "downloadType"))),
                       DisplayDate(GetValue(record, "expiration")),
                       ShortUrl(StringValue(GetValue(record, "uri"))))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderDocumentView(title As String, data As Dictionary(Of String, Object))
        lblViewTitle.Text = title
        Dim documents = FindDocumentUrls(data)
        lblViewSubtitle.Text = If(documents.Count = 0, "Document metadata returned", documents.Count.ToString(CultureInfo.InvariantCulture) & " downloadable document(s) returned")
        viewSplit.Panel1Collapsed = True

        AddViewDetailIfPresent("Document ID", FirstNonEmpty(StringValue(GetValue(data, "reportDocumentId")), StringValue(GetValue(data, "feedDocumentId"))))
        AddViewDetailIfPresent("Compression", FriendlyToken(StringValue(GetValue(data, "compressionAlgorithm"))))

        For i As Integer = 0 To documents.Count - 1
            AddViewDetail("Download " & (i + 1).ToString(CultureInfo.InvariantCulture), ShortUrl(documents(i).Value))
        Next

        Dim downloaded = AsDict(GetValue(data, "downloaded"))
        If downloaded.Count > 0 Then
            AddViewDetailIfPresent("Content type", StringValue(GetValue(downloaded, "contentType")))
            AddViewDetailIfPresent("Bytes read", ViewValue(GetValue(downloaded, "bytesRead")))
            If GetValue(downloaded, "truncated") IsNot Nothing Then AddViewDetail("Preview truncated", ViewValue(GetValue(downloaded, "truncated")))
            AddViewDetailIfPresent("Preview error", StringValue(GetValue(downloaded, "error")))
            Dim preview = StringValue(GetValue(downloaded, "content"))
            If preview <> "" Then AddViewDetail("Text preview", PreviewText(preview))
        End If

        If viewDetails.Rows.Count = 0 Then RenderDetailsFromObject(data)
    End Sub

    Private Sub RenderReportStatusView(title As String, data As Dictionary(Of String, Object))
        Dim reportId = StringValue(GetValue(data, "reportId"))
        Dim status = StringValue(GetValue(data, "processingStatus"))
        lblViewTitle.Text = title
        lblViewSubtitle.Text = String.Join(" · ", New String() {
            FriendlyToken(status),
            If(reportId = "", "", "Report " & reportId)
        }.Where(Function(value) value <> ""))
        viewSplit.Panel1Collapsed = True

        AddViewDetail("Report ID", reportId)
        AddViewDetail("Report type", StringValue(GetValue(data, "reportType")))
        AddViewDetail("Status", FriendlyToken(status))
        AddViewDetail("Created", DisplayDate(GetValue(data, "createdTime")))
        AddViewDetail("Processing started", DisplayDate(GetValue(data, "processingStartTime")))
        AddViewDetail("Processing ended", DisplayDate(GetValue(data, "processingEndTime")))
        AddViewDetail("Report document ID", StringValue(GetValue(data, "reportDocumentId")))
        AddViewDetail("Next step", StringValue(GetValue(data, "nextStep")))
        RemoveBlankViewDetails()
    End Sub

    Private Sub RenderFeedStatusView(title As String, data As Dictionary(Of String, Object))
        Dim feedId = StringValue(GetValue(data, "feedId"))
        Dim status = StringValue(GetValue(data, "processingStatus"))
        lblViewTitle.Text = title
        lblViewSubtitle.Text = String.Join(" · ", New String() {
            FriendlyToken(status),
            If(feedId = "", "", "Feed " & feedId)
        }.Where(Function(value) value <> ""))
        viewSplit.Panel1Collapsed = True

        AddViewDetail("Feed ID", feedId)
        AddViewDetail("Feed type", StringValue(GetValue(data, "feedType")))
        AddViewDetail("Status", FriendlyToken(status))
        AddViewDetail("Created", DisplayDate(GetValue(data, "createdTime")))
        AddViewDetail("Processing started", DisplayDate(GetValue(data, "processingStartTime")))
        AddViewDetail("Processing ended", DisplayDate(GetValue(data, "processingEndTime")))
        AddViewDetail("Input document ID", StringValue(GetValue(data, "inputFeedDocumentId")))
        AddViewDetail("Result document ID", StringValue(GetValue(data, "resultFeedDocumentId")))
        AddViewDetail("Next step", StringValue(GetValue(data, "nextStep")))
        RemoveBlankViewDetails()
    End Sub

    Private Sub RenderInboundOperationView(title As String, data As Dictionary(Of String, Object))
        Dim operationId = StringValue(GetValue(data, "operationId"))
        Dim status = StringValue(GetValue(data, "operationStatus"))
        lblViewTitle.Text = title
        lblViewSubtitle.Text = String.Join(" · ", New String() {
            FriendlyToken(status),
            If(operationId = "", "", "Operation " & operationId)
        }.Where(Function(value) value <> ""))
        viewSplit.Panel1Collapsed = True

        AddViewDetail("Operation ID", operationId)
        AddViewDetail("Operation", FriendlyToken(StringValue(GetValue(data, "operation"))))
        AddViewDetail("Status", FriendlyToken(status))
        AddViewDetail("Inbound plan ID", StringValue(GetValue(data, "inboundPlanId")))
        AddViewDetail("Next step", StringValue(GetValue(data, "nextStep")))

        Dim problems = ListValue(GetValue(data, "operationProblems"))
        For i As Integer = 0 To problems.Count - 1
            Dim problem = AsDict(problems(i))
            AddViewDetail("Problem " & (i + 1).ToString(CultureInfo.InvariantCulture),
                          String.Join(" · ", New String() {
                              FriendlyToken(StringValue(GetValue(problem, "severity"))),
                              StringValue(GetValue(problem, "code")),
                              StringValue(GetValue(problem, "message"))
                          }.Where(Function(value) value <> "")))
        Next
        RemoveBlankViewDetails()
    End Sub

    Private Sub RenderInboundShipmentView(data As Dictionary(Of String, Object))
        Dim shipmentId = StringValue(GetValue(data, "shipmentId"))
        Dim status = StringValue(GetValue(data, "status"))
        lblViewTitle.Text = FirstNonEmpty(StringValue(GetValue(data, "name")), If(shipmentId = "", "Inbound shipment", "Shipment " & shipmentId))
        lblViewSubtitle.Text = String.Join(" · ", New String() {
            FriendlyToken(status),
            If(shipmentId = "", "", shipmentId)
        }.Where(Function(value) value <> ""))
        viewSplit.Panel1Collapsed = True

        AddViewDetail("Shipment ID", shipmentId)
        AddViewDetail("Status", FriendlyToken(status))
        AddViewDetail("Amazon reference ID", StringValue(GetValue(data, "amazonReferenceId")))
        AddViewDetail("Placement option ID", StringValue(GetValue(data, "placementOptionId")))
        AddViewDetail("Transportation option ID", StringValue(GetValue(data, "selectedTransportationOptionId")))
        AddViewDetail("Shipment confirmation ID", StringValue(GetValue(data, "shipmentConfirmationId")))
        RemoveBlankViewDetails()

        Dim rows As New List(Of KeyValuePair(Of String, String))()
        FlattenViewValue(GetValue(data, "destination"), "destination", rows, 0)
        FlattenViewValue(GetValue(data, "source"), "source", rows, 0)
        FlattenViewValue(GetValue(data, "dates"), "dates", rows, 0)
        FlattenViewValue(GetValue(data, "trackingDetails"), "trackingDetails", rows, 0)
        For Each row In rows.Take(80)
            AddViewDetail(PrettyFieldPath(row.Key), FriendlyDetailValue(row.Key, row.Value))
        Next
    End Sub

    Private Sub RenderGenericListView(title As String, subtitle As String, records As List(Of Object), fallback As Object)
        lblViewTitle.Text = title
        lblViewSubtitle.Text = subtitle & If(records.Count > 0, " · " & records.Count.ToString(CultureInfo.InvariantCulture) & " record(s)", "")
        If records.Count = 0 Then
            viewSplit.Panel1Collapsed = True
            RenderDetailsFromObject(fallback)
            Return
        End If

        ConfigureViewTable(Tuple.Create("Record", 100))
        For i As Integer = 0 To records.Count - 1
            Dim record = records(i)
            AddViewRow(record, RecordSummary(record, i + 1))
        Next
        viewSplit.Panel1Collapsed = False
    End Sub

    Private Sub RenderObjectView(title As String, subtitle As String, value As Object)
        lblViewTitle.Text = If(String.IsNullOrWhiteSpace(title), "Result", title)
        lblViewSubtitle.Text = subtitle
        viewSplit.Panel1Collapsed = True
        RenderDetailsFromObject(value)
    End Sub

    Private Sub ConfigureViewTable(ParamArray columns() As Tuple(Of String, Integer))
        viewGrid.Rows.Clear()
        viewGrid.Columns.Clear()
        For Each column In columns
            Dim gridColumn As New DataGridViewTextBoxColumn With {
                .HeaderText = column.Item1,
                .Name = Regex.Replace(column.Item1, "[^A-Za-z0-9]", ""),
                .FillWeight = CSng(Math.Max(5, column.Item2)),
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            }
            viewGrid.Columns.Add(gridColumn)
        Next
    End Sub

    Private Sub AddViewRow(source As Object, ParamArray values() As Object)
        Dim displayValues = values.Select(Function(value) ViewValue(value)).Cast(Of Object)().ToArray()
        Dim index = viewGrid.Rows.Add(displayValues)
        viewGrid.Rows(index).Tag = source
    End Sub

    Private Sub ViewGridSelectionChanged(sender As Object, e As EventArgs)
        If RenderingView Then Return
        ShowSelectedViewRecord()
    End Sub

    Private Sub ShowSelectedViewRecord()
        If viewGrid.SelectedRows.Count = 0 Then Return
        Dim source = viewGrid.SelectedRows(0).Tag
        If source Is Nothing Then Return
        If CurrentViewOperation = "catalog" Then
            RenderCatalogProductDetail(AsDict(source))
        Else
            lblViewDetails.Text = "Selected record details"
            RenderDetailsFromObject(source)
        End If
    End Sub

    Private Sub ViewGridDoubleClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex < 0 Then Return
        If e.RowIndex < ReturnedRecordActions.Count Then
            cboReturnedRecords.SelectedIndex = e.RowIndex
            OpenReturnedRecord(Nothing, EventArgs.Empty)
        End If
    End Sub

    Private Sub RenderDetailsFromObject(value As Object)
        viewDetails.Rows.Clear()
        Dim rows As New List(Of KeyValuePair(Of String, String))()
        FlattenViewValue(value, "", rows, 0)
        If rows.Count = 0 Then
            AddViewDetail("Result", "No additional fields returned.")
            Return
        End If
        For Each row In rows.Take(1000)
            AddViewDetail(PrettyFieldPath(row.Key), FriendlyDetailValue(row.Key, row.Value))
        Next
        If rows.Count > 1000 Then AddViewDetail("More fields", (rows.Count - 1000).ToString(CultureInfo.InvariantCulture) & " additional fields are available in Raw response.")
    End Sub

    Private Sub FlattenViewValue(value As Object, path As String, rows As List(Of KeyValuePair(Of String, String)), depth As Integer)
        If rows.Count > 1400 OrElse depth > 10 Then Return

        If value Is Nothing Then
            If path <> "" Then rows.Add(New KeyValuePair(Of String, String)(path, ""))
            Return
        End If

        Dim dict = TryCast(value, Dictionary(Of String, Object))
        If dict IsNot Nothing Then
            If dict.Count = 0 AndAlso path <> "" Then rows.Add(New KeyValuePair(Of String, String)(path, "(empty)"))
            For Each pair In dict
                Dim nextPath = If(path = "", pair.Key, path & "." & pair.Key)
                FlattenViewValue(pair.Value, nextPath, rows, depth + 1)
            Next
            Return
        End If

        Dim list = ListValue(value)
        If TypeOf value Is Object() OrElse TypeOf value Is ArrayList OrElse TypeOf value Is IEnumerable(Of Object) Then
            If list.Count = 0 Then
                If path <> "" Then rows.Add(New KeyValuePair(Of String, String)(path, "(none)"))
                Return
            End If

            Dim simple = list.All(Function(item) item Is Nothing OrElse TypeOf item Is String OrElse TypeOf item Is ValueType)
            If simple Then
                rows.Add(New KeyValuePair(Of String, String)(path, String.Join(", ", list.Select(Function(item) ViewValue(item)))))
                Return
            End If

            For i As Integer = 0 To Math.Min(list.Count, 100) - 1
                FlattenViewValue(list(i), path & "[" & (i + 1).ToString(CultureInfo.InvariantCulture) & "]", rows, depth + 1)
            Next
            If list.Count > 100 Then rows.Add(New KeyValuePair(Of String, String)(path & ".more", (list.Count - 100).ToString(CultureInfo.InvariantCulture) & " more item(s)"))
            Return
        End If

        rows.Add(New KeyValuePair(Of String, String)(If(path = "", "Value", path), ViewValue(value)))
    End Sub

    Private Sub AddViewDetail(field As String, value As String)
        If String.IsNullOrWhiteSpace(field) Then field = "Value"
        If value Is Nothing Then value = ""
        viewDetails.Rows.Add(field, value)
    End Sub

    Private Sub AddViewDetailIfPresent(field As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        AddViewDetail(field, value)
    End Sub

    Private Function FriendlyDetailValue(path As String, value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim key = path.ToLowerInvariant()

        If key.EndsWith("time", StringComparison.Ordinal) OrElse
           key.EndsWith("date", StringComparison.Ordinal) OrElse
           key.EndsWith("createdat", StringComparison.Ordinal) OrElse
           key.EndsWith("updatedat", StringComparison.Ordinal) OrElse
           key.EndsWith("expiration", StringComparison.Ordinal) Then
            Return DisplayDate(value)
        End If

        If key.Contains("status") OrElse key.EndsWith("owner", StringComparison.Ordinal) OrElse
           key.EndsWith("category", StringComparison.Ordinal) OrElse key.EndsWith("severity", StringComparison.Ordinal) Then
            Return FriendlyToken(value)
        End If

        Return value
    End Function

    Private Function PreviewText(value As String) As String
        If String.IsNullOrEmpty(value) Then Return ""
        Const maxPreviewChars As Integer = 6000
        If value.Length <= maxPreviewChars Then Return value
        Return value.Substring(0, maxPreviewChars) & Environment.NewLine & "… preview shortened here; complete preview is available in Raw response."
    End Function

    Private Function PrettyFieldPath(path As String) As String
        If String.IsNullOrWhiteSpace(path) Then Return "Value"
        Dim value = Regex.Replace(path, "([a-z0-9])([A-Z])", "$1 $2")
        value = Regex.Replace(value, "\[(\d+)\]", " #$1")
        value = value.Replace(".", " › ")
        value = Regex.Replace(value, "\bId\b", "ID", RegexOptions.IgnoreCase)
        value = Regex.Replace(value, "\bAsin\b", "ASIN", RegexOptions.IgnoreCase)
        value = Regex.Replace(value, "\bSku\b", "SKU", RegexOptions.IgnoreCase)
        value = Regex.Replace(value, "\bMsku\b", "MSKU", RegexOptions.IgnoreCase)
        value = Regex.Replace(value, "\bFn sku\b", "FNSKU", RegexOptions.IgnoreCase)
        value = Regex.Replace(value, "\bUrl\b", "URL", RegexOptions.IgnoreCase)
        Return Char.ToUpperInvariant(value(0)) & value.Substring(1)
    End Function

    Private Function ViewValue(value As Object) As String
        If value Is Nothing Then Return ""
        If TypeOf value Is Boolean Then Return If(CBool(value), "Yes", "No")
        Dim text = Convert.ToString(value, CultureInfo.InvariantCulture)
        If text Is Nothing Then Return ""
        If text.Length > 1200 Then Return text.Substring(0, 1200) & " …"
        Return text
    End Function

    Private Function DisplayDate(value As Object) As String
        Dim text = ViewValue(value)
        If text = "" Then Return ""
        If Regex.IsMatch(text, "^\d{4}-\d{2}-\d{2}$") Then Return text
        Dim parsed As DateTimeOffset
        If DateTimeOffset.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, parsed) Then
            Return parsed.ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.InvariantCulture)
        End If
        Return text
    End Function

    Private Function FriendlyToken(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim spaced = value.Replace("_", " ").Trim()
        If spaced = "" Then Return ""
        If spaced.All(Function(ch) Not Char.IsLetter(ch) OrElse Char.IsUpper(ch)) Then
            Return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(spaced.ToLowerInvariant())
        End If
        Return spaced
    End Function

    Private Function FriendlyList(value As Object) As String
        Dim list = ListValue(value)
        If list.Count = 0 Then Return FriendlyToken(ViewValue(value))
        Return String.Join(", ", list.Select(Function(item) FriendlyToken(ViewValue(item))).Where(Function(item) item <> ""))
    End Function

    Private Function ShortUrl(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim uri As Uri = Nothing
        If Uri.TryCreate(value, UriKind.Absolute, uri) Then
            Dim path = uri.AbsolutePath
            If path.Length > 55 Then path = "…" & path.Substring(path.Length - 54)
            Return uri.Host & path
        End If
        If value.Length > 80 Then Return value.Substring(0, 77) & "…"
        Return value
    End Function

    Private Sub ApplyStatusStyle(row As DataGridViewRow, status As String)
        If row Is Nothing OrElse String.IsNullOrWhiteSpace(status) Then Return
        Dim normalized = status.ToUpperInvariant()
        If normalized.Contains("FATAL") OrElse normalized.Contains("FAILED") OrElse normalized.Contains("CANCEL") OrElse normalized.Contains("UNFULFILLABLE") Then
            row.DefaultCellStyle.BackColor = Color.MistyRose
        ElseIf normalized.Contains("DONE") OrElse normalized.Contains("SUCCESS") OrElse normalized.Contains("SHIPPED") Then
            row.DefaultCellStyle.BackColor = Color.Honeydew
        ElseIf normalized.Contains("IN_PROGRESS") OrElse normalized.Contains("IN_QUEUE") OrElse normalized.Contains("PENDING") OrElse normalized.Contains("UNSHIPPED") Then
            row.DefaultCellStyle.BackColor = Color.LemonChiffon
        End If
    End Sub

    Private Sub RemoveBlankViewDetails()
        For i As Integer = viewDetails.Rows.Count - 1 To 0 Step -1
            Dim value = Convert.ToString(viewDetails.Rows(i).Cells(1).Value, CultureInfo.InvariantCulture)
            If String.IsNullOrWhiteSpace(value) Then viewDetails.Rows.RemoveAt(i)
        Next
        If viewDetails.Rows.Count = 0 Then AddViewDetail("Result", "Amazon returned no additional fields.")
    End Sub

    Private Function MoneyText(value As Dictionary(Of String, Object)) As String
        If value Is Nothing OrElse value.Count = 0 Then Return ""
        Dim currency = FirstNonEmpty(StringValue(GetValue(value, "CurrencyCode")), StringValue(GetValue(value, "currencyCode")), StringValue(GetValue(value, "currency")))
        Dim amountObj = GetValue(value, "Amount")
        If amountObj Is Nothing Then amountObj = GetValue(value, "amount")
        Dim amount = ViewValue(amountObj)
        If currency = "" Then Return amount
        If amount = "" Then Return currency
        Return currency & " " & amount
    End Function

    Private Function FirstNonEmpty(ParamArray values() As String) As String
        For Each value In values
            If Not String.IsNullOrWhiteSpace(value) Then Return value
        Next
        Return ""
    End Function

    Private Function FirstDictionary(container As Dictionary(Of String, Object), key As String) As Dictionary(Of String, Object)
        Dim list = ListValue(GetValue(container, key))
        If list.Count = 0 Then Return New Dictionary(Of String, Object)()
        Return AsDict(list(0))
    End Function

    Private Function FirstList(container As Dictionary(Of String, Object), ParamArray keys() As String) As List(Of Object)
        For Each key In keys
            Dim list = ListValue(GetValue(container, key))
            If list.Count > 0 Then Return list
        Next
        Return New List(Of Object)()
    End Function

    Private Function RecordSummary(value As Object, index As Integer) As String
        Dim dict = AsDict(value)
        If dict.Count = 0 Then Return "Record " & index.ToString(CultureInfo.InvariantCulture)
        For Each key In {"asin", "orderId", "amazonOrderId", "reportId", "feedId", "inboundPlanId", "shipmentId", "msku", "sellerSku", "name"}
            Dim found = StringValue(GetValue(dict, key))
            If found <> "" Then Return found
        Next
        Return "Record " & index.ToString(CultureInfo.InvariantCulture)
    End Function

    ' -------------------- Results and follow-up UX --------------------
    Private Sub ShowResult(operation As String, result As ApiResult)
        LastResult = result
        lblMeta.Text = result.Status.ToString(CultureInfo.InvariantCulture) & " " & result.StatusText & If(result.DurationMs > 0, "  |  " & result.DurationMs.ToString(CultureInfo.InvariantCulture) & " ms", "") & If(result.RequestId <> "", "  |  Request " & result.RequestId, "") & If(result.RateLimit <> "", "  |  " & result.RateLimit & " req/s", "")
        Dim requestInfo As New Dictionary(Of String, Object) From {
            {"method", If(result.RequestMethod = "", Nothing, result.RequestMethod)},
            {"path", If(result.RequestPath = "", Nothing, result.RequestPath)},
            {"environment", EnvironmentName()},
            {"marketplaceId", If(cboMarketplace.SelectedItem Is Nothing, Nothing, SelectedMarketplace().Id)}
        }
        Dim envelope As New Dictionary(Of String, Object) From {
            {"ok", result.Ok}, {"status", result.Status}, {"statusText", result.StatusText}, {"requestId", If(result.RequestId = "", Nothing, result.RequestId)},
            {"gatewayId", If(result.GatewayId = "", Nothing, result.GatewayId)}, {"traceId", If(result.TraceId = "", Nothing, result.TraceId)},
            {"rateLimit", If(result.RateLimit = "", Nothing, result.RateLimit)}, {"durationMs", result.DurationMs}, {"attempts", result.Attempts},
            {"request", requestInfo}, {"data", result.Data}
        }
        If result.Problem IsNot Nothing Then envelope("problem") = New Dictionary(Of String, Object) From {{"code", result.Problem.Code}, {"message", result.Problem.Message}, {"details", result.Problem.Details}, {"action", result.Problem.Action}, {"retryable", result.Problem.Retryable}}
        txtRaw.Text = PrettyJson(envelope)
        txtResult.Text = BuildSummary(operation, result)
        RenderResultView(operation, result)
        DocumentUrls.Clear()
        DocumentUrls.AddRange(FindDocumentUrls(result.Data))
        LastDocumentUrl = If(DocumentUrls.Count > 0, DocumentUrls(0).Value, "")

        cboDocuments.Items.Clear()
        For Each document In DocumentUrls
            cboDocuments.Items.Add(document.Key)
        Next
        If cboDocuments.Items.Count > 0 Then cboDocuments.SelectedIndex = 0
        cboDocuments.Visible = DocumentUrls.Count > 1
        btnOpenDocument.Text = If(DocumentUrls.Count > 1, "Open selected document", "Open / download document")
        btnOpenDocument.Visible = DocumentUrls.Count > 0

        ConfigureReturnedRecords(operation, result)
        ConfigureNextStep(operation, result)
        tabs.SelectedIndex = 0
        If operation <> "connection" Then ShowResultWorkspace()
    End Sub

    Private Function BuildSummary(operation As String, result As ApiResult) As String
        Dim sb As New StringBuilder()
        If Not result.Ok Then
            sb.AppendLine("REQUEST FAILED")
            If result.Problem IsNot Nothing Then
                sb.AppendLine("Code: " & result.Problem.Code)
                sb.AppendLine("Message: " & result.Problem.Message)
                If result.Problem.Details <> "" Then sb.AppendLine("Details: " & result.Problem.Details)
                If result.Problem.Action <> "" Then sb.AppendLine("What to do: " & result.Problem.Action)
                sb.AppendLine("Automatic retry: " & If(result.Problem.Retryable, "Safe with backoff", "Not recommended until the cause/state is verified"))
            Else
                sb.AppendLine(result.ErrorMessage)
            End If
            If result.RequestId <> "" Then sb.AppendLine("Amazon request ID: " & result.RequestId)
            Return sb.ToString()
        End If

        sb.AppendLine("SUCCESS")
        Dim data = AsDict(result.Data)
        Select Case operation
            Case "connection"
                sb.AppendLine("LWA credentials accepted and the " & EnvironmentName() & " Sellers API endpoint responded successfully.")
                sb.AppendLine("Marketplace: " & SelectedMarketplace().Name & " (" & SelectedMarketplace().Id & ")")
                sb.AppendLine("Endpoint: " & Endpoint())
            Case "catalog"
                Dim items = ListValue(GetValue(data, "items"))
                sb.AppendLine(items.Count.ToString() & " catalogue record(s) returned")
                For Each raw In items.Take(30)
                    Dim item = AsDict(raw)
                    Dim asin = StringValue(GetValue(item, "asin"))
                    Dim title = CatalogTitle(item)
                    sb.AppendLine(asin & If(title = "", "", " - " & title))
                Next
                Dim family = AsDict(GetValue(data, "family"))
                If family.Count > 0 Then sb.AppendLine("Related family: " & Convert.ToString(GetValue(family, "returnedCount")) & " of " & Convert.ToString(GetValue(family, "requestedCount")) & " returned")
            Case "fees"
                Dim payload = AsDict(GetValue(data, "payload"))
                Dim fee = AsDict(GetValue(payload, "FeesEstimateResult"))
                Dim estimate = AsDict(GetValue(fee, "FeesEstimate"))
                Dim total = AsDict(GetValue(estimate, "TotalFeesEstimate"))
                sb.AppendLine("Status: " & StringValue(GetValue(fee, "Status")))
                If total.Count > 0 Then sb.AppendLine("Total fees: " & StringValue(GetValue(total, "CurrencyCode")) & " " & Convert.ToString(GetValue(total, "Amount"), CultureInfo.InvariantCulture))
                For Each raw In ListValue(GetValue(estimate, "FeeDetailList"))
                    Dim detail = AsDict(raw)
                    Dim amount = AsDict(GetValue(detail, "FinalFee"))
                    If amount.Count = 0 Then amount = AsDict(GetValue(detail, "FeeAmount"))
                    sb.AppendLine("- " & StringValue(GetValue(detail, "FeeType")) & ": " & StringValue(GetValue(amount, "CurrencyCode")) & " " & Convert.ToString(GetValue(amount, "Amount"), CultureInfo.InvariantCulture))
                Next
            Case "createReport"
                sb.AppendLine("Report ID: " & StringValue(GetValue(data, "reportId")))
            Case "report"
                sb.AppendLine("Report ID: " & StringValue(GetValue(data, "reportId")))
                sb.AppendLine("Status: " & StringValue(GetValue(data, "processingStatus")))
                If StringValue(GetValue(data, "reportDocumentId")) <> "" Then sb.AppendLine("Report document ID: " & StringValue(GetValue(data, "reportDocumentId")))
            Case "submitFeed"
                sb.AppendLine("Feed ID: " & StringValue(GetValue(data, "feedId")))
                sb.AppendLine("Input feed document ID: " & StringValue(GetValue(data, "inputFeedDocumentId")))
            Case "feed"
                sb.AppendLine("Feed ID: " & StringValue(GetValue(data, "feedId")))
                sb.AppendLine("Status: " & StringValue(GetValue(data, "processingStatus")))
                If StringValue(GetValue(data, "resultFeedDocumentId")) <> "" Then sb.AppendLine("Result feed document ID: " & StringValue(GetValue(data, "resultFeedDocumentId")))
            Case "createInboundPlan"
                sb.AppendLine("Inbound plan ID: " & StringValue(GetValue(data, "inboundPlanId")))
                sb.AppendLine("Operation ID: " & StringValue(GetValue(data, "operationId")))
            Case "inboundOperationStatus"
                sb.AppendLine("Operation ID: " & StringValue(GetValue(data, "operationId")))
                sb.AppendLine("Status: " & StringValue(GetValue(data, "operationStatus")))
            Case "reports"
                sb.AppendLine(ListValue(GetValue(data, "reports")).Count.ToString() & " report job(s) returned")
            Case "feeds"
                sb.AppendLine(ListValue(GetValue(data, "feeds")).Count.ToString() & " feed job(s) returned")
            Case "inboundPlans"
                sb.AppendLine(ListValue(GetValue(data, "inboundPlans")).Count.ToString() & " inbound plan(s) returned")
            Case "orders"
                sb.AppendLine(ListValue(GetValue(data, "orders")).Count.ToString() & " order(s) returned")
            Case "inventory"
                Dim payload = AsDict(GetValue(data, "payload"))
                sb.AppendLine(ListValue(GetValue(payload, "inventorySummaries")).Count.ToString() & " inventory record(s) returned")
            Case "reportDocument", "feedDocument", "itemLabels", "shipmentLabels", "billOfLading"
                Dim documentCount = FindDocumentUrls(data).Count
                If documentCount = 1 Then
                    sb.AppendLine("1 document URL returned and ready to open/download.")
                ElseIf documentCount > 1 Then
                    sb.AppendLine(documentCount.ToString(CultureInfo.InvariantCulture) & " document URLs returned. Choose the document beside the Open button.")
                End If
        End Select
        Dim nextStep = StringValue(GetValue(data, "nextStep"))
        If nextStep <> "" Then sb.AppendLine().AppendLine("Next: " & nextStep)
        Return sb.ToString().TrimEnd()
    End Function

    Private Function CatalogTitle(item As Dictionary(Of String, Object)) As String
        For Each raw In ListValue(GetValue(item, "summaries"))
            Dim summary = AsDict(raw)
            Dim t = StringValue(GetValue(summary, "itemName"))
            If t <> "" Then Return t
        Next
        Return ""
    End Function

    Private Sub ApplyReturnedIds(operation As String, dataObj As Object)
        Dim data = AsDict(dataObj)
        Select Case operation
            Case "createReport"
                If StringValue(GetValue(data, "reportId")) <> "" Then FieldValues("reportId") = StringValue(GetValue(data, "reportId"))
            Case "report"
                If StringValue(GetValue(data, "reportDocumentId")) <> "" Then FieldValues("reportDocumentId") = StringValue(GetValue(data, "reportDocumentId"))
            Case "submitFeed"
                If StringValue(GetValue(data, "feedId")) <> "" Then FieldValues("feedId") = StringValue(GetValue(data, "feedId"))
            Case "feed"
                If StringValue(GetValue(data, "resultFeedDocumentId")) <> "" Then FieldValues("feedDocumentId") = StringValue(GetValue(data, "resultFeedDocumentId"))
            Case "createInboundPlan"
                If StringValue(GetValue(data, "inboundPlanId")) <> "" Then FieldValues("inboundPlanId") = StringValue(GetValue(data, "inboundPlanId"))
                If StringValue(GetValue(data, "operationId")) <> "" Then FieldValues("operationId") = StringValue(GetValue(data, "operationId"))
            Case "inboundPlans"
                Dim plans = ListValue(GetValue(data, "inboundPlans"))
                If plans.Count > 0 Then
                    Dim id = StringValue(GetValue(AsDict(plans(0)), "inboundPlanId"))
                    If id <> "" Then FieldValues("inboundPlanId") = id
                End If
            Case "inboundPlan"
                Dim shipments = ListValue(GetValue(data, "shipments"))
                If shipments.Count > 0 Then
                    Dim id = StringValue(GetValue(AsDict(shipments(0)), "shipmentId"))
                    If id <> "" Then FieldValues("shipmentId") = id
                End If
            Case "orders"
                Dim orders = ListValue(GetValue(data, "orders"))
                If orders.Count > 0 Then
                    Dim id = StringValue(GetValue(AsDict(orders(0)), "orderId"))
                    If id <> "" Then FieldValues("orderId") = id
                End If
        End Select
    End Sub

    Private Function FindDocumentUrl(dataObj As Object) As String
        Dim documents = FindDocumentUrls(dataObj)
        If documents.Count = 0 Then Return ""
        Return documents(0).Value
    End Function

    Private Function FindDocumentUrls(dataObj As Object) As List(Of KeyValuePair(Of String, String))
        Dim output As New List(Of KeyValuePair(Of String, String))()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim data = AsDict(dataObj)

        Dim direct = SafeHttpsUrl(StringValue(GetValue(data, "url")))
        If direct <> "" AndAlso seen.Add(direct) Then
            Dim label = StringValue(GetValue(data, "reportDocumentId"))
            If label = "" Then label = StringValue(GetValue(data, "feedDocumentId"))
            If label = "" Then label = "Amazon document"
            output.Add(QPair(label, direct))
        End If

        Dim downloads = ListValue(GetValue(data, "documentDownloads"))
        For i As Integer = 0 To downloads.Count - 1
            Dim download = AsDict(downloads(i))
            Dim uri = SafeHttpsUrl(StringValue(GetValue(download, "uri")))
            If uri = "" OrElse Not seen.Add(uri) Then Continue For
            Dim label = StringValue(GetValue(download, "downloadType"))
            If label = "" Then label = "Document " & (i + 1).ToString(CultureInfo.InvariantCulture)
            output.Add(QPair(label, uri))
        Next

        Dim payload = AsDict(GetValue(data, "payload"))
        For Each key In {"DownloadURL", "downloadURL", "downloadUrl"}
            Dim url = SafeHttpsUrl(StringValue(GetValue(payload, key)))
            If url <> "" AndAlso seen.Add(url) Then
                output.Add(QPair("Amazon document", url))
            End If
        Next

        Return output
    End Function

    Private Function SafeHttpsUrl(value As String) As String
        If value = "" Then Return ""
        Dim uri As Uri = Nothing
        If Not Uri.TryCreate(value, UriKind.Absolute, uri) Then Return ""
        If uri.Scheme <> Uri.UriSchemeHttps Then Return ""
        Return uri.AbsoluteUri
    End Function

    Private Sub ConfigureReturnedRecords(operation As String, result As ApiResult)
        ReturnedRecordActions.Clear()
        cboReturnedRecords.Items.Clear()
        cboReturnedRecords.Visible = False
        btnOpenReturnedRecord.Visible = False
        If Not result.Ok Then Return

        Dim data = AsDict(result.Data)
        Select Case operation
            Case "orders"
                For Each raw In ListValue(GetValue(data, "orders"))
                    Dim record = AsDict(raw)
                    Dim id = StringValue(GetValue(record, "orderId"))
                    If id = "" Then id = StringValue(GetValue(record, "amazonOrderId"))
                    If id = "" Then Continue For
                    Dim fulfillment = AsDict(GetValue(record, "fulfillment"))
                    Dim status = StringValue(GetValue(fulfillment, "fulfillmentStatus"))
                    If status = "" Then status = StringValue(GetValue(record, "orderStatus"))
                    AddReturnedRecord("order", RecordLabel(id, "", status), New Dictionary(Of String, String) From {{"orderId", id}})
                Next

            Case "reports"
                For Each raw In ListValue(GetValue(data, "reports"))
                    Dim record = AsDict(raw)
                    Dim id = StringValue(GetValue(record, "reportId"))
                    If id = "" Then Continue For
                    AddReturnedRecord("report", RecordLabel(id, StringValue(GetValue(record, "reportType")), StringValue(GetValue(record, "processingStatus"))), New Dictionary(Of String, String) From {{"reportId", id}})
                Next

            Case "feeds"
                For Each raw In ListValue(GetValue(data, "feeds"))
                    Dim record = AsDict(raw)
                    Dim id = StringValue(GetValue(record, "feedId"))
                    If id = "" Then Continue For
                    AddReturnedRecord("feed", RecordLabel(id, StringValue(GetValue(record, "feedType")), StringValue(GetValue(record, "processingStatus"))), New Dictionary(Of String, String) From {{"feedId", id}})
                Next

            Case "inboundPlans"
                For Each raw In ListValue(GetValue(data, "inboundPlans"))
                    Dim record = AsDict(raw)
                    Dim id = StringValue(GetValue(record, "inboundPlanId"))
                    If id = "" Then Continue For
                    AddReturnedRecord("inboundPlan", RecordLabel(id, StringValue(GetValue(record, "name")), StringValue(GetValue(record, "status"))), New Dictionary(Of String, String) From {{"inboundPlanId", id}})
                Next

            Case "inboundPlan"
                Dim inboundPlanId = StringValue(GetValue(data, "inboundPlanId"))
                If inboundPlanId = "" Then inboundPlanId = S("inboundPlanId")
                For Each raw In ListValue(GetValue(data, "shipments"))
                    Dim record = AsDict(raw)
                    Dim shipmentId = StringValue(GetValue(record, "shipmentId"))
                    If shipmentId = "" OrElse inboundPlanId = "" Then Continue For
                    AddReturnedRecord("inboundShipment", RecordLabel(shipmentId, StringValue(GetValue(record, "name")), StringValue(GetValue(record, "status"))), New Dictionary(Of String, String) From {{"inboundPlanId", inboundPlanId}, {"shipmentId", shipmentId}})
                Next
        End Select

        For Each action In ReturnedRecordActions
            cboReturnedRecords.Items.Add(action)
        Next
        If cboReturnedRecords.Items.Count > 0 Then cboReturnedRecords.SelectedIndex = 0
        cboReturnedRecords.Visible = ReturnedRecordActions.Count > 0
        btnOpenReturnedRecord.Visible = ReturnedRecordActions.Count > 0
    End Sub

    Private Sub AddReturnedRecord(operationId As String, label As String, fields As Dictionary(Of String, String))
        ReturnedRecordActions.Add(New ReturnedRecordAction With {.OperationId = operationId, .Label = label, .Fields = fields})
    End Sub

    Private Function RecordLabel(id As String, secondary As String, status As String) As String
        Dim parts As New List(Of String)()
        If secondary <> "" Then parts.Add(secondary)
        If status <> "" Then parts.Add(status)
        parts.Add(id)
        Return String.Join(" · ", parts)
    End Function

    Private Sub OpenReturnedRecord(sender As Object, e As EventArgs)
        Dim index = cboReturnedRecords.SelectedIndex
        If index < 0 OrElse index >= ReturnedRecordActions.Count Then Return

        Dim action = ReturnedRecordActions(index)
        For Each pair In action.Fields
            FieldValues(pair.Key) = pair.Value
        Next
        NavigateToOperation(action.OperationId)
    End Sub

    Private Sub ConfigureNextStep(operation As String, result As ApiResult)
        btnNextStep.Visible = False
        NextOperationId = ""
        NextFieldKey = ""
        NextFieldValue = ""

        Dim data = AsDict(result.Data)

        ' A failed feed can still return a processing report that explains record-level errors.
        If operation = "feed" Then
            Dim resultDocumentId = StringValue(GetValue(data, "resultFeedDocumentId"))
            If resultDocumentId <> "" Then
                SetNextStep("feedDocument", "feedDocumentId", resultDocumentId, "Next: Open processing report")
            End If
        End If
        If Not result.Ok Then
            btnNextStep.Visible = NextOperationId <> ""
            Return
        End If

        Dim pagination = AsDict(GetValue(data, "pagination"))
        Select Case operation
            Case "catalog"
                Dim token = StringValue(GetValue(pagination, "nextToken"))
                If token <> "" Then SetNextStep("catalog", "pageToken", token, "Next: Prepare next page")
            Case "inventory"
                Dim token = StringValue(GetValue(pagination, "nextToken"))
                If token <> "" Then SetNextStep("inventory", "inventoryNextToken", token, "Next: Prepare next page")
            Case "orders"
                Dim token = StringValue(GetValue(pagination, "nextToken"))
                If token <> "" Then SetNextStep("orders", "orderPaginationToken", token, "Next: Prepare next page")
            Case "reports"
                Dim token = StringValue(GetValue(data, "nextToken"))
                If token <> "" Then SetNextStep("reports", "reportNextToken", token, "Next: Prepare next page")
            Case "feeds"
                Dim token = StringValue(GetValue(data, "nextToken"))
                If token <> "" Then SetNextStep("feeds", "feedNextToken", token, "Next: Prepare next page")
            Case "inboundPlans"
                Dim token = StringValue(GetValue(pagination, "nextToken"))
                If token <> "" Then SetNextStep("inboundPlans", "inboundPaginationToken", token, "Next: Prepare next page")
            Case "createReport"
                Dim reportId = StringValue(GetValue(data, "reportId"))
                If reportId <> "" Then SetNextStep("report", "reportId", reportId, "Next: Check report status")
            Case "report"
                Dim reportId = StringValue(GetValue(data, "reportId"))
                If reportId = "" Then reportId = S("reportId")
                Dim documentId = StringValue(GetValue(data, "reportDocumentId"))
                Dim status = StringValue(GetValue(data, "processingStatus"))
                If documentId <> "" Then
                    SetNextStep("reportDocument", "reportDocumentId", documentId, "Next: Open report document")
                ElseIf (status = "IN_QUEUE" OrElse status = "IN_PROGRESS") AndAlso reportId <> "" Then
                    SetNextStep("report", "reportId", reportId, "Next: Check report status again")
                End If
            Case "submitFeed"
                Dim feedId = StringValue(GetValue(data, "feedId"))
                If feedId <> "" Then SetNextStep("feed", "feedId", feedId, "Next: Check feed status")
            Case "feed"
                Dim feedId = StringValue(GetValue(data, "feedId"))
                If feedId = "" Then feedId = S("feedId")
                Dim resultDocumentId = StringValue(GetValue(data, "resultFeedDocumentId"))
                Dim status = StringValue(GetValue(data, "processingStatus"))
                If resultDocumentId <> "" Then
                    SetNextStep("feedDocument", "feedDocumentId", resultDocumentId, "Next: Open processing report")
                ElseIf (status = "IN_QUEUE" OrElse status = "IN_PROGRESS") AndAlso feedId <> "" Then
                    SetNextStep("feed", "feedId", feedId, "Next: Check feed status again")
                End If
            Case "createInboundPlan"
                Dim operationId = StringValue(GetValue(data, "operationId"))
                If operationId <> "" Then SetNextStep("inboundOperationStatus", "operationId", operationId, "Next: Check operation status")
            Case "inboundOperationStatus"
                Dim operationId = StringValue(GetValue(data, "operationId"))
                If operationId = "" Then operationId = S("operationId")
                Dim status = StringValue(GetValue(data, "operationStatus"))
                If status = "IN_PROGRESS" AndAlso operationId <> "" Then
                    SetNextStep("inboundOperationStatus", "operationId", operationId, "Next: Check operation status again")
                ElseIf status = "SUCCESS" AndAlso S("inboundPlanId") <> "" Then
                    SetNextStep("inboundPlan", "inboundPlanId", S("inboundPlanId"), "Next: Open inbound plan")
                End If
        End Select

        btnNextStep.Visible = NextOperationId <> ""
    End Sub

    Private Sub SetNextStep(operationId As String, fieldKey As String, fieldValue As String, label As String)
        NextOperationId = operationId
        NextFieldKey = fieldKey
        NextFieldValue = fieldValue
        btnNextStep.Text = label
    End Sub

    Private Sub OpenNextStep(sender As Object, e As EventArgs)
        If NextOperationId = "" Then Return

        Dim targetOperation = NextOperationId
        Dim patchKey = NextFieldKey
        Dim patchValue = NextFieldValue

        If String.Equals(targetOperation, CurrentOperation, StringComparison.Ordinal) Then
            SaveVisibleFieldValues()
            If patchKey <> "" Then FieldValues(patchKey) = patchValue
            ClearWriteConfirmation()
            BuildOperationFields()
            txtResult.Text = If(patchKey.EndsWith("Token", StringComparison.OrdinalIgnoreCase) OrElse patchKey = "pageToken",
                                "Next-page token loaded. Run the request to fetch the next page.",
                                "Follow-up values loaded. Run the request again.")
            txtRaw.Text = ""
            lblMeta.Text = ""
            btnOpenDocument.Visible = False
            cboDocuments.Visible = False
            cboDocuments.Items.Clear()
            DocumentUrls.Clear()
            cboReturnedRecords.Visible = False
            cboReturnedRecords.Items.Clear()
            btnOpenReturnedRecord.Visible = False
            ReturnedRecordActions.Clear()
            btnNextStep.Visible = False
            NextOperationId = ""
            NextFieldKey = ""
            NextFieldValue = ""
            Return
        End If

        If patchKey <> "" Then FieldValues(patchKey) = patchValue
        NavigateToOperation(targetOperation)
    End Sub

    Private Sub NavigateToOperation(operationId As String)
        For Each groupNode As TreeNode In operationTree.Nodes
            For Each node As TreeNode In groupNode.Nodes
                If node.Tag IsNot Nothing AndAlso String.Equals(CStr(node.Tag), operationId, StringComparison.Ordinal) Then
                    operationTree.SelectedNode = node
                    node.EnsureVisible()
                    Return
                End If
            Next
        Next
        SelectOperation(operationId)
    End Sub

    Private Sub CopyTextToClipboard(value As String, label As String)
        If String.IsNullOrWhiteSpace(value) Then Return
        Try
            Clipboard.SetText(value)
            lblMeta.Text = label & " copied to clipboard."
        Catch ex As Exception
            MessageBox.Show("Could not copy to the clipboard: " & ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub OpenDocument(sender As Object, e As EventArgs)
        If DocumentUrls.Count = 0 Then Return

        Dim index = cboDocuments.SelectedIndex
        If index < 0 OrElse index >= DocumentUrls.Count Then index = 0
        LastDocumentUrl = DocumentUrls(index).Value
        If LastDocumentUrl = "" Then Return

        Try
            Process.Start(LastDocumentUrl)
        Catch ex As Exception
            MessageBox.Show("Could not open the document URL: " & ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
