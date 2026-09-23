Option Explicit On
Option Strict On
Option Infer On

Imports System
Imports System.Collections
Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
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
    Private ReadOnly requestPanel As New FlowLayoutPanel()
    Private ReadOnly lblOperation As New Label()
    Private ReadOnly lblSandbox As New Label()
    Private ReadOnly btnRun As New Button()
    Private ReadOnly lblMeta As New Label()
    Private ReadOnly txtResult As New TextBox()
    Private ReadOnly txtRaw As New TextBox()
    Private ReadOnly btnOpenDocument As New Button()
    Private ReadOnly btnNextStep As New Button()
    Private ReadOnly cboDocuments As New ComboBox()
    Private ReadOnly cboReturnedRecords As New ComboBox()
    Private ReadOnly btnOpenReturnedRecord As New Button()
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

        Dim rightSplit As New SplitContainer With {.Dock = DockStyle.Fill, .Orientation = Orientation.Horizontal, .SplitterDistance = 430}
        mainSplit.Panel2.Controls.Add(rightSplit)

        Dim requestHost As New Panel With {.Dock = DockStyle.Fill, .Padding = New Padding(10)}
        rightSplit.Panel1.Controls.Add(requestHost)
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
        Dim resultTab As New TabPage("Result")
        Dim rawTab As New TabPage("Raw response")
        tabs.TabPages.Add(resultTab)
        tabs.TabPages.Add(rawTab)
        rightSplit.Panel2.Controls.Add(tabs)

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

        cboReturnedRecords.DropDownStyle = ComboBoxStyle.DropDownList
        cboReturnedRecords.Width = 285
        cboReturnedRecords.Visible = False
        resultActions.Controls.Add(cboReturnedRecords)

        btnOpenReturnedRecord.Text = "Open selected"
        btnOpenReturnedRecord.AutoSize = True
        btnOpenReturnedRecord.Padding = New Padding(8, 2, 8, 2)
        btnOpenReturnedRecord.Visible = False
        AddHandler btnOpenReturnedRecord.Click, AddressOf OpenReturnedRecord
        resultActions.Controls.Add(btnOpenReturnedRecord)

        btnNextStep.Text = "Next step"
        btnNextStep.AutoSize = True
        btnNextStep.Padding = New Padding(8, 2, 8, 2)
        btnNextStep.Visible = False
        AddHandler btnNextStep.Click, AddressOf OpenNextStep
        resultActions.Controls.Add(btnNextStep)

        cboDocuments.DropDownStyle = ComboBoxStyle.DropDownList
        cboDocuments.Width = 190
        cboDocuments.Visible = False
        resultActions.Controls.Add(cboDocuments)

        btnOpenDocument.Text = "Open / download document"
        btnOpenDocument.AutoSize = True
        btnOpenDocument.Padding = New Padding(8, 2, 8, 2)
        btnOpenDocument.Visible = False
        AddHandler btnOpenDocument.Click, AddressOf OpenDocument
        resultActions.Controls.Add(btnOpenDocument)
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

    Private Function Endpoint() As String
        Dim prefix = If(IsSandbox(), "https://sandbox.sellingpartnerapi-", "https://sellingpartnerapi-")
        Return prefix & SelectedMarketplace().Region & ".amazon.com"
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

    Private Async Function TestConnectionAsync() As Task
        If Not CredentialsReady() Then
            MessageBox.Show("Client ID, client secret, and refresh token are required.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        ToggleBusy(True, "Testing connection...")
        Try
            Dim token = Await GetAccessTokenAsync(True)
            Dim probe = Await CallSpApiAsync("/sellers/v1/marketplaceParticipations", HttpMethod.Get, Nothing, token.Item1)
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
        sb.AppendLine().AppendLine("Complete returned data:").AppendLine(PrettyJson(result.Data))
        Return sb.ToString()
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
