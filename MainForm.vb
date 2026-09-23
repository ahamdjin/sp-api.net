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
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
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

Public Class MainForm
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

    Private Class OperationInfo
        Public Property Id As String = ""
        Public Property Label As String = ""
        Public Property Kind As String = "read"
        Public Property Group As String = ""
        Public Overrides Function ToString() As String
            Return Label
        End Function
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
    Private ReadOnly tabs As New TabControl()

    Private CurrentOperation As String = "catalog"
    Private LastResult As ApiResult
    Private LastDocumentUrl As String = ""
    Private NextOperationId As String = ""
    Private ConnectionVerified As Boolean

    Private Const CoreOrderData As String = "PROCEEDS,EXPENSE,PROMOTION,CANCELLATION,FULFILLMENT,PACKAGES,TAX,PAYMENT,FULFILLMENT_ORDERS"
    Private Const DefaultCatalogData As String = "attributes,classifications,dimensions,identifiers,images,productTypes,relationships,salesRanks,summaries,vendorDetails"
    Private Const PreviewLimit As Integer = 2 * 1024 * 1024

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
    End Sub

    Public Sub RunCiSelfTest()
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

        lblMeta.Dock = DockStyle.Top
        lblMeta.Height = 28
        lblMeta.Padding = New Padding(6)
        resultTab.Controls.Add(lblMeta)

        Dim resultActions As New FlowLayoutPanel With {
            .Dock = DockStyle.Bottom,
            .Height = 40,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = False,
            .Padding = New Padding(3, 3, 3, 3)
        }

        btnNextStep.Text = "Next step"
        btnNextStep.AutoSize = True
        btnNextStep.Padding = New Padding(8, 2, 8, 2)
        btnNextStep.Visible = False
        AddHandler btnNextStep.Click, AddressOf OpenNextStep
        resultActions.Controls.Add(btnNextStep)

        btnOpenDocument.Text = "Open / download document"
        btnOpenDocument.AutoSize = True
        btnOpenDocument.Padding = New Padding(8, 2, 8, 2)
        btnOpenDocument.Visible = False
        AddHandler btnOpenDocument.Click, AddressOf OpenDocument
        resultActions.Controls.Add(btnOpenDocument)
        resultTab.Controls.Add(resultActions)

        txtResult.Dock = DockStyle.Fill
        txtResult.Multiline = True
        txtResult.ScrollBars = ScrollBars.Both
        txtResult.ReadOnly = True
        txtResult.WordWrap = False
        txtResult.Font = New Font("Consolas", 9.0F)
        resultTab.Controls.Add(txtResult)
        txtResult.BringToFront()

        txtRaw.Dock = DockStyle.Fill
        txtRaw.Multiline = True
        txtRaw.ScrollBars = ScrollBars.Both
        txtRaw.ReadOnly = True
        txtRaw.WordWrap = False
        txtRaw.Font = New Font("Consolas", 9.0F)
        rawTab.Controls.Add(txtRaw)
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

    Private Sub SelectOperation(id As String)
        SaveVisibleFieldValues()
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
        btnNextStep.Visible = False
        NextOperationId = ""
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
            FieldValues("confirmed") = False
            Dim confirmation As Control = Nothing
            If FieldControls.TryGetValue("confirmed", confirmation) Then
                Dim check = TryCast(confirmation, CheckBox)
                If check IsNot Nothing AndAlso check.Checked Then check.Checked = False
            End If
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

    Private Sub InvalidateConnectionState()
        ConnectionVerified = False
        FieldValues("confirmed") = False
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
            Dim token = Await GetAccessTokenAsync()
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

        ToggleBusy(True, "Waiting for Amazon...")
        LastDocumentUrl = ""
        btnOpenDocument.Visible = False
        Try
            Dim result = Await ExecuteOperationAsync(CurrentOperation)
            LastResult = result
            ApplyReturnedIds(CurrentOperation, result.Data)
            ShowResult(CurrentOperation, result)
        Catch ex As AppException
            ShowResult(CurrentOperation, LocalFailure(ex))
        Catch ex As Exception
            ShowResult(CurrentOperation, LocalFailure(New AppException(ex.Message, 500, "CLIENT_INTERNAL_ERROR", ex.ToString())))
        Finally
            ToggleBusy(False, "")
        End Try
    End Function

    Private Sub ToggleBusy(busy As Boolean, message As String)
        btnRun.Enabled = Not busy AndAlso Operations.First(Function(x) x.Id = CurrentOperation).Kind <> "legacy"
        btnTest.Enabled = Not busy
        operationTree.Enabled = Not busy
        Cursor = If(busy, Cursors.WaitCursor, Cursors.Default)
        If busy Then lblMeta.Text = message
    End Sub

    Private Function IsWriteOperation(operation As String) As Boolean
        Return operation = "createReport" OrElse operation = "submitFeed" OrElse operation = "createInboundPlan"
    End Function

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
                result.Problem = New ApiProblem With {.Code = "FEED_" & status, .Message = If(status = "FATAL", "Amazon aborted the feed during processing.", "Amazon cancelled the feed before processing completed."), .Details = If(StringValue(GetValue(data, "resultFeedDocumentId")) = "", "", "resultFeedDocumentId: " & StringValue(GetValue(data, "resultFeedDocumentId"))), .Action = "Review the feed processing report before resubmitting.", .Retryable = False}
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
                result.Ok = False : result.Status = 422 : result.StatusText = "Inbound operation failed"
                result.Problem = New ApiProblem With {.Code = "INBOUND_OPERATION_FAILED", .Message = "The asynchronous Fulfillment Inbound operation failed.", .Details = Json(GetValue(data, "operationProblems")), .Action = "Correct every operationProblem before starting another write.", .Retryable = False}
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

    Private Async Function GetAccessTokenAsync() As Task(Of Tuple(Of String, Integer))
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
            If Not response.IsSuccessStatusCode Then Throw New AppException(ExtractMessage(data, "Amazon rejected the supplied LWA credentials"), CInt(response.StatusCode), ExtractCode(data, "LWA_AUTH_FAILED"), Json(data))
            Dim dict = AsDict(data)
            Dim token = StringValue(GetValue(dict, "access_token"))
            If token = "" Then Throw New AppException("Amazon returned no access token", 502, "LWA_TOKEN_MISSING", Json(data))
            Dim expires As Integer = 3600
            Integer.TryParse(Convert.ToString(GetValue(dict, "expires_in"), CultureInfo.InvariantCulture), expires)
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
                .Data = data, .DurationMs = sw.ElapsedMilliseconds, .Attempts = outcome.Item2
            }
            If Not result.Ok Then result.Problem = BuildProblem(result.Status, result.StatusText, data, method, Header(response, "x-amzn-errortype"))
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
        If status = 401 OrElse status = 403 OrElse key.Contains("unauthorized") OrElse key.Contains("accessdenied") Then Return "Verify seller authorization and the Amazon role required by this operation."
        If key.Contains("could not match input arguments") OrElse key.Contains("sandbox request") Then Return "Amazon's static Sandbox accepts predefined request examples. Use the exact Sandbox values shown in the app."
        If status = 429 OrElse key.Contains("throttl") Then Return "Slow the request rate and retry after the throttle window."
        If status = 400 Then Return "Check required fields, IDs, date ranges, enum values, marketplace, and URL encoding."
        If status = 404 Then Return "Verify the resource ID, marketplace, and environment. Sandbox and Production IDs are not interchangeable."
        If status >= 500 AndAlso method <> HttpMethod.Get Then Return "Do not immediately resubmit this write. Verify the related Amazon resource/job first."
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
            If first.ContainsKey("details") Then Return Json(GetValue(first, "details"))
        End If
        If d.ContainsKey("details") Then Return Json(GetValue(d, "details"))
        Return ""
    End Function

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

        Return New ApiResult With {
            .Ok = False, .Status = ex.Status, .StatusText = If(ex.Status >= 500, "Local / network error", "Local validation error"), .ErrorMessage = ex.Message,
            .Problem = New ApiProblem With {.Code = ex.Code, .Message = ex.Message, .Details = ex.Details, .Action = action, .Retryable = retryable}
        }
    End Function

    Private Sub ShowResult(operation As String, result As ApiResult)
        LastResult = result
        lblMeta.Text = result.Status.ToString(CultureInfo.InvariantCulture) & " " & result.StatusText & If(result.DurationMs > 0, "  |  " & result.DurationMs.ToString(CultureInfo.InvariantCulture) & " ms", "") & If(result.RequestId <> "", "  |  Request " & result.RequestId, "") & If(result.RateLimit <> "", "  |  " & result.RateLimit & " req/s", "")
        Dim envelope As New Dictionary(Of String, Object) From {
            {"ok", result.Ok}, {"status", result.Status}, {"statusText", result.StatusText}, {"requestId", If(result.RequestId = "", Nothing, result.RequestId)},
            {"gatewayId", If(result.GatewayId = "", Nothing, result.GatewayId)}, {"traceId", If(result.TraceId = "", Nothing, result.TraceId)},
            {"rateLimit", If(result.RateLimit = "", Nothing, result.RateLimit)}, {"durationMs", result.DurationMs}, {"attempts", result.Attempts}, {"data", result.Data}
        }
        If result.Problem IsNot Nothing Then envelope("problem") = New Dictionary(Of String, Object) From {{"code", result.Problem.Code}, {"message", result.Problem.Message}, {"details", result.Problem.Details}, {"action", result.Problem.Action}, {"retryable", result.Problem.Retryable}}
        txtRaw.Text = PrettyJson(envelope)
        txtResult.Text = BuildSummary(operation, result)
        LastDocumentUrl = FindDocumentUrl(result.Data)
        btnOpenDocument.Visible = LastDocumentUrl <> ""
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
                Dim u = FindDocumentUrl(data)
                If u <> "" Then sb.AppendLine("Document URL returned and ready to open/download.")
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
        Dim data = AsDict(dataObj)
        Dim direct = SafeHttpsUrl(StringValue(GetValue(data, "url")))
        If direct <> "" Then Return direct

        Dim downloads = ListValue(GetValue(data, "documentDownloads"))
        If downloads.Count > 0 Then
            Dim uri = SafeHttpsUrl(StringValue(GetValue(AsDict(downloads(0)), "uri")))
            If uri <> "" Then Return uri
        End If

        Dim payload = AsDict(GetValue(data, "payload"))
        For Each key In {"DownloadURL", "downloadURL", "downloadUrl"}
            Dim url = SafeHttpsUrl(StringValue(GetValue(payload, key)))
            If url <> "" Then Return url
        Next
        Return ""
    End Function

    Private Function SafeHttpsUrl(value As String) As String
        If value = "" Then Return ""
        Dim uri As Uri = Nothing
        If Not Uri.TryCreate(value, UriKind.Absolute, uri) Then Return ""
        If uri.Scheme <> Uri.UriSchemeHttps Then Return ""
        Return uri.AbsoluteUri
    End Function

    Private Sub ConfigureNextStep(operation As String, result As ApiResult)
        btnNextStep.Visible = False
        NextOperationId = ""
        If Not result.Ok Then Return

        Dim data = AsDict(result.Data)
        Select Case operation
            Case "createReport"
                If StringValue(GetValue(data, "reportId")) <> "" Then
                    NextOperationId = "report"
                    btnNextStep.Text = "Next: Check report status"
                End If
            Case "report"
                If StringValue(GetValue(data, "reportDocumentId")) <> "" Then
                    NextOperationId = "reportDocument"
                    btnNextStep.Text = "Next: Open report document"
                End If
            Case "submitFeed"
                If StringValue(GetValue(data, "feedId")) <> "" Then
                    NextOperationId = "feed"
                    btnNextStep.Text = "Next: Check feed status"
                End If
            Case "feed"
                If StringValue(GetValue(data, "resultFeedDocumentId")) <> "" Then
                    NextOperationId = "feedDocument"
                    btnNextStep.Text = "Next: Open processing report"
                End If
            Case "createInboundPlan"
                If StringValue(GetValue(data, "operationId")) <> "" Then
                    NextOperationId = "inboundOperationStatus"
                    btnNextStep.Text = "Next: Check operation status"
                End If
        End Select

        btnNextStep.Visible = NextOperationId <> ""
    End Sub

    Private Sub OpenNextStep(sender As Object, e As EventArgs)
        If NextOperationId = "" Then Return
        NavigateToOperation(NextOperationId)
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

    Private Sub OpenDocument(sender As Object, e As EventArgs)
        If LastDocumentUrl = "" Then Return
        Try
            Process.Start(LastDocumentUrl)
        Catch ex As Exception
            MessageBox.Show("Could not open the document URL: " & ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
