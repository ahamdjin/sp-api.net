Option Explicit On
Option Strict On
Option Infer On

' Amazon SP-API Production-specific endpoint and request behavior.
Public Partial Class MainForm

    Private NotInheritable Class ProductionApiEnvironment
        Inherits ApiEnvironmentProfile

        Public Shared ReadOnly Instance As New ProductionApiEnvironment()

        Private Sub New()
        End Sub

        Public Overrides ReadOnly Property Name As String
            Get
                Return "production"
            End Get
        End Property

        Public Overrides ReadOnly Property TestConnectionPath As String
            Get
                Return "/sellers/v1/marketplaceParticipations"
            End Get
        End Property

        Public Overrides ReadOnly Property IncludeCatalogLocale As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property IncludeCatalogPageSize As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property UploadFeedContent As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property RejectRetiredListingFeeds As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property ValidateMarketplaceRegion As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property FeedVerificationMessage As String
            Get
                Return "Poll Feed status until DONE or FATAL, then inspect resultFeedDocumentId."
            End Get
        End Property

        Public Overrides Function Endpoint(region As String) As String
            Return "https://sellingpartnerapi-" & region & ".amazon.com"
        End Function

        Public Overrides Function DocumentPath(path As String) As String
            Return path & "?enableContentEncodingUrlHeader=true"
        End Function
    End Class

End Class
