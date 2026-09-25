Option Explicit On
Option Strict On
Option Infer On

' Amazon SP-API Sandbox-specific endpoint and request behavior.
Public Partial Class MainForm

    Private NotInheritable Class SandboxApiEnvironment
        Inherits ApiEnvironmentProfile

        Public Shared ReadOnly Instance As New SandboxApiEnvironment()

        Private Sub New()
        End Sub

        Public Overrides ReadOnly Property Name As String
            Get
                Return "sandbox"
            End Get
        End Property

        Public Overrides ReadOnly Property TestConnectionPath As String
            Get
                Return "/sellers/v1/marketplaceParticipations"
            End Get
        End Property

        Public Overrides ReadOnly Property IncludeCatalogLocale As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property IncludeCatalogPageSize As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property UploadFeedContent As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property RejectRetiredListingFeeds As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property ValidateMarketplaceRegion As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property FeedVerificationMessage As String
            Get
                Return "Static Sandbox validates Amazon's predefined examples and does not persist the upload like Production."
            End Get
        End Property

        Public Overrides Function Endpoint(region As String) As String
            Return "https://sandbox.sellingpartnerapi-" & region & ".amazon.com"
        End Function

        Public Overrides Function DocumentPath(path As String) As String
            Return path
        End Function
    End Class

End Class
