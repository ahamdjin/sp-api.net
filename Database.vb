Option Explicit On
Option Strict On
Option Infer On

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Linq

' All SQL Server access stays here. Existing production schema remains unchanged.
Friend NotInheritable Class InventoryDatabase

    Friend NotInheritable Class ProductRecord
        Public Property Asin As String = ""
        Public ReadOnly Values As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Public ReadOnly Units As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Public ReadOnly Notes As New List(Of String)()
    End Class

    Private Shared ReadOnly Fields As String() = {"Asin", "Weight", "Height", "Length", "Width", "LastUpdated", "ProductCatagory", "ImageURL"}
    Private Shared ReadOnly ProductFields As String() = {"Weight", "Height", "Length", "Width", "ProductCatagory", "ImageURL"}

    Private Sub New()
    End Sub

    #Region "01 - Amazon catalogue mapping (no database required)"
    Friend Shared Function MapCatalogItem(item As Dictionary(Of String, Object)) As ProductRecord
        Dim product As New ProductRecord()
        product.Asin = TextValue(ValueOf(item, "asin")).Trim().ToUpperInvariant()

        Dim groups = Items(ValueOf(item, "dimensions"))
        If groups.Count > 0 Then
            Dim measurements = ObjectOf(ValueOf(ObjectOf(groups(0)), "item"))
            For Each name In {"Weight", "Height", "Length", "Width"}
                AddMeasurement(product, name, ObjectOf(ValueOf(measurements, name.ToLowerInvariant())))
            Next
        End If

        ' Item dimensions are deliberately NOT replaced with package dimensions.
        If New String() {"Weight", "Height", "Length", "Width"}.Any(Function(field) Not product.Values.ContainsKey(field)) Then
            Dim attributes = ObjectOf(ValueOf(item, "attributes"))
            Dim dimensions = Items(ValueOf(attributes, "item_dimensions"))
            If dimensions.Count > 0 Then
                Dim itemDimensions = ObjectOf(dimensions(0))
                For Each name In {"Weight", "Height", "Length", "Width"}
                    If Not product.Values.ContainsKey(name) Then
                        AddMeasurement(product, name, ObjectOf(ValueOf(itemDimensions, name.ToLowerInvariant())))
                    End If
                Next
            End If
        End If

        ' Some product types supply item_weight separately from item_dimensions.
        If Not product.Values.ContainsKey("Weight") Then
            Dim attributes = ObjectOf(ValueOf(item, "attributes"))
            For Each weightObject In Items(ValueOf(attributes, "item_weight"))
                AddMeasurement(product, "Weight", ObjectOf(weightObject))
                If product.Values.ContainsKey("Weight") Then Exit For
            Next
        End If

        For Each groupObject In Items(ValueOf(item, "images"))
            Dim imageGroup = ObjectOf(groupObject)
            Dim images = Items(ValueOf(imageGroup, "images")).Select(Function(x) ObjectOf(x)).ToList()
            For Each imageData In images.OrderBy(Function(x) If(String.Equals(TextValue(ValueOf(x, "variant")), "MAIN", StringComparison.OrdinalIgnoreCase), 0, 1))
                Dim link = TextValue(ValueOf(imageData, "link"))
                Dim url As Uri = Nothing
                If Uri.TryCreate(link, UriKind.Absolute, url) AndAlso url.Scheme = Uri.UriSchemeHttps Then
                    product.Values("ImageURL") = link
                    Exit For
                End If
            Next
            If product.Values.ContainsKey("ImageURL") Then Exit For
        Next

        For Each groupObject In Items(ValueOf(item, "salesRanks"))
            Dim group = ObjectOf(groupObject)
            For Each rankObject In Items(ValueOf(group, "classificationRanks"))
                Dim category = TextValue(ValueOf(ObjectOf(rankObject), "classificationId"))
                If category <> "" Then
                    product.Values("ProductCatagory") = category
                    Exit For
                End If
            Next
            If product.Values.ContainsKey("ProductCatagory") Then Exit For
        Next
        If Not product.Values.ContainsKey("ProductCatagory") Then
            For Each groupObject In Items(ValueOf(item, "classifications"))
                For Each categoryObject In Items(ValueOf(ObjectOf(groupObject), "classifications"))
                    Dim category = TextValue(ValueOf(ObjectOf(categoryObject), "classificationId"))
                    If category <> "" Then
                        product.Values("ProductCatagory") = category
                        Exit For
                    End If
                Next
                If product.Values.ContainsKey("ProductCatagory") Then Exit For
            Next
        End If
        If product.Values.ContainsKey("ProductCatagory") Then
            product.Notes.Add("ProductCatagory uses the classification ID in the newer Catalog Items API. Confirm this matches the old ProductCategoryId.")
        End If
        For Each name In ProductFields
            If Not product.Values.ContainsKey(name) Then product.Notes.Add(name & " was not returned. An existing value will be preserved.")
        Next
        Return product
    End Function

    Private Shared Sub AddMeasurement(product As ProductRecord, name As String, measurement As Dictionary(Of String, Object))
        Dim raw = TextValue(ValueOf(measurement, "value"))
        Dim number As Decimal
        If raw = "" OrElse Not Decimal.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, number) Then Return
        product.Values(name) = number.ToString(CultureInfo.InvariantCulture)
        Dim unit = TextValue(ValueOf(measurement, "unit"))
        If unit <> "" Then product.Units(name) = unit
    End Sub

    Private Shared Function ValueOf(source As Dictionary(Of String, Object), key As String) As Object
        If source Is Nothing Then Return Nothing
        Dim value As Object = Nothing
        If source.TryGetValue(key, value) Then Return value
        For Each pair In source
            If String.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase) Then Return pair.Value
        Next
        Return Nothing
    End Function

    Private Shared Function ObjectOf(value As Object) As Dictionary(Of String, Object)
        Dim dict = TryCast(value, Dictionary(Of String, Object))
        Return If(dict, New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase))
    End Function

    Private Shared Function Items(value As Object) As List(Of Object)
        Dim array = TryCast(value, Object())
        If array IsNot Nothing Then Return array.ToList()
        Dim list = TryCast(value, ArrayList)
        If list IsNot Nothing Then Return list.Cast(Of Object)().ToList()
        Return New List(Of Object)()
    End Function

    Private Shared Function TextValue(value As Object) As String
        If value Is Nothing Then Return ""
        If TypeOf value Is String OrElse TypeOf value Is ValueType Then Return Convert.ToString(value, CultureInfo.InvariantCulture)
        Return ""
    End Function
    #End Region

    #Region "02 - SQL Server connection and table validation"
    Friend Shared Function TestConnection(connectionString As String) As String
        Using connection = OpenConnection(connectionString)
            ReadColumns(connection)
            Using command As New SqlCommand("SELECT TOP 1 [Asin] FROM [InventoryDimensions]", connection)
                Using reader = command.ExecuteReader()
                    If reader IsNot Nothing Then reader.Read()
                End Using
            End Using
            Return "Connected to SQL Server. InventoryDimensions has all eight expected columns and can be read. No records were changed."
        End Using
    End Function

    Private Shared Function OpenConnection(connectionString As String) As SqlConnection
        If String.IsNullOrWhiteSpace(connectionString) Then
            Throw New InvalidOperationException("Enter the existing SQL Server connection setting (FTSSQL) in the application.")
        End If
        Dim builder As SqlConnectionStringBuilder
        Try
            builder = New SqlConnectionStringBuilder(connectionString)
        Catch ex As ArgumentException
            Throw New InvalidOperationException("The SQL Server connection setting is not valid.", ex)
        End Try
        If String.IsNullOrWhiteSpace(builder.DataSource) OrElse String.IsNullOrWhiteSpace(builder.InitialCatalog) Then
            Throw New InvalidOperationException("The SQL Server connection must include a server name (Data Source) and database name (Initial Catalog).")
        End If
        Dim connection As New SqlConnection(builder.ConnectionString)
        Try
            connection.Open()
            Return connection
        Catch
            connection.Dispose()
            Throw
        End Try
    End Function

    Private Shared Function ReadColumns(connection As SqlConnection) As Dictionary(Of String, Type)
        Dim columns As New Dictionary(Of String, Type)(StringComparer.OrdinalIgnoreCase)
        Dim selected = String.Join(", ", Fields.Select(Function(field) "[" & field & "]"))
        Try
            Using command As New SqlCommand("SELECT TOP 0 " & selected & " FROM [InventoryDimensions]", connection)
                Using reader = command.ExecuteReader(CommandBehavior.SchemaOnly)
                    Dim schema = reader.GetSchemaTable()
                    If schema Is Nothing Then Throw New InvalidOperationException("Cannot inspect InventoryDimensions columns.")
                    For Each row As DataRow In schema.Rows
                        columns(Convert.ToString(row("ColumnName"), CultureInfo.InvariantCulture)) = DirectCast(row("DataType"), Type)
                    Next
                End Using
            End Using
        Catch ex As SqlException
            Throw New InvalidOperationException("SQL Server could not read InventoryDimensions and its expected columns. Check the selected database and table permissions. " & ex.Message, ex)
        End Try
        For Each field In Fields
            If Not columns.ContainsKey(field) Then Throw New InvalidOperationException("InventoryDimensions is missing the field " & field)
        Next
        Return columns
    End Function

    Private Shared Function ReadExisting(connection As SqlConnection, asin As String, Optional transaction As SqlTransaction = Nothing) As Dictionary(Of String, Object)
        Dim selected = String.Join(", ", Fields.Select(Function(field) "[" & field & "]"))
        Dim table = If(transaction Is Nothing, "[InventoryDimensions]", "[InventoryDimensions] WITH (UPDLOCK, HOLDLOCK)")
        Using command As New SqlCommand("SELECT " & selected & " FROM " & table & " WHERE [Asin] = @asin", connection)
            If transaction IsNot Nothing Then command.Transaction = transaction
            command.Parameters.Add("@asin", SqlDbType.NVarChar, 10).Value = asin
            Using reader = command.ExecuteReader()
                If Not reader.Read() Then Return Nothing
                Dim existing As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For index As Integer = 0 To reader.FieldCount - 1
                    existing(reader.GetName(index)) = If(reader.IsDBNull(index), Nothing, reader.GetValue(index))
                Next
                If reader.Read() Then Throw New InvalidOperationException("Multiple InventoryDimensions records have ASIN " & asin & ". Please correct the duplicates before saving.")
                Return existing
            End Using
        End Using
    End Function
    #End Region

    #Region "03 - Preview and confirmed SQL Server insert/update"
    Friend Shared Function Preview(connectionString As String, product As ProductRecord) As String
        ValidateProduct(product)
        Using connection = OpenConnection(connectionString)
            ReadColumns(connection)
            Dim current = ReadExisting(connection, product.Asin)
            Dim output As New System.Text.StringBuilder()
            output.AppendLine("Database: SQL Server")
            output.AppendLine("Table: InventoryDimensions")
            output.AppendLine("ASIN: " & product.Asin)
            output.AppendLine("Action: " & If(current Is Nothing, "INSERT", "UPDATE"))
            output.AppendLine()
            For Each field In ProductFields
                Dim newValue As String = Nothing
                If Not product.Values.TryGetValue(field, newValue) Then
                    output.AppendLine(field & ": not returned by Amazon (preserve existing data)")
                Else
                    Dim oldValue = If(current Is Nothing, "(new)", TextValue(current(field)))
                    Dim unit As String = Nothing
                    product.Units.TryGetValue(field, unit)
                    output.AppendLine(field & ": " & oldValue & "  ->  " & newValue & If(String.IsNullOrWhiteSpace(unit), "", " " & unit))
                End If
            Next
            output.AppendLine("LastUpdated: set on successful save")
            output.AppendLine()
            output.AppendLine("IMPORTANT: Amazon measurements have not been converted between units. Confirm what the existing database expects.")
            For Each note In product.Notes
                output.AppendLine(note)
            Next
            Return output.ToString()
        End Using
    End Function

    Friend Shared Function Save(connectionString As String, product As ProductRecord) As String
        ValidateProduct(product)
        Using connection = OpenConnection(connectionString)
            Dim columns = ReadColumns(connection)
            Using transaction = connection.BeginTransaction(IsolationLevel.ReadCommitted)
                Try
                    Dim current = ReadExisting(connection, product.Asin, transaction)
                    Dim inserting = current Is Nothing
                    Dim fieldsToSave As New List(Of String)()
                    Dim values As New List(Of Object)()
                    If inserting Then
                        fieldsToSave.Add("Asin")
                        values.Add(TypedValue("Asin", product.Asin, columns))
                    End If
                    For Each field In ProductFields
                        Dim newValue As String = Nothing
                        If product.Values.TryGetValue(field, newValue) Then
                            fieldsToSave.Add(field)
                            values.Add(TypedValue(field, newValue, columns))
                        End If
                    Next
                    fieldsToSave.Add("LastUpdated")
                    values.Add(TypedValue("LastUpdated", DateTime.Now, columns))

                    Dim sql As String
                    If inserting Then
                        sql = "INSERT INTO [InventoryDimensions] (" &
                              String.Join(", ", fieldsToSave.Select(Function(field) "[" & field & "]")) &
                              ") VALUES (" & String.Join(", ", fieldsToSave.Select(Function(field, index) "@p" & index.ToString(CultureInfo.InvariantCulture))) & ")"
                    Else
                        sql = "UPDATE [InventoryDimensions] SET " &
                              String.Join(", ", fieldsToSave.Select(Function(field, index) "[" & field & "] = @p" & index.ToString(CultureInfo.InvariantCulture))) &
                              " WHERE [Asin] = @p" & fieldsToSave.Count.ToString(CultureInfo.InvariantCulture)
                        values.Add(TypedValue("Asin", product.Asin, columns))
                    End If
                    Using command As New SqlCommand(sql, connection, transaction)
                        For index As Integer = 0 To values.Count - 1
                            command.Parameters.AddWithValue("@p" & index.ToString(CultureInfo.InvariantCulture), values(index))
                        Next
                        Dim affected = command.ExecuteNonQuery()
                        If affected <> 1 Then Throw New InvalidOperationException("The expected update count was 1, but SQL Server reported " & affected.ToString(CultureInfo.InvariantCulture) & ".")
                    End Using

                    Dim saved = ReadExisting(connection, product.Asin, transaction)
                    If saved Is Nothing Then Throw New InvalidOperationException("Could not verify the saved InventoryDimensions record.")
                    For index As Integer = 0 To fieldsToSave.Count - 1
                        Dim field = fieldsToSave(index)
                        If Not ExpectedMatches(saved(field), values(index), columns(field)) Then
                            Throw New InvalidOperationException("The stored " & field & " did not match the intended value. The change was rolled back.")
                        End If
                    Next
                    transaction.Commit()

                    Dim result As New System.Text.StringBuilder()
                    result.AppendLine(If(inserting, "Inserted ", "Updated ") & product.Asin & " in InventoryDimensions.")
                    result.AppendLine("Saved values verified:")
                    For Each field In fieldsToSave
                        result.AppendLine(field & ": " & TextValue(saved(field)))
                    Next
                    Return result.ToString()
                Catch
                    Try
                        transaction.Rollback()
                    Catch
                        ' A failed connection may already have rolled back the transaction.
                    End Try
                    Throw
                End Try
            End Using
        End Using
    End Function

    Private Shared Function ExpectedMatches(actual As Object, expected As Object, fieldType As Type) As Boolean
        If actual Is Nothing Then Return expected Is Nothing
        If fieldType Is GetType(DateTime) Then
            Return Math.Abs((CDate(actual) - CDate(expected)).TotalSeconds) < 1
        End If
        If fieldType Is GetType(String) Then
            Return String.Equals(Convert.ToString(actual, CultureInfo.InvariantCulture), Convert.ToString(expected, CultureInfo.InvariantCulture), StringComparison.Ordinal)
        End If
        Return Convert.ToDecimal(actual, CultureInfo.InvariantCulture) = Convert.ToDecimal(expected, CultureInfo.InvariantCulture)
    End Function

    Private Shared Sub ValidateProduct(product As ProductRecord)
        If product Is Nothing OrElse Not System.Text.RegularExpressions.Regex.IsMatch(product.Asin, "^[A-Z0-9]{10}$") Then
            Throw New InvalidOperationException("A valid ten-character ASIN is required.")
        End If
        If product.Values.Count = 0 Then Throw New InvalidOperationException("Amazon returned no supported dimensions, category or image values to save.")
    End Sub

    Private Shared Function TypedValue(field As String, value As Object, columns As Dictionary(Of String, Type)) As Object
        Dim fieldType = columns(field)
        If fieldType Is GetType(String) Then Return Convert.ToString(value, CultureInfo.InvariantCulture)
        If fieldType Is GetType(DateTime) Then
            If TypeOf value Is DateTime Then Return value
            Throw New InvalidOperationException(field & " needs a Date/Time value.")
        End If
        If fieldType Is GetType(Decimal) Then Return Decimal.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If fieldType Is GetType(Double) Then Return Double.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If fieldType Is GetType(Single) Then Return Single.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If fieldType Is GetType(Integer) Then Return Integer.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If fieldType Is GetType(Long) Then Return Long.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If fieldType Is GetType(Short) Then Return Short.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        Throw New InvalidOperationException(field & " is stored as unsupported SQL Server data type " & fieldType.Name & ". No values were changed.")
    End Function
    #End Region

    #Region "04 - Disposable SQL Server integration test (CI only)"
    Friend Shared Sub RunCiSqlIntegrationTests()
        If Environment.GetEnvironmentVariable("SP_API_CI_SQL_SMOKE") <> "1" Then
            Throw New InvalidOperationException("The isolated SQL integration test is only available in CI test mode.")
        End If

        Dim name = "SPAPI_Test_" & Guid.NewGuid().ToString("N")
        Dim server = "(localdb)\MSSQLLocalDB"
        Dim masterBuilder As New SqlConnectionStringBuilder With {
            .DataSource = server, .InitialCatalog = "master", .IntegratedSecurity = True, .ConnectTimeout = 60
        }
        Dim testBuilder As New SqlConnectionStringBuilder(masterBuilder.ConnectionString) With {.InitialCatalog = name}
        Dim created As Boolean = False

        Try
            Using master As New SqlConnection(masterBuilder.ConnectionString)
                master.Open()
                Using command As New SqlCommand("CREATE DATABASE [" & name & "]", master)
                    command.CommandTimeout = 90
                    command.ExecuteNonQuery()
                End Using
                created = True
            End Using

            Using connection As New SqlConnection(testBuilder.ConnectionString)
                connection.Open()
                Using command As New SqlCommand(
                    "CREATE TABLE [dbo].[InventoryDimensions] (" &
                    "[Asin] nvarchar(10) NOT NULL PRIMARY KEY, [Weight] nvarchar(50) NULL, " &
                    "[Height] nvarchar(50) NULL, [Length] nvarchar(50) NULL, [Width] nvarchar(50) NULL, " &
                    "[LastUpdated] datetime2(3) NOT NULL, [ProductCatagory] nvarchar(100) NULL, " &
                    "[ImageURL] nvarchar(1000) NULL)", connection)
                    command.ExecuteNonQuery()
                End Using
            End Using

            Dim connectionString = testBuilder.ConnectionString
            If Not TestConnection(connectionString).Contains("Connected") Then
                Throw New InvalidOperationException("SQL Server read-only connection test failed.")
            End If

            Dim catalog As New Dictionary(Of String, Object) From {
                {"asin", "B07N4M94X4"},
                {"dimensions", New Object() {
                    New Dictionary(Of String, Object) From {
                        {"item", New Dictionary(Of String, Object) From {
                            {"weight", New Dictionary(Of String, Object) From {{"value", 1.25D}, {"unit", "pounds"}}},
                            {"height", New Dictionary(Of String, Object) From {{"value", 3.0D}, {"unit", "inches"}}},
                            {"length", New Dictionary(Of String, Object) From {{"value", 5.0D}, {"unit", "inches"}}},
                            {"width", New Dictionary(Of String, Object) From {{"value", 4.0D}, {"unit", "inches"}}}
                        }}
                    }
                }},
                {"images", New Object() {
                    New Dictionary(Of String, Object) From {
                        {"images", New Object() {
                            New Dictionary(Of String, Object) From {{"variant", "MAIN"}, {"link", "https://example.com/photo.jpg"}}
                        }}
                    }
                }},
                {"salesRanks", New Object() {
                    New Dictionary(Of String, Object) From {
                        {"classificationRanks", New Object() {
                            New Dictionary(Of String, Object) From {{"classificationId", "12345"}}
                        }}
                    }
                }}
            }
            Dim product = MapCatalogItem(catalog)
            Dim preview = Preview(connectionString, product)
            If Not preview.Contains("Action: INSERT") Then Throw New InvalidOperationException("New ASIN did not preview an INSERT.")
            If Not Save(connectionString, product).Contains("Inserted") Then Throw New InvalidOperationException("INSERT did not complete.")

            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim saved = ReadExisting(connection, product.Asin)
                If saved Is Nothing OrElse TextValue(saved("Weight")) <> "1.25" OrElse TextValue(saved("Width")) <> "4" OrElse
                   TextValue(saved("ProductCatagory")) <> "12345" Then
                    Throw New InvalidOperationException("INSERT did not store the expected dimensions and category.")
                End If
            End Using

            Dim partial As New ProductRecord With {.Asin = "B07N4M94X4"}
            partial.Values("Weight") = "2.5"
            partial.Values("ImageURL") = "https://example.com/O'Brien-image.jpg"
            If Not Preview(connectionString, partial).Contains("Action: UPDATE") Then
                Throw New InvalidOperationException("Existing ASIN did not preview an UPDATE.")
            End If
            If Not Save(connectionString, partial).Contains("Updated") Then Throw New InvalidOperationException("UPDATE failed.")

            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim saved = ReadExisting(connection, partial.Asin)
                If TextValue(saved("Weight")) <> "2.5" OrElse TextValue(saved("Height")) <> "3" OrElse
                   TextValue(saved("Width")) <> "4" OrElse TextValue(saved("ProductCatagory")) <> "12345" OrElse
                   Not TextValue(saved("ImageURL")).Contains("O'Brien") Then
                    Throw New InvalidOperationException("UPDATE overwrote missing fields or mishandled SQL quotes.")
                End If
            End Using

            Dim invalid As New ProductRecord With {.Asin = "B07N4M94X4"}
            invalid.Values("Weight") = "3.25"
            invalid.Values("ImageURL") = New String("X"c, 1050)
            Dim refused As Boolean = False
            Try
                Save(connectionString, invalid)
            Catch ex As SqlException
                refused = True
            End Try
            If Not refused Then Throw New InvalidOperationException("Expected to reject overlong ImageURL.")
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Dim saved = ReadExisting(connection, invalid.Asin)
                If TextValue(saved("Weight")) <> "2.5" Then
                    Throw New InvalidOperationException("A failed update modified the record instead of rolling back.")
                End If
            End Using

            Dim badAsin As New ProductRecord With {.Asin = "B07N4M94X4' OR 1=1"}
            badAsin.Values("Weight") = "99"
            refused = False
            Try
                Save(connectionString, badAsin)
            Catch ex As InvalidOperationException
                refused = True
            End Try
            If Not refused Then Throw New InvalidOperationException("ASIN validation failed to reject unsafe identifiers.")
        Finally
            If created Then
                Using master As New SqlConnection(masterBuilder.ConnectionString)
                    master.Open()
                    Using command As New SqlCommand(
                        "ALTER DATABASE [" & name & "] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; DROP DATABASE [" & name & "]", master)
                        command.CommandTimeout = 90
                        command.ExecuteNonQuery()
                    End Using
                End Using
            End If
        End Try
    End Sub
    #End Region

End Class
