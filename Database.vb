Option Explicit On
Option Strict On
Option Infer On

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports System.IO
Imports System.Linq

' All database access stays here. No changes to Joel's existing Access schema.
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
        If product.Values.Count < 4 Then
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

    #Region "02 - Access connection and table validation"
    Friend Shared Function TestConnection(filePath As String) As String
        Using connection = OpenAccess(filePath)
            Dim schema = ReadColumns(connection)
            Using command As New OleDbCommand("SELECT TOP 1 [Asin] FROM [InventoryDimensions]", connection)
                Using reader = command.ExecuteReader()
                    If reader IsNot Nothing Then reader.Read()
                End Using
            End Using
            Return "Connected. InventoryDimensions has all eight expected columns. Read access verified."
        End Using
    End Function

    Private Shared Function OpenAccess(filePath As String) As OleDbConnection
        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
            Throw New InvalidOperationException("Select an existing Microsoft Access database file first.")
        End If
        Dim extension = Path.GetExtension(filePath).ToLowerInvariant()
        If extension <> ".accdb" AndAlso extension <> ".mdb" Then
            Throw New InvalidOperationException("The selected file must be an Access .accdb or .mdb database.")
        End If
        Dim firstError As Exception = Nothing
        For Each provider In {"Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0"}
            Dim builder As New OleDbConnectionStringBuilder()
            builder.Provider = provider
            builder.DataSource = Path.GetFullPath(filePath)
            Dim connection As New OleDbConnection(builder.ConnectionString)
            Try
                connection.Open()
                Return connection
            Catch ex As Exception When TypeOf ex Is OleDbException OrElse TypeOf ex Is InvalidOperationException
                connection.Dispose()
                If firstError Is Nothing Then firstError = ex
            End Try
        Next
        Throw New InvalidOperationException("Could not open this Access database. Check that Microsoft Access Database Engine (ACE) is installed in the same bitness as the application, that the file is accessible, and that any linked tables can connect. " &
                                            If(firstError Is Nothing, "", "Details: " & firstError.Message), firstError)
    End Function

    Private Shared Function ReadColumns(connection As OleDbConnection) As Dictionary(Of String, Type)
        Dim columns As New Dictionary(Of String, Type)(StringComparer.OrdinalIgnoreCase)
        Dim fieldSql = String.Join(", ", Fields.Select(Function(field) "[" & field & "]"))
        Try
            Using command As New OleDbCommand("SELECT TOP 0 " & fieldSql & " FROM [InventoryDimensions]", connection)
                Using reader = command.ExecuteReader(CommandBehavior.SchemaOnly)
                    Dim schema = reader.GetSchemaTable()
                    If schema Is Nothing Then Throw New InvalidOperationException("Could not read the Access table schema.")
                    For Each row As DataRow In schema.Rows
                        columns(Convert.ToString(row("ColumnName"), CultureInfo.InvariantCulture)) = DirectCast(row("DataType"), Type)
                    Next
                End Using
            End Using
        Catch ex As OleDbException
            Throw New InvalidOperationException("InventoryDimensions must be readable and contain Asin, Weight, Height, Length, Width, LastUpdated, ProductCatagory, and ImageURL. The table may be missing, have different column names, or be a disconnected linked table. " & ex.Message, ex)
        End Try
        For Each field In Fields
            If Not columns.ContainsKey(field) Then Throw New InvalidOperationException("The Access table is missing the column: " & field)
        Next
        Return columns
    End Function

    Private Shared Function ReadExisting(connection As OleDbConnection, asin As String, Optional transaction As OleDbTransaction = Nothing) As Dictionary(Of String, Object)
        Dim fieldsSql = String.Join(", ", Fields.Select(Function(field) "[" & field & "]"))
        Using command As New OleDbCommand("SELECT " & fieldsSql & " FROM [InventoryDimensions] WHERE [Asin] = ?", connection)
            If transaction IsNot Nothing Then command.Transaction = transaction
            command.Parameters.AddWithValue("@asin", asin)
            Using reader = command.ExecuteReader()
                If Not reader.Read() Then Return Nothing
                Dim previous As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                For i As Integer = 0 To reader.FieldCount - 1
                    previous(reader.GetName(i)) = If(reader.IsDBNull(i), Nothing, reader.GetValue(i))
                Next
                If reader.Read() Then Throw New InvalidOperationException("Multiple InventoryDimensions records share ASIN " & asin & ". Resolve duplicates before saving.")
                Return previous
            End Using
        End Using
    End Function
    #End Region

    #Region "03 - Read-only preview and explicit insert/update"
    Friend Shared Function Preview(filePath As String, product As ProductRecord) As String
        ValidateProduct(product)
        Using connection = OpenAccess(filePath)
            ReadColumns(connection)
            Dim current = ReadExisting(connection, product.Asin)
            Dim message As New System.Text.StringBuilder()
            message.AppendLine("Access file: " & Path.GetFileName(filePath))
            message.AppendLine("Table: InventoryDimensions")
            message.AppendLine("ASIN: " & product.Asin)
            message.AppendLine("Action: " & If(current Is Nothing, "INSERT", "UPDATE"))
            message.AppendLine()
            For Each field In ProductFields
                Dim newValue As String = Nothing
                If Not product.Values.TryGetValue(field, newValue) Then
                    message.AppendLine(field & ": not returned by Amazon (preserve existing value)")
                Else
                    Dim oldValue = If(current Is Nothing, "(new)", TextValue(current(field)))
                    Dim unit As String = Nothing
                    product.Units.TryGetValue(field, unit)
                    message.AppendLine(field & ": " & oldValue & "  ->  " & newValue & If(String.IsNullOrWhiteSpace(unit), "", " " & unit))
                End If
            Next
            message.AppendLine("LastUpdated: set only if the save succeeds")
            message.AppendLine()
            message.AppendLine("Review units carefully: the existing table has no unit columns. Values are saved as Amazon returns them, without conversion.")
            For Each note In product.Notes
                message.AppendLine(note)
            Next
            Return message.ToString()
        End Using
    End Function

    Friend Shared Function Save(filePath As String, product As ProductRecord) As String
        ValidateProduct(product)
        Using connection = OpenAccess(filePath)
            Dim columns = ReadColumns(connection)
            Using transaction = connection.BeginTransaction()
                Try
                    Dim current = ReadExisting(connection, product.Asin, transaction)
                    Dim isInsert = current Is Nothing
                    Dim fieldNames As New List(Of String)()
                    Dim values As New List(Of Object)()
                    If isInsert Then
                        fieldNames.Add("Asin")
                        values.Add(TypedValue("Asin", product.Asin, columns))
                    End If
                    For Each field In ProductFields
                        Dim value As String = Nothing
                        If product.Values.TryGetValue(field, value) Then
                            fieldNames.Add(field)
                            values.Add(TypedValue(field, value, columns))
                        End If
                    Next
                    fieldNames.Add("LastUpdated")
                    values.Add(TypedValue("LastUpdated", DateTime.Now, columns))

                    Dim query As String
                    If isInsert Then
                        query = "INSERT INTO [InventoryDimensions] (" &
                                String.Join(", ", fieldNames.Select(Function(field) "[" & field & "]")) &
                                ") VALUES (" & String.Join(", ", fieldNames.Select(Function(field) "?")) & ")"
                    Else
                        query = "UPDATE [InventoryDimensions] SET " &
                                String.Join(", ", fieldNames.Select(Function(field) "[" & field & "] = ?")) &
                                " WHERE [Asin] = ?"
                        values.Add(TypedValue("Asin", product.Asin, columns))
                    End If
                    Using command As New OleDbCommand(query, connection, transaction)
                        For Each value In values
                            command.Parameters.AddWithValue("@p" & command.Parameters.Count.ToString(CultureInfo.InvariantCulture), value)
                        Next
                        Dim affected = command.ExecuteNonQuery()
                        If affected <> 1 Then Throw New InvalidOperationException("Expected one database record to be changed; Access reported " & affected.ToString(CultureInfo.InvariantCulture) & ".")
                    End Using
                    transaction.Commit()
                    Return If(isInsert, "Inserted", "Updated") & " " & product.Asin & " in InventoryDimensions."
                Catch
                    transaction.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Function

    Private Shared Sub ValidateProduct(product As ProductRecord)
        If product Is Nothing OrElse Not System.Text.RegularExpressions.Regex.IsMatch(product.Asin, "^[A-Z0-9]{10}$") Then
            Throw New InvalidOperationException("A valid ten-character ASIN is required before saving.")
        End If
        If product.Values.Count = 0 Then Throw New InvalidOperationException("Amazon returned no supported inventory data to save.")
    End Sub

    Private Shared Function TypedValue(field As String, value As Object, columns As Dictionary(Of String, Type)) As Object
        Dim typeOfColumn = columns(field)
        If typeOfColumn Is GetType(String) Then Return Convert.ToString(value, CultureInfo.InvariantCulture)
        If typeOfColumn Is GetType(DateTime) Then
            If TypeOf value Is DateTime Then Return value
            Throw New InvalidOperationException(field & " is expected to be a Date/Time column.")
        End If
        If typeOfColumn Is GetType(Decimal) Then Return Decimal.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If typeOfColumn Is GetType(Double) Then Return Double.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If typeOfColumn Is GetType(Single) Then Return Single.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If typeOfColumn Is GetType(Integer) Then Return Integer.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If typeOfColumn Is GetType(Long) Then Return Long.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        If typeOfColumn Is GetType(Short) Then Return Short.Parse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
        Throw New InvalidOperationException(field & " has an unsupported Access data type (" & typeOfColumn.Name & "). No changes were saved.")
    End Function
    #End Region

End Class
