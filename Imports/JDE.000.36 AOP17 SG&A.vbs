Sub XLCode()
    Dim row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String, periodTo As String
    Dim connection As Object, country As String, startPeriod As Integer, endPeriod As Integer, rngInput As Range
    Dim vSGACategory As Variant, vSGAField As Variant, colSGA As Collection, i As Integer
    
    vSGACategory = Array("Selling - Field", "Selling - Office", "Marketing", "Research & Development expenses", "Central design expenses" _
        , "Finance expenses", "Strategy expenses", "Human resources expenses", "Communication expenses", "IT expenses" _
        , "Legal expenses", "General Management expenses", "Operations expenses")
    vSGAField = Array("SellingField", "SellingOffice", "Marketing", "RAndD", "CentralDesign", "Finance", "Strategy", "HR", "Communication" _
        , "IT", "Legal", "GeneralManagement", "Operations")
        
    Set colSGA = New Collection
    
    For i = LBound(vSGACategory) To UBound(vSGACategory)
        colSGA.Add vSGAField(i), vSGACategory(i)
    Next i

    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    periodFrom = GetPar([A1], "Period From=")
    periodTo = GetPar([A1], "Period To=")
    startPeriod = CInt(Right(periodFrom, 2))
    endPeriod = CInt(Right(periodTo, 2))
    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")

    Set rngInput = Range("XLRInput")

    With rngInput
        For row = 3 To .Rows.Count
            If Not IsEmpty(.Cells(row, 1)) Then
                For period = startPeriod To endPeriod
                    If .Cells(row, period + 1) <> 0 Then
                        rs.AddNew
                        rs.Fields("Country") = country
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                        rs.Fields("SourceType") = "AOPSGA"
                        rs.Fields("Forecast") = "yes"
                        rs.Fields("sku") = .Cells(row, 5)
                        rs.Fields("Customer") = .[D1]
                        rs.Fields("PromoNonPromo") = "NonPromo"
                        If .Cells(row, 4) = "Selling" Then
                            rs.Fields("Selling") = -1 * IIf(IsError(.Cells(row, period + 5)), 0, .Cells(row, period + 5))
                        ElseIf .Cells(row, 4) = "Other G&A" Then
                            rs.Fields("OtherGA") = -1 * IIf(IsError(.Cells(row, period + 5)), 0, .Cells(row, period + 5))
                        End If
                    End If
                Next period
            End If
        Next row
    End With

    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'AOPSGA' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & _
        Quot(country) & " AND Period BETWEEN " & Quot(periodFrom) & " AND " & Quot(periodTo)
    rs.ActiveConnection = connection
    rs.UpdateBatch
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    connection.Close
End Sub
Function GetEmptyRecordSet(ByVal sTable As String) As Object
    Dim rsData As Object, connection As Object
    
    Set connection = GetDBConnection()
    connection.Open
    Set rsData = CreateObject("ADODB.Recordset")
    With rsData
        .CursorLocation = 3 'adUseClient
        .CursorType = 1 'adOpenKeyset
        .LockType = 4 'adLockBatchOptimistic
        .Open sTable, connection
        .ActiveConnection = Nothing
    End With
    
    connection.Close
    Set GetEmptyRecordSet = rsData
End Function
Function GetDBConnection() As Object
    Dim pw As String, connectionString As String, dbConnection As Object, sDbName As String
    
    pw = "xlsysjs14"
    sDbName = GetSQL("SELECT ParValue FROM XLControl WHERE Code = 'Database'")
    connectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & GetPref(9) & sDbName & "; Jet OLEDB:Database password=" & pw
    Set dbConnection = CreateObject("ADODB.Connection")
    dbConnection.Open connectionString: dbConnection.Close
    Set GetDBConnection = dbConnection
End Function
