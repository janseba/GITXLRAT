Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String, periodTo As String
    Dim connection As Object, country As String, startPeriod As Integer, bladen As Variant, sht As Variant, endPeriod As Integer

    bladen = Array("Distribution costs per kg", "Warehousing costs per kg")

    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    periodFrom = GetPar([A1], "Period From=")
    periodTo = GetPar([A1], "Period To=")
    startPeriod = CInt(Right(periodFrom, 2))
    endPeriod = CInt(Right(periodTo, 2))
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")

    For Each sht In bladen
        Set wks = ActiveWorkbook.Sheets(sht)
        With wks
            For row = 3 To wks.UsedRange.Rows.Count
                If Not IsEmpty(.Cells(row, 1)) Then
                    If Left(.Cells(row, 1), 5) = "Total" Then Exit For
                    For period = startPeriod To endPeriod
                        If .Cells(row, 3 + period) <> 0 Then
                            rs.AddNew
                            rs.Fields("Country") = country
                            rs.Fields("PlanVersion") = planVersion
                            rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                            rs.Fields("SourceType") = "WD AOP17"
                            rs.Fields("Forecast") = "yes"
                            rs.Fields("sku") = .Cells(row, 1)
                            rs.Fields("Customer") = .[B1]
                            rs.Fields("PromoNonPromo") = "NonPromo"
                            If sht = "Distribution costs per kg" Then
                                rs.Fields("Shipping") = -1 * IIf(IsError(.Cells(row, 3 + period)), 0, .Cells(row, 3 + period))
                            ElseIf sht = "Warehousing costs per kg" Then
                                rs.Fields("Warehouse") = -1 * IIf(IsError(.Cells(row, 3 + period)), 0, .Cells(row, 3 + period))
                            End If
                        End If
                    Next period
                End If
            Next row
        End With
    Next sht

    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'WD AOP17' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & _
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