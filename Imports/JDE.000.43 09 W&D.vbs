Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String
    Dim connection As Object, country As String, startPeriod As Integer, sheetnames As Variant, sheetName As Variant
    sheetnames = Array("Warehousing costs per kg", "Distribution costs per kg")
    
    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    periodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion))
    startPeriod = CInt(Right(periodFrom, 2))
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblWD WHERE PlanVersion IS NULL")
    
    For Each sheetName In sheetnames
        Set wks = ActiveWorkbook.Sheets(sheetName)
        With wks
            For row = 3 To wks.UsedRange.Rows.Count
                If Not IsEmpty(.Cells(row, 3)) Then
                    For period = startPeriod To 12
                        rs.AddNew
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("LineType") = .Name
                        rs.Fields("ProfitCenter") = .Cells(row, 3)
                        rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                        rs.Fields("Amount") = .Cells(row, period + 3)
                    Next period
                End If
            Next row
        End With
    Next sheetName
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblWD WHERE PlanVersion = " & Quot(planVersion)
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
