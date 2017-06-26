Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String
    Dim connection As Object, startPeriod As Integer
    Set wks = ActiveSheet
    planVersion = GetPar([A1], "Plan Version=")
    periodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion))
    startPeriod = CInt(Right(periodFrom, 2))
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    If wks.Cells(2, startPeriod + 7) <> periodFrom Then
        XLImp "ERROR", "The periods in the file do not correspond with the periods of the planversion.": Exit Sub
    End If

Set rs = GetEmptyRecordSet("SELECT * FROM tblVolumeSales WHERE PlanVersion IS NULL")
    With wks
        For row = 4 To wks.UsedRange.Rows.Count
            If Not IsEmpty(.Cells(row, 5)) Then 'Customer Number
                For period = startPeriod To 12
                    If .Cells(row, period + 7) <> 0 Then
                        rs.AddNew
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("PlanningCustomer") = .Cells(row, 5)
                        rs.Fields("PlanningCategory") = .Cells(row, 7)
                        rs.Fields("PromoNonPromo") = "NonPromo"
                        rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                        rs.Fields("Volume") = .Cells(row, period + 7) * 1000
                    End If
                Next period
            End If
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblVolumeSales WHERE PlanVersion = " & Quot(planVersion)
    rs.ActiveConnection = connection
    rs.UpdateBatch
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