Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String
    Dim connection As Object, country As String, startPeriod As Integer
    Set wks = ActiveSheet
    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    periodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion))
If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
    XLImp "ERROR", "The plan version has been locked for input": Exit Sub
End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblFactsTemp WHERE PlanVersion IS NULL")
    startPeriod = CInt(Right(periodFrom, 2))
    
    With wks
        For row = 2 To wks.UsedRange.Rows.Count
            If Not IsEmpty(.Cells(row, 5)) Then
                For period = startPeriod To 12
                    If .Cells(row, period + 20) <> 0 Then
                        rs.AddNew
                        rs.Fields("Country") = country
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                        rs.Fields("Forecast") = "yes"
                        rs.Fields("Product") = .Cells(row, 4)
                        rs.Fields("PlanningCustomer") = .[D1]
                        rs.Fields("PromoNonPromo") = "NonPromo"
                        rs.Fields("VolumeSplit") = .Cells(row, period + 6)
                        rs.Fields("Volume") = .Cells(row, period + 20)
                    End If
                Next period
            End If
        Next row
    End With
	
	Dim salesCustomer As String
	salesCustomer = wks.[D1]
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFactsTemp WHERE PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & _
        " AND PlanningCustomer = " & Quot(salesCustomer)
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
    Dim pw As String, connectionString As String, dbConnection As Object
    
    pw = "xlsysjs14"
    'connectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & GetPref(9) & "XLReporting_JDE_Retail_DE.dat; Jet OLEDB:Database password=" & pw
    Set dbConnection = CreateObject("ADODB.Connection")
    dbConnection.Open connectionString: dbConnection.Close
    Set GetDBConnection = dbConnection
End Function