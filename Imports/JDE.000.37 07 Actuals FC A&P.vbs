Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String, periodTo As String
    Dim connection As Object, country As String, startPeriod As Integer, endPeriod As Integer, forecastOrActual As String
    Dim sourceType As String, startPlanVersion As String
    
    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    periodFrom = GetPar([A1], "Period From=")
    periodTo = GetPar([A1], "Period To=")
    forecastOrActual = GetPar([A1], "Forecast / Actuals=")
    startPeriod = CInt(Right(periodFrom, 2))
    endPeriod = CInt(Right(periodTo, 2))
    startPlanVersion = GetSQL("SELECT FromPeriod FROM sources WHERE Source = " & Quot(planVersion))
    sourceType = GetSQL("SELECT SourceType FROM Sources WHERE Source = " & Quot(planVersion))
    
    If forecastOrActual = "Forecast" And sourceType = "ACTUALS" Then
        XLImp "ERROR", "Your trying to upload forecast data in a plan version for Actuals": Exit Sub
    End If
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")

    Set wks = ActiveSheet

    With wks
        For row = 3 To wks.UsedRange.Rows.Count
            If Not IsEmpty(.Cells(row, 5)) Then
                For period = startPeriod To endPeriod
                    If .Cells(row, period + 5) <> 0 And .Cells(row, period + 5) <> "" Then
                        rs.AddNew
                        rs.Fields("Country") = country
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                        rs.Fields("SourceType") = "AP ACT & FC"
                        If CLng(Left(periodFrom, 4)) * 100 + period < startPlanVersion Or forecastOrActual = "Actuals" Then
                            rs.Fields("Forecast") = "no"
                        Else
                            rs.Fields("Forecast") = "yes"
                        End If
                        rs.Fields("sku") = .Cells(row, 5)
                        rs.Fields("Customer") = .[D1]
                        rs.Fields("PromoNonPromo") = "NonPromo"
                        If .Cells(row, 4) = "Adv Work Media" Then
                            rs.Fields("AdvWorkingMedia") = -1 * IIf(IsError(.Cells(row, period + 5)), 0, .Cells(row, period + 5))
                        ElseIf .Cells(row, 4) = "Adv Non Work Med" Then
                            rs.Fields("AdvNonWorkingMedia") = -1 * IIf(IsError(.Cells(row, period + 5)), 0, .Cells(row, period + 5))
                        ElseIf .Cells(row, 4) = "Brewer Support" Then
                            rs.Fields("BrewerSupport") = -1 * IIf(IsError(.Cells(row, period + 5)), 0, .Cells(row, period + 5))
                        ElseIf .Cells(row, 4) = "Promotion Other" Then
                            rs.Fields("PromotionOther") = -1 * IIf(IsError(.Cells(row, period + 5)), 0, .Cells(row, period + 5))
                        End If
                    End If
                Next period
            End If
        Next row
    End With

    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'AP ACT & FC' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & _
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