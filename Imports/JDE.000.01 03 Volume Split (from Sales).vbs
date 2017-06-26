Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, period As Integer, planVersion As String, periodFrom As String
    Dim connection As Object, country As String, startPeriod As Integer, wksData As Worksheet, rsTonLtr As Object
    Dim rsFacts As Object
    Set wksData = Worksheets("XLRep 02 Distribute Volume fro")
    planVersion = GetPar(wksData.[A1], "Plan Version=")
    periodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion))
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblDistributionKeys WHERE PlanVersion IS NULL")
    Set rsTonLtr = GetEmptyRecordSet("SELECT DISTINCT SKU FROM tblSKU WHERE VolumeType = 'TonLtr'")
    Set rsFacts = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")
    startPeriod = CInt(Right(periodFrom, 2))
        
    For Each wks In ActiveWorkbook.Sheets
        If Left(wks.name, 5) <> "XLRep" Then
            With wks
                For row = 3 To wks.UsedRange.Rows.Count
                    If Not IsEmpty(.Cells(row, 3)) Then
                        For period = startPeriod To 12
                            If .Cells(row, period + 4) <> 0 Then
                                rs.AddNew
                                rs.Fields("PlanVersion") = planVersion
                                rs.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                                rs.Fields("SKU") = Left(.Cells(row, 2), InStr(.Cells(row, 2), " |") - 1)
                                rs.Fields("AlternativeSKU") = .Cells(row, 1)
                                rs.Fields("Customer") = .[A1]
                                rs.Fields("VolumeSplit") = .Cells(row, period + 4)
                                rsFacts.AddNew
                                rsFacts.Fields("Country") = "AT"
                                rsFacts.Fields("PlanVersion") = planVersion
                                rsFacts.Fields("Period") = CLng(Left(periodFrom, 4)) * 100 + period
                                rsFacts.Fields("SourceType") = "FCVolume"
                                rsFacts.Fields("Forecast") = "yes"
                                rsFacts.Fields("SKU") = Left(.Cells(row, 2), InStr(.Cells(row, 2), " |") - 1)
                                rsFacts.Fields("Customer") = .[A1]
                                rsFacts.Fields("PromoNonPromo") = "NonPromo"
                                rsFacts.Fields("OnOffInvoice") = "On"
                                If IsTonLtr(rsFacts.Fields("SKU"), rsTonLtr) Then rsFacts.Fields("VolumeType") = "TonLtr" Else rsFacts.Fields("VolumeType") = "Ton"
                                rsFacts.Fields("Volume") = .Cells(row, period + 18)
                            End If
                        Next period
                    End If
                Next row
            End With
        End If
    Next wks
    
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblDistributionKeys WHERE PlanVersion = " & Quot(planVersion)
    connection.Execute "DELETE FROM tblFacts WHERE Forecast = 'yes' AND PlanVersion = " & Quot(planVersion)
    rs.ActiveConnection = connection
    rs.UpdateBatch
    rsFacts.ActiveConnection = connection
    rsFacts.UpdateBatch
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

Function IsTonLtr(ByVal sku As String, ByRef rs As Object)
    Dim result As Boolean
    result = False
    rs.Filter = ""
    rs.MoveFirst
    rs.Filter = "SKU=" & sku
    If Not rs.EOF Then result = True
    IsTonLtr = result
End Function