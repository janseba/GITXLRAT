Sub XLCode()
    Dim vWorksheets As Variant, w As Variant, wks As Worksheet, iCol As Integer, periodFrom As String, planVersion As String
    Dim iRow As Integer, sSalesConditionLevel As String, sql As String, dealFrom As String, dealTo As String, noPeriods As Integer, period As Integer
    Dim NIS As Double, rsNIS As Object, rsFacts As Object, connection As Object, AllocatedDiscount As Double, rsSKU As Object
    Dim sSKU As String, vWsData As Variant, colWsData As Collection, i As Integer, valueField As String, discountCode As String, discountInfo As Variant
    
    vWorksheets = Array("Folders-Ads_XL-Reporting", "Listing fees_XL-Reporting", "TPRs Lump sum_XL-Reporting")
    vWsData = Array("Fol|discount3eur", "LF|107_TABDFOffinvTAS", "TPR|discount4fix")
    Set colWsData = New Collection
    For i = LBound(vWorksheets) To UBound(vWorksheets)
        colWsData.Add vWsData(i), vWorksheets(i)
    Next i
    planVersion = GetPar([A1], "Plan Version=")
    If planVersion = "" Then XLImp "Error", "There was no planversion selected": Exit Sub
    periodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion))
    
    RunSQL "DELETE FROM tblFacts WHERE LEFT(SourceType,2) = 'FD' AND PlanVersion = " & Quot(planVersion)
    
    For Each w In vWorksheets
        discountInfo = Split(colWsData.Item(w), "|")
        valueField = discountInfo(1)
        discountCode = discountInfo(0)
        Set wks = ActiveWorkbook.Worksheets(w)
        With wks
            For iCol = 3 To wks.UsedRange.Columns.Count
                If Not IsError(.Cells(2, iCol)) And Not IsError(.Cells(10, iCol)) And Not IsError(.Cells(5, iCol)) Then
                    dealTo = .Cells(10, iCol)
                    'Check if Customer Nr and 5 are filled and if forecast periodFrom <= end period of deal
                    If Not IsEmpty(.Cells(2, iCol)) And .Cells(5, iCol) <> 0 And periodFrom <= dealTo Then
                        sSalesConditionLevel = ""
                        'make a list of SalesConditionLevel
                        For iRow = 13 To .UsedRange.Rows.Count
                            If UCase(.Cells(iRow, iCol)) = "X" Then sSalesConditionLevel = sSalesConditionLevel & "," & Quot(.Cells(iRow, 2))
                        Next iRow
                        sSalesConditionLevel = Mid(sSalesConditionLevel, 2)
                        
                        'get number of periods
                        If periodFrom < .Cells(9, iCol) Then dealFrom = .Cells(9, iCol) Else dealFrom = periodFrom
                        noPeriods = Right(dealTo, 2) - Right(dealFrom, 2) + 1
                        
                        'get recordset with NIS details and totals
                        sql = "SELECT 'Detail' AS Type, b.Period, b.Customer, b.SKU, b.PromoNonPromo, SUM(b.NIS) AS NIS FROM View_PLBase AS b WHERE b.PlanVersion = " & Quot(planVersion) & " AND b.Period BETWEEN " & Quot(dealFrom) & " AND " & Quot(dealTo) & " AND Customer = " & Quot(.Cells(2, iCol)) & " AND SalesConditionLevel IN (" & sSalesConditionLevel & ") GROUP BY b.Period, b.PromoNonPromo, b.Customer, b.SKU UNION ALL SELECT 'Total' AS Type, a.Period, " & Quot(.Cells(2, iCol)) & " AS Customer,'TOTAL' AS SKU, 'TOTAL' AS PromoNonPromo, SUM(a.NIS) AS NIS FROM View_PLBase AS a WHERE a.PlanVersion = " & Quot(planVersion) & " AND a.Period BETWEEN " & _
                            Quot(dealFrom) & " AND " & Quot(dealTo) & " AND a.Customer = " & Quot(.Cells(2, iCol)) & " AND a.SalesConditionLevel IN (" & sSalesConditionLevel & ")" & _
                            " GROUP BY Period"
                        Set rsNIS = GetRecordSet(sql)
                        Set rsFacts = GetRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")
                        
                        
                        'add records to tblFacts for each Period
                        For period = 1 To noPeriods
                            AllocatedDiscount = 0
                            rsNIS.Filter = "Type = 'Total' AND Period = " & Quot(dealFrom + period - 1)
                            If Not (rsNIS.BOF Or rsNIS.EOF) Then 'Er moet wel NIS in de periode zijn
                                NIS = rsNIS.Fields("NIS").Value
                                rsNIS.Filter = 0
                                rsNIS.Filter = "Type = 'Detail' AND Period = " & Quot(dealFrom + period - 1)
                                rsNIS.MoveFirst
                                Do Until rsNIS.EOF
                                    rsFacts.AddNew
                                    rsFacts.Fields("Country") = "AT"
                                    rsFacts.Fields("PlanVersion") = planVersion
                                    rsFacts.Fields("Period") = rsNIS.Fields("Period").Value
                                    rsFacts.Fields("SourceType") = "FD" & discountCode & Right(.Cells(3, iCol),7)
                                    rsFacts.Fields("Forecast") = "yes"
                                    rsFacts.Fields("SKU") = rsNIS.Fields("SKU").Value
                                    rsFacts.Fields("Customer") = rsNIS.Fields("Customer").Value
                                    rsFacts.Fields("PromoNonPromo") = rsNIS.Fields("PromoNonPromo").Value
                                    If NIS = 0 Then
                                        rsFacts.Fields(valueField) = 0
                                    Else
                                        rsFacts.Fields(valueField) = -(rsNIS.Fields("NIS") / NIS) * (.Cells(5, iCol) / .Cells(11, iCol))
                                    End If
                                    AllocatedDiscount = AllocatedDiscount + rsFacts.Fields(valueField)
                                    rsNIS.MoveNext
                                Loop
                            End If
                            If Round(-(.Cells(5, iCol) / .Cells(11, iCol)), 2) <> Round(AllocatedDiscount, 2) Then

                                'Determine number of profit centers
                                Set rsSKU = GetRecordSet("SELECT DISTINCT ProfitCenter FROM tblSKU WHERE SalesConditionLevel IN (" & sSalesConditionLevel & ")")
                                'Get pseudo SKU for Beans
                                sSKU = GetRecordSet("SELECT TOP 1 sku FROM tblSKU WHERE EUProductHierarchy = 'DEMB Planning & Budgeting' AND ProfitCenter = '22042  RETAIL AT MSERVE BNS'").Fields("SKU").Value
                                'More than one profit center than put unallocated fixed discount on Beans
                                If rsSKU.RecordCount > 1 Then
                                    rsFacts.AddNew
                                    rsFacts.Fields("Country") = "AT"
                                    rsFacts.Fields("PlanVersion") = planVersion
                                    rsFacts.Fields("Period") = dealFrom + period - 1
                                    rsFacts.Fields("SourceType") = "FD" & discountCode & Right(.Cells(3, iCol),7)
                                    rsFacts.Fields("Forecast") = "yes"
                                    rsFacts.Fields("SKU") = sSKU
                                    rsFacts.Fields("Customer") = .Cells(2, iCol)
                                    rsFacts.Fields("PromoNonPromo") = "NonPromo"
                                    rsFacts.Fields(valueField) = -(.Cells(5, iCol) / .Cells(11, iCol)) - AllocatedDiscount
                                Else
                                    'Get dummy sku for profit center
                                    Set rsSKU = GetRecordSet("SELECT  TOP 1 SKU FROM tblSKU WHERE ProfitCenter = " & Quot(rsSKU.Fields("ProfitCenter").Value) & " AND EUProductHierarchy = 'DEMB Planning & Budgeting'")
                                    If rsSKU.RecordCount = 1 Then sSKU = rsSKU.Fields("SKU").Value
                                    rsFacts.AddNew
                                    rsFacts.Fields("Country") = "AT"
                                    rsFacts.Fields("PlanVersion") = planVersion
                                    rsFacts.Fields("Period") = dealFrom + period - 1
                                    rsFacts.Fields("SourceType") = "FD" & discountCode & Right(.Cells(3, iCol),7)
                                    rsFacts.Fields("Forecast") = "yes"
                                    rsFacts.Fields("SKU") = sSKU
                                    rsFacts.Fields("Customer") = .Cells(2, iCol)
                                    rsFacts.Fields("PromoNonPromo") = "NonPromo"
                                    rsFacts.Fields(valueField) = -(.Cells(5, iCol) / .Cells(11, iCol)) - AllocatedDiscount
                                End If
                            End If
                        Next period
                    Set connection = GetDBConnection: connection.Open
                    rsFacts.ActiveConnection = connection
                    rsFacts.UpdateBatch
                    connection.Close
                    End If
                End If
            Next iCol
        End With
    Next w

    'Copy Listing Fees to seperate field
    RunSQL "UPDATE tblFacts SET [17_1OneListFee] = [107_TABDFOffinvTAS] WHERE LEFT(SourceType, 4) = 'FDLF' AND Forecast = 'yes' AND PlanVersion = " & Quot(planVersion)End Sub

Function GetRecordSet(ByVal sTable As String) As Object
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
    Set GetRecordSet = rsData
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