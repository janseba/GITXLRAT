Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, planVersion As String, period As String
    Dim connection As Object, country As String, bladen As Variant, sht As Variant, col As Long

    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If

    Set rs = GetEmptyRecordSet("SELECT * FROM tblFactsAOP WHERE PlanVersion IS NULL")

    For Each wks In ActiveWorkbook.Worksheets

        With wks
            If Left(.Name, 16) = "Import-File_XLR-" Then
                For row = 2 To wks.UsedRange.Rows.Count
                    For col = 6 To 17
                        If Not IsEmpty(.Cells(row, 3)) And .Cells(row, col) <> 0 Then
                            rs.AddNew
                            rs.Fields("Country") = country
                            rs.Fields("PlanVersion") = planVersion
                            rs.Fields("Period") = Left(.Cells(1, col), 4) & Right(.Cells(1, col), 2)
                            rs.Fields("SourceType") = "GP AOP17"
                            rs.Fields("Forecast") = "yes"
                            rs.Fields("SKU") = .Cells(row, 4)
                            rs.Fields("Customer") = .Cells(row, 5)
                            rs.Fields("PromoNonPromo") = "NonPromo"
                            rs.Fields("OnOffInvoice") = ""
                            rs.Fields(CStr(.Cells(row, 3))) = .Cells(row, col)
                        End If
                    Next col
                Next row
            End If
        End With
    Next wks
    
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFactsAOP WHERE SourceType = 'GP AOP17' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country)
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'GP AOP17' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country)
    rs.ActiveConnection = connection
    rs.UpdateBatch
    Application.Wait DateAdd("s", 5, Now) 'Wait for 1 second
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    XLImp "INSERT INTO tblFacts(Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice, " & _
        "Volume, FAP1, [14_3TermofPayment], lpa, discount1eur, discount4eur, discount3eur, [107_TABDFOffinvTAS], [17_1OneListFee], discount5eur, " & _
        "discount2fix, mb, DisplayCosts, ecoTax) " & _
        "SELECT Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice, SUM(Volume), SUM(FAP1)" & _
        ", -SUM([14_3TermofPayment]), -SUM(lpa), -SUM(discount1eur), -SUM(discount4eur), -SUM(discount3eur), -SUM([107_TABDFOffinvTAS]), -SUM([17_1OneListFee])" & _
        ", SUM(discount5eur), -SUM(discount2fix), -SUM(mb) + SUM(DisplayCosts), -SUM(DisplayCosts), -Sum(ecoTax)" & _
        "FROM tblFactsAOP " & _
        "WHERE PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & _
        " GROUP BY Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice", "Insert AOP in database"
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

