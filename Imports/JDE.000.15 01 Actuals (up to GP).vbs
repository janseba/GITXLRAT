Sub XLCode()
    Dim wks As Worksheet, row As Long, rs As Object, planVersion As String, period As String
    Dim connection As Object, country As String, bladen As Variant, sht As Variant, periodFrom As String, sVolumeType As String
    Dim sql As String

    bladen = Array("EcoTax € kg", "CTax € kg", "MB € kg", "Display € kg")

    planVersion = GetPar([A1], "Plan Version=")
    country = GetPar([A1], "Country=")
    period = GetPar([A1], "Period=")
    periodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion))
If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
    XLImp "ERROR", "The plan version has been locked for input": Exit Sub
End If
    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")

    Set wks = ActiveSheet

    With wks
        For row = 5 To wks.UsedRange.Rows.Count
            If Not IsEmpty(.Cells(row, 3)) Then
                if IsEmpty(.Cells(row, 15)) AND Not IsEmpty(.Cells(row, 16)) Then
                  sVolumeType = "TonLtr"
                ElseIf IsEmpty(.Cells(row, 16)) AND Not IsEmpty(.Cells(row, 15)) Then
                  sVolumeType = "Ton"
                Else
                  sVolumeType = "Ton"
                End If
                rs.AddNew
                rs.Fields("Country") = country
                rs.Fields("PlanVersion") = planVersion
                rs.Fields("Period") = period
                rs.Fields("SourceType") = "ActualsGP"
                rs.Fields("Forecast") = "no"
                rs.Fields("SKU") = .Cells(row, 4)
                rs.Fields("Customer") =  IIf(.Cells(row, 11)="#","D99997",.Cells(row,11)) 'If not assigned then dummy finance customer'
                rs.Fields("PromoNonPromo") = "NonPromo"
                rs.Fields("OnOffInvoice") = IIf(.Cells(row,6)="On Invoice","On",IIf(.Cells(row,6)="Off Invoice FIN","Off_F","Off_T"))
                rs.Fields("VolumeType") = sVolumeType
                rs.Fields("NoDisplays") = .Cells(row, 19) 'KIT-DIS BU Total Sales
                rs.Fields("Volume") = .Cells(row, 14) 'Vol RU Total Sales (KG)
                            rs.Fields("105_GOSFAP1") = .Cells(row, 23) + .Cells(row, 24) + .Cells(row, 25) + .Cells(row, 26) '10.1 Normal Sales + 10.2 Ret Gift Coupon + 10.3 Corr Seas Prod + 10.5 Excise
                            rs.Fields("14_3TermofPayment") = .Cells(row, 34) '14.3 Term of Payment
                            If .Cells(row, 6) = "On Invoice" Then
                                   rs.Fields("102_PPROnInv") = .Cells(row, 39)  + .Cells(row, 40) '11.1 PPR Standard + 11.2 List Price Adj
                                   rs.Fields("101_SalesReturnsOnInv") = .Cells(row, 29) '10.4 Sales Returns
                                   rs.Fields("104_TAEfficiencyOnInv") =.Cells(row, 32) + .Cells(row, 33) + .Cells(row, 35) '14.1 Logistic + 14.2 ECR + 14.4 Cost Avoid Ret
                                   rs.Fields("103_TPROnInv") = .Cells(row, 44) + .Cells(row, 45) + .Cells(row, 46) '12.1 TPR Extra Vol + 12.2 TPR Free Charge + 12.3 TPR Other
                       rs.Fields("CPIncentivesOnInv") = .Cells(row, 49) '12.4 Price Coupon/LC
                                   rs.Fields("106_TABMCOnInv") = .Cells(row, 54) + .Cells(row, 55) + .Cells(row, 56) '16.1 Folders/Ads + 16.2 Placements + 16.3 Other/Consumer Advertising Trade
                                   rs.Fields("107_TABDFOnInv") = .Cells(row, 60) + .Cells(row, 61) + .Cells(row, 62) + .Cells(row, 63) + .Cells(row, 64) + _
                                          .Cells(row, 65) + .Cells(row, 66) + .Cells(row, 67) + .Cells(row, 68) + .Cells(row, 69) + .Cells(row, 70) + .Cells(row, 71) +  _
                                          .Cells(row, 72) '13.1 Year End Rebate + 13.2 Eurobonus + 13.3 Headquarters + 13.4 Growth Bonus + 15.1 Stores + 15.2 Special Occasion
                                          '15.3 Merchandising + 15.4 Assortment + 15.5 Redistribution + 15.6 Shelf + 15.7 Category Mangmt + 17.1 One List Fee
                                          '17.2 One off Price Equalization
                            ElseIf .Cells(row, 6) = "Off Invoice TAS" Then
                                   rs.Fields("102_PPROffInvTAS") = .Cells(row, 39)  + .Cells(row, 40) '11.1 PPR Standard + 11.2 List Price Adj
                                   rs.Fields("101_SalesReturnsOffInvTAS") = .Cells(row, 29) '10.4 Sales Returns
                                   rs.Fields("104_TAEfficiencyOffInvTAS") =.Cells(row, 32) + .Cells(row, 33) + .Cells(row, 35) '14.1 Logistic + 14.2 ECR + 14.4 Cost Avoid Ret
                                   rs.Fields("103_TPROffInvTAS") = .Cells(row, 44) + .Cells(row, 45) + .Cells(row, 46) '12.1 TPR Extra Vol + 12.2 TPR Free Charge + 12.3 TPR Other
                       rs.Fields("CPIncentivesOffInvTAS") = .Cells(row, 49) '12.4 Price Coupon/LC
                                   rs.Fields("106_TABMCOffInvTAS") = .Cells(row, 54) + .Cells(row, 55) + .Cells(row, 56) '16.1 Folders/Ads + 16.2 Placements + 16.3 Other/Consumer Advertising Trade
                                   rs.Fields("107_TABDFOffInvTAS") = .Cells(row, 60) + .Cells(row, 61) + .Cells(row, 62) + .Cells(row, 63) + .Cells(row, 64) + _
                                          .Cells(row, 65) + .Cells(row, 66) + .Cells(row, 67) + .Cells(row, 68) + .Cells(row, 69) + .Cells(row, 70) + .Cells(row, 71) +  _
                                          .Cells(row, 72) '13.1 Year End Rebate + 13.2 Eurobonus + 13.3 Headquarters + 13.4 Growth Bonus + 15.1 Stores + 15.2 Special Occasion
                                          '15.3 Merchandising + 15.4 Assortment + 15.5 Redistribution + 15.6 Shelf + 15.7 Category Mangmt + 17.1 One List Fee
                                          '17.2 One off Price Equalization
                            ElseIf .Cells(row, 6) = "Off Invoice FIN" Then
                                   rs.Fields("102_PPROffInvFIN") = .Cells(row, 39)  + .Cells(row, 40) '11.1 PPR Standard + 11.2 List Price Adj
                                   rs.Fields("101_SalesReturnsOffInvFIN") = .Cells(row, 29) '10.4 Sales Returns
                                   rs.Fields("104_TAEfficiencyOffInvFIN") =.Cells(row, 32) + .Cells(row, 33) + .Cells(row, 35) '14.1 Logistic + 14.2 ECR + 14.4 Cost Avoid Ret
                                   rs.Fields("103_TPROffInvFIN") = .Cells(row, 44) + .Cells(row, 45) + .Cells(row, 46) '12.1 TPR Extra Vol + 12.2 TPR Free Charge + 12.3 TPR Other
                       rs.Fields("CPIncentivesOffInvFIN") = .Cells(row, 49) '12.4 Price Coupon/LC
                                   rs.Fields("106_TABMCOffInvFIN") = .Cells(row, 54) + .Cells(row, 55) + .Cells(row, 56) '16.1 Folders/Ads + 16.2 Placements + 16.3 Other/Consumer Advertising Trade
                                   rs.Fields("107_TABDFOffInvFIN") = .Cells(row, 60) + .Cells(row, 61) + .Cells(row, 62) + .Cells(row, 63) + .Cells(row, 64) + _
                                          .Cells(row, 65) + .Cells(row, 66) + .Cells(row, 67) + .Cells(row, 68) + .Cells(row, 69) + .Cells(row, 70) + .Cells(row, 71) +  _
                                          .Cells(row, 72) '13.1 Year End Rebate + 13.2 Eurobonus + 13.3 Headquarters + 13.4 Growth Bonus + 15.1 Stores + 15.2 Special Occasion
                                          '15.3 Merchandising + 15.4 Assortment + 15.5 Redistribution + 15.6 Shelf + 15.7 Category Mangmt + 17.1 One List Fee
                                          '17.2 One off Price Equalization
                            End If             
                            rs.Fields("18Royaltieslncome3rdParty") = .Cells(row, 76) '18 Royalties lncome 3rd Party
                            rs.Fields("CostOfSales") = .Cells(row, 81) + .Cells(row, 82) + .Cells(row, 83)  + .Cells(row, 85) + .Cells(row, 88) + .Cells(row, 101) + .Cells(row, 106) + _
                                   .Cells(row, 102) + .Cells(row, 103) + .Cells(row, 107) + .Cells(row, 108) + .Cells(row, 114) + .Cells(row, 117) + .Cells(row, 118) + _
                                   .Cells(row, 119) + .Cells(row, 120) '22.3.1 Cost of Sales + 22.3.3 Rec Seas Prod + 
                                   '22.3.4 Adjustment Seasonal Products Merchandise Bought + 22.4 Buying Results + 22.7 Cost Of Goods Sold Split + 24.1.1 Mat. Used Price Var Commodities + 
                                   '24.1.2 Mat. Used Price Var Raw Materials + 24.1.3 Mat. Used Price Var Packaging Materials +24.2.1 Mat EV Com +
                                   '24.2.2 Mat EV Raw + 24.2.3 Mat EV Pack + 28.1 Royalties + 28.5 Out of Home Operations Cost +30.1 Mfg Cost Capacity Variances + 
                                   '30.2 Mfg Cost Efficiency Variances + 30.3 Mfg Cost Budget Variances + 30.4 Mfg Cost Other Variances
                            rs.Fields("DisplayCosts") = .Cells(row, 84) '22.3.5 Extra costs
                rs.Fields("GreenDot") = .Cells(row, 86) '22.5 Eco Tax
                rs.Fields("CoffeeTax") = -1 * .Cells(row, 87) '22.6 Excise Tax
                rs.Fields("17_1OneListFee") = .Cells(row, 71) '17.1 One List Fee
                rs.Fields("15_2SpecialOccasion") = .Cells(row, 65) '15.2 Special Occasion
            End If
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'ActualsGP' AND Forecast = 'no' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & _
       " AND Period = " & Quot(period)
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'Overlay' AND Forecast = 'no' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) 
    If period < periodFrom Then
       connection.Execute "DELETE FROM tblFacts WHERE Forecast = 'yes' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & _
       " AND Period = " & Quot(period)
    End If
    rs.ActiveConnection = connection
    rs.UpdateBatch
    XLImp "SELECT COUNT(code) FROM Companies", rs.RecordCount & " lines were added to database in 1 batch update"
    connection.Execute "DELETE FROM tblFacts WHERE SourceType = 'DisplayCor' AND Forecast = 'no' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & _
      " AND Period = " & Quot(period)
    XLImp "INSERT INTO tblFacts (Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice, VolumeType, NoDisplays, Volume) " & _
      "SELECT a.Country, a.PlanVersion, a.Period, 'DisplayCor', a.Forecast, a.SKU, a.Customer, a.PromoNonPromo, a.OnOffInvoice, a.VolumeType, a.NoDisplays, " & _
      "a.NoDisplays * b.NetWeight AS Volume FROM tblFacts AS a INNER JOIN tblMixedDisplays AS b ON a.SKU = b.SKU WHERE b.Type = 'Display' AND a.NoDisplays <> 0 " & _
      "AND a.planVersion = " & Quot(planVersion) & " AND a.Country = " & Quot(country) & " AND Period = " & Quot(period), "Increase volume for displays"
    XLImp "INSERT INTO tblFacts (Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice, VolumeType, NoDisplays, Volume) " & _
      "SELECT a.Country, a.PlanVersion, a.Period, 'DisplayCor' AS SourceType, a.Forecast, b.SKU, a.Customer, a.PromoNonPromo, a.OnOffInvoice, a.VolumeType, a.NoDisplays, " & _
      "-1 * a.NoDisplays * b.Pieces * b.NetWeight AS Volume FROM tblFacts AS a LEFT JOIN tblMixedDisplays AS b ON a.SKU = b.DisplaySKU WHERE a.SourceType='DisplayCor' " & _
      "AND b.Type = 'SKU' AND a.PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & " AND Period = " & Quot(period), "correct volume of dsiplay SKUs"
    XLImp "DELETE FROM tblFacts WHERE SourceType = 'DsplCOGS' AND Forecast = 'no' AND PlanVersion = " & Quot(planVersion) & " AND Country = " & Quot(country) & _
      " AND Period = " & Quot(period)
    XLImp "INSERT INTO tblFacts(Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice, VolumeType, CostOfSales, DisplayCosts) " & _
      "SELECT a.Country, a.PlanVersion, a.Period, 'DsplCOGS', a.Forecast, a.SKU, a.Customer, a.PromoNonPromo, a.OnOffInvoice, a.VolumeType, " & _
      "IIf(c.VolFact = 0, 0, a.Volume / c.VolFact) * c.COGSFact AS COGSAdj, IIf(c.VolFact = 0, 0, a.Volume / c.VolFact) * c.DisplayFact " & _
      "FROM (tblFacts AS a LEFT JOIN tblMixedDisplays AS b ON a.SKU = b.SKU) LEFT JOIN " & _
      "(SELECT PlanVersion, Period, SKU, Customer, IIf(ISNULL(SUM(Volume)), 0, SUM(Volume)) AS VolFact, IIf(ISNULL(SUM(CostOfSales)), 0, SUM(CostOfSales)) AS COGSFact, IIf(ISNULL(SUM(DisplayCosts)), 0, SUM(DisplayCosts)) AS DisplayFact " & _
      "FROM tblFacts WHERE SourceType <> 'DisplayCor' GROUP BY PlanVersion, Period, SKU, Customer) AS c " & _
      "ON a.PlanVersion = c.PlanVersion AND a.Period = c.Period AND a.SKU = c.SKU AND a.Customer = c.Customer " & _
      "WHERE b.Type <> 'Display' AND a.SourceType = 'DisplayCor' AND a.PlanVersion = " & Quot(planVersion) & " AND a.Country = " & Quot(country) & " AND a.Period = " & Quot(period), "correct cogs of skus in case of display"
    XLImp "INSERT INTO tblFacts(Country, PlanVersion, Period, SourceType, Forecast, SKU, Customer, PromoNonPromo, OnOffInvoice, VolumeType, CostOfSales, DisplayCosts) " & _
      "SELECT a.Country, a.PlanVersion, a.Period, 'DsplCOGS', a.Forecast, b.DisplaySKU, a.Customer, a.PromoNonPromo, a.OnOffInvoice, a.VolumeType, - 1 * a.CostOfSales, - 1 * a.DisplayCosts " & _
      "FROM tblFacts AS a INNER JOIN tblMixedDisplays AS b ON a.SKU = b.SKU WHERE a.SourceType = 'DsplCOGS' AND a.PlanVersion = " & Quot(planVersion) & " AND a.Country = " & Quot(country) & " AND a.Period = " & Quot(period), "correct cogs of skus in case of display"
    sql = "UPDATE tblFacts LEFT JOIN tblSKU ON tblFacts.SKU = tblSKU.SKU " & _
        "SET tblFacts.Pieces = IIf(tblSKU.WeightInKg = 0 OR ISNULL(tblSKU.WeightInKg), 0, tblFacts.Volume / tblSKU.WeightInKg), " & _
        "tblFacts.Drinks = IIf(tblSKU.CupsPerKg = 0 OR ISNULL(tblSKU.CupsPerKg), 0, tblFacts.Volume * tblSKU.CupsPerKg), " & _
        "tblFacts.tDiscs = IIf(tblSKU.tDiscPerKg = 0 OR ISNULL(tblSKU.tDiscPerKg), 0, tblFacts.Volume * tblSKU.tDiscPerKg) " & _
        "WHERE tblFacts.PlanVersion = " & Quot(PlanVersion) & " AND tblFacts.Forecast = 'no' AND Period = " & Quot(period)
    XLImp sql, "Calculate Drinks, tDiscs and Pieces ..."
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