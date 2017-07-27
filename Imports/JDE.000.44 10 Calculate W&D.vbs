Sub XLCode()
    Dim sql As String, planVersion As String
    planVersion = GetPar([A1], "Plan Version=")
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    
    'Calculate Shipping
    sql = "UPDATE tblFacts SET tblFacts.Shipping = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset Shipping to 0..."
    sql = "UPDATE (tblFacts INNER JOIN tblWD " & _
        "ON tblFacts.PlanVersion=tblWD.PlanVersion AND tblFacts.Period = tblWD.Period) " & _
        "INNER JOIN tblSKU ON tblFacts.SKU = tblSKU.SKU AND tblWD.ProfitCenter = LEFT(tblSKU.ProfitCenter, 5) " & _
        "SET tblFacts.Shipping = tblFacts.Volume * - tblWD.Amount " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblFacts.PlanVersion = " & Quot(planVersion) & " AND tblWD.LineType = 'Distribution costs per kg'"
    XLImp sql, "Calculate Shipping..."

    'Calculate Warehouse
    sql = "UPDATE tblFacts SET tblFacts.Warehouse = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset Warehouse to 0..."
    sql = "UPDATE (tblFacts INNER JOIN tblWD " & _
        "ON tblFacts.PlanVersion=tblWD.PlanVersion AND tblFacts.Period = tblWD.Period) " & _
        "INNER JOIN tblSKU ON tblFacts.SKU = tblSKU.SKU AND tblWD.ProfitCenter = LEFT(tblSKU.ProfitCenter, 5) " & _
        "SET tblFacts.Warehouse = tblFacts.Volume * - tblWD.Amount " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblFacts.PlanVersion = " & Quot(planVersion) & " AND tblWD.LineType = 'Warehousing costs per kg'"
    XLImp sql, "Calculate Warehouse..."

End Sub
