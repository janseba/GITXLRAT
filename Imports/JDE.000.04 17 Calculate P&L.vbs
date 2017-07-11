Sub XLCode()
    Dim sql As String, planVersion As String
    planVersion = GetPar([A1], "Plan Version=")
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    
    'Calculate FAP1
    sql = "UPDATE tblFacts SET tblFacts.FAP1 = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset FAP1 to 0..."
    sql = "UPDATE (tblFacts INNER JOIN tblFAP " & _
        "ON tblFacts.PlanVersion=tblFAP.PlanVersion AND tblFacts.Period = tblFAP.Period) " & _
        "INNER JOIN tblSKU ON tblFacts.SKU = tblSKU.SKU AND tblFAP.SalesCondition = tblSKU.SalesConditionLevel " & _
        "SET tblFacts.FAP1 = tblFacts.Pieces * tblFAP.FAPPerPiece " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Calculate GOS..."
    
End Sub
