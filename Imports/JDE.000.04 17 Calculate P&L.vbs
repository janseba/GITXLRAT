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

    'Calculate PPP (LPA)
    sql = "UPDATE tblFacts SET tblFacts.LPA = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset LPA to 0..."
    sql = "UPDATE (tblFacts INNER JOIN tblPPP " & _
        "ON tblFacts.PlanVersion=tblPPP.PlanVersion AND tblFacts.Period = tblPPP.Period AND tblFacts.Customer = tblPPP.Customer) " & _
        "INNER JOIN tblSKU ON tblFacts.SKU = tblSKU.SKU AND tblPPP.SalesConditionLevel = tblSKU.SalesConditionLevel " & _
        "SET tblFacts.LPA = -1 * tblFacts.Pieces * tblPPP.PPPPerPiece " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Calculate PPP..."

    'Calculate TPR On-Invoice (discount1eur)
    sql = "UPDATE tblFacts SET tblFacts.discount1eur = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset discount1eur to 0..."
    sql = "UPDATE ((tblFacts INNER JOIN tblTPRPromoShare " & _
        "ON tblFacts.PlanVersion=tblTPRPromoShare.PlanVersion AND tblFacts.Period = tblTPRPromoShare.Period AND tblFacts.Customer = tblTPRPromoShare.Customer) " & _
        "INNER JOIN tblTPROnInvoice ON tblFacts.PlanVersion = tblTPROnInvoice.planVersion AND tblFacts.Period = tblTPROnInvoice.Period AND tblFacts.Customer = tblTPROnInvoice.Customer) " & _
        "INNER JOIN tblSKU ON tblFacts.SKU = tblSKU.SKU AND tblTPRPromoShare.SalesConditionLevel = tblSKU.SalesConditionLevel AND tblTPROnInvoice.SalesConditionLevel = tblSKU.SalesConditionLevel " & _
        "SET tblFacts.discount1eur = -1 * tblFacts.Pieces * tblTPRPromoShare.PromoShare * tblTPROnInvoice.TPROnInvoice " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Calculate TPR On-Invoice..."

    'Calculate TPR Off-Invoice (discount1eur)
    sql = "UPDATE tblFacts SET tblFacts.discount4eur = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset discount4eur to 0..."
    sql = "UPDATE ((tblFacts INNER JOIN tblTPRPromoShare " & _
        "ON tblFacts.PlanVersion=tblTPRPromoShare.PlanVersion AND tblFacts.Period = tblTPRPromoShare.Period AND tblFacts.Customer = tblTPRPromoShare.Customer) " & _
        "INNER JOIN tblTPROffInvoice ON tblFacts.PlanVersion = tblTPROffInvoice.planVersion AND tblFacts.Period = tblTPROffInvoice.Period AND tblFacts.Customer = tblTPROffInvoice.Customer) " & _
        "INNER JOIN tblSKU ON tblFacts.SKU = tblSKU.SKU AND tblTPRPromoShare.SalesConditionLevel = tblSKU.SalesConditionLevel AND tblTPROffInvoice.SalesConditionLevel = tblSKU.SalesConditionLevel " & _
        "SET tblFacts.discount4eur = -1 * tblFacts.Pieces * tblTPRPromoShare.PromoShare * tblTPROffInvoice.TPROffInvoice " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Calculate TPR Off-Invoice..."      
End Sub
