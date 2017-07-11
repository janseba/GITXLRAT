Sub XLCode()
    Dim wks As Worksheet, planVersion As String, periodFrom As String, periodTo As String
    Dim country As String, startPeriod As Integer, sql As String, planVersionTarget As String

    Set wks = ActiveSheet
    planVersion = GetPar([A1], "Plan Version Source=")
    country = GetPar([A1], "Country=")
    periodFrom = GetPar([A1], "Period From=")
    periodTo = GetPar([A1], "Period To=")
    planVersionTarget = GetPar([A1], "Plan Version Target=")
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersionTarget)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    sql = "DELETE FROM tblReferenceVolume WHERE PlanVersion = " & Quot(planVersionTarget)
    XLImp sql, "Delete old reference data"
    sql = "INSERT INTO tblReferenceVolume(PlanVersion, SKU, Customer, Volume) " & _
        "SELECT " & Quot(planVersionTarget) & ", a.SKU, a.Customer, SUM(a.Volume) " & _
        "FROM tblFacts AS a INNER JOIN tblSKU AS b ON a.SKU = b.SKU " & _
        "WHERE PlanVersion = " & Quot(planVersion) & " AND Forecast = 'no' AND Period Between " & Quot(periodFrom) & " AND " & Quot(periodTo) & " AND b.Active = 'yes' " & _
        "GROUP BY a.PlanVersion, a.SKU, a.Customer, Period " & _
        "HAVING SUM(Volume) > 0"
    XLImp sql, "Add reference volumes to database"
    sql = "DELETE FROM tblrefdistributionkeys WHERE planversion = " & Quot(planVersionTarget)
    XLImp sql, "Delete old distribution keys..."
    sql = "INSERT INTO tblrefdistributionkeys (planversion, alternativesku, customer, salesplanning, distrkey) " & _
        "SELECT PlanVersion, AlternativeSKU, Customer, SalesPlanning, ROUND(DistrKey * 100, 0) AS DKey FROM View_VolumeDistributionKeys " & _
        "WHERE DistrKey <> 0 AND PlanVersion = " & Quot(planVersionTarget)
    XLImp sql, "Adding distribution keys..."
    sql = "INSERT INTO tblrefdistributionkeys (planversion, alternativesku, customer, salesplanning, distrkey) " & _
        "SELECT planversion, Max(alternativesku), d.customer, d.salesplanning, d.verschil " & _
        "FROM (SELECT a.planversion, a.customer, a.salesplanning, a.alternativesku, c.verschil " & _
        "FROM (tblrefdistributionkeys AS a INNER JOIN view_maxdistributionkey AS b ON a.planversion = b.planversion " & _
        "AND a.customer = b.customer AND a.salesplanning = b.salesplanning AND a.distrkey = b.maxkey) " & _
        "INNER JOIN view_distributionkeydifference AS c ON a.planversion = c.planversion AND a.customer = c.customer " & _
        "AND a.salesplanning = c.salesplanning) AS d " & _
        "GROUP BY d.planversion, d.customer, d.salesplanning, d.verschil"
    XLImp sql, "Correct rounding errors in distribution keys..."
End Sub