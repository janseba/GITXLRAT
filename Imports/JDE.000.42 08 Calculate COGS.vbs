Sub XLCode()
    Dim sql As String, planVersion As String, rsFacts As Object, rsCostAdj As Object, connection As Object
    
    planVersion = GetPar([A1], "Plan Version=")
    If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(planVersion)) = "y" Then
        XLImp "ERROR", "The plan version has been locked for input": Exit Sub
    End If
    
    'Calculate EcoTax
    sql = "UPDATE tblFacts SET tblFacts.ecotax = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset EcoTax to 0..."
    sql = "UPDATE tblFacts INNER JOIN tblCostInput " & _
        "ON tblFacts.PlanVersion=tblCostInput.PlanVersion AND tblFacts.Period = tblCostInput.Period AND tblFacts.SKU = tblCostInput.sku " & _
        "SET tblFacts.ecotax = tblFacts.Volume * - tblCostInput.Amount " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblCostInput.LineType = 'EcoTax € kg' AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Calculate EcoTax..."
    
    'Calculate MB
    sql = "UPDATE tblFacts SET tblFacts.mb = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset MB to 0..."
    sql = "UPDATE tblFacts INNER JOIN tblCostInput " & _
        "ON tblFacts.PlanVersion=tblCostInput.PlanVersion AND tblFacts.Period = tblCostInput.Period AND tblFacts.SKU = tblCostInput.sku " & _
        "SET tblFacts.mb = tblFacts.Volume * - tblCostInput.Amount " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblCostInput.LineType = 'MB € kg ex display-costs' AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Calculate MB..."
    
    'Calculate Display
    sql = "UPDATE tblFacts SET tblFacts.displaycosts = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset Display Costs to 0..."
    sql = "UPDATE tblFacts INNER JOIN tblCostInput " & _
        "ON tblFacts.PlanVersion=tblCostInput.PlanVersion AND tblFacts.Period = tblCostInput.Period AND tblFacts.SKU = tblCostInput.sku " & _
        "SET tblFacts.displaycosts = tblFacts.Volume * - tblCostInput.Amount " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblCostInput.LineType = 'Display € kg' AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Calculate Display Costs..."

    'Calculate Cost ADJ per kg
    sql = "UPDATE tblFacts SET tblFacts.mbadjustments = 0 WHERE tblFacts.Forecast = 'yes' " & _
       " AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Reset Cost ADJ per kg to 0..."
    sql = "UPDATE tblFacts INNER JOIN tblCostInput " & _
        "ON tblFacts.PlanVersion=tblCostInput.PlanVersion AND tblFacts.Period = tblCostInput.Period AND tblFacts.SKU = tblCostInput.sku " & _
        "SET tblFacts.mbadjustments = tblFacts.Volume * - tblCostInput.Amount " & _
        "WHERE tblFacts.Forecast = 'yes' AND tblCostInput.LineType = 'cost ADJ € kg' AND tblFacts.PlanVersion = " & Quot(planVersion)
    XLImp sql, "Calculate Cost ADJ per kg..."
    
    Set rsFacts = GetRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")
    Set rsCostAdj = GetRecordSet("SELECT * FROM tblCostInput WHERE LineType = 'cost ADJ abs' AND PlanVersion = " & Quot(planVersion))
    rsCostAdj.MoveFirst
    Do Until rsCostAdj.EOF
        rsFacts.AddNew
        rsFacts.Fields("Country") = "AT"
        rsFacts.Fields("PlanVersion") = planVersion
        rsFacts.Fields("Period") = rsCostAdj.Fields("Period").Value
        rsFacts.Fields("SourceType") = "Cost ADJ abs"
        rsFacts.Fields("Forecast") = "yes"
        rsFacts.Fields("SKU") = rsCostAdj.Fields("SKU").Value
        rsFacts.Fields("Customer") = rsCostAdj.Fields("Customer").Value
        rsFacts.Fields("mbadjustments") = -rsCostAdj.Fields("Amount").Value
        rsCostAdj.MoveNext
    Loop
    
    RunSQL "DELETE FROM tblFacts WHERE SourceType  = 'Cost ADJ abs' AND Forecast = 'yes' AND  PlanVersion = " & Quot(planVersion)

    Set connection = GetDBConnection: connection.Open
    rsFacts.ActiveConnection = connection
    rsFacts.UpdateBatch
    connection.Close
    
End Sub

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