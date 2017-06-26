Sub XLCode()
    Dim template As String, wkbTemplate As Workbook, wksData As Worksheet, wksTemplate As Worksheet, sql As String, customers As Variant
    Dim wkbReport As Workbook, volumeSales As Object, planversion As String, pos As Long, period As Integer, sPeriodFrom As String, n As name
    Dim wks As Worksheet, c As Range, proposeSplit As String

    'Switch of automatic calculation
    Application.Calculation = xlCalculationManual
        
    Set wksData = ActiveSheet
    Set wkbReport = ActiveWorkbook
    planversion = GetPar(wksData.[A2], "Plan Version=")
    sPeriodFrom = GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planversion))
    proposeSplit = GetPar(wksData.[A2], "Propose Split ?=")

    DoEvents: Application.StatusBar = "Get sales volume"
    AddDataSheets "SalesVolume", "SELECT * FROM tblVolumeSales WHERE PlanVersion = " & Quot(planversion), wksData.name
    DoEvents: Application.StatusBar = "Get new distribution keys"
    AddDataSheets "DistributionKeys", "SELECT * FROM View_VolumeDistributionKeys WHERE PlanVersion = " & Quot(planversion), wksData.name
    DoEvents: Application.StatusBar = "Get previous distribution keys"
    AddDataSheets "SDK", "SELECT * FROM tblDistributionKeys WHERE PlanVersion = " & Quot(planversion), wksData.name
        
    'Add a template for every customer
    template = GetPref(9) & "Templates\TemplateVolume.xlsx"
    Set wkbTemplate = Application.Workbooks.Open(Filename:=template, ReadOnly:=True)
    wkbTemplate.Sheets("Volume").Move Before:=wksData
    Set wksTemplate = ActiveWorkbook.Sheets("Volume")
    wksTemplate.[A2] = Left(sPeriodFrom, 4)
    wksTemplate.[A3] = proposeSplit
    For Each c In wksTemplate.Range("rngSalesVolumeBody")
        c.Formula = c.Formula
    Next c
    
    For Each c In wksTemplate.Range("rngSplitBody")
        c.Formula = c.Formula
    Next c
           
    'Get planningCategories
    Dim planningCategories As Object
    sql = "SELECT SalesPlanning FROM tblSKU WHERE Active = 'yes' GROUP BY SalesPlanning, SortOrder ORDER BY SortOrder"
    Set planningCategories = GetRecordSet(sql)
    
    'Fill summarytable
    If planningCategories.RecordCount > 3 Then
        ResizeRngByRows wksTemplate.Range("rngPlanningCategory").Cells(3, 1), _
            planningCategories.RecordCount - 3
    End If
    planningCategories.MoveFirst
    Do Until planningCategories.EOF
        wksTemplate.Range("rngPlanningCategory").Cells(planningCategories.AbsolutePosition + 1, 1) = planningCategories.Fields("SalesPlanning")
        planningCategories.MoveNext
    Loop
    
    'Copy planningblocks
    Dim i As Integer
    If planningCategories.RecordCount > 1 Then
        For i = planningCategories.RecordCount To 2 Step -1
            CopyRange wksTemplate.Names("rngCategory").RefersToRange
        Next i
    End If
    
    'Define first level names
    Dim rngCategory As Range, rngBody As Range
    planningCategories.MoveFirst
    Do Until planningCategories.EOF
        wksTemplate.Names.Add "C_" & CleanName(planningCategories.Fields("SalesPlanning")), wksTemplate.Names("rngCategory").RefersToRange.Offset((planningCategories.AbsolutePosition - 1) * 6)
        Set rngCategory = wksTemplate.Names("C_" & CleanName(planningCategories.Fields("SalesPlanning"))).RefersToRange
        wksTemplate.Names.Add "B_" & CleanName(planningCategories.Fields("SalesPlanning")), rngCategory.Offset(1, 2).Resize(rngCategory.Rows.Count - 3, rngCategory.Columns.Count - 2)
        'rngCategory.Offset(1, 3).Resize(rngCategory.Rows.Count - 3, rngCategory.Columns.Count - 3).FormulaR1C1 = "= SUMIFS(SDK.VolumeSplit,SDK.AlternativeSKU,RC1,SDK.Customer,R1C1,SDK.Period,R2C1 & RIGHT(R9C,2))"
        rngCategory.Cells(1, 1) = planningCategories.Fields("SalesPlanning")
        planningCategories.MoveNext
    Loop
    
    'Fill SKUs per planning category
    Dim SKUs As Object, currentCategory As String, currentRange As Range
    sql = "SELECT DISTINCT SKU, AlternativeSKU, Description, SalesPlanning FROM tblSKU WHERE Active = 'yes'"
    Set SKUs = GetRecordSet(sql)
    
    planningCategories.MoveFirst
    Do Until planningCategories.EOF
        currentCategory = planningCategories.Fields("SalesPlanning")
        SKUs.Filter = "SalesPlanning = " & Quot(currentCategory)
        SKUs.MoveFirst
        Set currentRange = wksTemplate.Names("C_" & CleanName(currentCategory)).RefersToRange
        If SKUs.RecordCount > 3 Then
            ResizeRngByRows currentRange.Cells(3, 1), _
                SKUs.RecordCount - 3
        End If
        Do Until SKUs.EOF
            currentRange.Cells(1 + SKUs.AbsolutePosition, 1).Offset(, -1).Value = SKUs.Fields("AlternativeSKU")
            currentRange.Cells(1 + SKUs.AbsolutePosition, 1) = SKUs.Fields("SKU") & " | " & SKUs.Fields("Description")
            currentRange.Cells(1 + SKUs.AbsolutePosition, 2) = currentCategory
            SKUs.MoveNext
        Loop
        planningCategories.MoveNext
    Loop
    
    'Get a list of customers
    sql = "SELECT DISTINCT Customer, CustomerName FROM tblCustomer WHERE PlanningCustomer IS NOT NULL"
    Set customers = GetRecordSet(sql)
    
    'sql = "SELECT * FROM tblVolumeSales WHERE PlanVersion = " & Quot(planversion)
    'Set volumeSales = GetRecordSet(sql)
    
    customers.MoveFirst
    Do Until customers.EOF
        Application.StatusBar = customers.Fields("CustomerName")
        DoEvents
        wkbReport.Sheets("Volume").Copy Before:=wksData
        ActiveSheet.name = "C_" & CleanName(customers.Fields("CustomerName"))
        ActiveSheet.[A1] = customers.Fields("Customer")
        customers.MoveNext
    Loop
    
    Debug.Print "Start " & Now
    Application.Calculate
    Debug.Print "Stop " & Now
    
    For Each wks In wkbReport.Sheets
        If Left(wks.name, 2) = "C_" Then
            wks.Range("rngSalesVolumeBody").Copy: wks.Range("rngSalesVolumeBody").PasteSpecial xlPasteValues
            For Each n In wks.Names
                If InStr(n.name, "B_") > 0 Then
                    n.RefersToRange.Value = n.RefersToRange.Value
                End If
            Next n
        End If
    Next wks

    Application.DisplayAlerts = False
    wkbReport.Sheets("Volume").Delete
    wksData.Visible = xlSheetHidden
    wkbReport.Sheets("DistributionKeys").Delete
    wkbReport.Sheets("SalesVolume").Delete
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    Exit Sub
    
    'Get reference volume
    Dim volumeDistribution As Object, j As Integer, customer As String
    wksTemplate.[B1] = customer
    sql = "SELECT SKUs,Volume,SegmentTotal FROM View_VolumeDistribution WHERE PlanningCustomer =" & Quot(customer) & " AND" & Quot(planversion)
    Set volumeDistribution = GetRecordSet(sql)
    volumeDistribution.MoveFirst
    Do Until volumeDistribution.EOF
        pos = FindPos(wksTemplate.Range("lstSKUs"), volumeDistribution.Fields("SKUs"), True)
        If pos <> 0 Then
            If volumeDistribution.Fields("SegmentTotal") = 0 Then
                wksTemplate.Cells(pos, 4) = 0
            Else
                wksTemplate.Cells(pos, 4) = Round(volumeDistribution.Fields("Volume") / volumeDistribution.Fields("SegmentTotal"), 3)
            End If
        End If
        volumeDistribution.MoveNext
    Loop
    
    'Get previous split (if proposeSplit = no)
    Dim volumeSplit As Object
    If proposeSplit = "no" Then
        sql = "SELECT Product, VolumeSplit, Period FROM tblFactsTemp WHERE PlanningCustomer =" & Quot(customer) & _
            " AND PlanVersion = " & Quot(planversion)
        Set volumeSplit = GetRecordSet(sql)
        volumeSplit.MoveFirst
        Do Until volumeSplit.EOF
            period = CInt(Right(volumeSplit.Fields("Period"), 2))
            pos = FindPos(wksTemplate.Range("lstSKUs"), volumeSplit.Fields("Product"), True)
            wksTemplate.Cells(pos, period + 4) = volumeSplit.Fields("VolumeSplit")
            volumeSplit.MoveNext
        Loop
    End If
    
    'Check if split adds up to 100%, only by initial split
    Dim splitTotal As Double, maxSplitRow As Integer, maxSplit As Double, row As Integer
    If proposeSplit = "yes" Then
        planningCategories.MoveFirst
        Do Until planningCategories.EOF
            splitTotal = 0
            maxSplit = 0
            maxSplitRow = 0
            row = 0
            currentCategory = planningCategories.Fields("SalesPlanning")
            Set currentRange = wksTemplate.Names("C_" & CleanName(currentCategory)).RefersToRange.Offset(, 2).Resize(, 1)
            For Each c In currentRange
                row = row + 1
                If Not IsEmpty(c.Offset(, -1)) Then
                    splitTotal = splitTotal + c.Value
                    If c.Value > maxSplit Then
                        maxSplit = c.Value
                        maxSplitRow = row
                    End If
                End If
            Next c
            If splitTotal <> 1 And splitTotal <> 0 Then
                currentRange(maxSplitRow) = currentRange(maxSplitRow) + (1 - splitTotal)
            End If
            'copy split to right
            For Each c In currentRange
                If Not IsEmpty(c.Offset(, -1)) Then
                    For j = 1 To 12
                        c.Offset(, j) = c.Value
                    Next j
                End If
            Next c
            planningCategories.MoveNext
        Loop
    End If
    Application.Calculation = xlCalculationAutomatic

End Sub
Function GetRecordSet(ByVal sql As String) As Object
    Dim rsData As Object, connection As Object
    
    Set connection = GetDBConnection()
    connection.Open
    Set rsData = CreateObject("ADODB.Recordset")
    With rsData
        .CursorLocation = 3 'adUseClient
        .CursorType = 1 'adOpenKeyset
        .LockType = 4 'adLockBatchOptimistic
        .Open sql, connection
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
Sub CopyRange(ByVal rng As Range)
    rng.EntireRow.Copy
    rng.EntireRow.Offset(rng.Rows.Count).Insert Shift:=xlShiftDown
    Application.CutCopyMode = False

On Error GoTo 0
End Sub
Function CleanName(ByVal str As String) As String
    str = Replace(str, " ", "")
    str = Replace(str, "&", "_")
    str = Replace(str, "'", "")
    str = Replace(str, "/", "_")
    CleanName = str
End Function
Sub ResizeRngByRows(singleCell As Range, extraRows As Long)
    singleCell.EntireRow.Copy
    singleCell.Resize(extraRows).EntireRow.Insert Shift:=xlShiftDown
    Application.CutCopyMode = False

On Error GoTo 0
End Sub
Function FindPos(ByRef rng As Range, ByVal id As String, findRow As Boolean) As Long
    Dim fnd As Variant
    Set fnd = rng.Find(What:=id, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False, SearchOrder:=xlByColumns)
    If Not fnd Is Nothing Then
        If findRow Then
            FindPos = fnd.row
        Else
            FindPos = fnd.Column
        End If
    Else
        FindPos = 0
    End If
End Function
Sub ResetReport()
    Dim wks As Worksheet
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    For Each wks In ActiveWorkbook.Sheets
        If Left(wks.name, 5) <> "XLRep" Then wks.Delete
    Next wks
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub AddDataSheets(ByVal name As String, ByVal sql As String, ByVal sBeforeSheet As String)
    Dim wks As Worksheet, vPlan As Object, i As Integer, vNames As Variant, categoryFilter As String, customerFilter As String
    'categoryFilter = GetPar(wksData.[A2], "ProfitCenter=")
    'customerFilter = GetPar(wksData.[A2], "CustomerName=")
    'vNames = Intersect(wksData.UsedRange, wksData.Range("5:5"))
    Set wks = ActiveWorkbook.Sheets.Add(Before:=ActiveWorkbook.Sheets(sBeforeSheet)): wks.name = name
    Set vPlan = GetRecordSet(sql)
    wks.[A1].CopyFromRecordset vPlan
    For i = 1 To vPlan.Fields.Count
        Names.Add name & "." & vPlan.Fields(i - 1).name, Intersect(wks.UsedRange, wks.Cells(1, i).EntireColumn).Resize(wks.UsedRange.Rows.Count)
    Next i
End Sub