Sub XLCode()

Dim template As String, wksData As Worksheet, vNames As Variant, i As Long, wksReport As Worksheet
Dim period As Long, wkbTemplate As Workbook, wksValidatie As Worksheet, plan As String, ref1 As String, ref2 As String
Dim wksLE As Worksheet, vPlan As Variant, entity As String, categoryFilter As String, customerFilter As String
Dim planYear As String, ref1Year As String
Application.EnableEvents = False: Application.Calculation = xlCalculationManual

entity = GetPar([A2], "Country=")
plan = GetPar([A2], "Plan Version=")
ref1 = GetPar([A2], "Reference 1=")
categoryFilter = GetPar([A2], "ProfitCenter=")
customerFilter = GetPar([A2], "CustomerName=")
planYear = Left(GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(plan)), 4)
ref1Year = Left(GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(ref1)), 4)

Set wksData = ActiveSheet

template = GetPref(9) & "Templates\Template_CategoryPL.xlsx"
Set wkbTemplate = Application.Workbooks.Open(Filename:=template, ReadOnly:=True)
wkbTemplate.Sheets("P&L By Category").Move Before:=wksData
wkbTemplate.Sheets("Validation").Move Before:=wksData

Set wksReport = Sheets("P&L By Category")
Set wksValidatie = Sheets("Validation")

wksReport.Range("ptr.Plan").Value = plan
wksReport.Range("ptr.Ref1").Value = ref1
wksReport.Range("ptr.PlanYear").Value = planYear
wksReport.Range("ptr.Ref1Year").Value = ref1Year
wksReport.Range("ptr.PlanningCustomer").Value = "*"
wksReport.Range("ptr.CustomerName").Value = "*"

If categoryFilter <> "" Then wksReport.[Z4].Value = "Selected Categories: " & categoryFilter
If customerFilter <> "" Then wksReport.[Z5].Value = "Selected Customers: " & customerFilter

AddDataSheets "plan", plan, entity, wksData
AddDataSheets "ref1", ref1, entity, wksData

vulValidatie wksValidatie

'wksValidatie.Visible = xlSheetHidden: wksData.Visible = xlSheetHidden: wksLE.Visible = xlSheetHidden
wksReport.Activate
Application.Calculate
Application.EnableEvents = True: Application.Calculation = xlCalculationAutomatic

End Sub
Sub vulValidatie(ByRef wksValidatie As Worksheet)

    Dim vData As Variant, i As Integer, pgData As Variant, rng As Range, c As Range, rangeNames() As Variant
    
    'Period validation
    Set rng = wksValidatie.[A2]
    rangeNames = Array("plan.Month", "ref1.Month")
    vData = GetValidation(rangeNames)
    rng.Resize(UBound(vData) + 1) = Application.Transpose(vData)
    Names.Add "lst.Period", rng.Resize(UBound(vData) + 1)
    
    'Condition customer validation
    Set rng = wksValidatie.[B3]
    rangeNames = Array("plan.PlanningCustomer", "ref1.PlanningCustomer")
    vData = GetValidation(rangeNames)
    rng.Resize(UBound(vData) + 1) = Application.Transpose(vData)
    Names.Add "lst.PlanningCustomer", rng.Offset(-1).Resize(UBound(vData) + 2)
    
    'Customer name validation
    Set rng = wksValidatie.[C3]
    rangeNames = Array("plan.CustomerName", "ref1.CustomerName")
    vData = GetValidation(rangeNames)
    rng.Resize(UBound(vData) + 1) = Application.Transpose(vData)
    Names.Add "lst.CustomerName", rng.Offset(-1).Resize(UBound(vData) + 2)
    
End Sub

Sub AddDataSheets(ByVal name As String, ByVal planversion As String, ByVal entity As String, ByRef wksData As Worksheet)
    Dim wks As Worksheet, vPlan As Variant, i As Integer, vNames As Variant, PlanningCustomer As String, customer As String, sql As String
    Dim period As Variant, monthFrom As Integer, monthTo As Integer
    period = GetPar(wksData.[A2], "Period=")
    If period <> "" Then
        period = Split(period, "-")
        monthFrom = CInt(period(0))
        monthTo = CInt(period(1))
    Else
        monthFrom = 1
        monthTo = 12
    End If
    ActiveWorkbook.Sheets("P&L By Category").Range("ptr.PeriodFrom") = monthFrom
    ActiveWorkbook.Sheets("P&L By Category").Range("ptr.PeriodTo") = monthTo
    PlanningCustomer = GetPar(wksData.[A2], "Planning Customer=")
    customer = GetPar(wksData.[A2], "Customer Name=")
    vNames = Intersect(wksData.UsedRange, wksData.Range("5:5"))
    Set wks = ActiveWorkbook.Sheets.Add(Before:=wksData): wks.name = name
    PlanningCustomer = Replace(PlanningCustomer, ",", "','")
    customer = Replace(customer, ",", "','")
    If PlanningCustomer = "" Then
        If customer = "" Then
            sql = "SELECT * FROM View_CustomerCategoryPL WHERE Planversion = " & Quot(planversion) & " AND Country = " & Quot(entity) & " AND Month BETWEEN " & monthFrom & " AND " & monthTo
        Else
            sql = "SELECT * FROM View_CustomerCategoryPL WHERE Planversion = " & Quot(planversion) & " AND Country = " & Quot(entity) & " AND Month BETWEEN " & monthFrom & " AND " & monthTo & " AND CustomerName IN ('" & customer & "')"
        End If
    Else
        If customer = "" Then
            sql = "SELECT * FROM View_CustomerCategoryPL WHERE Planversion = " & Quot(planversion) & " AND Country = " & Quot(entity) & " AND Month BETWEEN " & monthFrom & " AND " & monthTo & " AND PlanningCustomer IN ('" & PlanningCustomer & "')"
        Else
            sql = "SELECT * FROM View_CustomerCategoryPL WHERE Planversion = " & Quot(planversion) & " AND Country = " & Quot(entity) & " AND Month BETWEEN " & monthFrom & " AND " & monthTo & " AND PlanningCustomer  IN ('" & PlanningCustomer & "')" & " AND CustomerName IN ('" & customer & "')"
        End If
    End If
    vPlan = GetDBData(sql)
    If IsArray(vPlan) Then
        wks.[A1].Resize(UBound(vPlan, 2) + 1, UBound(vPlan, 1) + 1) = Application.Transpose(vPlan)
        For i = 1 To UBound(vNames, 2)
            Names.Add name & "." & vNames(1, i), Intersect(wks.UsedRange, wks.Cells(1, i).EntireColumn).Resize(wks.UsedRange.Rows.Count)
        Next i
    Else
        For i = 1 To UBound(vNames, 2)
            Names.Add name & "." & vNames(1, i), wks.Cells(1, i)
        Next i
    End If

End Sub
Function GetValidation(ByRef rangeNames() As Variant) As Variant
    Dim dctValues As Object, c As Range, i As Integer, rng As Range
    Set dctValues = CreateObject("Scripting.Dictionary")
    For i = 0 To UBound(rangeNames)
        Set rng = Range(rangeNames(i))
        For Each c In rng
            If Not dctValues.Exists(c.Value) Then dctValues.Add c.Value, c.Value
        Next c
    Next i
    GetValidation = dctValues.Keys
End Function
Function GetDBData(ByVal sql As String) As Variant
    Dim pw As String, connectionString As String, dbConnection As Object, rst As Object, vResult As Variant, sDbName As String
    
    pw = "xlsysjs14"
    sDbName = GetSQL("SELECT ParValue FROM XLControl WHERE Code = 'Database'")
    connectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & GetPref(9) & sDbName & "; Jet OLEDB:Database password=" & pw
    Set dbConnection = CreateObject("ADODB.Connection")
    dbConnection.Open connectionString
    Set rst = CreateObject("ADODB.Recordset")
    rst.Open sql, dbConnection, 3, 1
    If Not rst.EOF Then
        vResult = rst.GetRows
    Else
        vResult = ""
    End If
    dbConnection.Close
    Set dbConnection = Nothing
    GetDBData = vResult
End Function
Sub RestartSheet()
    Dim n As name
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    For Each n In ActiveWorkbook.Names
        On Error Resume Next
        n.Delete
        On Error GoTo 0
    Next n
        
    Dim s As Worksheet
    For Each s In ActiveWorkbook.Sheets
        If Left(s.name, 3) <> "XLR" Then
            Application.DisplayAlerts = False
            s.Delete
            Application.DisplayAlerts = True
        End If
    Next s
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Function CleanName(ByVal str As String) As String
    str = Replace(str, " ", "")
    str = Replace(str, "&", "_")
    str = Replace(str, "'", "")
    str = Replace(str, "/", "_")
    str = Replace(str, "-", "")
    CleanName = str
End Function
