Option Explicit

Private Function Num(ByVal amt As Variant, Optional ByVal dec As Integer = 2) As String
    Num = Run("XLReporting.xlam!Num", amt, dec)
End Function

Private Function Quot(ByVal txt As Variant, Optional ByVal key As Boolean = False) As String
    Quot = Run("XLReporting.xlam!Quot", txt, key)
End Function

Private Function GetPar(ByVal txt As String, ByVal key As String, Optional ByVal sep As String = ";") As String
    GetPar = Run("XLReporting.xlam!GetPar", txt, key, sep)
End Function
    
Private Function GetSQL(ByVal sql As String) As Variant
    GetSQL = Run("XLReporting.xlam!GetSQL", sql)
End Function

Private Function RunSQL(ByVal sql As String) As Variant
    Run "XLReporting.xlam!RunSQL", sql
End Function

Public Function XLRep(ByVal From As String, ByVal Retrieve As String, ByVal Where As String) As Variant
    XLRep = Run("XLReporting.xlam!XLRep", From, Retrieve, Where)
End Function

Private Function XLImp(ByVal sql As String, Optional ByVal msg As String = "Working...") As Long
    XLImp = Run("XLReporting.xlam!XLImp", sql, msg)
End Function

Private Function GetPref(ByVal id As Integer) As String
    GetPref = Run("XLReporting.xlam!GetPref", id)
End Function

    Sub XLCode()

    Dim template As String, wksData As Worksheet, vNames As Variant, i As Long, wksReport As Worksheet
    Dim Period As Long, wkbTemplate As Workbook, wksValidatie As Worksheet, plan As String, ref1 As String, ref2 As String
    Dim wksLE As Worksheet, vPlan As Variant, entity As String, categoryFilter As String, customerFilter As String
    Application.EnableEvents = False: Application.Calculation = xlCalculationManual

    entity = GetPar([A2], "Country=")
    plan = GetPar([A2], "Plan Version=")
    ref1 = GetPar([A2], "Reference 1=")
    ref2 = GetPar([A2], "Reference 2=")
    categoryFilter = GetPar([A2], "ProfitCenter=")
    customerFilter = GetPar([A2], "CustomerName=")


    Set wksData = ActiveSheet

    template = GetPref(9) & "Templates\Template_PLByPeriod.xlsm"
    Set wkbTemplate = Application.Workbooks.Open(Filename:=template, ReadOnly:=True)
    wkbTemplate.Sheets("P&L By Period").Move Before:=wksData
    wkbTemplate.Sheets("Validatie").Move Before:=wksData

    Set wksReport = Sheets("P&L By Period")
    Set wksValidatie = Sheets("Validatie")

    If categoryFilter <> "" Then wksReport.[Z4].Value = "Selected Categories: " & categoryFilter
    If customerFilter <> "" Then wksReport.[Z5].Value = "Selected Customers: " & customerFilter

    AddDataSheets "plan", plan, entity, wksData
    AddDataSheets "ref1", ref1, entity, wksData
    AddDataSheets "ref2", ref2, entity, wksData

    vulValidatie wksValidatie
    'wksValidatie.Visible = xlSheetHidden: wksData.Visible = xlSheetHidden: wksLE.Visible = xlSheetHidden
    wksReport.Activate
    Names("ptr.Plan").RefersToRange.Value = plan
    Names("ptr.Ref1").RefersToRange.Value = ref1
    Names("ptr.Ref2").RefersToRange.Value = ref2
    
    Application.Calculate
    Application.EnableEvents = True: Application.Calculation = xlCalculationAutomatic

    End Sub
    Sub AddDataSheets(ByVal name As String, ByVal planversion As String, ByVal entity As String, ByRef wksData As Worksheet)
        Dim wks As Worksheet, vPlan As Variant, i As Integer, vNames As Variant, categoryFilter As String, customerFilter As String, sql As String
        categoryFilter = GetPar(wksData.[A2], "ProfitCenter=")
        customerFilter = GetPar(wksData.[A2], "CustomerName=")
        vNames = Intersect(wksData.UsedRange, wksData.Range("5:5"))
        Set wks = ActiveWorkbook.Sheets.Add(Before:=wksData): wks.name = name
        categoryFilter = Replace(categoryFilter, ",", "','")
        customerFilter = Replace(customerFilter, ",", "','")
        If categoryFilter = "" Then
            If customerFilter = "" Then
                sql = "SELECT * FROM View_CustomerCategoryPL WHERE Planversion = " & Quot(planversion) & " AND Country = " & Quot(entity)
            Else
                sql = "SELECT * FROM View_CustomerCategoryPL WHERE Planversion = " & Quot(planversion) & " AND Country = " & Quot(entity) & " AND ConditionCustomer IN ('" & customerFilter & "')"
            End If
        Else
            If customerFilter = "" Then
                sql = "SELECT * FROM View_CustomerCategoryPL WHERE Planversion = " & Quot(planversion) & " AND Country = " & Quot(entity) & " AND ProfitCenter IN ('" & categoryFilter & "')"
            Else
                sql = "SELECT * FROM View_CustomerCategoryPL WHERE Planversion = " & Quot(planversion) & " AND Country = " & Quot(entity) & " AND ProfitCenter  IN ('" & categoryFilter & "')" & " AND ConditionCustomer IN ('" & customerFilter & "')"
            End If
        End If
        vPlan = GetDBData(sql)
        wks.[A1].Resize(UBound(vPlan, 2) + 1, UBound(vPlan, 1) + 1) = Application.Transpose(vPlan)
        For i = 1 To UBound(vNames, 2)
            Names.Add name & "." & vNames(1, i), Intersect(wks.UsedRange, wks.Cells(1, i).EntireColumn).Resize(wks.UsedRange.Rows.Count)
        Next i
    End Sub
    Sub vulValidatie(ByRef wksValidatie As Worksheet)

        Dim vData As Variant, i As Integer, pgData As Variant, rng As Range, c As Range
        
        'Customer validation
        Set rng = wksValidatie.[B4]
        vData = GetDBData("SELECT DISTINCT CustomerName FROM tblCustomer WHERE CustomerName IS NOT NULL Order BY CustomerName ")
        rng.Resize(UBound(vData, 2) + 1) = Application.Transpose(vData)
        Names.Add "lst.Customer", rng.Offset(-2).Resize(UBound(vData, 2) + 3)

        'Brand validation
        Set rng = wksValidatie.[C4]
        vData = GetDBData("SELECT DISTINCT Brand FROM tblSKU WHERE Brand IS NOT NULL Order BY Brand")
        rng.Resize(UBound(vData, 2) + 1) = Application.Transpose(vData)
        Names.Add "lst.Brand", rng.Offset(-2).Resize(UBound(vData, 2) + 3)
        
        'SKU validation
        Set rng = wksValidatie.[D4]
        vData = GetDBData("SELECT DISTINCT ProfitCenter, Prdha2, EUProductHierarchy FROM tblSKU WHERE ProfitCenter IS NOT NULL AND Prdha2 IS NOT NULL AND EUProductHierarchy IS NOT NULL ORDER BY ProfitCenter, Prdha2, EUProductHierarchy")
        rng.Resize(UBound(vData, 2) + 1, UBound(vData, 1) + 1) = Application.Transpose(vData)
        Names.Add "tbl.SKUGroups", rng.Offset(-2).Resize(UBound(vData, 2) + 3, UBound(vData, 1) + 1)
        
        'Level1
        Set rng = wksValidatie.[G4]
        vData = GetDBData("SELECT DISTINCT ProfitCenter FROM tblSKU WHERE ProfitCenter IS NOT NULL ORDER BY ProfitCenter")
        rng.Resize(UBound(vData, 2) + 1) = Application.Transpose(vData)
        Names.Add "lst.Level1", rng.Offset(-2).Resize(UBound(vData, 2) + 3)
        
        'Level2
        Set rng = wksValidatie.[H4]
        vData = GetDBData("SELECT DISTINCT Prdha2 FROM tblSKU WHERE Prdha2 IS NOT NULL ORDER BY Prdha2")
        rng.Resize(UBound(vData, 2) + 1) = Application.Transpose(vData)
        Names.Add "lst.Level2", rng.Offset(-2).Resize(UBound(vData, 2) + 3)

        'Level3
        Set rng = wksValidatie.[I4]
        vData = GetDBData("SELECT DISTINCT EUProductHierarchy FROM tblSKU WHERE EUProductHierarchy IS NOT NULL ORDER BY EUProductHierarchy")
        rng.Resize(UBound(vData, 2) + 1) = Application.Transpose(vData)
        Names.Add "lst.Level3", rng.Offset(-2).Resize(UBound(vData, 2) + 3)
        
    End Sub

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
