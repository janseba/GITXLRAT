Sub XLCode()
    ImportGOS
    ImportPPP
    ImportTPRPromoShare
    ImportTPROnInvoice
End Sub

Sub ImportGOS()
    Dim wks As Worksheet, row As Long, rs As Object, periodFrom As Integer, period As Integer, periodTo As Integer
    Dim year As Long, planVersion As String, connection As Object
    Set wks = ActiveWorkbook.Sheets("GOS national")
    Set rs = GetEmptyRecordSet("SELECT * FROM tblFAP WHERE PlanVersion IS NULL")
    periodFrom = GetPeriodFrom()
    periodTo = GetPeriodTo()
    planVersion = GetPlanVersion()
    year = GetYear()
    With wks
        For row = 6 To .UsedRange.Rows.Count
            If Not IsEmpty(.Cells(row, 2)) Then
                For period = periodFrom To periodTo
                    If .Cells(row, period + 2) <> 0 Then
                        rs.AddNew
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("SalesCondition") = .Cells(row, 2)
                        rs.Fields("Period") = year * 100 + period
                        rs.Fields("FAPPerPiece") = .Cells(row, period + 2)
                    End If
                Next period
            End If
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblFAP WHERE PlanVersion = " & Quot(planVersion)
    rs.ActiveConnection = connection
    rs.UpdateBatch
End Sub
Sub ImportPPP()
    Dim wks As Worksheet, row As Long, rs As Object, periodFrom As Integer, period As Integer, periodTo As Integer
    Dim year As Long, planVersion As String, connection As Object, col As Integer
    Set wks = ActiveWorkbook.Sheets("Std PPP per customer")
    Set rs = GetEmptyRecordSet("SELECT * FROM tblPPP WHERE PlanVersion IS NULL")
    periodFrom = GetPeriodFrom()
    periodTo = GetPeriodTo()
    planVersion = GetPlanVersion()
    year = GetYear()
    With wks
        For row = 6 To .UsedRange.Rows.Count
            For col = 3 To .UsedRange.Columns.Count
                If Not IsEmpty(.Cells(row, 2)) Then
                    If Not IsEmpty(.Cells(4, col)) And .Cells(row, col) <> 0 And IsNumeric(.Cells(row, col)) Then
                        rs.AddNew
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("Customer") = .Cells(2, col)
                        rs.Fields("SalesConditionLevel") = .Cells(row, 2)
                        rs.Fields("Period") = .Cells(4, col)
                        rs.Fields("PPPPerPiece") = .Cells(row, col)
                    End If
                End If
            Next col
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblPPP WHERE PlanVersion = " & Quot(planVersion)
    rs.ActiveConnection = connection
    rs.UpdateBatch
End Sub
Sub ImportTPRPromoShare()
    Dim wks As Worksheet, row As Long, rs As Object, periodFrom As Integer, period As Integer, periodTo As Integer
    Dim year As Long, planVersion As String, connection As Object, col As Integer
    Set wks = ActiveWorkbook.Sheets("TPR-% promo-share")
    Set rs = GetEmptyRecordSet("SELECT * FROM tblTPRPromoShare WHERE PlanVersion IS NULL")
    periodFrom = GetPeriodFrom()
    periodTo = GetPeriodTo()
    planVersion = GetPlanVersion()
    year = GetYear()
    With wks
        For row = 6 To .UsedRange.Rows.Count
            For col = 3 To .UsedRange.Columns.Count
                If Not IsEmpty(.Cells(row, 2)) Then
                    If Not IsEmpty(.Cells(4, col)) And .Cells(row, col) <> 0 And IsNumeric(.Cells(row, col)) Then
                        rs.AddNew
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("Customer") = .Cells(2, col)
                        rs.Fields("SalesConditionLevel") = .Cells(row, 2)
                        rs.Fields("Period") = .Cells(4, col)
                        rs.Fields("PromoShare") = .Cells(row, col)
                    End If
                End If
            Next col
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblTPRPromoShare WHERE PlanVersion = " & Quot(planVersion)
    rs.ActiveConnection = connection
    rs.UpdateBatch
End Sub
Sub ImportTPROnInvoice()
    Dim wks As Worksheet, row As Long, rs As Object, periodFrom As Integer, period As Integer, periodTo As Integer
    Dim year As Long, planVersion As String, connection As Object, col As Integer
    Set wks = ActiveWorkbook.Sheets("TPR-on invoice")
    Set rs = GetEmptyRecordSet("SELECT * FROM tblTPROnInvoice WHERE PlanVersion IS NULL")
    periodFrom = GetPeriodFrom()
    periodTo = GetPeriodTo()
    planVersion = GetPlanVersion()
    year = GetYear()
    With wks
        For row = 6 To .UsedRange.Rows.Count
            For col = 3 To .UsedRange.Columns.Count
                If Not IsEmpty(.Cells(row, 2)) Then
                    If Not IsEmpty(.Cells(4, col)) And .Cells(row, col) <> 0 And IsNumeric(.Cells(row, col)) Then
                        rs.AddNew
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("Customer") = .Cells(2, col)
                        rs.Fields("SalesConditionLevel") = .Cells(row, 2)
                        rs.Fields("Period") = .Cells(4, col)
                        rs.Fields("TPROnInvoice") = .Cells(row, col)
                    End If
                End If
            Next col
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblTPROnInvoice WHERE PlanVersion = " & Quot(planVersion)
    rs.ActiveConnection = connection
    rs.UpdateBatch
End Sub
Sub ImportTPROffInvoice()
    Dim wks As Worksheet, row As Long, rs As Object, periodFrom As Integer, period As Integer, periodTo As Integer
    Dim year As Long, planVersion As String, connection As Object, col As Integer
    Set wks = ActiveWorkbook.Sheets("TPR-on invoice")
    Set rs = GetEmptyRecordSet("SELECT * FROM tblTPROffInvoice WHERE PlanVersion IS NULL")
    periodFrom = GetPeriodFrom()
    periodTo = GetPeriodTo()
    planVersion = GetPlanVersion()
    year = GetYear()
    With wks
        For row = 6 To .UsedRange.Rows.Count
            For col = 3 To .UsedRange.Columns.Count
                If Not IsEmpty(.Cells(row, 2)) Then
                    If Not IsEmpty(.Cells(4, col)) And .Cells(row, col) <> 0 And IsNumeric(.Cells(row, col)) Then
                        rs.AddNew
                        rs.Fields("PlanVersion") = planVersion
                        rs.Fields("Customer") = .Cells(2, col)
                        rs.Fields("SalesConditionLevel") = .Cells(row, 2)
                        rs.Fields("Period") = .Cells(4, col)
                        rs.Fields("TPROffInvoice") = .Cells(row, col)
                    End If
                End If
            Next col
        Next row
    End With
    Set connection = GetDBConnection: connection.Open
    connection.Execute "DELETE FROM tblTPROffInvoice WHERE PlanVersion = " & Quot(planVersion)
    rs.ActiveConnection = connection
    rs.UpdateBatch
End Sub
Function GetPlanVersion() As String
    GetPlanVersion = GetPar(ActiveWorkbook.Sheets("What do do").[A1], "Plan Version=")
End Function
Function GetPeriodFrom() As Integer
    Dim planVersion As String
    planVersion = GetPar(ActiveWorkbook.Sheets("What do do").[A1], "Plan Version=")
    GetPeriodFrom = Right(GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion)), 2)
End Function
Function GetYear() As Integer
    Dim planVersion As String
    planVersion = GetPar(ActiveWorkbook.Sheets("What do do").[A1], "Plan Version=")
    GetYear = Left(GetSQL("SELECT FromPeriod FROM Sources WHERE Source = " & Quot(planVersion)), 4)
End Function
Function GetPeriodTo() As Integer
    Dim planVersion As String
    planVersion = GetPar(ActiveWorkbook.Sheets("What do do").[A1], "Plan Version=")
    GetPeriodTo = Right(GetSQL("SELECT ToPeriod FROM Sources WHERE Source = " & Quot(planVersion)), 2)
End Function

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
