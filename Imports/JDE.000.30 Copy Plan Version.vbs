Sub XLCode()
Dim sql As String, fromPlan As String, toPlan As String, country As String
fromPlan = GetPar([A1], "From Plan Version=")
toPlan = GetPar([A1], "To Plan Version=")
If GetSQL("SELECT Locked FROM sources WHERE Source = " & Quot(toPlan)) = "y" Then
    XLImp "ERROR", "The plan version has been locked for input": Exit Sub
End If
country = GetPar([A1], "Country=")
sql = "DELETE FROM tblFacts WHERE PlanVersion = " & Quot(toPlan) & " AND Country = " & Quot(country)
XLImp sql, "Delete old plan data ..."

sql = GenerateCopyFunction(fromPlan, toPlan)
XLImp sql, "Copy plan ..."

End Sub

Function GenerateCopyFunction(fromPlan as String, toPlan as String) As String
    Dim cn As Object, rs As Object, rs2 As Object, f As Variant, fieldNames As String, sql As String
    Set cn = GetDBConnection
    cn.Open
    
    Set rs = GetEmptyRecordSet("SELECT * FROM tblFacts WHERE PlanVersion IS NULL")
    
    For Each f In rs.Fields
        fieldNames = fieldNames & ",[" & f.Name & "]"
    Next f
    fieldNames = Right(fieldNames, Len(fieldNames) - 1)
    
    sql = "INSERT INTO tblFacts (" & fieldNames & ") SELECT "
    
    fieldNames = ""
    For Each f In rs.Fields
        If f.Name = "PlanVersion" Then
            fieldNames = fieldNames & "," & Quot(toPlan)
        Else
            fieldNames = fieldNames & ",[" & f.Name & "]"
        End If
    Next f
    fieldNames = Right(fieldNames, Len(fieldNames) - 1)
    
    sql = sql & fieldNames & "FROM tblFacts WHERE PlanVersion = " & Quot(fromPlan)
        
    Set cn = Nothing
    Set rs = Nothing
    
    GenerateCopyFunction = sql
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