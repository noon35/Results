Attribute VB_Name = "DataMining"
Option Compare Database
Option Explicit
Sub CreateResultsTable(ByVal tableName As String)

    Dim cat As New ADOX.Catalog
    Dim t As ADOX.Table
    Dim f As Field
    Dim tableExists As Boolean
    Dim col As ADOX.Column

    Dim c As ADODB.Connection

    'get and open connection
    Set c = GetConnection
    c.Open
    Set cat.ActiveConnection = c

    'check for table
    tableExists = False
    For Each t In cat.Tables
        If t.Name = tableName Then
            tableExists = True
            Exit For
        End If
    Next

    'drop the table if it exists
    If tableExists = True Then cat.Tables.Delete tableName

    Set t = New ADOX.Table

    Set col = New ADOX.Column

    With col
        .Name = "FileName"
        .Type = adVarChar
        .Attributes = adColNullable
        .DefinedSize = 500
    End With
    t.Columns.Append col

    t.Columns.Append CreateField("ErrorMessage", adVarChar, 500)
    t.Columns.Append CreateField("FileDate", adDBTimeStamp)
    t.Columns.Append CreateField("LastTrackerEntry", adDBTimeStamp)
    t.Columns.Append CreateField("Notes", adVarChar, 500)
    t.Columns.Append CreateField("InsertDate", adDBTimeStamp)


    t.Name = tableName
    cat.Tables.Append t

End Sub


Function CreateField(fieldName As String, datatype As ADOX.DataTypeEnum, Optional FieldSize = 0) As ADOX.Column
    Dim col As New ADOX.Column

    col.Name = fieldName
    col.Type = datatype
    If FieldSize > 0 Then col.DefinedSize = FieldSize
    col.Attributes = adColNullable

    Set CreateField = col

End Function


Function GetConnection() As ADODB.Connection

    Dim c As New ADODB.Connection
    'this is my local sql express
    'c.ConnectionString = "Provider=SQLNCLI11;Server=Shane-CB\SQLExpress;Database=Results;Integrated Security=SSPI;DataTypeCompatibility=80"
    'c.ConnectionString = "Provider=SQLNCLI11;Server=DC1;Database=Results;Integrated Security=SSPI;DataTypeCompatibility=80"
    c.ConnectionString = "Provider=SQLNCLI11;Server=DC1;Database=Results;Integrated Security=SSPI"
    Set GetConnection = c
End Function



Sub ExamineDatabases(ByVal resultsTable As String, ByVal SQL As String, Optional ByVal maxfileCount As Long = 50000)
    Dim f As File
    Dim fileName As String
    Dim fso As New FileSystemObject
    Dim db As Database
    Dim tableAlreadyCreated As Boolean
    Dim fd As Field
    Dim fileCount As Long
    Dim tempval
    Dim rsErrorLog As New ADODB.Recordset

    Dim dbResults As Database
    Dim rs As Recordset
    Dim rsResults As New ADODB.Recordset

    Dim conn As ADODB.Connection


    On Error GoTo egghead

    Set dbResults = CurrentDb

    Set conn = GetConnection
    conn.Open


    'tableAlreadyCreated = True
    
    For Each f In fso.GetFolder(GetDatabaseStorageFolder).Files
        'uncomment to skip databases
        'If f.Name < "shoreline.mdb" Then GoTo nextFile
        
        If f.Path Like "*.mdb" Or f.Path Like "*.accdb" Then
            fileCount = fileCount + 1
            fileName = Replace(f.Name, "'", "")

            Forms!form1!lblFileName.Caption = f.Name
            Forms!form1.Refresh

            Set db = OpenDatabase(f.Path)
            Set rs = db.OpenRecordset(SQL)
            

            If Not tableAlreadyCreated Then
                CreateResultsTable resultsTable
                tableAlreadyCreated = True
            End If

            
            UpdateSchema rs, resultsTable

            rsResults.Open resultsTable, conn, adOpenKeyset, adLockBatchOptimistic, adCmdTable


            Do While Not rs.EOF
                With rsResults
                    .AddNew
                    !fileName = fileName
                    !insertdate = Now()
                    For Each fd In rs.Fields
                        'On Error Resume Next
                        tempval = fd.Value
                        If fd.Type = dbDate And tempval < #1/1/1900# Then tempval = #1/1/1900#
                        rsResults.Fields(Replace(fd.Name, ".", "_")).Value = tempval
                    Next
                    On Error GoTo egghead
                    .Update
                End With


                rs.MoveNext
            Loop
            rsResults.UpdateBatch
            rsResults.Close

        End If

nextFile:
        Set db = Nothing
        If fileCount = maxfileCount Then Exit For
        Forms!form1.Refresh
        DoEvents
    Next


    Exit Sub
egghead:

    With rsErrorLog
        If rsResults.State = adStateOpen Then rsResults.Close
        'Set conn = GetConnection()
        'conn.Open
        .Open "Errors", conn, adOpenKeyset, adLockOptimistic
        .AddNew
            !fileName = f.Name
            !ErrorMessage = Err.Number & ": " & Err.Description
            !DigName = resultsTable
            !insertdate = Now()
        .Update
        .Close
    End With
    'Resume Next
    Resume nextFile

End Sub


Function GetLastTrackerEntry(db As Database)
    Dim rs As Recordset
    On Error GoTo egghead

    Set rs = db.OpenRecordset("SELECT Max(Date) as LastTrackerEntry FROM tblTracker WHERE NOT DatabasePath Like 'G:\Camps\*'")
    GetLastTrackerEntry = rs!LastTrackerEntry
    Set rs = Nothing

    Exit Function
egghead:
    GetLastTrackerEntry = DateSerial(1899, 12, 31)
    Exit Function

End Function



Function GetDatabaseStorageFolder() As String
    GetDatabaseStorageFolder = "I:\All Camps"
    'GetDatabaseStorageFolder = "I:\AllCamps2"
End Function


Sub UpdateSchema(source As dao.Recordset, targetTableName)
    Dim f As dao.Field
    Dim cat As ADOX.Catalog
    Dim col As ADOX.Column
    Dim t As ADOX.Table
    Dim conn As ADODB.Connection
    Dim comm As New ADODB.Command
    Dim colName As String



    Set conn = GetConnection
    conn.Open
    Set cat = New ADOX.Catalog
    Set cat.ActiveConnection = conn
    Set t = cat.Tables(targetTableName)

    For Each f In source.Fields
        colName = Replace(f.Name, ".", "_")
        If ColumnExists(t, colName) = False Then
            If f.Type = dbDate Then
                comm.ActiveConnection = conn
                comm.CommandText = "ALTER TABLE " & targetTableName & " ADD " & colName & " DateTime2 NULL"
                comm.CommandType = adCmdText
                comm.Execute
            ElseIf f.Type = dbText Then
                comm.ActiveConnection = conn
                comm.CommandText = "ALTER TABLE " & targetTableName & " ADD " & colName & " varchar(max) NULL"
                comm.CommandType = adCmdText
                comm.Execute
            Else

                Set col = New ADOX.Column
                col.Name = "[" & colName & "]"
                col.Attributes = adColNullable

                col.Type = GetSQLDataType(f.Type)
                If col.Type = adVarChar Then
                    col.DefinedSize = 7000
                End If
                t.Columns.Append col
            End If

        End If
    Next

End Sub

Function ColumnExists(t As ADOX.Table, colName As String) As Boolean
    Dim col As ADOX.Column
    Dim exists As Boolean

    exists = False
    For Each col In t.Columns
        If LCase(col.Name) = LCase(colName) Then
            exists = True
            Exit For
        End If
    Next

    ColumnExists = exists


End Function


Function GetSQLDataType(ByVal AccessDataType As dao.DataTypeEnum) As ADODB.DataTypeEnum

    Select Case AccessDataType
        Case dbText, dbMemo
            GetSQLDataType = adLongVarWChar
        Case dbDate
            GetSQLDataType = adDBTimeStamp
        Case dbCurrency
            GetSQLDataType = adCurrency
        Case dbLong, dbInteger
            GetSQLDataType = adInteger
        Case dbBoolean
            GetSQLDataType = dbInteger
        Case Else
            GetSQLDataType = AccessDataType
    End Select
End Function

