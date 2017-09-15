Attribute VB_Name = "mod_Temp_Tables"
Option Compare Database
Option Explicit

Public Function UpdateTempTable(Tablename As String, TableQueryOrSQL As String, _
                                Optional InCurrentDb As Boolean = False, _
                                Optional ValidMinutes As Integer = 0, _
                                Optional PK_Field As Variant = Null) As Boolean

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strTempFile As String
    Dim strSQL As String, strMsg As String
    Dim strError As String
    Dim varExtDB As Variant
    Dim intMousePointer As Integer
    
    On Error GoTo ProcError
    
    intMousePointer = Screen.MousePointer
    DoCmd.Hourglass True
    
    'set the default return value for the function
    UpdateTempTable = False

    'If the table exists in the local db, then check to see how long it has been since
    'the temp table was created.  If more than ValidMinutes, then drop the table
    strError = "Deleting existing table in currendtb"
    If TableExists(Tablename) Then
        If DateDiff("n", CurrentDb.TableDefs(Tablename).Properties("DateCreated"), Now()) <= ValidMinutes Then
            UpdateTempTable = True
            GoTo ProcExit
        Else
            DropTable Tablename
        End If
    End If
    
    'If the temp table is supposed to be created in an external database, then make sure it exists
    'Make sure the temp database exists in the same folder as the current project
    If InCurrentDb = False Then
        strError = "Creating the temp database"
        strTempFile = CurrentProject.Path & "\" _
                    & Left(CurrentProject.Name, InStrRev(CurrentProject.Name, ".") - 1) _
                    & "_Temp.accdb"
        If FileExists(strTempFile) = False Then
            DBEngine.CreateDatabase strTempFile, dbLangGeneral, dbVersion120
        End If
    
        'Check to see whether the table already exists in the temp.accdb file.  If so, delete it
        strError = "Dropping the table in the temp database"
        Set db = DBEngine.OpenDatabase(strTempFile)
        If TableExists(Tablename, db) = True Then
            db.Execute "Drop Table [" & Tablename & "]", dbFailOnError
        End If
        Set db = Nothing
    End If
    
    'Define the SQL to insert the records from TableQueryOrSQL into the temp table
    strSQL = "SELECT zz.* INTO [" & Tablename & "] "
    If InCurrentDb = False Then
        strSQL = strSQL & "IN " & Wrap(strTempFile)
    End If
    strSQL = strSQL & " FROM "
           
    'If the TableQueryOrSQL contains a SELECT INTO statement, then display message and exit
    'If the TableQueryOrSQL contains a SELECT statement, then wrap it in () as a subquery
    'If the TableQueryOrSQL is a query or table then just insert the value of TableQueryOrSQL
    'in the SQL string.
    'However, if the table exists and it is a SharePoint list (database field in mSysObjects
    'contains http:// or https:// then ignore) then ignore the table
    If InStr(TableQueryOrSQL, "SELECT") > 0 And InStr(TableQueryOrSQL, "INTO") > 0 Then
        strMsg = "Cannot pass a Maketable or Append query to this function"
        MsgBox strMsg, vbOKOnly, "Invalid argument for TableQueryOrSQL"
        strSQL = ""
    ElseIf InStr(TableQueryOrSQL, "SELECT") = 1 And InStr(TableQueryOrSQL, "INTO") = 0 Then
        strSQL = strSQL & "(" & TableQueryOrSQL & ") as zz"
    ElseIf QueryExists(TableQueryOrSQL) Then
        strSQL = strSQL & TableQueryOrSQL & " as zz"
    ElseIf TableExists(TableQueryOrSQL) Then
        varExtDB = DLookup("Database", "mSysObjects", "[Name] = " & Wrap(TableQueryOrSQL))
        If InStr(Nz(varExtDB, ""), "http") Then
            strMsg = "Unable to use Sharepoint list names directly because of potential " _
                   & "field type conflicts with earlier versions of Access.  To include " _
                   & "as SharePoint list in this function, pass it a SELECT query that " _
                   & "includes the specific fields to be used from the list."
            MsgBox strMsg, vbOKOnly, "Invalid argument for TableQueryorsQL"
            strSQL = ""
        Else
            strSQL = strSQL & TableQueryOrSQL & " as zz"
        End If
    Else
        MsgBox "Invalid syntax for the SQL string", vbOKOnly, "Invalid argument for TableQueryOrSQL"
        strSQL = ""
    End If

    'If strSQL = "" then exit the function
    If strSQL = "" Then GoTo ProcExit
            
    'Otherwise, execute the SQL to create an empty table in temp.accdb
    strError = "Writing data to the temp table in the temp db"
    CurrentDb.Execute strSQL, dbFailOnError
            
    'If a primary key field was defined, then alter the structure of the table
    If IsNull(PK_Field) = False Then
        If InCurrentDb = True Then
            Set db = CurrentDb
        Else
            Set db = DBEngine.OpenDatabase(strTempFile)
        End If
        strSQL = "ALTER TABLE [" & Tablename & "] " _
               & "ALTER COLUMN [" & PK_Field & "] Long " _
               & "CONSTRAINT PrimaryKey PRIMARY KEY;"
        db.Execute strSQL, dbFailOnError
        Set db = Nothing
    End If
    
    'If the temp table was created in an external db then link the table to the current project
    If InCurrentDb = False Then
        strError = "Linking table to the current database"
        
        Set tdf = CurrentDb.CreateTableDef(Tablename)
        tdf.Connect = ";DATABASE=" & strTempFile
        tdf.SourceTableName = Tablename
        CurrentDb.TableDefs.Append tdf
        'DisplayNavPane (False)
    End If
    CurrentDb.TableDefs.Refresh
    
    UpdateTempTable = True
ProcExit:
    If Not db Is Nothing Then Set db = Nothing
    DoCmd.Hourglass False
    Screen.MousePointer = intMousePointer
    
    Exit Function
ProcError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbOKOnly, "UpdateTempTable error"
    Debug.Print "UpdateTempTable error", Err.Number, Err.Description
    Resume ProcExit
    
End Function

Public Sub DropTable(Tablename As String, Optional DropFromTemp As Boolean = False)

    Dim db As DAO.database
    Dim strTempFile As String
    Dim strError As String
    Dim intMousePointer As Integer
    
    On Error GoTo ProcError
    
    intMousePointer = Screen.MousePointer
    DoCmd.Hourglass True
    
    If DropFromTemp Then
        strError = "Verify the temp database"
        strTempFile = CurrentProject.Path & "\" _
                    & Left(CurrentProject.name, InStrRev(CurrentProject.name, ".") - 1) _
                    & "_Temp.accdb"
        If FileExists(strTempFile) = True Then
            strError = "Dropping the table in the temp database"
            Set db = DBEngine.OpenDatabase(strTempFile)
            If TableExists(Tablename, db) = True Then
                db.Execute "Drop Table [" & Tablename & "]", dbFailOnError
            End If
            Set db = Nothing
        End If
    End If
    
    strError = "Dropping the table from current database"
    If TableExists(Tablename) Then
        DoCmd.DeleteObject acTable, Tablename
    End If

ProcExit:
    If Not db Is Nothing Then Set db = Nothing
    DoCmd.Hourglass False
    Screen.MousePointer = intMousePointer
    
    Exit Sub
ProcError:
    MsgBox Err.Number & vbCrLf & Err.Description, vbOKOnly, "DropTable error"
    Debug.Print "DropTable error", Err.Number, Err.Description
    Resume ProcExit
    
End Sub

Public Function TableExists(Tablename As String, Optional db As DAO.Database) As Boolean

    Dim intFields As Integer
    Dim ReleaseDB As Boolean
    
    On Error GoTo ProcError
    
    'The default database is the currentdb, but if one is passed, use it
    If db Is Nothing Then
        Set db = CurrentDb
        ReleaseDB = True
    End If
    
    'If the table exists, then the next line will determine how many fields are in the table
    'If it doesn't exist, then this will raise an error
    intFields = db.TableDefs(Tablename).Fields.Count
    TableExists = True
    
ProcExit:
    If ReleaseDB Then Set db = Nothing
    Exit Function
    
ProcError:
    TableExists = False
    Resume ProcExit
    
End Function

Public Function QueryExists(QueryName As String, Optional db As DAO.Database) As Boolean

    Dim intFields As Integer
    Dim ReleaseDB As Boolean
    
    On Error GoTo ProcError
    
    'The default database is the currentdb, but if one is passed, use it
    If db Is Nothing Then
        Set db = CurrentDb
        ReleaseDB = True
    End If
    
    'If the query exists, then the next line will determine how many fields are in the table
    'If it doesn't exist, then this will raise an error
    intFields = db.QueryDefs(QueryName).Fields.Count
    QueryExists = True
    
ProcExit:
    If ReleaseDB Then Set db = Nothing
    Exit Function
    
ProcError:
    QueryExists = False
    Resume ProcExit
    
End Function

Public Function FileExists(FileName As String) As Boolean

    FileExists = Len(Dir(FileName, vbNormal + vbHidden + vbSystem + vbReadOnly)) > 0
    
End Function

Public Function Wrap(WrapWhat As Variant, Optional WrapWith As String = """") As String

    'Created by Dale Fye, 2000-03-10
    
    'This function is used to wrap some value with some sort of wrapping character.
    
    'Accepts a variant and wraps that with whatever character or group of characters
    'are passed in the optional WrapWith argument.
    
    'It also replaces all values equal to the WrapWith text with duplicates of that character
    'which enables wrapping a text string that contains quotes.
    
    'If the WrapWhat value is NULL, then the function returns a empty string wrapped in quotes
    'I generally use this to wrap text in quotes or date values in the #
    
    Wrap = WrapWith & Replace(WrapWhat & "", WrapWith, WrapWith & WrapWith) & WrapWith
    
End Function

