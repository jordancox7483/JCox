Attribute VB_Name = "ImportAndParseConversionData"
Option Compare Database
Public Sub BuildQuery()
MsgBox "Remember to setup the Code Conversion tables", vbOKOnly, "Reminder"

 ImportTables
 
 BuildEmpl
 BuildEdd
 BuildEETO
 BuildEREC
'BuildADIS No Data
 BuildYTD
 
 DoCmd.OpenQuery "ConvertEarningsAndDeductions", acViewNormal
 DoCmd.OpenQuery "ConvertEarningsAndDeductionsYTD", acViewNormal
 
 
 DoCmd.RunMacro "padeeid"
 
 CheckColumns
End Sub
Public Function GetFileName(vPath) As String
    Dim x As Integer
    Dim vFileName
    x = 5 ' Start at 5 b/c file should be as least x.ext
    While 1 < Len(vPath)
        'Check for "\" to determine where file name stops
        vFileName = Right(vPath, x)
        If InStr(1, vFileName, "\") Then
            GetFileName = Right(vPath, x - 1)
            Exit Function 'found file name
        End If
        x = x + 1
    Wend
End Function
Public Function GetFilePath(vPath) As String
    Dim x As Integer
    Dim iTmpFileName As Variant
    Dim vFilepath
    x = 5 ' Start at 5 b/c file should be as least x.ext
    While 1 < Len(vPath)
        'Check for "\" to determine where file name stops
        vFilepath = Right(vPath, x)
        If InStr(1, vFilepath, "\") Then
            iTmpFileName = Replace(vPath, Right(vPath, x - 1), "")
            GetFilePath = iTmpFileName
            Exit Function 'found file name
        End If
        x = x + 1
    Wend
End Function


Public Sub ImportTables()
    Dim OpenDlg As FileDialog
    Dim vSelectedFile As Variant
    Dim db As Database
    Dim SQL, vOverwrite, vNewPath, iDate, iTable, MultiComp, iLen As String
    Dim vAppend As Boolean
    Dim File(1 To 11) As String
    
    

    On Error GoTo ErrTrap

 
    
    vAppend = False
    Set OpenDlg = Application.FileDialog(msoFileDialogFilePicker)
    Set db = CurrentDb
    vPath = "C:\UltiPro\"


    DoCmd.SetWarnings False
    OpenDlg.AllowMultiSelect = True
    'set OpenDlg.Filters (0)
    'Add a filter that includes only .rdb and make it the first item in the list.
    OpenDlg.Filters.Add "PayChex Conversion Files", "*.rdb", 1
    OpenDlg.Title = "Select Conversion Data"
    
    OpenDlg.InitialFileName = vPath
    
    OpenDlg.InitialView = msoFileDialogViewDetails
    'OpenDlg.Show
     If OpenDlg.Show = -1 Then
    
        For Each vSelectedFile In OpenDlg.SelectedItems
           
       'Determine table name from file and import the file (changing to .txt in the process)
                If Len(GetFileName(vSelectedFile)) = 11 Then
                    iLen = 3
                Else
                    iLen = 4
                End If
            iTable = Mid(GetFileName(vSelectedFile), 5, iLen)
            db.TableDefs.Delete iTable
            FileCopy vSelectedFile, GetFilePath(vSelectedFile) & Replace(GetFileName(vSelectedFile), ".rdb", ".txt")
            vSelectedFile = Replace(vSelectedFile, ".rdb", ".txt")
            DoCmd.TransferText acImportFixed, iTable, iTable, vSelectedFile, False, ""
            DoEvents
            'SQL = "DELETE * FROM " & iTable & " WHERE [EEID] is Null"
            'db.Execute SQL, dbFailOnError
            Kill vSelectedFile
        Next
    Else
        MsgBox "You clicked Cancel in the file dialog box."
        db.Close
        Set db = Nothing
        Exit Sub
    End If
    DoCmd.SetWarnings True
    db.Close
    Set db = Nothing
        MsgBox "File import completed!", , "Import Process"
        Exit Sub
CleanClose:
    MsgBox "Process Cancelled", vbInformation + vbOKOnly, "Import Multiple Companies"
    db.Close
    Set db = Nothing
    Exit Sub

ErrTrap:
    If Err.Number = 3265 Then
        Resume Next
    End If
    If Err.Number = 3011 Then
        MsgBox Err.Number
        MsgBox Err.Description
    End If
    MsgBox Err.Number
    MsgBox Err.Description
    





End Sub

Public Sub BuildEmpl()
    Dim dbs As Database
    Dim Standard As Boolean
    Dim SQL, Inbox, InputMsg, TopValue, Source As String
    Dim Tablename(1 To 50) As String
    Dim k, x, y, z, Numb, SP, LP As Integer
    Dim RS, rsAnal As DAO.Recordset
    Dim PAYCHEX As Recordset
    Dim tdf As TableDef
    Dim fld As Field
    Dim Newtdf, tabletdf  As TableDef
    Dim test As String
    Dim CHOOSESQL, CHOOSESQL2, INSERTSQL, SELECTSQL As String
    Inbox = ""
    'Set Source Table Here:
    Source = "EMPL"
    
    'Different RecordCodes
                    Tablename(1) = "EmployeeData"
                    
                    
    'Because I want to automate I am hardcodinc vbYes as the answer
    InputMsg = vbYes
    Select Case InputMsg
        Case Is = vbCancel
            Exit Sub
 'When choosing to reload a specific table (no in msgbox):
        Case Is = vbNo
EnterTableName:
            Inbox = InputBox("Enter the Table Name", "Table", Inbox)
            If Inbox = "" Then
                Exit Sub
            Else
                Set dbs = CurrentDb
                k = 1
                Tablename(k) = Inbox
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name.  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    If x > 6 Then
                        LP = tdf(x).Size
                        If x = Numb - 1 Then
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                        Else
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                        End If
                        SP = SP + LP
                    Else
                        'Do Nothing because the first 7 field names are hard coded
                    End If
                  x = x + 1
                Wend
                SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,2) = LEFT('" & Tablename(k) & "',2)"
                
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            End If
   'When loading all tables (yes in msgbox):
        Case Is = vbYes
            Set dbs = CurrentDb
            k = 1
            For k = k To 1
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name, " & Tablename(k) & ".  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    
                    LP = tdf(x).Size
                    If x = Numb - 1 Then
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                    Else
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                    End If
                    SP = SP + LP

                  x = x + 1
          
                Wend
                
                If Mid(Tablename(k), 1, 2) = "3X" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,3) = LEFT('" & Tablename(k) & "',3)"
                    
                ElseIf Mid(Tablename(k), 1, 2) = "61" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE (MID([" & Source & "].Data,14,2) + SPACE(1) + MID([" & Source & "].Data,80,1)) = LEFT('" & Tablename(k) & "',4) "
                
                Else
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] "
                End If
                        
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            Next k
        End Select
        

CleanClose:
    Set Newtdf = Nothing
    Set tdf = Nothing
    dbs.Close
    Set dbs = Nothing
    InputMsg = vbNo
    If InputMsg = vbYes Then GoTo EnterTableName
    Exit Sub
        
TableExists:
        Select Case Err.Number
            Case Is = 3010 '3010 Table already exists.
                'Delete existing table
                DoCmd.DeleteObject acTable, Tablename(k) & "_Analysis"
                Resume
            Case Is = 3146 '3146 ODBC TIME OUT
                MsgBox Err.Number
                MsgBox Err.Description
                InputMsg = MsgBox("ODBC TimeOut.  Would you like to contiunue the this process?", vbYesNo + vbCritical, "Timeout Error")
                If InputMsg = vbYes Then Resume
                GoTo CleanClose
            Case Is = 3265 '3265 Standard Table does not exists.
                k = k + 1
                Resume
            Case Else
                MsgBox Tablename(k)
                MsgBox Err.Number
                MsgBox Err.Description
                GoTo CleanClose
        End Select
End Sub

Public Sub BuildEdd()
    Dim dbs As Database
    Dim Standard As Boolean
    Dim SQL, Inbox, InputMsg, TopValue, Source As String
    Dim Tablename(1 To 50) As String
    Dim k, x, y, z, Numb, SP, LP As Integer
    Dim RS, rsAnal As DAO.Recordset
    Dim PAYCHEX As Recordset
    Dim tdf As TableDef
    Dim fld As Field
    Dim Newtdf, tabletdf  As TableDef
    Dim test As String
    Dim CHOOSESQL, CHOOSESQL2, INSERTSQL, SELECTSQLedd As String
    Inbox = ""
    'Set Source Table Here:
    Source = "EDD"
    
    'Different RecordCodes
                    Tablename(1) = "EmployeeDirectDeposit"
                    
                            
           
    InputMsg = vbYes
    Select Case InputMsg
        Case Is = vbCancel
            Exit Sub
 'When choosing to reload a specific table (no in msgbox):
        Case Is = vbNo
EnterTableName:
            Inbox = InputBox("Enter the Table Name", "Table", Inbox)
            If Inbox = "" Then
                Exit Sub
            Else
                Set dbs = CurrentDb
                k = 1
                Tablename(k) = Inbox
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name.  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    If x > 6 Then
                        LP = tdf(x).Size
                        If x = Numb - 1 Then
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                        Else
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                        End If
                        SP = SP + LP
                    Else
                        'Do Nothing because the first 7 field names are hard coded
                    End If
                  x = x + 1
                Wend
                SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,2) = LEFT('" & Tablename(k) & "',2)"
                
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            End If
   'When loading all tables (yes in msgbox):
        Case Is = vbYes
            Set dbs = CurrentDb
            k = 1
            For k = k To 1
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name, " & Tablename(k) & ".  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    
                    LP = tdf(x).Size
                    If x = Numb - 1 Then
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                    Else
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                    End If
                    SP = SP + LP

                  x = x + 1
          
                Wend
                
                If Mid(Tablename(k), 1, 2) = "3X" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,3) = LEFT('" & Tablename(k) & "',3)"
                    
                ElseIf Mid(Tablename(k), 1, 2) = "61" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE (MID([" & Source & "].Data,14,2) + SPACE(1) + MID([" & Source & "].Data,80,1)) = LEFT('" & Tablename(k) & "',4) "
                
                Else
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] "
                End If
                        
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            Next k
        End Select
        

CleanClose:
    Set Newtdf = Nothing
    Set tdf = Nothing
    dbs.Close
    Set dbs = Nothing
    InputMsg = vbNo
    If InputMsg = vbYes Then GoTo EnterTableName
    Exit Sub
        
TableExists:
        Select Case Err.Number
            Case Is = 3010 '3010 Table already exists.
                'Delete existing table
                DoCmd.DeleteObject acTable, Tablename(k) & "_Analysis"
                Resume
            Case Is = 3146 '3146 ODBC TIME OUT
                MsgBox Err.Number
                MsgBox Err.Description
                InputMsg = MsgBox("ODBC TimeOut.  Would you like to contiunue the this process?", vbYesNo + vbCritical, "Timeout Error")
                If InputMsg = vbYes Then Resume
                GoTo CleanClose
            Case Is = 3265 '3265 Standard Table does not exists.
                k = k + 1
                Resume
            Case Else
                MsgBox Tablename(k)
                MsgBox Err.Number
                MsgBox Err.Description
                GoTo CleanClose
        End Select
End Sub

Public Sub BuildEETO()
    Dim dbs As Database
    Dim Standard As Boolean
    Dim SQL, Inbox, InputMsg, TopValue, Source As String
    Dim Tablename(1 To 50) As String
    Dim k, x, y, z, Numb, SP, LP As Integer
    Dim RS, rsAnal As DAO.Recordset
    Dim PAYCHEX As Recordset
    Dim tdf As TableDef
    Dim fld As Field
    Dim Newtdf, tabletdf  As TableDef
    Dim test As String
    Dim CHOOSESQL, CHOOSESQL2, INSERTSQL, SELECTSQLedd As String
    Inbox = ""
    'Set Source Table Here:
    Source = "EETO"
    
    'Different RecordCodes
                    Tablename(1) = "EmployeeTimeOff"
                    
                            
           
    InputMsg = vbYes
    Select Case InputMsg
        Case Is = vbCancel
            Exit Sub
 'When choosing to reload a specific table (no in msgbox):
        Case Is = vbNo
EnterTableName:
            Inbox = InputBox("Enter the Table Name", "Table", Inbox)
            If Inbox = "" Then
                Exit Sub
            Else
                Set dbs = CurrentDb
                k = 1
                Tablename(k) = Inbox
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name.  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    If x > 6 Then
                        LP = tdf(x).Size
                        If x = Numb - 1 Then
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                        Else
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                        End If
                        SP = SP + LP
                    Else
                        'Do Nothing because the first 7 field names are hard coded
                    End If
                  x = x + 1
                Wend
                SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,2) = LEFT('" & Tablename(k) & "',2)"
                
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            End If
   'When loading all tables (yes in msgbox):
        Case Is = vbYes
            Set dbs = CurrentDb
            k = 1
            For k = k To 1
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name, " & Tablename(k) & ".  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    
                    LP = tdf(x).Size
                    If x = Numb - 1 Then
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                    Else
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                    End If
                    SP = SP + LP

                  x = x + 1
          
                Wend
                
                If Mid(Tablename(k), 1, 2) = "3X" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,3) = LEFT('" & Tablename(k) & "',3)"
                    
                ElseIf Mid(Tablename(k), 1, 2) = "61" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE (MID([" & Source & "].Data,14,2) + SPACE(1) + MID([" & Source & "].Data,80,1)) = LEFT('" & Tablename(k) & "',4) "
                
                Else
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] "
                End If
                        
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            Next k
        End Select
        

CleanClose:
    Set Newtdf = Nothing
    Set tdf = Nothing
    dbs.Close
    Set dbs = Nothing
    InputMsg = vbNo
    If InputMsg = vbYes Then GoTo EnterTableName
    Exit Sub
        
TableExists:
        Select Case Err.Number
            Case Is = 3010 '3010 Table already exists.
                'Delete existing table
                DoCmd.DeleteObject acTable, Tablename(k) & "_Analysis"
                Resume
            Case Is = 3146 '3146 ODBC TIME OUT
                MsgBox Err.Number
                MsgBox Err.Description
                InputMsg = MsgBox("ODBC TimeOut.  Would you like to contiunue the this process?", vbYesNo + vbCritical, "Timeout Error")
                If InputMsg = vbYes Then Resume
                GoTo CleanClose
            Case Is = 3265 '3265 Standard Table does not exists.
                k = k + 1
                Resume
            Case Else
                MsgBox Tablename(k)
                MsgBox Err.Number
                MsgBox Err.Description
                GoTo CleanClose
        End Select
End Sub

Public Sub BuildEREC()
    Dim dbs As Database
    Dim Standard As Boolean
    Dim SQL, Inbox, InputMsg, TopValue, Source As String
    Dim Tablename(1 To 50) As String
    Dim k, x, y, z, Numb, SP, LP As Integer
    Dim RS, rsAnal As DAO.Recordset
    Dim PAYCHEX As Recordset
    Dim tdf As TableDef
    Dim fld As Field
    Dim Newtdf, tabletdf  As TableDef
    Dim test As String
    Dim CHOOSESQL, CHOOSESQL2, INSERTSQL, SELECTSQLedd As String
    Inbox = ""
    'Set Source Table Here:
    Source = "EREC"
    
    'Different RecordCodes
                    Tablename(1) = "EmployeeScheduledEarnOrDed"
                    
                            
           
    InputMsg = vbYes
    Select Case InputMsg
        Case Is = vbCancel
            Exit Sub
 'When choosing to reload a specific table (no in msgbox):
        Case Is = vbNo
EnterTableName:
            Inbox = InputBox("Enter the Table Name", "Table", Inbox)
            If Inbox = "" Then
                Exit Sub
            Else
                Set dbs = CurrentDb
                k = 1
                Tablename(k) = Inbox
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name.  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    If x > 6 Then
                        LP = tdf(x).Size
                        If x = Numb - 1 Then
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                        Else
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                        End If
                        SP = SP + LP
                    Else
                        'Do Nothing because the first 7 field names are hard coded
                    End If
                  x = x + 1
                Wend
                SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,2) = LEFT('" & Tablename(k) & "',2)"
                
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            End If
   'When loading all tables (yes in msgbox):
        Case Is = vbYes
            Set dbs = CurrentDb
            k = 1
            For k = k To 1
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name, " & Tablename(k) & ".  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    
                    LP = tdf(x).Size
                    If x = Numb - 1 Then
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                    Else
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                    End If
                    SP = SP + LP

                  x = x + 1
          
                Wend
                
                If Mid(Tablename(k), 1, 2) = "3X" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,3) = LEFT('" & Tablename(k) & "',3)"
                    
                ElseIf Mid(Tablename(k), 1, 2) = "61" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE (MID([" & Source & "].Data,14,2) + SPACE(1) + MID([" & Source & "].Data,80,1)) = LEFT('" & Tablename(k) & "',4) "
                
                Else
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] "
                End If
                        
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            Next k
        End Select
        

CleanClose:
    Set Newtdf = Nothing
    Set tdf = Nothing
    dbs.Close
    Set dbs = Nothing
    InputMsg = vbNo
    If InputMsg = vbYes Then GoTo EnterTableName
    Exit Sub
        
TableExists:
        Select Case Err.Number
            Case Is = 3010 '3010 Table already exists.
                'Delete existing table
                DoCmd.DeleteObject acTable, Tablename(k) & "_Analysis"
                Resume
            Case Is = 3146 '3146 ODBC TIME OUT
                MsgBox Err.Number
                MsgBox Err.Description
                InputMsg = MsgBox("ODBC TimeOut.  Would you like to contiunue the this process?", vbYesNo + vbCritical, "Timeout Error")
                If InputMsg = vbYes Then Resume
                GoTo CleanClose
            Case Is = 3265 '3265 Standard Table does not exists.
                k = k + 1
                Resume
            Case Else
                MsgBox Tablename(k)
                MsgBox Err.Number
                MsgBox Err.Description
                GoTo CleanClose
        End Select
End Sub

Public Sub BuildADIS()
    Dim dbs As Database
    Dim Standard As Boolean
    Dim SQL, Inbox, InputMsg, TopValue, Source As String
    Dim Tablename(1 To 50) As String
    Dim k, x, y, z, Numb, SP, LP As Integer
    Dim RS, rsAnal As DAO.Recordset
    Dim PAYCHEX As Recordset
    Dim tdf As TableDef
    Dim fld As Field
    Dim Newtdf, tabletdf  As TableDef
    Dim test As String
    Dim CHOOSESQL, CHOOSESQL2, INSERTSQL, SELECTSQLedd As String
    Inbox = ""
    'Set Source Table Here:
    Source = "ADIS"
    
    'Different RecordCodes
                    Tablename(1) = "EmployeeAllocDistribution"
                    
                            
           
    InputMsg = vbYes
    Select Case InputMsg
        Case Is = vbCancel
            Exit Sub
 'When choosing to reload a specific table (no in msgbox):
        Case Is = vbNo
EnterTableName:
            Inbox = InputBox("Enter the Table Name", "Table", Inbox)
            If Inbox = "" Then
                Exit Sub
            Else
                Set dbs = CurrentDb
                k = 1
                Tablename(k) = Inbox
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name.  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    If x > 6 Then
                        LP = tdf(x).Size
                        If x = Numb - 1 Then
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                        Else
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                        End If
                        SP = SP + LP
                    Else
                        'Do Nothing because the first 7 field names are hard coded
                    End If
                  x = x + 1
                Wend
                SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,2) = LEFT('" & Tablename(k) & "',2)"
                
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            End If
   'When loading all tables (yes in msgbox):
        Case Is = vbYes
            Set dbs = CurrentDb
            k = 1
            For k = k To 1
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name, " & Tablename(k) & ".  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    
                    LP = tdf(x).Size
                    If x = Numb - 1 Then
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                    Else
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                    End If
                    SP = SP + LP

                  x = x + 1
          
                Wend
                
                If Mid(Tablename(k), 1, 2) = "3X" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,3) = LEFT('" & Tablename(k) & "',3)"
                    
                ElseIf Mid(Tablename(k), 1, 2) = "61" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE (MID([" & Source & "].Data,14,2) + SPACE(1) + MID([" & Source & "].Data,80,1)) = LEFT('" & Tablename(k) & "',4) "
                
                Else
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] "
                End If
                        
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            Next k
        End Select
        

CleanClose:
    Set Newtdf = Nothing
    Set tdf = Nothing
    dbs.Close
    Set dbs = Nothing
    InputMsg = vbNo
    If InputMsg = vbYes Then GoTo EnterTableName
    Exit Sub
        
TableExists:
        Select Case Err.Number
            Case Is = 3010 '3010 Table already exists.
                'Delete existing table
                DoCmd.DeleteObject acTable, Tablename(k) & "_Analysis"
                Resume
            Case Is = 3146 '3146 ODBC TIME OUT
                MsgBox Err.Number
                MsgBox Err.Description
                InputMsg = MsgBox("ODBC TimeOut.  Would you like to contiunue the this process?", vbYesNo + vbCritical, "Timeout Error")
                If InputMsg = vbYes Then Resume
                GoTo CleanClose
            Case Is = 3265 '3265 Standard Table does not exists.
                k = k + 1
                Resume
            Case Else
                MsgBox Tablename(k)
                MsgBox Err.Number
                MsgBox Err.Description
                GoTo CleanClose
        End Select
End Sub

Public Sub BuildYTD()
    Dim dbs As Database
    Dim Standard As Boolean
    Dim SQL, Inbox, InputMsg, TopValue, Source As String
    Dim Tablename(1 To 50) As String
    Dim k, x, y, z, Numb, SP, LP As Integer
    Dim RS, rsAnal As DAO.Recordset
    Dim PAYCHEX As Recordset
    Dim tdf As TableDef
    Dim fld As Field
    Dim Newtdf, tabletdf  As TableDef
    Dim test As String
    Dim CHOOSESQL, CHOOSESQL2, INSERTSQL, SELECTSQLedd As String
    Inbox = ""
    'Set Source Table Here:
    Source = "YTD"
    
    'Different RecordCodes
                    Tablename(1) = "EmployeeYTD"
                    
                            
           
    InputMsg = vbYes
    Select Case InputMsg
        Case Is = vbCancel
            Exit Sub
 'When choosing to reload a specific table (no in msgbox):
        Case Is = vbNo
EnterTableName:
            Inbox = InputBox("Enter the Table Name", "Table", Inbox)
            If Inbox = "" Then
                Exit Sub
            Else
                Set dbs = CurrentDb
                k = 1
                Tablename(k) = Inbox
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name.  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    If x > 6 Then
                        LP = tdf(x).Size
                        If x = Numb - 1 Then
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                        Else
                            SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                        End If
                        SP = SP + LP
                    Else
                        'Do Nothing because the first 7 field names are hard coded
                    End If
                  x = x + 1
                Wend
                SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,2) = LEFT('" & Tablename(k) & "',2)"
                
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            End If
   'When loading all tables (yes in msgbox):
        Case Is = vbYes
            Set dbs = CurrentDb
            k = 1
            For k = k To 1
                '1 = Access Table and 4 = Linked Table
                SQL = "SELECT Name INTO TmpTable FROM MSYSOBJECTS WHERE TYPE IN (1, 4) AND NAME = '" & Tablename(k) & "'"
                dbs.Execute (SQL)
                dbs.TableDefs.Refresh
                Set tabletdf = dbs.TableDefs("TmpTable")
                If tabletdf.RecordCount = 0 Then
                    MsgBox "You have entered an invalid table name, " & Tablename(k) & ".  Please Re-Enter.", vbOKOnly, "Invalid Table"
                    DoCmd.DeleteObject acTable, "TmpTable"
                    GoTo EnterTableName
                    
                End If
                Set tabletdf = Nothing
                DoCmd.DeleteObject acTable, "TmpTable"
                On Error GoTo TableExists
                Set tdf = dbs.TableDefs(Tablename(k))
                SQL = "DELETE * FROM [" & Tablename(k) & "]"
                dbs.Execute SQL, dbFailOnError
                
                'Set tdf = dbs.TableDefs(Tablename(k))
                Numb = tdf.Fields.Count
                'initialize variables
                x = 0
                SP = 1
                LP = 1
                INSERTSQL = ""
                SELECTSQL = ""
                INSERTSQL = "INSERT INTO [" & Tablename(k) & "] ( "
                SELECTSQL = "SELECT "
                
                While x < Numb
                    'MsgBox tdf(x).Name
                    If x = Numb - 1 Then
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "]) "
                    Else
                        INSERTSQL = INSERTSQL & "[" & tdf(x).Name & "], "
                    End If
                    
                    
                    LP = tdf(x).Size
                    If x = Numb - 1 Then
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ) "
                    Else
                        SELECTSQL = SELECTSQL & "Mid([" & Source & "].[Data], " & SP & ", " & LP & " ), "
                    End If
                    SP = SP + LP

                  x = x + 1
          
                Wend
                
                If Mid(Tablename(k), 1, 2) = "3X" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE MID([" & Source & "].Data,14,3) = LEFT('" & Tablename(k) & "',3)"
                    
                ElseIf Mid(Tablename(k), 1, 2) = "61" Then
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] WHERE (MID([" & Source & "].Data,14,2) + SPACE(1) + MID([" & Source & "].Data,80,1)) = LEFT('" & Tablename(k) & "',4) "
                
                Else
                    SELECTSQL = SELECTSQL & " FROM [" & Source & "] "
                End If
                        
                SQL = INSERTSQL & SELECTSQL
                Debug.Print SQL
                dbs.Execute SQL, dbFailOnError
            Next k
        End Select
        

CleanClose:
    Set Newtdf = Nothing
    Set tdf = Nothing
    dbs.Close
    Set dbs = Nothing
    InputMsg = vbNo
    If InputMsg = vbYes Then GoTo EnterTableName
    Exit Sub
        
TableExists:
        Select Case Err.Number
            Case Is = 3010 '3010 Table already exists.
                'Delete existing table
                DoCmd.DeleteObject acTable, Tablename(k) & "_Analysis"
                Resume
            Case Is = 3146 '3146 ODBC TIME OUT
                MsgBox Err.Number
                MsgBox Err.Description
                InputMsg = MsgBox("ODBC TimeOut.  Would you like to contiunue the this process?", vbYesNo + vbCritical, "Timeout Error")
                If InputMsg = vbYes Then Resume
                GoTo CleanClose
            Case Is = 3265 '3265 Standard Table does not exists.
                k = k + 1
                Resume
            Case Else
                MsgBox Tablename(k)
                MsgBox Err.Number
                MsgBox Err.Description
                GoTo CleanClose
        End Select
End Sub

