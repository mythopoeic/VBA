Attribute VB_Name = "modDBConnection"
Option Explicit

Public DBCONT As Object

Public Function connectDatabase()
    Set DBCONT = CreateObject("ADODB.Connection")

    Dim Server_Name As String
    Dim Database_Name As String
    Dim Driver_Name As String
    Dim User_ID As String
    Dim Password As String
    
    Dim sConn As String
    
    Server_Name = wksAccess.Range("DBServerName")
    User_ID = wksAccess.Range("DBUserID")
    Password = wksAccess.Range("DBPassword")
    Database_Name = wksAccess.Range("DBDatabaseName")
    Driver_Name = wksAccess.Range("DBDriverName")
    
    sConn = "Driver=" & Driver_Name & ";Server=" & _
                                Server_Name & ";Database=" & Database_Name & _
                                ";UID=" & User_ID & ";PWD=" & Password & ";"
                                
    '''''''''''''''''''''''''''
    ' string for Access 2010 DB
    '''''''''''''''''''''''''''
        
        'Dim strDBPath As String
        'strDBPath = "c:/users/mytho/Desktop/SampleDatabase.accdb"
        'sConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                     "Data Source=" & strDBPath & ";" & _
                                     "Jet OLEDB:Engine Type=5;" & _
                                     "Persist Security Info=False;"
    '    On Error GoTo 2003
    '    DBCONT.Open sConn
    '    DBCONT.cursorlocation = 3
    '
    '    Exit Function
    '2003:
    '    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '                                 "Data Source=" & strDBPath & ";" & _
    '                                 "Persist Security Info=False;"
    On Error GoTo Error
    DBCONT.Open sConn
    DBCONT.cursorlocation = 3
    
    Exit Function
Error:
    Call closeDatabase
End Function

Public Function closeDatabase()
    On Error Resume Next
    DBCONT.Close
    Set DBCONT = Nothing
    On Error GoTo 0
End Function

Public Sub AccessSheetRefreshFromDB()
    Dim sqlstr As String
    Dim rs As Object
    Dim intCount3 As Integer
    Dim intCount2 As Integer
    Dim intCount As Integer
    Dim strArray() As String
    
    ' Clear AccessSheet
    wksAccess.Range("UserIDs").ClearContents
    wksAccess.Range("PwdHash").ClearContents
    wksAccess.Range("AccessListSheets").ClearContents
    wksAccess.Range("AccessListUsers").ClearContents
    wksAccess.Range("SheetList").ClearContents
    
    ' Create RecordSet
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Connect to Database
    Call connectDatabase
    rs.Open "SELECT * FROM [User]", DBCONT
    
    ' Check to see if the recordset contains rows
    If Not (rs.EOF And rs.BOF) Then
        intCount = 1
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        Do Until rs.EOF = True
            
            ' Update Hash Table
            If rs!UserID = "admin" Then
                wksAccess.Range("MasterUserName").Value = rs!UserID
                wksAccess.Range("MasterUserHash").Value = rs!Hash
                intCount = intCount - 1
            Else
                wksAccess.Range("UserIDsStart").Offset(intCount - 1).Value = rs!UserID
                wksAccess.Range("PwdHashStart").Offset(intCount - 1).Value = rs!Hash
            End If
            ' Update Sheet Association Table and Sheets Table
            strArray = Split(rs!Sheets, ",")
            For intCount2 = LBound(strArray) To UBound(strArray)
                If rs!UserID <> "admin" Then
                    ' Update access list
                    AccessListAddRow rs!UserID, Trim(strArray(intCount2))
                    
                    ' Update SheetList
                    If wksAccess.Range("SheetListStart") = "" Then
                        ' First Sheet, add it to Sheet List
                        wksAccess.Range("SheetListStart") = Trim(strArray(intCount2))
                    Else
                        If (wksAccess.Range("SheetList").Find(strArray(intCount2), LookIn:=xlValues) Is Nothing) Then
                            ' Sheet doesn't already exist in Sheet List, so add it
                            wksAccess.Range("SheetListStart").Offset(wksAccess.Range("SheetList").Count) = Trim(strArray(intCount2))
                        End If
                    End If
                End If
            Next
            
            
            'Move to the next record.
            rs.MoveNext
            intCount = intCount + 1
        Loop
    Else
        'MsgBox "There are no records in the recordset."
    End If
    
    rs.Close
    Set rs = Nothing
    Call closeDatabase
    
    ' Refresh Sheet List and Sheet Association

End Sub


Public Function readDBRecord(UserID As String, strOutput As String) As String
    
    Dim sqlstr As String
    Dim rs As Object
    
    ' Create RecordSet
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Connect to Database
    Call connectDatabase
    

    sqlstr = "SELECT " & strOutput & " FROM [User] WHERE UserID='" & UserID & "'"
    rs.Open sqlstr, DBCONT

    If rs.RecordCount > 0 Then
        readDBRecord = rs(0)
    End If

    rs.Close
    Set rs = Nothing
    Call closeDatabase
End Function

Public Sub insertDBRecord(UserID As String, Hash As String, Sheets As String)

    Dim sqlstr As String
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    sqlstr = "SELECT UserID FROM [User] WHERE UserID='" & UserID & "'"
    
    
    Call connectDatabase
    'check if record already exists. If not, add new record
    rs.Open sqlstr, DBCONT
    If rs.RecordCount = 0 Then
    sqlstr = "INSERT INTO [User](UserID, Hash, Sheets) VALUES('" & UserID & "','" & Hash & "','" & Sheets & "')"
    DBCONT.Execute sqlstr
    End If
    
    rs.Close
    Set rs = Nothing
    Call closeDatabase
End Sub

Public Sub deleteDBRecord(DeletingID As String)

    Dim sqlstr As String
    
    sqlstr = "DELETE FROM [User] WHERE UserID='" & DeletingID & "'"
    
    Call connectDatabase
    DBCONT.Execute sqlstr
    Call closeDatabase

End Sub

Public Sub updateDBRecord(field As String, UserID As String, newValue As String)

    Dim sqlstr As String
    sqlstr = "UPDATE [User] SET " & field & "='" & newValue & _
            "' WHERE UserID='" & UserID & "'"
            
    Call connectDatabase
    DBCONT.Execute sqlstr
    Call closeDatabase

End Sub

Public Function UserExists(strUserID As String) As Boolean
    Dim sqlstr As String
    Dim rs As Object
    
    ' Create RecordSet
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Connect to Database
    Call connectDatabase
    

    sqlstr = "SELECT COUNT(Hash) FROM [User] WHERE UserID='" & strUserID & "'"
    rs.Open sqlstr, DBCONT

    If rs.RecordCount > 0 Then
        UserExists = CInt(rs(0)) > 0
    End If

    rs.Close
    Set rs = Nothing
    Call closeDatabase
End Function
