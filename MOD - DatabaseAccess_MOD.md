<pre>
Attribute VB_Name = "DatabaseAccess_MOD"
Sub InitDatabaseAndTables()
    ReDim Preserve GLBDatabaseNameList(1)
    GLBDatabaseNameList(UBound(GLBDatabaseNameList)) = "dl_oge_analytics"
    ReDim Preserve GLBDatabaseNameList(UBound(GLBDatabaseNameList) + 1)
    GLBDatabaseNameList(UBound(GLBDatabaseNameList)) = "da_customer_vw"
    
    ReDim Preserve GLBTableNameList(1)
    GLBTableNameList(UBound(GLBTableNameList)) = "billing_statement_charge"
End Sub

Function DBMakeConnection(DBConn)
    If (Len(userName) = 0 Or Len(Password) = 0) And (Not DBConn.State = adStateOpen) Then
        LoginForm.Show
    End If
    If formCancel Then
        If Not DBConn Is Nothing Then Set DBConn = Nothing
        Exit Function
    End If
    '
    ' Connection string
    '
    s = "DSN=OGE;Databasename=dbc;Uid=" & userName & ";PWD=" & Password & ";Authentication Mechanism=LDAP;"

    Debug_Print s
    Call StatusbarShow("DBMakeConnection: Open")
    DBConn.Open s
    If DBConnect.State = adStateOpen Then 'If connection is success, continue
        Call StatusbarDisplay("DBMakeConnection: Connected to Database")
        Application.ODBCTimeout = 900
    End If
End Function

Sub DBConnectionProperties()
Dim DBCn As ADODB.Connection
    Set DBCn = DBCheckConnection(DBCn)
    
    For i = 0 To DBCn.Properties.count - 1
        Cells(i + 1, 1) = DBCn.Properties(i).Name
        Cells(i + 1, 2) = DBCn.Properties(i).Attributes
        Cells(i + 1, 3) = DBCn.Properties(i).Value
    Next i
    
    Cells(i + 1, 1) = "Command Timeout"
    Cells(i + 1, 2) = ""
    Cells(i + 1, 3) = DBCn.CommandTimeout
    
    Application.odbc
    
End Sub
</pre>
