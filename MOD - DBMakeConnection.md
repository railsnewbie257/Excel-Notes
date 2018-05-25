<pre>
Function DBMakeConnection(DBConn)
    If (Len(UserName) = 0 Or Len(Password) = 0) And (Not DBConn.State = adStateOpen) Then
        LoginForm.Show
    End If
    If formCancel Then
        If Not DBConn Is Nothing Then Set DBConn = Nothing
        Exit Function
    End If
    '
    ' Connection string
    '
    s = "DSN=OGE;Databasename=dbc;Uid=" & UserName & ";PWD=" & Password & ";Authentication Mechanism=LDAP;"

    Debug_Print s
    Call StatusbarShow("DBMakeConnection: Open")
    DBConn.Open s
    If DBConnect.State = adStateOpen Then 'If connection is success, continue
        Call StatusbarDisplay("DBMakeConnection: Connected to Database")
        Application.ODBCTimeout = 900
    End If
End Function
</pre>
