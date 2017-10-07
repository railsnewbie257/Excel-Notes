
<h2>In Global Module</h2>

<pre>
Public Username As String
Public Password As String
Public formCancel As Boolean
Public DBConnect As ADODB.Connection
</pre>

<h2>Make LoginForm Userform</h2>

- TextBox: <b>tstUserName</b>
- TextBox: <b>txtPassword</b>
- CommandButton: <b>btnSubmit</b>
- CommandButton: <b>btnCancel</b>

Userform Code
<pre>

Private Sub btnSubmit_Click()
    Username = txtUsername.Text
    Username_save = txtUsername.Text
    
    Password = txtPassword.Text
    Password_save = txtPassword.Text
    
    formCancel = False
    Unload LoginForm
End Sub
Private Sub btnCancel_Click()
    formCancel = True
    Unload LoginForm
End Sub

Private Sub UserForm_Initialize()
    txtUsername.Text = LCase(Environ$("Username"))
    'txtUsername.Enabled = False
    txtPassword.SetFocus
End Sub
</pre>


<h2> In DBConnect Module</h2>

<pre>
Sub ConnectDB()
    'Set DBConnect = New ADODB.Connection
    'Set Recordset = New ADODB.Recordset
    tries = 0
TryAgain:
    On Error Resume Next
    If (Len(Username) = 0 Or Len(Password) = 0) And (Not DBConnect.State = adStateOpen) Then
        Set DBConnect = New ADODB.Connection
        DBConnect.ConnectionTimeout = 0 'To wait till the query finishes without generating error
        DBConnect.CommandTimeout = 1200
        
        LoginForm.Show
    End If
    If formCancel Then
        On Error Resume Next
            DBConnect.Close
        Set DBConnect = Nothing
        Exit Sub
    End If

    s = "DSN=OGE;Databasename=dbc;Uid=" & Username & ";PWD=" & Password & ";Authentication Mechanism=LDAP;"
'=============================================================================
    On Error GoTo Goterror
    Debug.Print s
    If Not (DBConnect.State = adStateOpen) Then
        On Error GoTo Goterror
            DBConnect.Open s
    End If
    
    If DBConnect.State = adStateOpen Then 'If connection is success, continue
        MsgBox "Connected to Teradata"
        Application.ODBCTimeout = 900
    Else
        MsgBox "Could not connect to Teradata"
    End If
    
    Exit Sub
    
Goterror:
    Debug.Print Err.Description
    retCode = MsgBox(Err.Description & vbNewLine & vbNewLine & "Try Again?", vbYesNoCancel)
    If (retCode = vbNo) Or (retCode = vbCancel) Then
        Exit Sub
    Else
        'On Error GoTo 0
        Password = ""
        GoTo TryAgain
    End If

End Sub
</pre>
