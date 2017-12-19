From: https://stackoverflow.com/questions/28379191/unable-to-connect-to-teradata-from-excel-using-vba-code-teradata-server-cant

<pre>
Function OpenConn() As Object 

Set OpenConn = New ADODB.Connection 

Dim myConnectionString As String 

myConnectionString = "Provider=TDOLEDB;Data Source=MyTeradataServerName;Persist Security Info=True;User ID=MyTeradataUserID;Password=MyTeradataPass;Session Mode=ANSI;DefaultDatabase=GRP_BCE_FINANCE_IM;MaxResponseSize=65477;" 

OpenConn.Open myConnectionString 

End Function
</pre>

<pre>
Sub PushCCHier()
Dim TeraObjCmd As New ADODB.Command
Dim TeraObjRs As ADODB.Recordset
Dim TeraObjRs2 As ADODB.Recordset

Dim TeraCnxn As Object
Set TeraCnxn = OpenConn()

TeraObjCmd.ActiveConnection = TeraCnxn

'Clear Previous Data
TeraObjCmd.ActiveConnection = TeraCnxn
TeraObjCmd.CommandText = "delete from MyTable"
TeraObjCmd.Execute

'Load New Data
Set TeraObjRs2 = New ADODB.Recordset
TeraObjRs2.Open "SELECT * FROM MyTable where 1 = 2 ", TeraCnxn, adOpenStatic, adLockOptimistic
With TeraObjRs2
    For irow = 7 To 8 'loading results from rows in my spredsheet
        If Len(Trim(Range("B" & irow).Value)) <> 0 Then 'Avoid blank rows
            .AddNew
            .Fields(0) = Range("B" & irow).Value
        End If
    Next
   .UpdateBatch
   .Close
End With

' clean up objects
Set objCmd = Nothing
Call CloseConn(TeraCnxn)
MsgBox "Update Complete!"

	

This connection script worked for me.

' Add Microsoft ActiveX Data Objects 2.8 Library in References ' When installing Teradata SQL Assistant, include the ODBC Driver for Teradata will install the TDOLEDB provider ' This example connects to Teradata, deletes the contents of MyTable & inserts row 7- 8 from the active spreadsheet

Function OpenConn() As Object Set OpenConn = New ADODB.Connection Dim myConnectionString As String myConnectionString = "Provider=TDOLEDB;Data Source=MyTeradataServerName;Persist Security Info=True;User ID=MyTeradataUserID;Password=MyTeradataPass;Session Mode=ANSI;DefaultDatabase=GRP_BCE_FINANCE_IM;MaxResponseSize=65477;" OpenConn.Open myConnectionString End Function

Sub CloseConn(conn As Object) conn.Close Set conn = Nothing End Sub

Sub PushCCHier()

Dim TeraObjCmd As New ADODB.Command
Dim TeraObjRs As ADODB.Recordset
Dim TeraObjRs2 As ADODB.Recordset

Dim TeraCnxn As Object
Set TeraCnxn = OpenConn()

TeraObjCmd.ActiveConnection = TeraCnxn

'Clear Previous Data
TeraObjCmd.ActiveConnection = TeraCnxn
TeraObjCmd.CommandText = "delete from MyTable"
TeraObjCmd.Execute

'Load New Data
Set TeraObjRs2 = New ADODB.Recordset
TeraObjRs2.Open "SELECT * FROM MyTable where 1 = 2 ", TeraCnxn, adOpenStatic, adLockOptimistic
With TeraObjRs2
    For irow = 7 To 8 'loading results from rows in my spredsheet
        If Len(Trim(Range("B" & irow).Value)) <> 0 Then 'Avoid blank rows
            .AddNew
            .Fields(0) = Range("B" & irow).Value
        End If
    Next
   .UpdateBatch
   .Close
End With

' clean up objects
Set objCmd = Nothing
Call CloseConn(TeraCnxn)
MsgBox "Update Complete!"
End Sub
</pre>

<pre>
Sub CloseConn(conn As Object) 
  conn.Close 
  Set conn = Nothing 
End Sub
</pre>
