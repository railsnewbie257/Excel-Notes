<h2>Tools | References</h2>

in VBA editor under <b>Tools | References</b>, need to checkbox
- Microsoft Forms 2.0 Object Library
- Microsoft ActiveX Data Objects (Multidimensional) 2.8 Library
- Microsoft ActiveX Data Objects 2.7 Library

<h2>VBA Code</h2>
<pre>
Set DBConnect = New ADODB.Connection
Set RecordSet = New ADODB.Recordset

DBConnect.ConnectionTimeout = 0 'To wait till the query finishes without generating error
DBConnect.CommandTimeout = 1200

userName = "YOUR USERNAME"
userPassword = "your password"

s = "DSN=OGE;Databasename=dbc;Uid=" & UserName & ";PWD=" & userPassword & ";Authentication Mechanism=LDAP;"

DBConnect.Open s
If DBConnect.State = adStateOpen Then
    MsgBox "Connected to database"
End If

RecordSet.open "select * from dbc.dbcinfo;", DBConnect
Debug.Print RecordSet.fields.count
</pre>

<h2>Copy column headers to spreadsheet</h2>
<pre>
For i = 0 To Recordset.Fields.Count - 1
    Cells(1, i + 1) = Recordset.Fields(i).Name
Next i
</pre>

