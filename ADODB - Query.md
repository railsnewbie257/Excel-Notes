<h2>ADODB Setup</h2>

<pre>
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
    With DBRs
        .CursorLocation = adUseClient <em>' adUseServer, adUseClient</em>
        .CursorType = adUseClient <em>' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly</em>
        .LockType = adLockOptimistic <em>' adLockReadOnly</em>
        Set .ActiveConnection = DBCn
    End With
</pre>
.CursorLocation:
- adUseServer
- adUseClient

.CursorType:
- adUseClient
- adOpenStatic
- adOpenDynamic
- adOpenForwardOnly

.LockType:
- adLockOptimistic
- adLockReadOnly

<h2>Query Call</h2>

<pre>
On Error GoTo gotError
<em>useQuery</em> = "SELECT * FROM <em>Table</em>"
DBRs.Open <em>useQuery</em>
</pre>

<h2>Error Catching Routine</h2>

<pre>
gotError:
    MsgBox "DBQuery Error (" & Err.Number & "): " & Err.Description, vbOKOnly, Title:="DBQuery ERROR"
    Set DBCn = DBCheckConnection(DBCn)
    Set DBRs = DBCheckRecordset(DBRs)
</pre>
