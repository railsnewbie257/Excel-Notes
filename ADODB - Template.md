<pre>
Sub ReadDatabase()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim useQuery as String
Dim i as Integer, j as Integer

<b>On Error GoTo gotError</b>
10    useQuery = "SELECT TOP 1 * from " & tableName

20    Set DBCn = DBCheckConnection(DBCn)
30    Set DBRs = DBCheckRecordset(DBRs)

40    With DBRs
50        .CursorLocation = adUseClient ' adUseServer
60        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
70        .LockType = adLockReadOnly ' adLockOptimistic
80        Set .ActiveConnection = DBCn
90    End With

100   <b>DBRs.Open useQuery, DBCn</b>

'----------------------------------------------------------------------------------------
110   <b>recordCount = DBRs.recordCount</b> ' see if something returned
120   For row = 1 To recordCount         ' down the sheet
        
130      fieldCount = DBRs.Fields.count
140      For j = 0 To fieldCount - 1     ' across the row
150          Cells(row, j+1) = DBRs.Fields(j).Value
160      Next j
170      <b>DBRs.MoveNext</b>                   ' next row of data
        
180  Next i
190  <b>DBRs.Close</b>
200  Exit Sub
    
<b>gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:=" "
    Stop
    Resume Next</b>
End Sub
