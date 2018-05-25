<pre>
<em>Sub ReadDatabase()</em>
<b>Dim DBCn As ADODB.Connection</b>
<b>Dim DBRs As ADODB.Recordset</b>
<b>Dim useQuery as String</b>
<b>Dim fieldCount as Integer</b>
<b>Dim i as Integer, j as Integer</b>

<b>On Error GoTo gotError</b>
10    <em>useQuery = "SELECT TOP 1 * from " & tableName</em>

20    <b>Set DBCn = DBCheckConnection(DBCn)</b>
30    <b>Set DBRs = DBCheckRecordset(DBRs)</b>

40    <b>With DBRs</b>
50        <b>.CursorLocation = adUseClient ' adUseServer</b>
60        <b>.CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly</b>
70        <b>.LockType = adLockReadOnly ' adLockOptimistic</b>
80        <b>Set .ActiveConnection = DBCn</b>
90    <b>End With</b>

100   <b>DBRs.Open useQuery, DBCn</b>

'----------------------------------------------------------------------------------------
110   <b>recordCount = DBRs.recordCount</b> <em>' see if something returned</em>
120   <b>For row = 1 To recordCount</b>         <em>' down the sheet</em>
        
130      <b>fieldCount = DBRs.Fields.count</b>
140      <b>For j = 0 To fieldCount - 1</b>     <em>' across the row</em>
150          <em>Cells(row, j+1) = DBRs.Fields(j).Value</em>
160      <b>Next j</b>
170      <b>DBRs.MoveNext</b>                   <em>' next row of data</em>
        
180  <b>Next i</b>
190  <b>DBRs.Close</b>
200  <b>Exit Sub</b>
    
<b>gotError:</b>
    <b>MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:=" "</b>
    <b>Stop</b>
    <b>Resume Next</b>
End Sub
