<pre>
Sub RunMyQuery()

Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset
Dim useQuery As String


    On Error GoTo gotError

    '
    ' Setup Connection
    '
    Set DBCn = DBCheckConnection(DBCn)
    '
    Setup ResultSet
    '
    Set DBRs = DBCheckRecordset(DBRs)
    '
    ' Other parameters
    '
    With DBRs
        .CursorLocation = adUseClient ' adUseServer
        .CursorType = adOpenDynamic ' adUseClient ' adOpenStatic ' adOpenDynamic ' adOpenForwardOnly
        .LockType = adLockReadOnly ' adLockOptimistic
        ' Set .ActiveConnection = DBCn
    End With
    '
    ' Run the Query
    '
    DBRs.Open useQuery, DBCn
    
    recordCount = DBRs.recordCount
    
    If recordCount = 0 Then 
        ' no recrods found
    End If
    '
    ' Get Data
    '
    For j = 0 To recordCount - 1
            = DBRs.Fields(0).Value
        Cells(j + startRow, useCol + 1) = DBRs.Fields(1).Value
                
        DBRs.MoveNext
    Next j
End Sub
</pre>
