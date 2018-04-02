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
    
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="VKWH_Download"
    Stop
    Resume Next
End Sub
</pre>

<h2>References</h2>

- Microsoft Office 15.0 Object Library
- Microsoft ActiveX Data Objects (Multi-dimensional)
- Microsoft ActiveX Data Objects 2.7 Library
- Microsoft  DOemq 1.o0 Object Library
- Microsoft Windows Common Controls-2 6.0 (SP6)
- RefEdit Control

- Change MACROWORKBOOK to name of current workbook

<h2>Other Modules</h2>

- Globals
- DatabaseAccess
- Display (StatusbarDisplay)
