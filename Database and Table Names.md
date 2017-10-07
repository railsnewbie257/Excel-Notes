<pre>
    '
    ' Database and Table Names
    '
    If IsArrayAllocated(GLBDatabaseList) Then
        For i = UBound(GLBDatabaseList) To 1 Step -1
            cbDatabaseName.AddItem GLBDatabaseList(i)
        Next i
        cbDatabaseName = GLBDatabaseList(UBound(GLBDatabaseList))
    End If
    If IsArrayAllocated(GLBTableList) Then
        For i = UBound(GLBTableList) To 1 Step -1
            cbTableName.AddItem GLBTableList(i)
        Next i
        cbTableName = GLBTableList(UBound(GLBTableList))
    End If
</pre>
    
