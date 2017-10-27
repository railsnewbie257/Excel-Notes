<pre>
  If (Not columnLen) = True Then
    ReDim Preserve columnLen(1)
  Else
    ReDim Preserve columnLen(UBound(columnLen) + 1)
  End If  
    columnLen(UBound(columnLen)) = useRange.Areas(i).Rows.count
</pre>
