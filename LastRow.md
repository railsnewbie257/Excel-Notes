Gets around SpecialCells problem of remembring previous maximum last row

<pre>Sub FindLastRow()

Dim LastRow As Long

	If WorksheetFunction.CountA(Cells) > 0 Then

		'Search for any entry, by searching backwards by Rows.

		LastRow = Cells.Find(What:="*", After:=[A1], _
			  SearchOrder:=xlByRows, _
			  SearchDirection:=xlPrevious).Row
	End If

End Sub
</pre>
