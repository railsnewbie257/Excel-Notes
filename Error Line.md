<h2>Use <b>Erl</b> to report line error

<pre>
Sub sample()
Dim i As Long

On Error GoTo Whoa

10    Debug.Print "A"
20    Debug.Print "B"
30    i = "Sid"
40    Debug.Print "A"

50    Exit Sub
Whoa:
    MsgBox "Error on Line : " & Erl
End Sub
</pre>
