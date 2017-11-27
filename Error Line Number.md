<h2>Use <b>Erl</b> to report line error line number</h2>

<pre>
Sub sample()
Dim i As Long

<b>On Error GoTo gotError</b>

10    Debug.Print "A"
20    Debug.Print "B"
30    i = "Sid"  ' error on this line
40    Debug.Print "A"

50    Exit Sub

<b>gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl
    Stop
    Resume Next</b>
</pre>

reports error on line (<b>30</b>)

<h2>Also possible is</h2>

<pre>
Sub sample()
Dim i As Long

<b>On Error GoTo gotError</b>

10    Debug.Print "A"
      Debug.Print "B"
      i = "Sid"
      Debug.Print "A"

50    Exit Sub

<b>gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl
    Stop
    Resume Next</b>
</pre>

will report the last line number is saw (<b>10</b>), otherwise <b>0</b>
