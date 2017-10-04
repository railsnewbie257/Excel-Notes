[Call Excel VBA Shell Command Object Output](https://officetricks.com/execute-shell-read-output-vba/)

<pre>
Function wshell(runCmd)
Dim wsh As Object

Set wsh = CreateObject("WScript.Shell")

Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1 'or whatever suits you best
Dim errorCode As Integer

Set wshOut = wsh.exec(runCmd).stdout

While Not wshOut.AtEndOfStream
    outputLine = wshOut.ReadLine
    If outputLine <> "" Then
        output = output & outputLine & vbCrLf
    End If
Wend
    
wshell = output
Set wsh = Nothing

End Function
</pre>
