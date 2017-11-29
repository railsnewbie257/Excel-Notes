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

<pre>
Public Function ShellRun(sCmd As String) As String
    
    On Error GoTo gotError
    'Run a shell command, returning the output as a string'

      Dim oShell As Object
10    Set oShell = CreateObject("WScript.Shell")

    'run command'
      Dim oExec As Object
      Dim oOutput As Object
20    Set oExec = oShell.Exec(sCmd)
30    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object'
      Dim s As String
      Dim sLine As String
40    While Not oOutput.AtEndOfStream
50        sLine = oOutput.ReadLine
60        If sLine <> "" Then s = s & sLine & vbCrLf
70    Wend

80    ShellRun = s

      Set oExec = Nothing
      Set oShell = Nothing
      
      Exit Function
      
gotError:
    MsgBox Err.Number & " " & Err.Description & vbNewLine & vbNewLine & "Error on line: " & Erl, Title:="ShellRun"
    Stop
    Resume Next
      
End Function
</pre>
