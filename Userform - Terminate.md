https://stackoverflow.com/questions/3511903/execute-code-when-form-is-closed-in-vba-excel-2007

<h2>To Prevent Userform to Close Using (X) In Top Right Corner</h2>
<pre>
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ...
        <b>Cancel = True</b>
    End If
End Sub
</pre>

<pre>
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ...
    End If
End Sub
</pre>

<h2>Terminate Is Executed Whenever The Form Is Closed</h2>

<pre>
' any closing of Userform will come here
Private Sub UserForm_Terminate()
    ...
    Unload Me
End Sub
</pre>

