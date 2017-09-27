https://stackoverflow.com/questions/3511903/execute-code-when-form-is-closed-in-vba-excel-2007

<pre>
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        ' Your codes
        ' Tip: If you want to prevent closing UserForm by Close (×) button in the right-top corner of the UserForm, just uncomment the following line:
        ' Cancel = True
    End If
End Sub
</pre>

<pre>
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        'Your code goes here
    End If
End Sub
</pre>
