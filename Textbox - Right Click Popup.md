<b>In Userform</b>  
NOTE: Button=2 is right click, Button=1 is left click

<pre>
Private Sub TextBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MakePopUp
    If Button = 2 Then
        Application.CommandBars("MyPopUp").ShowPopup
    End If
End Sub
</pre>


<b>In standard module</b>

<pre>
Sub MakePopUp()
     'Remove any old instance of MyPopUp
    On Error Resume Next
    CommandBars("MyPopUp").Delete
    On Error GoTo 0
     
    With CommandBars.Add(Name:="MyPopUp", Position:=msoBarPopup)
        .Controls.Add(Type:=msoControlButton, ID:=19).OnAction = "Textbox_Copy"
        .Controls.Add(Type:=msoControlButton, ID:=22).OnAction = "Textbox_Paste"
    End With
End Sub

Sub Textbox_copy()
    UserForm1.TextBox1.Copy
End Sub

Sub Textbox_paste()
    UserForm1.TextBox1.Paste
End Sub

</pre>
