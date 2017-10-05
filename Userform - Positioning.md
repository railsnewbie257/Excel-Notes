<h2>To Center On The Calling Window</h2>

<pre>
Private Sub UserForm_Initialize()
    Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2
    Me.Left = Application.Left + Application.width / 2 - Me.width / 2
End Sub
</pre>
