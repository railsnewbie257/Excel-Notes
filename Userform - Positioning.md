<h2>To Center On The Calling Window</h2>

<pre>
Private Sub UserForm_Initialize()
    <em>Yourform</em>.Top = Application.Top + Application.Height / 2 - <em>Yourform</em>.Height / 2
    <em>Yourform</em>r.Left = Application.Left + Application.width / 2 - <em>Yourform</em>.width / 2
End Sub
</pre>
