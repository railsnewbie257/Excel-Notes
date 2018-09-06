Requirements:



<pre>
<b>Private Sub UserForm_Initialize</b>()
    '
    ' Center form on ActiveWindow
    '
    Me.top = Application.top + Application.Height / 2 - Me.Height / 2
    Me.left = Application.left + Application.width / 2 - Me.width / 2
<b>End Sub</b>
</pre>
