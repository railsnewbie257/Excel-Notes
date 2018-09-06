<h2>Useful for UserForm_Initialize()</h2>

Requires:

[Userform - Globals](https://github.com/ppihoge/Excel-Notes/blob/master/Userform%20-%20Globals.md)

<pre>
    formCancel = false
    '
    ' <em>Right click Copy / Paste</em>
    '
    Set cBar = New clsBar
    cBar.Initialize Me
    '
    ' <em>Center form on ActiveWindow</em>
    '
    Me.Top = Application.Top + Application.Height / 2 - Me.Height / 2
    Me.Left = Application.Left + Application.width / 2 - Me.width / 2
</pre>
