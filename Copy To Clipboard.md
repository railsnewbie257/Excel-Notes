<pre>
Sub CopyToClipboard()
Dim clipboard As MSForms.DataObject

    Set clipboard = New MSForms.DataObject
    clipboard.SetText <em>"Text to go in clipboard"</em>
    clipboard.PutInClipboard
End Sub
</pre>

Then use <b>Paste</b> to get it from the Clipboard
