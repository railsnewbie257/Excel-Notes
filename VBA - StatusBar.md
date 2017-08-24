<h2>Status Bar On</h2>
<pre>
Function StatusBarOn()
    Application.DisplayStatusBar = True
End Function
</pre>

<h2>Status Bar Off</h2>
<pre>
Function StatusBarOff()
    Application.DisplayStatusBar = False
End Function
</pre>

<h2>Status Bar Display</h2>
<pre>
Sub StatusbarDisplay(Optional s)
    Application.DisplayStatusBar = True
    If IsMissing(s) Then s = "testing..."
        Application.StatusBar = s
        <b>DoEvents</b> ' necessary to force display update
End Sub
</pre>
