<h2>Copy UI</h2>
<pre>
Sub DeployUI()
Dim FSO As Object

    Set FSO = CreateObject("Scripting.Filesystemobject")
    Username = LCase(Environ$("Username"))

    toFile = "C:\Users\" & Username & "\AppData\Local\Microsoft\Office\Excel.officeUI"
    fromFile = "C:\OGE\Excel.officeUI"
    
    If Dir(toFile) <> "" Then
        Call FSO.CopyFile(toFile, toFile & "_old")
    End If

    If Dir(fromFile) <> "" Then
    Debug.Print toFile
        Call FSO.CopyFile(fromFile, toFile, True)
        Call FSO.DeleteFile(fromFile)
        MsgBox "New Menus Deployed."
    End If
End Sub
</pre>

<h2>Get the UI</h2>
<pre>
Sub GetUI()
Dim FSO As Object

    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    Username = LCase(Environ$("Username"))
    
    fromFile = "C:\Users\" & Username & "\AppData\Local\Microsoft\Office\Excel.officeUI"
    toFile = "C:\OGE\Excel.officeUI"
    
    Call FSO.CopyFile(fromFile, toFile)
    
End Sub
</pre>
