
<pre>
Dim FSO As Object

    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    Call FSO.CopyFile(fromFile, toFile)
    Call FSO.MoveFile(fromFile, toFile)
    Call FSO.DeleteFile(fileName)
</pre>
