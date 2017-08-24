
<pre>
Dim <b>FSO As Object</b>

    Set <b>FSO = CreateObject("Scripting.Filesystemobject")</b>
    
    Call <b>FSO.CopyFile</b>(fromFile, toFile)
    Call <b>FSO.MoveFile</b>(fromFile, toFile)
    Call <b>FSO.DeleteFile</b>(fileName)
</pre>
