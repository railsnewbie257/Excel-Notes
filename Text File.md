From: https://stackoverflow.com/questions/11503174/how-to-create-and-write-to-a-txt-file-using-vba

<pre>
Dim fso as Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile as Object
Set oFile = FSO.CreateTextFile(strPath)
oFile.WriteLine "test" 
oFile.Close
Set fso = Nothing
Set oFile = Nothing    
</pre>
