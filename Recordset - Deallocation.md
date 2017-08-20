[Why should I close and destroy a recordset?](https://stackoverflow.com/questions/22304994/why-should-i-close-and-destroy-a-recordset)
Always:
<pre>
<b>Recordset.Close</b>  ' without .Close will still consume resources
<b>Set Recordset = Nothing</b>  ' deallocates memory for garbage collection
</pre>
