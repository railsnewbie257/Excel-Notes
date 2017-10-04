<h2>There are 2 different InputBoxes</h2>
<pre>
<b>InputBox</b>("prompt", "title", "default")

- if <b>OK</b> clicked then "default" will be returned, can use " " (single empty space)
- if <b>CANCEL</b> is clicked then "" will be returned
</pre>

<h2>This one can return a range</h2>

[Application.InputBox Method (Excel)](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/application-inputbox-method-excel)

<pre>
On Error Resume Next
set aRange = <b>Application.Inputbox</b>(<em>"prompt"</em>, title:=<em>"title"</em>, type:=<em>8</em>)
If Isempty(aRange) then Exit
</pre>

InputBox types:
- 0 - A formula
- 1 - A number
- 2 - Text (a string)
- 4 - A logical value ( True or False )
- 8 - A cell reference, as a Range object
- 16 - An error value, such as #N/A
- 32 - An array of values


For <b>Text</b> input use
<pre>

</pre>
