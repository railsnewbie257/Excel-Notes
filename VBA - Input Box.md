<h2>There are 2 different InputBoxes</h2>
<pre>
InputBox("prompt", "title", "default")

- if <b>OK</b> clicked then "default" will be returned, can use " " (single empty space)
- if <b>CANCEL</b> is clicked then "" will be returned
</pre>

<h2>This one can return a range</h2>

<pre>
On Error Resume Next
set aRange = Application.Inputbox(<em>"prompt"</em>, title:=<em>"title"</em>, type:=<em>8</em>)
If Isempty(aRange) then Exit
</pre>


For <b>Text</b> input use
<pre>

</pre>
