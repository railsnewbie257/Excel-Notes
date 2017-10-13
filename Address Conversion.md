<h2>Convert Column Number To Column Letter</h2>

<pre>
<colLet</em> = Split(Split(Columns(<em>colNum</em>).Address, ":")(1), "$")(1)
</pre>

<h2>Get Column Letter From Address</h2>

<pre>
<em>colLet</em> = Split(ActiveCell.Address, "$")(1)
</pre>

<h2>Get Row NUmber From Address</h2>

<pre>
<em>rowNum</em> = Split(ActiveCell.Address, "$")(2)
</pre>
