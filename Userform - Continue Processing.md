<h2>Return Control To Caller</h2>

Calling .Show using <b>False</b> will display the form but return control to the caller

<pre>
<em>Yourform</em>.Show <b>False</b>
</pre>

Then any properties which are updated will be displayed

<pre>
<em>Yourform</em>.CheckBox1 = True
</pre>

When finished be sure to

<pre>
<em>Yourform</em>.Close
</pre>
