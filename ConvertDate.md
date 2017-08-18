<h2>Convert a date to datetime serial number</h2>
<pre>
["G13"] = <b>2017-07-12</b> <b>08:04:10</b>.467-05:00
</pre>

<pre>
=LEFT(G13,10)+MID(G13,12,8)
</pre>
