<pre>
Sub TestDBCheckConnection()
Dim DBCn As ADODB.Connection
Dim DBRs As ADODB.Recordset

  Set DBCn = DBCheckConnection(DBCn)
  Set DBRs = DBCheckRecordset(DBRs)
  
  DBRs.Open 
  
</pre>
