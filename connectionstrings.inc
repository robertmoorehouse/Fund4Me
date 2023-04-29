<%
'objConn1 cashback database
Dim objConn1
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")'connection, user, ID	

'objConn2 invoices database
Dim objConn2
Set objConn2 = Server.CreateObject("ADODB.Connection")
objConn2.Open Application("Connection2_ConnectionString"), Application("Connection2_RuntimeUserName"), Application("Connection2_RuntimePassword")
	
'objConn3 stores database
Dim objConn3
Set objConn3 = Server.CreateObject("ADODB.Connection")
objConn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
%>		