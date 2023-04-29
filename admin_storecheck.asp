<%@ Language = VBScript%>
Merchant value
<%
Set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set Lookup = Server.CreateObject("ADODB.recordset")
			
checkval = request.Form("merchantcheck")
			
SQL = "SELECT * FROM stores where merchant like '%" & checkval & "%' order by Merchant"

Lookup.open SQL, conn3 , 1, 3

do while not lookup.EOF

	response.Write( "<br><a href=""storeadmin.asp?vid=" & (lookup("id")) & """> " & lookup("id") & "</a> | " &lookup("merchant") & "<br>")
	
	lookup.MoveNext
	
Loop
Lookup.close	


%>
<hr>

Notes Value
<br><Br>
<%
SQL = "SELECT * FROM stores where notes like '%" & checkval & "%' order by Merchant"

Lookup.open SQL, conn3 , 1, 3

do while not lookup.EOF

	response.Write(lookup("merchant") & "<br>" & lookup("notes") & "<br><br><hr>" )
	
	lookup.MoveNext
	
Loop
Lookup.close	


%>
