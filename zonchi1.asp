<%@ Language = VBScript%>
<%
on error resume next
testa = split(request.ServerVariables("HTTP_REFERER"),"/")

ad = request.QueryString("ad")



Set conn1 = Server.CreateObject("ADODB.Connection")
conn1.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")

Set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")


set Lookup = Server.CreateObject("ADODB.recordset")
Lookup.CursorLocation = 3 
Lookup.CursorType = 1
				
zonchiname = testa(4)
if zonchiname = "" then zonchiname = "funds4me"


SQL = "SELECT * from members "
SQL = SQL + "WHERE ((zonchiname = '" & zonchiname & "');"

Lookup.open SQL, conn1,1,3

if lookup.bof and lookup.eof then
SQL = "SELECT * from members "
SQL = SQL + "WHERE (zonchiname = 'funds4me');"
CookieUserID = "zonchi"
else
CookieUserID = "f4m" & lookup("memID") & "(zonchi:" & testa(4) & ")"
end if
lookup.close

Lookup.open SQL, conn1,1,3


	stores = split(lookup("zonchistores" & ad),",")
	
	lookup.close
	
	for a = 0 to 3
					
	
		SQL = "SELECT stores.* "
		SQL = SQL + " FROM stores "
		SQL = SQL + "WHERE (merchant='" & stores(a) & "');"
		
		Lookup.open SQL, conn3,1,3
		'response.Write(a & "|" & SQL)
		if Lookup.BOF and lookup.EOF then
		
		else
			lookup.movefirst
			
			do while not lookup.eof
			linkit = Replace(lookup("Link"),"CookieUserID",CookieUserID,1, -1, 1)
				%>
				&nbsp;<a href="<%=linkit%>"><%=lookup("Banner")%></a>&nbsp;
				<%
					
				lookup.movenext
				
			loop
		
		end if
		Lookup.Close	
	next
%>
