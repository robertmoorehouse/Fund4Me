<%@ Language = VBScript%>
<%

Set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set Lookup = Server.CreateObject("ADODB.recordset")

Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
	set oRs = Server.CreateObject("ADODB.recordset")			
SQL = "SELECT * from stores order by id desc;"

Lookup.open SQL, conn3 , 1, 3
lookup.movefirst
do while not lookup.EOF

		
	sSQL = "SELECT ProductName FROM stores WHERE (ID = " & lookup("ID") & ")"
	
	
	
	
	
	oRs.open sSQL, conn, 1,3
	
	oRs.MoveFirst
	
	
	if Lookup("shoppin") = "Coupons Only" or Lookup("shoppin") = "None" then
		ProductName = "<span id=""merchantlink"" title=""Find coupons and discounts for " & Lookup("merchant") & " with funds4me""><a href=""" & Lookup("merchant") & """>"& Lookup("merchant") &"</a></span>"
	else
		ProductName = "<span id=""merchantlink"" title=""" & Lookup("shoppin") & " Cashback from " & Lookup("merchant") & """><a href=""" & Lookup("merchant") & """>"& Lookup("merchant") &"</a></span>"
	end if

	oRs("ProductName") = ProductName
	response.Write(oRs("ProductName") & " - " & ProductName & "<br />" & chr(13))
	
	
	'linkit = Replace(Lookup("Link"),"383304","1292087",1, -1, 1)
	'oRs("Link") = linkit
	'response.Write(oRs("Link") & " - " & linkit & "<br />" & chr(13))
	
	oRs.Update
	oRs.close
	
	lookup.MoveNext
Loop
Lookup.close	

	
	'oRs("ProductName") = ProductName
	
	
	'
	
  	'oRs.close



%>
