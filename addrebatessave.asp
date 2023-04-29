<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	memberID = Trim(Request.QueryString("memID"))
	if memberID > "" then
		sUpdate = "yes"
	else
		sUpdate = "no"
	end if
	
	Dim conn, oRs, oRs2
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	set oRs = Server.CreateObject("ADODB.recordset")

	Set conn2 = Server.CreateObject("ADODB.Connection")
	conn2.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	set oRs2 = Server.CreateObject("ADODB.recordset")
	
	Set conn3 = Server.CreateObject("ADODB.Connection")
	conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")

	store = Trim(Request.Form("store"))
	sql = "select * from stores WHERE (ID = " & store & ");" 

	set rs = Server.CreateObject("ADODB.recordset")
	rs.CursorLocation = 3 
	rs.CursorType = 1

	rs.open SQL, conn3, 1,3

	storelink = "<a href='shoprequired.asp?shop=" & store & "'>" & rs("Merchant") &"</a>"
	rs.Close
	
	memID = Trim(Request.Form("member"))
	product = Trim(Request.Form("Product"))
	if product = "" then product = "Order"
	fundsraised = Trim(Request.Form("fundsraised"))
	currencypaid = Trim(Request.Form("currency"))
	shopDate = Trim(Request.Form("shopdate"))
	
	
	invoicemonth = month(date)
	if len(cint(invoicemonth)) = 1 then
	invoicemonth = "0" & invoicemonth
	end if
	invoiceyear = year(date)
	if len(invoiceyear) = 4 then
	invoiceyear = right(invoiceyear,2)
	end if
	
	invoiceuse = invoicemonth & invoiceyear
	invoicecalc = invoiceyear & invoicemonth
	'netraised = Num2Pounds((fundsraised / 117.5)*100)
	netraised = fundsraised
	
	
		
		sSQL = "SELECT * FROM fundsraised WHERE memID = 999999999"
		oRs.open sSQL, conn, 1,3
					
		oRs.AddNew
	
		oRs("store") = storelink
		oRs("memID") = memID
		oRs("schoolID") = "Funds4Me.co.uk"
		oRs("fundsraised") = netraised
		oRs("product") = product
		oRs("site") = "Funds4Me.co.uk"
		oRs("currencypaid") = currencypaid
		oRs("shopDate") = shopDate
		oRs("invoicemonth") = invoiceuse
		oRs("invoicecalc") = invoicecalc
		oRs("DateEntered") = day(date) & "/" & month(date) & "/" & year(date)
		oRs("invoiced") = "No"
		oRs.update()
	
  		oRs.close()
  		text = memID & " - " & schoolID & " - " & storelink & " - " & currencypiad & netraised
	
	
  	Response.Redirect("addrebates.asp?adminidid=Saved "& text)

%>
<!--#INCLUDE FILE="displaypounds.inc"-->