<%
'
' Get the order number and password
'

memID = session("memID")
%>
<html>
<HEAD>
	<title>CashBack Shopping - Funds4me (UK)</title>
	
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<META NAME="description" CONTENT="Join the internets oldest Cashback shopping site. Funds4me cash back was Britains first cashback site and rated as one of the top payers."><META NAME="keywords" CONTENT="cash back, cashback, shopping">
	<meta name="author" content="Inka Maier, TiF Pure Solutions Limited, + 44 (0) 1482 611021">
	
	<meta name="rating" content="General">
	<link rel="stylesheet" type="text/css" href="styles.css">
</HEAD>	
<BODY BGCOLOR=#ffffff TOPMARGIN=28 LEFTMARGIN=18 MARGINWIDTH="18" MARGINHEIGHT="28">
<a href="javascript:history.back(1);"><img src="images/buttons/back.gif" border="0"></a><br>
<%
InvoiceNo = Request.QueryString("invoice")
sql = "select * from members Where memID = '" & memID & "'"
Dim objConn2
Set objConn2 = Server.CreateObject("ADODB.Connection")
objConn2.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
		
set Lookup = Server.CreateObject("ADODB.recordset")
Lookup.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
Lookup.CursorType = 1
Lookup.open SQL, objConn2



if Lookup.BOF and Lookup.EOF then
	Response.redirect("account.asp")
end if 
%>
<table width=100%>
<tr>
	<td valign=top>
		<FONT face=Arial size=2><b>Invoice From:</B><br><br>
		<FONT face=Arial size=2><%= Lookup("firstname") %> <%= Lookup("lastname") %></FONT><br>
		<FONT face=Arial size=2><%= Lookup("Address1") %><br>
		<FONT face=Arial size=2><%= Lookup("Address2") %><br>
		<FONT face=Arial size=2><%= Lookup("city") %><br>
		<FONT face=Arial size=2><%= Lookup("County") %><br>
		<FONT face=Arial size=2><%= Lookup("Country") %><br>
		<FONT face=Arial size=2><%= Lookup("Postcode") %><br>
		<FONT face=Arial size=2><%= Lookup("Email") %></FONT><br>
		<FONT face=Arial size=2><%= Lookup("phone") %></FONT><br><br>
	</td>
	<td valign=top>
		<FONT face=Arial size=2><b>Invoice To:</b><br><br>
		Shoppin.net<br>
		<br>
		<br>
		Historic Imported Data.<br><br>
	</td>
</tr>
</table>
<%
payto = Lookup("firstname") & " " & Lookup("lastname")
username = Lookup("username")
Lookup.close
		
SQL = "SELECT * FROM fundsraised WHERE ((Invoicemonth = '" & InvoiceNo & "') and (memID = " & cint(Session("memID")) & "))"
Lookup.open SQL, objConn2,1,3											
			
%>
<table border="0" width="800" >
  <tr>
    <td align="left"><FONT face=Arial size=2><b>Date: 16/<%= left(Lookup("Invoicemonth"),2)%>/<%= right(Lookup("Invoicemonth"),2)%></b></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
  </tr>
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2>&nbsp;</td>
  </tr>
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2>&nbsp;</td>
  </tr>
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2><b>RE: www.Funds4Me.co.uk </B>: <%=username%></td>
  </tr>

  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2>&nbsp;</td>
  </tr>
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2><hr></td>
  </tr><!--#INCLUDE FILE="displaypounds.inc"-->
  <tr>
  <%
  if Lookup.BOF and Lookup.EOF then
  else
  %>
  <td align="left" ><FONT face=Arial size=2><b>Date</b></td>
  <td align="left" ><FONT face=Arial size=2><b>Store</b></td>
  <td align="left" ><FONT face=Arial size=2><b>Product</b></td>
  <td align="center" ><FONT face=Arial size=2><b>Cash Back</b></td>
  </tr>
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2><hr></td>
  </tr>
  
  
  <tr>
  <%
  lookup.movefirst
  do while not lookup.eof
  %>
  <td align="left" ><FONT face=Arial size=2><%=Lookup("shopdate")%></td>
  <td align="left" ><FONT face=Arial size=2><%=Lookup("store")%></td>
  <td align="left" ><FONT face=Arial size=2><%=Lookup("product")%></td>
  <td align="center" ><FONT face=Arial size=2><b><%= Num2Pounds(trim(Lookup("fundsraised")))%></b></td>
  </tr>
  
  <tr>
  <%
  lookup.MoveNext
  loop
  end if
  
  Lookup.close
		
  sSQLFunds = "SELECT Invoicemonth, sum(fundsRaised) as Fund FROM fundsRaised WHERE (memID = " & cint(Session("memID")) & " and Invoicemonth = '" & InvoiceNo & "') GROUP BY Invoicemonth"

  Lookup.open sSQLFunds, objConn2,1,3	
  
  %>
  
  
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2><hr></td>
  </tr>
  <td align="Right" Colspan=3><FONT face=Arial size=2>Cash Back earned in invoice period.</td>
	<td align="center" ><FONT face=Arial size=2><b><%= Num2Pounds(trim(Lookup("Fund")))%></b></td>
  </tr>
  <tr>
    <td align="Right"><FONT face=Arial size=2>&nbsp;</td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2><hr></td>
  </tr>
 
 <tr>
	<td colspan=4>&nbsp;</td>
 </tr>
 
 
</table>



<% 
Lookup.Close

Session("memID") = memID

 %>
  
</center>
</BODY>

</html>