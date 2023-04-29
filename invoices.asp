<%
'
' Get the order number and password
'

memID = session("memID")
%>
<html>
<HEAD>
	<title>CashBack Shopping - Funds4me (UK)</title>
<!-- books, music, toys, videos, CDs, games &amp; electronics, computers, fashion, gifts, food wine, office &amp; business stores and more.Shopping on-line through Funds4Schools is easy. With over 200 well-known high street stores and on-line shops you'll find anything you want and with so many different special offers, you'll save money as well as time.What's more, you'll be raising funds for your favourite school without leaving your home. -->
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<META NAME="description" CONTENT="books, music, toys, videos, CDs, games &amp; electronics, computers, fashion, gifts, food wine, office &amp; business stores and more.Shopping on-line through Funds4Schools is easy. With over 200 well-known high street stores and on-line shops you'll find anything you want and with so many different special offers, you'll save money as well as time.What's more, you'll be raising funds for your favourite school without leaving your home."><META NAME="keywords" CONTENT="cash back, cashback, shopping">
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
payto = Lookup("firstname") & " " & Lookup("lastname")
username = Lookup("username")

%>
<table width=700 ID="Table1">
<tr>
	<td valign=top width=65%>
		<FONT face=Arial size=2><b>Invoice From:</B><br><br>
		<FONT face=Arial size=2><%= Lookup("firstname") %>&nbsp;<%= Lookup("lastname") %></FONT><br>
		<FONT face=Arial size=2><%= Lookup("Address1") %><br>
		<FONT face=Arial size=2><%= Lookup("Address2") %><br>
		<FONT face=Arial size=2><%= Lookup("city") %><br>
		<FONT face=Arial size=2><%= Lookup("County") %><br>
		<FONT face=Arial size=2><%= Lookup("Country") %><br>
		<FONT face=Arial size=2><%= Lookup("Postcode") %><br>
	</td>
	<td valign=top>
		<FONT face=Arial size=2><b>Invoice To:</b><br><br>
		Funds4Me.co.uk<br>TiF Pure Solutions Limited<br>
		Commerce House<br>
		62 Paragon Street<br>
		Kingston Upon Hull<br>
		East Yorkshire<br>
		HU1 3PW<br>
	</td>
</tr>
</table>
<%

Lookup.close

Dim objConn1
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("Connection2_ConnectionString"), Application("Connection2_RuntimeUserName"), Application("Connection2_RuntimePassword")
		
SQL = "SELECT * FROM invoices WHERE ((InvoiceNo = " & request.QueryString("invoice") & ") and id = '" & memID & "')"
Lookup.open SQL, objConn1,1,3											
			
%>
<table border="0" width="700" >
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2>&nbsp;</td>
  </tr>
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2>&nbsp;</td>
  </tr>  
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2><b>RE: Funds4Me.co.uk </B>: <%=username%></td>
  </tr>

  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2>&nbsp;</td>
  </tr>
  <tr>
    <td align="left" valign=top><FONT face=Arial size=2><b>Invoice No: </td>
    <td align="left" valign=top><FONT face=Arial size=2>F4MEUK-<%= Lookup("InvoiceNo") %></td>
    <td align="left" valign=top><FONT face=Arial size=2><b>Payment Date :</b><%= Lookup("Payment_Date") %></td>
    <td align="left"><FONT face=Arial size=2></td>
  </tr>
  <tr>
    <td align="left" ><FONT face=Arial size=2><b>Cheque payable to </b></td>
    <td align="left"><FONT face=Arial size=2><%= payto%></td>
    <td align="left" valign=top><FONT face=Arial size=2><b>Cheque No :</b><%= Lookup("Cheque_Number") %></b></td>
    <td align="left"><FONT face=Arial size=2></td>
  </tr>
   <tr>
    <td align="left" colspan=4><FONT face=Arial size=2>Last purchase date: <%= Lookup("InvoiceDate") %></td>
  </tr>
  <tr>
    <td align="left" colspan=4><FONT face=Arial size=2><hr></td>
    </tr><!--#INCLUDE FILE="displaypounds.inc"-->
  <tr>
  <%
  
  
  set Lookup1 = Server.CreateObject("ADODB.recordset")

  SQL = "SELECT * FROM fundsraised WHERE ((InvNo = '" & InvoiceNo & "') and (memID = " & cint(Session("memID")) & "))"
  Lookup1.open SQL, objConn2,1,3
  if Lookup1.BOF and Lookup1.EOF then
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
  lookup1.movefirst
  do while not lookup1.eof
  %>
  <td align="left" ><FONT face=Arial size=2><%=Lookup1("shopdate")%></td>
  <td align="left" ><FONT face=Arial size=2><%=Lookup1("store")%></td>
  <td align="left" ><FONT face=Arial size=2><%=Lookup1("product")%></td>
  <td align="right" ><FONT face=Arial size=2><b><%= Num2Pounds(trim(Lookup1("fundsraised")))%></b></td>
  </tr>
  
  <tr>
  <%
  lookup1.MoveNext
  loop
  end if
  
  Lookup1.close
  
  %>
  
  
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2><hr></td>
  </tr>
  
  
  
  
  
  <td align="Right" Colspan=3><FONT face=Arial size=2>Funds earned in invoice period</td>
	<td align="right" ><FONT face=Arial size=2><b><%= Num2Pounds(trim(Lookup("rebate_Amount")))%></b></td>
  </tr>
  
  <tr>
    <td align="Right" Colspan=3><FONT face=Arial size=2>Admin Fee</td>
	<td align="right" ><FONT face=Arial size=2><b><%= num2pounds(trim(Lookup("Bill_Amount")))%></b></td>
  </tr>
  <tr>
    <td align="left"><FONT face=Arial size=2>&nbsp;</td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2><hr>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
  </tr>
  <tr>
    <td align="Right" Colspan=3><FONT face=Arial size=2>Sub Total</td>
	<td align="right" ><FONT face=Arial size=2><b><%= num2pounds(trim(Lookup("Amount")))%></b></td>
  </tr>
  <tr>
    <td align="Right"><FONT face=Arial size=2>&nbsp;</td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2><hr></td>
  </tr>
  <tr>
    <td align="Right" Colspan=3><FONT face=Arial size=2>
    <%if trim(Lookup("Cheque_Number")) = "Carried Forward" then%>
    <b>Carried Forward</b></td>
    <%else%>
    <b>Cheque Issued For</b></td>
    <%end if%>
	<td align="right" ><FONT face=Arial size=2><b><%= num2pounds(trim(Lookup("Invoice_Amount")))%></b></td>
  </tr>
  <tr>
    <td align="left"><FONT face=Arial size=2>&nbsp;</td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2><hr></td>
  </tr>
	<tr>
    <td align="left"><FONT face=Arial size=2>&nbsp;</td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
    <td align="left"><FONT face=Arial size=2></td>
  </tr>
  
 
</table>



<% 
Lookup.Close

 %>
  
</center>
</BODY>

</html>