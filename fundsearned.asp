<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">
								<a href="javascript:history.back(1);"><img src="images/buttons/back.gif" border="0"></a><br><Br>
<!--#INCLUDE FILE="displaypounds.inc"-->
<%
sql = "select * from fundsraised Where schoolID = '" & Session("SchoolID") & "'"
Dim objConn2
Set objConn2 = Server.CreateObject("ADODB.Connection")
objConn2.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
		
set Lookup = Server.CreateObject("ADODB.recordset")
Lookup.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
Lookup.CursorType = 1

set Lookup1 = Server.CreateObject("ADODB.recordset")
Lookup1.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
Lookup1.CursorType = 1

SQL = "select * from fundsraised Where schoolID = '" & Session("SchoolID") & "'"
mon= Request.QueryString("mon")
if mon > "" then
SQL = SQL & " AND InvoiceMonth = '" & mon & "'"
end if
Lookup.open SQL,objConn2

If Lookup.BOF And Lookup.EOF Then %> <P></P>

<h3><FONT face=Arial >Sorry <br></FONT></h3>As yet your school has not earned any funds.
<% else %>


<center>
<table border="0" width="100%" >
  <tr>
    <td align="left" class=standart valign=top><b>Date:</b></td>
    <td align="left" class=standart valign=top><b>Store:</b></td>
    <td align="left" class=standart valign=top><b>Funds Raised:</b></td>
    <td align="left" class=standart valign=top><b>Invoiced:</b></td>
    <td align="left" class=standart valign=top><b>Raised By:</b></td>
  </tr>
  
  <% 
  runtotalcom = 0
  runtotal = 0
  Lookup.MoveFirst
  do while not Lookup.EOF
  %>
  <tr>
    <td align="left" class=standart valign=top><%= Lookup("shopDate") %></td>
    <td align="left" class=standart valign=top><%= Lookup("Store") %></td>
    <td align="left" class=standart valign=top><%= num2pounds(Lookup("fundsraised")) %></td>
    <td align="left" class=standart valign=top><%= Lookup("Invoiced") %></td>
  <%
  
	SQL = "select * from members Where memID = '" & Lookup("memID") & "'"
	Lookup1.open SQL,objConn2
	if Lookup1.BOF and Lookup1.EOF then
	raisedby = "Hidden"
	else
		if (Lookup1("report") = "yes") or (Lookup1("report") = "Yes") then
		rasiedby = Lookup1("FirstName") & " " & Lookup1("LastName") '& " (" & Lookup1("postcode") & ")"
		else
		rasiedby = "Hidden"
		end if
	end if
	Lookup1.Close
  %>  
  <td align="left" class=standart valign=top><%=rasiedby%></td>
  </tr>
 <%

 runtotalcom = runtotalcom + Lookup("fundsraised")
 Lookup.MoveNext
 loop
 %>
 <tr>
	<td colspan=5>&nbsp;</td>
 </tr>
 <tr>
	<td align="left"></td>
    <td align="left">Totals</td>
    <td align="left"><%= num2pounds(runtotalcom) %></td>
    <td align="left"></td>
    <td align="left"></td>
  </tr>
 
<% 
Session("SchoolID") = SchoolID


SQL = "select * from schools Where ID = '" & Session("SchoolID") & "'"
Lookup1.open SQL,objConn2

carriedforward = Lookup1("CarriedForward")

Lookup1.Close
	
if len(carriedforward) > 0 then
	mon= Request.QueryString("mon")
	if mon > "" then
	else
	%>
	<tr>
		<td colspan=5>&nbsp;</td>
	 </tr>
	 <tr>
		<td colspan=5>&nbsp;</td>
	 </tr>
	<tr>
		<td colspan=5>From your Invoiced Items above, you have <%= num2pounds(carriedforward) %> Carried forward, we only raise invoices when you are due a cheque now.
	</td>
	 </tr>
	<%
	end if
end if
	
%>

</table>
<% End If %>
  
	
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->