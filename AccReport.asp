<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	iMainMenu = 6
	iSchoolMenu = 0 
	
	if Session("memID") = "" then
		Response.Redirect("account.asp")
	end if
	detailed = "no"
	if Request.QueryString("report") > "" then
		detailed = Request.QueryString("report")
		if Request.QueryString("report") = "all" then
			detsql = "" 
		else
			detsql = " and invno = '" & detailed & "'"
		end if
	end if
	if Request.QueryString("report") = "shop" then
		detailed = Request.QueryString("report")
		detsql = " and len(invoiced) = 2 "
	end if
	Dim conn, oRs, oRs2
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	
	Set SchoolConn = Server.CreateObject("ADODB.Connection")
	SchoolConn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	
	
	Set RebateConn = Server.CreateObject("ADODB.Connection")
	RebateConn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	
	set oRsRebates = Server.CreateObject("ADODB.recordset")
	
	if session("memID") = "29" then
		sSQLRebates = "SELECT sum(fundsraised) as Com FROM fundsraised"
	else
		sSQLRebates = "SELECT sum(fundsraised) as Com FROM fundsraised WHERE memID = '" & Session("memID") & "'"
	end if
	oRsRebates.open sSQLRebates, RebateConn, 1,3
	total = oRsRebates("com")
	oRsRebates.close
	set oRsRebates=nothing
	
	
	set oRsFunds = Server.CreateObject("ADODB.recordset")	
			
	if detailed = "no" then	
		if session("memID") = "29" then
			sSQLFunds = "SELECT invno, sum(fundsRaised) as Fund FROM fundsRaised GROUP BY invno"
		else
			sSQLFunds = "SELECT invno, sum(fundsRaised) as Fund FROM fundsRaised WHERE memID = " & cint(Session("memID")) & " GROUP BY invno"
		end if					
		oRsFunds.open sSQLFunds, conn, 1,3
		
	else
		if session("memID") = "29" then
			sSQLFunds = "SELECT * FROM fundsraised Order BY shopdate"
		else
			sSQLFunds = "SELECT * FROM fundsraised WHERE (memID = " & cint(Session("memID")) & detsql & ") Order BY shopdate"
		end if
		
		oRsFunds.open sSQLFunds, RebateConn, 1,3
		
	end if		
	
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
<script Language="JavaScript"><!--

	function RollOver(imgName,num)
	{ 
		document.images[imgName].src = "images/buttons/accMenu" + num + "On.gif"
	}
	function RollOut(imgName,num)
	{ 
		document.images[imgName].src = "images/buttons/accMenu" + num + ".gif"
	}
-->
</script>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<!--#INCLUDE FILE="header.inc"-->
<table width="800" cellspacing="0" cellpadding="0">
	<tr>
		<td><img src="images/pix.gif" width="517" height="1"></td>
		<td bgcolor="#99CCFF"><img src="images/pix.gif" width="142" height="1"></td>
		<td bgcolor="#CDE4F6"><img src="images/pix.gif" width="141" height="1"></td>
	</tr>
	<tr>
		<td valign="top" align="left" valign="top">
			<table width="498" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<img src="images/pix.gif" width="1" height="450">
					</td>
					<td align="center" valign="top">
<!-- Main Text area -->
	<table border="0" cellpadding="0" cellspacing="0" width="498" valign="top">
      <tr>
        <td width="498" class="standart" valign="top" align="center">
			<table width="490" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ffffff">
				<td class="WhiteText" width="90" valign="top" align="left"><!--<a href="AccSchools.asp"><img border="0" src="images/buttons/accMenu1.gif" name="men1" OnMouseOver="javascript:RollOver('men1',1);" OnMouseOut="javascript:RollOut('men1',1);" width="90"></a><a href="AccProfile.asp"><img border="0" src="images/buttons/accMenu2.gif" width="90" name="men2" OnMouseOver="javascript:RollOver('men2',2);" OnMouseOut="javascript:RollOut('men2',2);"></a><a href="AccReport.asp"><img border="0" src="images/buttons/accMenu3.gif" width="90" name="men3" OnMouseOver="javascript:RollOver('men3',3);" OnMouseOut="javascript:RollOut('men3',3);"></a>-->
</td>
				<td colspan="2" class="BlueText" width="408" valign="top" align="center">
					<img src="images/pix.gif" width="408" height="10"><br>
					<b>Welcome <%=Session("memName")%></b><br><br><br>
					<font size="2"><b><i>My Report</i></b></font>
				</td>
			</tr>
			</table>
			</table>
	<form name="SchoolsForm" action="AccSchools.asp" method="post">
		<table width="440" cellpadding="0" cellspacing="0">
		<tr bgcolor="#ffffff">
				<td colspan="5" class="standart" valign="middle" align="left">
					<b>Summary of your Cash Back raised to date.</b><br>
					Please note that it can take up to 6 weeks until we get the reports from the stores so your funds raised will not immediately appear on your statement after you shop online.
					<br><br>
				<% if detailed = "no" then %>
					<a href="AccReport.asp?report=shop"><img border="0" src="images/buttons/detailed.gif"></a><br><br>
					<% else %>
					<a href="AccReport.asp?"><img border="0" src="images/buttons/summary.gif"></a><br><br>
					<% end if %>
				</td>
				
			</tr>
			<tr bgcolor="#ff9900">
				<% if detailed = "no" then%>
				
				
						<td colspan="4" class="WhiteText" valign="top" align="left">
							
							<b><i>Invoice</i></b>
							
						</td>
						
						
																
						<td class="WhiteText" valign="top" align="center">
							<b><i>Cash Back Earned</i></b>
						</td>
						
				</tr>
				
	<%			
				if oRsFunds.bof and oRsFunds.eof then
					
				else
					oRsFunds.MoveFirst
					i = 0
					Do while not oRsFunds.EOF
						i = i + 1
						
							
						
		%>			
						<tr bgcolor="#ffcc99">
							
							<td colspan="4" class="LargeBlack" valign="top" align="left">
								<a href="AccReport.asp?report=<%=oRsFunds("invno")%>"><%=oRsFunds("invno")%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
							
							<td class="LargeBlack" valign="top" align="center">
								<%=num2pounds(oRsFunds("Fund"))%>
							</td>
												
						</tr>
						
		<%				
					oRsFunds.MoveNext	
					Loop			
					if i > 0 then
					%>
					<tr bgcolor="#ffcc99">
							
					<td colspan="4" class="LargeBlack" valign="top" align="left">
						<br><a href="AccReport.asp?report=all">View-All</a>
					<Br><Br></td>
							
					<td class="LargeBlack" valign="top" align="center">
						
					</td>
												
					</tr>
					
					<%
					end if
					
				end if
			else
				%>
				
				<td class="WhiteText" valign="top" align="left">
					<b><i>Store</i></b>
				</td>
				<td class="WhiteText" >
					&nbsp;&nbsp;&nbsp;
				</td>
				<td class="WhiteText" valign="top" align="left">
					<b><i>Purchased</b></i>
				</td>
				
				<td class="WhiteText" valign="top" align="center">
					<b><i>Cash Back </i></b>
				</td>
				<td class="WhiteText" valign="top" align="center">
					<b><i>Paid</i></b>
				</td>
				</tr>
				
	<%			
				if oRsFunds.bof and oRsFunds.eof then
					
				else
					oRsFunds.MoveFirst
					i = 0
					Do while not oRsFunds.EOF
						i = i + 1
					%>			
						<tr bgcolor="#ffcc99">
							
							<td class="LargeBlack" valign="top" align="left">
								<%=oRsFunds("Store")%>
							</td>
							<td class="WhiteText" >
								&nbsp;&nbsp;&nbsp;
							</td>
							<td class="LargeBlack" valign="top" align="left">
								<%=oRsFunds("shopdate")%>
							</td>
							<td class="LargeBlack" valign="top" align="center">
<%
							temp = oRsFunds("fundsraised")
							%>
								<%=oRsFunds("currencypaid")%><%=temp%>

							</td>
							<td class="LargeBlack" valign="top" align="center">
							<%	if len(oRsFunds("invoiced")) = 3 then
									if oRsFunds("dateentered") = "21/09/2003" then	%>
										<a href="invoicesshoppin.asp?invoice=<%=oRsFunds("Invoicemonth")%>">Yes
							<%		else 
										
							%>
										<a href="invoices.asp?invoice=<%=oRsFunds("invno")%>">Yes
							<%		end if
							
								else %>
									No
								<%
								
								end if
							%>
							</td>
						
						</tr>
						
				
		<%				
					oRsFunds.MoveNext	
					Loop			
					
				end if
			end if
	%>			
				
	
						<tr bgcolor="#ffffff">
				<td colspan="5" class="standart" valign="middle" align="left">
					<br><br><b>Commission Payment Dates Subject to Minimum cheque amount.</b><br>
					 Earned January is paid 5th May.<br>
					 Earned February is paid 5th June.<br>
					 Earned March is paid 5th July.<br>
					 Earned April is paid 5th August.<br>
					 Earned May is paid 5th September.<br>
					 Earned June is paid 5th October.<br>
					 Earned July is paid 5th November.<br>
					 Earned August is paid 5th December.<br>
					 Earned September is paid 5th January.<br>
					 Earned October is paid 5th February.<br>
					 Earned November is paid 5th March.<br>
					 Earned December is paid 5th April.<br>
					<br><br>
				</td>
				
			</tr>
			
		</table>		
	</form> 			
      
      
<%
	oRsFunds.Close()
	Set oRsFunds = nothing
	conn.Close()
	SchoolConn.close
	RebateConn.close
%>						
<!--#INCLUDE FILE="displaypounds.inc"-->
<!-- End of Main Text -->
					</td>
				</tr>
			</table>

		</td>
		<td bgcolor="#99CCFF" valign="top" align="center">
			<%  ' Supporters/Retailers/Specials
			
				if iRightMenu = 1 then
			%>	  <!--#INCLUDE FILE="coupons.inc"-->
			<%	end if
				if iRightMenu = 2 then
			%>	  <!--#INCLUDE FILE="retailer.inc"-->
			<%	end if
				if iRightMenu = 3 then
			%>	  <!--#INCLUDE FILE="specials.inc"-->
			<%	end if			
			%>
		</td>
		<td bgcolor="#CDE4F6" valign="top" align="center">
			<!--#INCLUDE FILE="adverts.inc"-->
		</td>
	</tr>
	<tr>		
		<td colspan="4">
			<!--#INCLUDE FILE="footer.inc"-->
		</td>
	</tr>
</table>	
</body>
</html>
