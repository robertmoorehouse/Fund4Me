<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	iMainMenu = 0
	iSchoolMenu = 0 
	
if session("memID") = "29" then

	
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
	<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

	<!--#INCLUDE FILE="header.inc"-->
	<table width="800" cellspacing="0" cellpadding="0">
		<tr>
			<td>
				<img src="images/pix.gif" width="517" height="10">
			</td>
			<td bgcolor="#99ccff">
				<img src="images/pix.gif" width="142" height="10">
			</td>
			<td bgcolor="#CDE4F6">
				<img src="images/pix.gif" width="141" height="10">
			</td>
		</tr>
		<tr>
			<td valign="top" align="left" valign="top">
				<table width="517" cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<img src="images/pix.gif" width="20" height="450">
						</td>
						<td align="center" valign="top">
	<!-- Main Text area -->
<%
id = Request.QueryString("adminidid")

	if id > "" then

%>
		
		<table border="0" cellpadding="0" cellspacing="0" width="490" valign="top">
	      <tr>
	        <td width="490" class="standart" valign="top" align="center">
	          <span class="blueText"><font size="2">
				<b>Register Commissions</b></font>
			  </span>
	        </td>
	      </tr>

		  <tr>
	        <td width="100%" class="standart" align="left">
	          <br>
	         <span class="standart">

			</span><br><br>
	        </td>
	      </tr>
	      <tr><td>
	      <%
			Dim conn, oRs, oRs2
			SQL = "select * from stores order by merchant"

			Set conn = Server.CreateObject("ADODB.Connection")
			conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
			set rs = Server.CreateObject("ADODB.recordset")
			rs.CursorLocation = 3 
			rs.CursorType = 1
							
			rs.open SQL, conn, 1,3
			
			Set conn = Server.CreateObject("ADODB.Connection")
			conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
			set oRs2 = Server.CreateObject("ADODB.recordset")
			oRs2.CursorLocation = 3 
			oRs2.CursorType = 1
			sSQL = "SELECT * FROM members where site = 'Funds4Me.co.uk'"
			oRs2.open sSQL, conn, 1,3	
			%>
	      <form name="regiForm" action="addrebatessave.asp" method="post" >
	                        <table cellSpacing="0" cellPadding="0" width="490" border="0">
	                          <tr id="small" vAlign="middle" bgcolor="#ffcc99">
	                            <td><img src="images/pix.gif" width="8" height="1" border="0"></td>
	                            <td class="standart">Store:&nbsp;</td>
	                            <td class="standart">
								  <select size="1" name="store" ID="Select1"> 
										<%
										do while not rs.EOF
							            
										%>
										<option value="<%=rs("ID")%>"><%=trim(rs("Merchant"))%></option>
										<%
										rs.MoveNext
										loop
										rs.Close
										%>
									</select>
	                              </td>
	                            </tr>
	                            <tr id="small" vAlign="middle" bgcolor="#ffcc99">
	                            <td><img src="images/pix.gif" width="8" height="1" border="0"></td>
	                            <td class="standart">Member:&nbsp;</td>
	                            <td class="standart">
								  <select size="1" name="member"> 
								    <%
	                                do while not ors2.EOF
	                                %>
	                                <option value="<%=ors2("memID")%>"><%=ors2("memID")%></option>
	                                <%
	                                ors2.MoveNext
	                                loop
	                                ors2.close
									%>
	                              </select>
	                              </td>
	                            </tr>
	                           
								<tr id="small" vAlign="middle" bgcolor="#ffcc99">
	                            <td><img src="images/pix.gif" width="8" height="1" border="0"></td>
	                            <td class="standart">Product:&nbsp;</td>
	                            <td class="standart">
								  <input maxLength="8" size="30" name="Product" Value="" ID="Text1">
	                              </td>
	                            </tr>
	                            <tr id="small" bgcolor="#ffcc99">
	                              <td><img src="images/pix.gif" width="8" height="1" border="0"></td>
	                              <td class="standart">Funds Raised:</td>
	                              <td class="standart"><input maxLength="8" size="30" name="fundsraised" Value="0.00"></td>
	                            </tr>
	                            <tr id="small" bgcolor="#ffcc99">
	                              <td><img src="images/pix.gif" width="8" height="1" border="0"></td>
	                              <td class="standart">currency:</td>
	                              <td class="standart"><select size="1" name="currency">
								  <option value="£" SELECTED>£</option>
								  <option value="$">$</option></td>
	                            </tr>
	                            <tr id="small" bgcolor="#ffcc99">
	                              <td><img src="images/pix.gif" width="8" height="1" border="0"></td>
	                              <td class="small">Date: dd/mm/yyyy</td>
	                              <td class="small"><input maxLength="30" size="30" name="shopdate" Value="<%=date()%>"></td>
	                            </tr>                           
	                            <tr id="small" bgcolor="#FFFFFF">
	                              <td><img src="images/pix.gif" width="8" height="1" border="0"></td>
									<td class="small" align="left" colspan="2">
									<input type="image" alt="Search" hspace="0" src="images/buttons/searchOrange.gif" border="0" id="image1" name="image1">
								  </td>
								</tr>
	                     </table>   
						 </form>
	      
	      
	<%

	%>
	<table width="490" cellpadding="0" cellspacing="0">
			 <tr><td colspan="3" class="standart" bgcolor="#FFFFFF"> 
			
			</td></tr>
			<tr><td colspan="3" class="standart" bgcolor="#FFFFFF">
			</td></tr>
		<tr>
			<td class="WhiteText" bgcolor="#ff9900">
				<b><i>Last Item</b></i>
			</td>
			<td class="WhiteText" bgcolor="#ff9900">
				<img src="images/pix.gif" width="20" height="1" border="0">
			</td>
			<td class="WhiteText" bgcolor="#ff9900">
				<font color="white"></font>
			</td>
		</tr>
		<tr><td colspan="3" class="standart" bgcolor="#FFFFFF"><%=Request.QueryString("adminidid")%>
			</td></tr>
		</table>
		

	    
	      </td></tr>
	      
	      </table>

	<br><br>

							

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
	<%
						else
	%>
							<table border="0" cellpadding="0" cellspacing="0" width="490" valign="top">
							<tr>
		  <td colspan="3" align="left" valign="top">
			<a href="javascript:history.back(1);"><img src="images/buttons/back.gif" border="0"></a>
			<br><br>
		  </td>
		</tr>	
							<tr>
								<td width="490" class="standart" valign="top" align="center">
								<span class="blueText"><font size="2">
									<b>Amend Uninvoiced Commissions</b></font>
								</span>
								</td>
							</tr>

							<tr>
								<td width="100%" class="standart" align="left">
								<br>
								<span class="standart">

								</span><br><br>
								</td>
							</tr>
							<tr><td>
							<%
								idamend = request.QueryString("idamend")
								if idamend > "" then
									
									
									SQL = "select * from stores order by merchant"

									Set conn = Server.CreateObject("ADODB.Connection")
									conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
									set rs = Server.CreateObject("ADODB.recordset")
									rs.CursorLocation = 3 
									rs.CursorType = 1
													
									rs.open SQL, conn, 1,3
									
									
									Set conn1 = Server.CreateObject("ADODB.Connection")
									conn1.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
									
									set oRs3 = Server.CreateObject("ADODB.recordset")
									oRs3.CursorLocation = 3 
									oRs3.CursorType = 1
									sSQL = "SELECT * FROM fundsraised where ID= " & idamend
									oRs3.open sSQL, conn1, 1,3
									
									%>
									<form name="regiForm" action="addrebatessave.asp?adminidid=adminamend" method="post" ID="Form1">
												<table cellSpacing="0" cellPadding="0" width="490" border="0" ID="Table1">
													<tr id="small" bgcolor="#ffcc99">
													<td><img src="images/pix.gif" width="8" height="1" border="0"></td>
													<td class="standart">memID:</td>
													<td class="standart"><input type=hidden name=idamend value="<%=idamend%>"><input maxLength="8" size="30" name="memID" value="<%=oRs3("memID")%>" ID="Text4"></td>
													</tr>
													
													<tr id="small" vAlign="middle" bgcolor="#ffcc99">
													<td><img src="images/pix.gif" width="8" height="1" border="0"></td>
													<td class="standart">Store:&nbsp;</td>
													<td class="standart">
													<select size="1" name="store" ID="Select1">
														<option value="" selected>RE SELECT <%=trim(oRs3("store"))%></option>
														<%
														do while not rs.EOF
							                            
														%>
														<option value="<%=rs("ID")%>"><%=trim(rs("Merchant"))%></option>
														<%
														rs.MoveNext
														loop
														rs.Close
														%>
													</select>
													</td>
													</tr>
													<tr id="small" bgcolor="#ffcc99">
													<td><img src="images/pix.gif" width="8" height="1" border="0"></td>
													<td class="standart">Funds Raised:</td>
													<td class="standart"><input maxLength="8" size="30" name="fundsraised" value="<%=oRs3("fundsraised")%>" ID="Text2"></td>
													</tr>
													<tr id="small" bgcolor="#ffcc99">
													<td><img src="images/pix.gif" width="8" height="1" border="0"></td>
													<td class="standart">currency:</td>
													<td class="standart"><select size="1" name="currency" ID="Select4">
													<option value="<%=oRs3("currencypaid")%>" selected><%=oRs3("currencypaid")%></option>
													<option value="£" >£</option>
													<option value="$">$</option></td>
													</tr>
													<tr id="small" bgcolor="#ffcc99">
													<td><img src="images/pix.gif" width="8" height="1" border="0"></td>
													<td class="small">Date: dd/mm/yyyy</td>
													<td class="small"><input maxLength="30" size="30" name="shopdate" Value="<%=oRs3("shopdate")%>" ID="Text3"></td>
													</tr>
													<tr id="small" bgcolor="#ffcc99">
													<td><img src="images/pix.gif" width="8" height="1" border="0"></td>
													<td class="small">InvoiceMonth: mmyy</td>
													<td class="small"><input maxLength="30" size="30" name="InvoiceMonth" Value="<%=oRs3("InvoiceMonth")%>" ID="Text5"></td>
													</tr>                           
													<tr id="small" bgcolor="#FFFFFF">
													<td><img src="images/pix.gif" width="8" height="1" border="0"></td>
														<td class="small" align="left" colspan="2">
														<input type="image" alt="Search" hspace="0" src="images/buttons/searchOrange.gif" border="0" id="Image2" name="image1">
													</td>
													</tr>
											</table>   
											</form>
											<%
	end if
	end if
	%>
<%
else
Response.Redirect("default.asp")
end if
%>		
</body>
</html>