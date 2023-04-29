<%@ Language = VBScript%>
<%
	iMainMenu = 1
	iSchoolMenu = 0 

%>
<HTML>
<HEAD>
	<title>CashBack Shopping - Funds4me (UK)</title>
	<!-- books, music, toys, videos, CDs, games &amp; electronics, computers, fashion, gifts, food wine, office &amp; business stores and more.Shopping on-line through Funds4Me is easy. With over 200 well-known high street stores and on-line shops you'll find anything you want and with so many different special offers, you'll save money as well as time.What's more, you'll be raising funds for your favourite school without leaving your home. -->
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<META NAME="description" CONTENT="Join the internets oldest Cashback shopping site. Funds4me cash back was Britains first cashback site and rated as one of the top payers.">
	<META NAME="keywords" CONTENT="cash back, cashback, shopping">
	<meta name="author" content="Robert Moorehouse">
	<meta name="rating" content="General">
	<link rel="stylesheet" type="text/css" href="styles.css">
<style type="text/css">
	body { font-family: Arial, Helvetica, Sans-serif; font-size: 9px; color: #000000; }
</style>
<script src="js/funds4_nav.js"></script>
<script src="js/funds4_cookie_detect.js"></script>

</HEAD>
	<body bgcolor="#ffffff" link="#006600" alink="#339900" vlink="#CC3300" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"><!--#INCLUDE FILE="headerdefault.inc"-->
		<table width="800" cellspacing="0" cellpadding="0">
			<tr>
				<td>
				</td>
				<td bgcolor="#99ccff">
					<IMG height=10 src="images/pix.gif" width=142>
				</td>
				<td bgcolor="#cde4f6">
					<IMG height=10 src="images/pix.gif" width=141>
				</td>
			</tr>
			<tr>
				<td valign="top" align="left">
					<table width="517" cellpadding="0" cellspacing="0">
						<tr><td colspan=2><IMG src="images/pix.gif" alt="cash back shopping" width=517 height=10></td>
						<tr>
						<tr>
							<td>
								<IMG src="images/pix.gif" alt="cash back shopping" width=20 height=450>							</td>
							<td align="middle" valign="top"><!-- Main Text area -->
								<span class="blueText">
								<% if Session("memID") > "" then %>
									              <span class="blueText">          Welcome <%=Session("memName")%>
								<% else
										if id > "" then %>
										<span class="blueText"><b>You currently raise funds for <%=name%>
										</b>
									<%	else %>          
									<span class="Largeblue"><b>Please Log in or register to earn Cash Back</b>
									<%	end if 
									end if%>
								</span>
								<br>
								<br>
								<table width="497" cellpadding="0" cellspacing="0">
									<tr>
										<td class="standart" width="235" valign="top" align="left"><img src="images/freebalance.jpg"><br><br>
											Shopping on-line through <b>Funds4Me</b> is easy. With over 450 well-known <b>high 
												street stores</b>&nbsp;and <b>on-line shops</b> 
                  you'll find anything you want and with so many different special
											      offers, you'll save money as well 
											as time.<br>
											<br>
											<form name="LoginForm" action="account.asp" method="post">
												<table width="230" cellpadding="0" cellspacing="0">
													<tr bgcolor="#0099cc">
														<td class="WhiteText" width="8" valign="top" align="left">
															<IMG height=1 src="images/pix.gif" width=8>
														</td>
														<td colspan="2" class="WhiteText" width="222" valign="top" align="left">
															<b><i>Member Login</i></b>
														</td>
													</tr>
													<tr bgcolor="#99ccff">
														<td class="WhiteText" width="8" valign="top" align="left">
															<IMG height=1 src="images/pix.gif" width=8>
														</td>
														<td class="Standart" width="111" valign="top" align="left">
															Email:<br>
															<input name="Username" size="13">
														</td>
														<td class="Standart" width="111" valign="top" align="left">
															Password:<br>
															<input type="password" name="Password" size="13">
														</td>
													</tr>
													<tr bgcolor="#99ccff">
														<td class="WhiteText" width="5" valign="top" align="left">
															<IMG height=1 src="images/pix.gif" width=5>
														</td>
														<td class="WhiteText" width="113" valign="top" align="left">
															<A href="JavaScript:document.LoginForm.submit();"><IMG src="images/buttons/login.gif" border=0></A>
														</td>
														<td class="WhiteText" width="112" valign="top" align="right">
															<A href="register2.asp"><IMG src="images/buttons/register.gif" border=0></A>
														</td>
													</tr>
												</table>
											</form>
											<IMG src="images/home1.gif" border=0>
										</td>
										<td width="17" valign="top" align="left">
											<IMG height=10 src="images/pix.gif" width=17>
										</td>
										<td class="standart" width="235" valign="top" align="left">
											   What's more, you'll be&nbsp;earning 
                  <STRONG>CASH&nbsp;BACK</STRONG> on your normal purchases from 
                  over 450 stores.<BR><BR><STRONG>Membership is 
                  FREE!</STRONG>&nbsp;Complete the quick register box below to 
                  start earning Cash Back<BR><BR>
											   There is no catch: every time you shop on-line, a cash rebate, 
											worth up to 30% of the amount you have spent, will be paid back to you.<br>
											<br>
											<font color="#ff9900"><b><i>To&nbsp;Earn CASH BACK when you shop:</i></b></font><br>
											All 
                  you have to do is follow our links every time you go shopping. 
                  By following our links we are paid a sales commission, this is 
                  then paid back to you!<BR>
											             
											      &nbsp;
											<br>
											               
											        
											<form name="SearchForm" action="Stores.asp?sub=yes" method="post">
												<table width="230" cellpadding="0" cellspacing="0">
													<tr bgcolor="#ff9900">
														<td class="WhiteText" width="8" valign="top" align="left">
															<IMG height=1 src="images/pix.gif" width=8>
														</td>
														<td colspan="2" class="WhiteText" width="222" valign="top" align="left">
															<b><i>Store&nbsp;Search</i></b>
														</td>
													</tr>
													<tr bgcolor="#ffcc99">
														<td colspan="3" class="WhiteText" width="230" valign="top" align="left"><IMG height=4 src="images/pix.gif" width=230></td>
													</tr>
													<tr bgcolor="#ffcc99">
														<td class="WhiteText" width="8" valign="top" align="left">
															<IMG height=1 src="images/pix.gif" width=8>
														</td>
														<td class="Standart" width="90" valign="top" align="left">
															 Store&nbsp;Name:
														</td>
														<td class="WhiteText" width="132" valign="top" align="left">
															<input name="formStore" size="17">
														</td>
													</tr>
																										
													<tr bgcolor="#ffcc99">
														<td class="WhiteText" width="8" valign="top" align="left">
															<IMG height=1 src="images/pix.gif" width=8>
														</td>
														<td class="WhiteText" width="8" valign="top" align="left">
															<A href="JavaScript:document.SearchForm.submit();"><IMG src="images/buttons/search.gif" border=0></A>
														</td>
														<td class="WhiteText" width="222" valign="top" align="right">
															<A href="stores.asp"><IMG src="images/buttons/searchAdv.gif" border=0></A>
														</td>
													</tr>
												</table>
											</form>
										</td>
										<td width="10" valign="top" align="left">
											<IMG height=10 src="images/pix.gif" width=10>
										</td>
									</tr>
								</table><!-- End of Main Text -->
								
							</td>
						</tr>
					</table>
				</td>
				<td bgcolor="#99ccff" valign="top" align="left">
					<center><b><u>Top Stores</u></b></center>
					<%
					Set conn3 = Server.CreateObject("ADODB.Connection")
					conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
		
								
					set Lookup5 = Server.CreateObject("ADODB.recordset")
					Lookup5.CursorLocation = 3 
					Lookup5.CursorType = 1
												
								

					SQL = "SELECT top 10 id,merchant,shoppin, ListCountry,ProductName FROM stores where ((stores.ListCountry like '%,uk,%') and (shoppin <> 'Coupons Only')) order by ID;"
					'Response.Write(SQL)
					Lookup5.open SQL, conn3
								
					%>
					<br><table>
					<% do while not Lookup5.EOF
						temp = trim(Lookup5("Merchant")) + " " + trim(Lookup5("shoppin"))			
						
						if len(temp) < 20 then
						%>
						<tr>
							<td class="tinyblack"><%=trim(Lookup5("ProductName"))%></td>
							<td class="tinyblack1"><b><%=trim(Lookup5("shoppin"))%></b></td>
						</tr>
						<%
						end if
						Lookup5.MoveNext
						loop
						Lookup5.Close
						%>
						</table>
						<br>
						<center><b><u>New Stores</u></b></center>
					<%
					Set conn3 = Server.CreateObject("ADODB.Connection")
					conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
		
								
					set Lookup5 = Server.CreateObject("ADODB.recordset")
					Lookup5.CursorLocation = 3 
					Lookup5.CursorType = 1
												
								

					SQL = "SELECT top 10 id,merchant,shoppin, ListCountry,ProductName FROM stores where ((stores.ListCountry like '%,uk,%') and (shoppin <> 'Coupons Only')) order by ID Desc; "
					'Response.Write(SQL)
					Lookup5.open SQL, conn3
								
					%>
					<br><table>
					<% do while not Lookup5.EOF
						temp = trim(Lookup5("Merchant")) + " " + trim(Lookup5("shoppin"))			
						
						if len(temp) < 20 then
						%>
						<tr>
							<td class="tinyblack"><%=trim(Lookup5("ProductName"))%></td>
							<td class="tinyblack1"><b><%=trim(Lookup5("shoppin"))%></b></td>
						</tr>
						<%
						end if
						Lookup5.MoveNext
						loop
						Lookup5.Close
						%>
						</table>
						
						
						<br>
						<center><b><u>Categories (UK)</u></b></center>
						
									<%
						
									Set conn = Server.CreateObject("ADODB.Connection")
									conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
									set Lookup2 = Server.CreateObject("ADODB.recordset")
									Lookup2.CursorLocation = 3 
									Lookup2.CursorType = 1
													
									SQL = "SELECT category.* FROM category where Master = 'yes' order by categoryname"
									Lookup2.open SQL, conn, 1,3
									
									set Lookup3 = Server.CreateObject("ADODB.recordset")
									Lookup3.CursorLocation = 3 
									Lookup3.CursorType = 1
																		
									a=0
									b=1
									%>
									<table width=100%>
									<tr><td valign=top class="tinyblack">
										<br>
									<%
									
									do while not lookup2.EOF
										
										catval = ""
												          
										%>
										<a href="stores.asp?catID=<%=lookup2("CategoryID")%>&catname=<%=lookup2("CategoryName")%>&country=uk"><%=lookup2("CategoryName")%></a><%
										
										
										SQL = "SELECT Count(Merchant) AS CountOfMerchant FROM stores where category like '%," & lookup2("CategoryID") & ",%'"
										Lookup3.open SQL,conn , 0, 1
										if Lookup3.bof and Lookup3.eof then
										temp = 0
										else
										temp = Lookup3("CountOfMerchant")
										end if
										%><font class="tinyblack1">&nbsp;(<%=temp%>)</font><br><br><%
										
										lookup3.close
												
										lookup2.MoveNext
									loop				
													
						
						
									%>
									</td></tr></table>		
						
						
						
						
				</td>
				<td bgcolor="#cde4f6" valign="top" align="middle"><!--#INCLUDE FILE="adverts.inc"-->
				</td>
			</tr>
			
			<tr>
				<td colspan="4"><!--#INCLUDE FILE="footer.inc"-->
				</td>
			</tr>
		</table>
		

	</body>
</HTML>
