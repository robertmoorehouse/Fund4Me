<%@ Language = VBScript%>
<%
	iMainMenu = 1
	iSchoolMenu = 0 

%>
<HTML>
<HEAD>
	<title>Funds 4 Me: the shopping portal providing rebates, savings, cash back, coupons, The incentive shopping site</title>
	<!-- books, music, toys, videos, CDs, games &amp; electronics, computers, fashion, gifts, food wine, office &amp; business stores and more.Shopping on-line through Funds4Me is easy. With over 200 well-known high street stores and on-line shops you'll find anything you want and with so many different special offers, you'll save money as well as time.What's more, you'll be raising funds for your favourite school without leaving your home. -->
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<META NAME="description" CONTENT="books, music, toys, videos, CDs, games &amp; electronics, computers, fashion, gifts, food wine, office &amp; business stores and more.Shopping on-line through Funds4Me is easy. With over 200 well-known high street stores and on-line shops you'll find anything you want and with so many different special offers, you'll save money as well as time.What's more, you'll be raising funds for your favourite school without leaving your home.">
	<META NAME="keywords" CONTENT="online store, retailer, store, shopping store, rebates, coupons, savings, bargain, sales, sale, shopping portal, books, games, CD, CD-ROM, DVD, wine, toys, electronics, computers, consumer electronics, gifts, fashion, clothes, videos, movies, home, garden, borders, shopping reward, incentive shopping, stores, shopping online, shop, bargain, internet coupon, shopping reward, incentive shopping, shop, stores, shopping online, online shopping community, shopping, fund raising, charity, school, primary school, secondary school, upper school, raise funds, grants, grant aid, PTA, parent teacher, parent association, governors, summer faite, schools, primary schools, secondary schools, upper schools, high school, high schools">
	<meta name="author" content="Inka Maier, TiF Pure Solutions Limited, + 44 (0) 1482 611021">
	<meta name="revisit-after" content="30 Days">
	<meta name="rating" content="General">
	<link rel="stylesheet" type="text/css" href="styles.css">
</HEAD>
	<body bgcolor="#ffffff" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0"><!--#INCLUDE FILE="header.inc"-->
		<table width="800" cellspacing="0" cellpadding="0">
			<tr>
				<td><!--<a href="didyouknow.asp"><img src="images/Funds4Me_V1.gif" border=0></a>-->
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
						<tr><td colspan=2><IMG height=10 src="images/pix.gif" width=517></td><tr>
						<tr>
							<td>
								<IMG height=450 src="images/pix.gif" width=20>
							</td>
							<td align="middle" valign="top"><!-- Main Text area -->
								<span class="blueText">
								<% if Session("memID") > "" then %>
									              <span class="blueText">          Welcome <%=Session("memName")%>,<br>
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
										<td class="standart" width="235" valign="top" align="left"><STRONG>Welcome Shoppin.net 
                  members.</STRONG> If you were a registered member on 
                  shoppin.net your account has been transferred to here. Simply 
                  login using your email address and your shoppin.net 
                  password.<BR><BR>For people not in the <STRONG>United 
                  Kingdom</STRONG> you need our .com Site <A 
                  href="http://www.funds4me.com">http://www.funds4me.com</A>. 
                  None UK registered shoppin.net members have their account there. If 
                  you have any questions please ask.<BR><BR>
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
				<td bgcolor="#99ccff" valign="top" align="middle">
					<%  ' Supporters/Retailers/Specials
			
				if iRightMenu = 1 then
			%> <!--#INCLUDE FILE="coupons.inc"-->
					<%	end if
				if iRightMenu = 2 then
			%> <!--#INCLUDE FILE="retailer.inc"-->
					<%	end if
				if iRightMenu = 3 then
			%> <!--#INCLUDE FILE="specials.inc"-->
					<%	end if			
			%>
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
