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
<style type="text/css">
	body { font-family: Arial, Helvetica, Sans-serif; font-size: 9px; color: #000000; }
</style>
<script src="js/funds4_nav.js"></script>
<script src="js/funds4_cookie_detect.js"></script>
<script >
function indexInit()
{
	init990();
}
</script>

</HEAD>
	<body bgcolor="#ffffff" link="#006600" alink="#339900" vlink="#CC3300" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="init990();"><!--#INCLUDE FILE="headerdefault.inc"-->
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
						<tr><td colspan=2><IMG height=10 src="images/pix.gif" width=517></td><tr>
						<tr>
							<td>
								<IMG height=450 src="images/pix.gif" width=20>
							</td>
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
												
								

					SQL = "SELECT top 20 id,merchant,shoppin, ListCountry FROM stores where ((stores.ListCountry like '%,uk,%') and (shoppin <> 'Coupons Only')) order by ID "
					'Response.Write(SQL)
					Lookup5.open SQL, conn3
								
					%>
					<br><table>
					<% do while not Lookup5.EOF
						temp = trim(Lookup5("Merchant")) + " " + trim(Lookup5("shoppin"))			
						
						if len(temp) < 20 then
						%>
						<tr>
							<td class="tinyblack">&nbsp;<a href="shoprequired.asp?shop=<%=Lookup5("id")%>"><%=trim(Lookup5("Merchant"))%></a></td>
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
												
								

					SQL = "SELECT top 20 id,merchant,shoppin, ListCountry FROM stores where ((stores.ListCountry like '%,uk,%') and (shoppin <> 'Coupons Only')) order by ID Desc; "
					'Response.Write(SQL)
					Lookup5.open SQL, conn3
								
					%>
					<br><table>
					<% do while not Lookup5.EOF
						temp = trim(Lookup5("Merchant")) + " " + trim(Lookup5("shoppin"))			
						
						if len(temp) < 20 then
						%>
						<tr>
							<td class="tinyblack">&nbsp;<a href="shoprequired.asp?shop=<%=Lookup5("id")%>"><%=trim(Lookup5("Merchant"))%></a></td>
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
		
		


<%

if (Session("memID") > "") then
else
	%>

<script language="JavaScript"><!--
function verify() {
submit = true;
		
if (document.regiForm.email_address.value == "") {
alert("You must enter an Email Address");
document.regiForm.email_address.focus();
submit = false;
}	else {	
if (document.regiForm.password.value == "") {
alert("You must enter a Password");
document.regiForm.password.focus();
submit = false;
}	else {
pw = document.regiForm.password.value;
if (pw.length < 6) {
alert("Your password must have at least 6 characters.");
document.regiForm.password.value="";
document.regiForm.reenter_password.value="";
document.regiForm.password.focus();
submit = false;
}	else {
emailaddress = document.regiForm.email_address.value;
at = emailaddress.indexOf("@");
dot = emailaddress.indexOf(".");
if (at == -1){
alert("You email address is not valid");
submit = false;
} else {
if (dot == -1){
alert("You email address is not valid");
submit = false;
}	else {
if (document.regiForm.reenter_password.value == "") {
}	else {	
if (document.regiForm.reenter_password.value != document.regiForm.password.value) {
alert("The password confirmation is not the same as the password!");
document.regiForm.password.value="";
document.regiForm.reenter_password.value="";
document.regiForm.password.focus();
submit = false;
}	else {	
ch = document.regiForm.terms.value;
if (!regiForm.terms.checked) {
alert("You must agree to our terms and conditions.");
document.regiForm.terms.focus();
submit = false;
}
}	}	}	}	}	}	}
		
		
		if (submit == true) {
		document.regiForm.submit();
		}	
	
		
	}
-->
</script>

	<script src="js/how_funds4.js"></script>

	<div id="howFunds4Works" style="position:absolute;z-index:20;top:350px;left:0px;">
	<TABLE width="800" cellspacing="0" cellpadding="0" border='0' bgcolor="#ffcc99">
	<TR>
	<TD width='750' bgcolor='#ff9900' align='center' class="WhiteText" height='25'><b> &nbsp;&nbsp;&nbsp;&nbsp;How Funds 4 Me Works : Already a member then LOG IN!</b></TD>
	<TD  align='center' bgcolor='#ff9900'  valign='middle'><A href="javascript:closeHowFunds4Layers()" ><img src='/images/close.gif' border='0'></A></TD>
	</TR>
	<TR><TD colspan='2' bgcolor="#ffcc99">
	<TABLE cellspacing="0" width='100%' cellpadding="0" border='0' bgcolor="#ffcc99">
	        <tr> 
	          <form name="regiForm" method="POST" action="funds4getmenot_install.asp" target="_top"> 
        <tr> 
          <td  class="black-text" valign="top"  align='left'  bgcolor="#FEF4EA">
            <table height="129" cellspacing="0" cellpadding="0" border="0" bgcolor="#FEF4EA" >
                <tr><td rowspan='10'><img src='/images/spacer.gif' border='0' width='2' height='1'></td>
        		<td valign="top">
        		
        		<table cellspacing="0" cellpadding="2" border='0' bgcolor="#FEF4EA" >
        		<tr><td colspan='3' align='left'>
                    <p class="standart"><b>1: Sign Up or Log in Now!</b>                    
                </td></tr>
                <tr><td colspan='3' align='left'>  
                    <img src='images/email.gif' border='0' width='66' height='9'>
                </td></tr>
                <tr><td colspan='3' align='left'>
                    <input type="text" name="email_address" maxlength="255" size="18" value="">
                </td></tr>
                <tr height="22" valign="bottom">
                	<td align='left' >
                	    <img src='images/pass.gif' border='0' width='80' height='21'  align="ABSBOTTOM">
                	</td>
                	<td>
                		<img src='images/repass.gif' border='0' width='46' height='21'  align="ABSBOTTOM">
           		</td>
           		<td>
           		</td>
                </tr>
                <tr>
                	<td>
                	    <input type="password" name="password" maxlength="20" size="8" value="">
                	</td>
                	<td>
 					<input type="password" name="reenter_password" maxlength="20" size="8" value="">
                	</td></tr>
                	<tr><td colspan=2>
                		<A href="JavaScript:verify();"><IMG src="images/buttons/login1.gif" border=0></a> <A href="JavaScript:verify();"><IMG src="images/buttons/register1.gif" border=0></A>
                	</td>
                	
                </tr>
                </table>
	          </td>
	          <td  class="black-text" valign="top"  align='left'  bgcolor="#FEF4EA">
	            <table height="129" cellspacing="0" cellpadding="0" border="0" bgcolor="#FEF4EA" >
	                <tr>
	        		<td valign="top" width=180>
	        		<p class="standart"><b>2: Select a merchant to shop at.</b><br><br>
					<img src="images/how3.gif">
					</p>
	               	</td>          			          			
	        		</tr>
	            </table>
	          </td>
	          <td  class="black-text" valign="top"  align='left'  bgcolor="#FEF4EA">
	            <table height="129" cellspacing="0" cellpadding="0" border="0" bgcolor="#FEF4EA" >
	                <tr>
	        		<td valign="top" width=200>
	        		<p class="standart"><b>3: Buy from that merchant</b><br><br>
					<img src="images/how5.gif">
					</p>
	               	</td>          			          			
	        		</tr>
	            </table>
	          </td>
	          <td class="black-text" valign="top"  align='left'  bgcolor="#FEF4EA">
	            <table height="129" cellspacing="0" cellpadding="0" border="0" bgcolor="#FEF4EA" >
	                <tr>
	        		<td valign="top" width=200>
	        		<p class="standart"><b>4: Funds4 pays the Cash Back to your account and sends a cheque in the post.</b><br>
					&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/how6.gif">
					</p>
	               	</td>          			          			
	        		</tr>
	            </table>
	          </td>
	        </tr>
	        <tr>
	        <td class="tinyblack1" colspan=4>
	        <input type="checkbox" name="terms" checked>
 
            I have read and agreed to the <a href="useragreement.asp">Terms &amp; 
            Conditions</a> and <a href="privacy.asp">Privacy Policy</a> &nbsp;<input type="checkbox" name="F4GMN_download" checked>
I would like to download the Funds4-get-me-not reminder tool.</td></tr>
	 </TABLE>
	 </TD>
	 </TR>
	</TABLE>
	</div>

<% 
end if %>



	</body>
</HTML>
