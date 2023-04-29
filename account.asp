
<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">


	<table border="0" cellpadding="0" cellspacing="0" width="498" valign="top">
              <TBODY>
      <tr>
        <td width="498" class="standart" valign="top" align="middle">
  <%if loggedIn = "yes" then%><table width="490" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ffffff">
				<td class="WhiteText" width="90" valign="top" align="left">
</td>
				<td colspan="2" class="BlueText" width="408" valign="top" align="middle">
					<IMG height=10 src="images/pix.gif" width=408><br>
					<br>
					<font size="2"><b><i>My Account</i></b></font>
				</td>
			</tr>
			</table>
			<table width="380" cellpadding="0" cellspacing="0">
			<%
			if session("memID") = "29" then
			%>
			<tr bgcolor="#ffffff">
				<td class="LargeBlack" width="80" valign="center" align="middle">
					<A href="addrebates.asp?adminidid=admin"><IMG src="images/schools.gif" border=0></a>
				</td>
							
				<td class="LargeBlack" width="300" valign="middle" align="left">
					<span class="BlueText"><b>Administration</b></span><br>
					<a href="addrebates.asp?adminidid=adminadd">Add Rebates</a><br>
					<a href="addrebates.asp?adminidid=adminamend">Amend Rebates</a><br><br>
					<a href="promoaccount.asp">Promo Details</a><br>
					<a href="pressadmin.asp">Press Details</a><br>
					<a href="specialsadmin.asp?id=1">Specials Details</a><br><br>
					<a href="Storeadmin.asp">Amend Store Details</a><br>
					<a href="Storeadmin.asp?add=add">Add Store Details</a><br>
					<a href="invoice1.asp" >Invoice</a><br><br>
				</td>
			</tr>
			<tr bgcolor="#ffffff">
				<td colspan="2" class="WhiteText" width="8" valign="top" align="left">
					<IMG height=4 src="images/pix.gif" width=380>
					<IMG height=1 src="images/orangeSquare.gif" width=380>
					<IMG height=4 src="images/pix.gif" width=380>
				</td>
			</tr>
			<% end if %>
			<tr bgcolor="#ffffff">
				<td colspan="2" class="WhiteText" width="380" valign="top" align="left">
					<IMG height=20 src="images/pix.gif" width=380>
				</td>
			</tr>
			<tr bgcolor="#ffffff">
				<td class="LargeBlack" width="80" valign="center" align="middle">
					<A href="AccReport.asp"><IMG src="images/schools.gif" border=0></a>
				</td>
				<td class="LargeBlack" width="300" valign="center" align="left">
					<span class="BlueText"><b>My Report</b></span><br>
					<%
					
					set Recordset1 = Server.CreateObject("ADODB.recordset")
					Recordset1.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
					Recordset1.CursorType = 1
					
					SQL = "SELECT fundsraised "
					SQL = SQL & "FROM fundsraised "
					SQL = SQL & "WHERE (memID ='" & session("memID") & "');"
			
					Recordset1.open SQL, objConn1
	
					if Recordset1.bof and Recordset1.eof then
					%>
					You have no purchases on your account yet. When you have earned some cashback you will be able to see the details here.
					<%
					else
																		
						Recordset1.movefirst
						comm = 0
						do while not Recordset1.eof
							comm = comm + recordset1("fundsraised")
							recordset1.movenext
						loop
						%>
						<b>Total Cash Back Earned</b>  <a href="AccReport.asp">
								<%=num2pounds(comm)%>
							</a></b>
						<%
					end if
						Recordset1.close%>
							
					
					Click to view a summary of your individual purchases.
				</td>
			</tr>
			
				
			<tr bgcolor="#ffffff">
				<td colspan="2" class="WhiteText" width="8" valign="top" align="left">
					<IMG height=4 src="images/pix.gif" width=380>
					<IMG height=1 src="images/orangeSquare.gif" width=380>
					<IMG height=4 src="images/pix.gif" width=380>
				</td>
			</tr>
			<tr bgcolor="#ffffff">
				<td class="LargeBlack" width="80" valign="center" align="middle">
					<A href="AccProfile.asp"><IMG src="images/schools.gif" border=0></a>
				</td>
				<td class="LargeBlack" width="300" valign="center" align="left">
					<span class="BlueText"><b>My Profile</b></span><br>
					Manage your account profile including email, password, address and phone number.
				</td>
			</tr>
			
			
			<tr bgcolor="#ffffff">
				<td colspan="2" class="WhiteText" width="8" valign="top" align="left">
					<IMG height=4 src="images/pix.gif" width=380>
					<IMG height=1 src="images/orangeSquare.gif" width=380>
					<IMG height=4 src="images/pix.gif" width=380>
				</td>
			</tr>
					<tr bgcolor="#ffffff">
				<td class="LargeBlack" width="80" valign="center" align="middle">
					<A href="AccReport.asp"><IMG src="images/schools.gif" border=0></a>
				</td>
				<td class="LargeBlack" width="300" valign="center" align="left">
					<span class="BlueText"><b>Log Off</b></span><br>
					To remove all Funds4Me settings from this computer <A href="account.asp?do=off">click here</a> or to keep your user details on this computer while 
disabling access to your account <A href="account.asp?do=part">click here</a>. 
				</td>
			</tr>
			<tr bgcolor="#ffffff">
				<td colspan="2" class="WhiteText" width="8" valign="top" align="left">
					<IMG height=4 src="images/pix.gif" width=380>
					<IMG height=1 src="images/orangeSquare.gif" width=380>
					<IMG height=4 src="images/pix.gif" width=380>
				</td>
			</tr>
			
		</table>
			
			
			
			
  <%		else
  
				  If Request.QueryString("email") = "known" then
					Response.Write("<br><b>Your email address has already been registered!<br>If you have forgotten your password, please enter your email address below and press <A href='JavaScript:pwreminder();'>remind me</a>.<br>Alternatively, <A href='javascript:history.back(1);'>click here</a> to return to the registration page.</b>")
					Response.Write("<br><br><img src='images/blueSquare.gif' width='490' height='1' border='0'><br><br>")
                  else
                  
	                   If Request.QueryString("text") = "password" then
						Response.Write("<br><b>You will receive an email with your accound details shortly.</b>")
						Response.Write("<br><br><img src='images/blueSquare.gif' width='490' height='1' border='0'><br><br>")
	                   else
						
						If (loggedIn = "no") then
							Response.Write("<br><b>Your details have not been found, please try again or <A href='register2.asp'>register here</a></b>")
							Response.Write("<br><br><img src='images/blueSquare.gif' width='490' height='1' border='0'><br><br>")
						else 
						
						 If (loggedIn = "wrong site") then
							Response.Write("<br><b>You are on the wrong site, None UK members should use <a href=''http://www.funds4me.com'> funds4me.com</a>, please use that site or <A href='register2.asp'>register here</a> with a UK address.</b>")
							Response.Write("<br><br><img src='images/blueSquare.gif' width='490' height='1' border='0'><br><br>")
							else %>
													
						<table width="360" cellpadding="0" cellspacing="0">
							<tr bgcolor="#ffffff">
								<td class="standart" valign="top" align="left"><br>
                        <CENTER></CENTER>When you shop Funds4Me as a registered 
                        user, you'll receive access to online reports of your 
                        purchases, receive optional email shopping reminders 
                        with valuable merchant offers, and be able to <b>earn CASH 
                        BACK</b> for your purchases. If you do not register, or 
                        login the commission we earn from your purchases will be 
                        paid to a school on the funds4schools.co.uk or rebates4schools.com websites. <br><br>
	<center></center>
						</td></tr></table>
					<%	
						  end if
						end if
	                  end if
	                 end if
	                 
	%>				
					<span class="BlueText"><b>Please enter your email address and password to get your account details.</b>
					</span><br><br>
					<form name="LoginForm" action="account.asp" method="post">
					<table width="360" cellpadding="0" cellspacing="0">
						<tr bgcolor="#0099cc">
							<td class="WhiteText" width="8" valign="top" align="left">
								<IMG height=1 src="images/pix.gif" width=8>
							</td>
							<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
								<b><i>Member Login</i></b>
							</td>
						</tr>
						<tr bgcolor="#99ccff">
							<td class="WhiteText" width="8" valign="top" align="left">
								<IMG height=1 src="images/pix.gif" width=8>
							</td>
							<td class="Standart" width="176" valign="top" align="left">
								Email:<br>
								<input name="username" >
							</td>
							<td class="Standart" width="176" valign="top" align="left">
								Password:<br>
								<input type="password" id="password" name="password" >
							</td>
						</tr>
						<tr bgcolor="#99ccff">
							<td class="WhiteText" width="8" valign="top" align="left">
								<IMG height=1 src="images/pix.gif" width=8>
							</td>
							<td class="WhiteText" width="176" valign="top" align="left">
								<A href="JavaScript:document.LoginForm.submit();"><IMG src="images/buttons/login.gif" border=0></a>
							</td>
							<td class="WhiteText" width="176" valign="top" align="right">
								<A href="register2.asp"><IMG src="images/buttons/register.gif" border=0></a>&nbsp;&nbsp;&nbsp;&nbsp;
							</td>
						</tr>
					</table><br>
					<span class="standart">To receive a password reminder, enter your email address and <A href="JavaScript:pwreminder();">click here</a>.</span>
					</form> 			
							
<%					
  				end if
  			
%>
		
        </td>
      </tr>

      </tr>
    </table>
	
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->