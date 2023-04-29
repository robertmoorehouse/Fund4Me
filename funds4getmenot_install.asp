<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	iMainMenu = 6
	iSchoolMenu = 0 
	temp = Request.Form("F4GMN_download")
	if temp = "" then
		
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
		set oRs = Server.CreateObject("ADODB.recordset")
			oRs.CursorLocation = 3 
			oRs.CursorType = 1
			
			
		sSQL = "SELECT * FROM members WHERE email = '" & Trim(Request.Form("email_address")) & "' AND site = 'Funds4Me.co.uk'"
		oRs.open sSQL, conn, 1,3
			if oRs.EOF and oRs.BOF then
				oRs.AddNew
					oRs("firstName") = "Please complete"
					oRs("lastName") = "your account details"
					oRs("email") = Trim(Request.Form("email_address"))
					oRs("username") = Trim(Request.Form("email_address"))
					oRs("password") = Trim(Request.Form("password"))
					oRs("RegisterDate") = day(date) & "/" & month(date) & "/" & year(date)
					oRs("site") = "Funds4Me.co.uk"
					report = "yes"
					oRs("report") = report
					newsletter = "yes"					
					oRs("newsletter") = newsletter
					ipad = trim(request.servervariables("REMOTE_ADDR"))
					oRs("ipaddress") = ipad
				oRs.update()
	
  				oRs.close()
  				sSQL = "SELECT * FROM members WHERE email = '" & Trim(Request.Form("email_address")) & "' AND site = 'Funds4Me.co.uk'"
				oRs.open sSQL, conn, 1,3
				CookieUserID = oRs("memid")
				
				Session("memName") = "Complete your account"
  				Session("memID") = oRs("memid")
  				temp1 = oRs("memid")
				
				response.cookies("details")("memberID") = CookieUserID
				response.cookies("details")("UserName") = Trim(Request.Form("email_address"))
				response.cookies("details")("Password") = Trim(Request.Form("password"))
				response.cookies("details").expires = "July 31,2008"
				response.cookies("details")("site") = "Funds4Me.co.uk"
				
				BodyText = "Dear Member<br><br>Welcome to Funds 4 Me"
				BodyText = BodyText & "<br><br>When you login as a member at Funds 4 Me and then follow the links on our site you are able to earn Cash Back from over 500 stores world wide."
				BodyText = BodyText & "<br><br>Your login details are:<br>Email: " & Trim(Request.Form("email_address")) & "<br>Password: " & Trim(Request.Form("password")) & "<br><br>Please remember to login to your account and complete your personal details, otherwise we will not know where to send your check!"
				BodyText = BodyText & "Do you have children and want to protect them when they are online, as a thank you for joining Funds4me you can purchase KidiSafe (see kidisafe.com) for just £14.95. Order by phone and quote funds4me.com offer <br><br>We have also just started your account with a £5 gift! now go shop and earn some cash back!<br><br><b>Funds 4 Me</b>"
					
			
				
				Set objCDO = Server.CreateObject("CDONTS.NewMail")
				objCDO.BodyFormat = 0 
				objCDO.MailFormat = 0 
				objCDO.To = email
				objCDO.Bcc = "rsm@funds4schools.co.uk"
				objCDO.From = "admin@Funds4Me.co.uk"
				objCDO.Subject = "Registration for Funds 4 Me"
				objCDO.Body = BodyText
	
				objCDO.Send
	
				Set objCDO = Nothing
				
				
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
		
		
				oRs.close()
				sSQL = "SELECT * FROM fundsraised WHERE memID = 999999999"

				oRs.open sSQL, conn, 1,3
							
				oRs.AddNew
			
				oRs("store") = "Funds4"
				oRs("memID") = temp1
				oRs("schoolID") = "Funds4Me.co.uk"
				oRs("fundsraised") = "5.00"
				oRs("product") = "£5.00 Opening offer"
				oRs("site") = "Funds4Me.co.uk"
				oRs("currencypaid") = "£"
				oRs("shopDate") = day(date) & "/" & month(date) & "/" & year(date)
				oRs("invoicemonth") = invoiceuse
				oRs("invoicecalc") = invoicecalc
				oRs("DateEntered") = day(date) & "/" & month(date) & "/" & year(date)
				oRs("invoiced") = "No"
				oRs.update()
			
				
			else
				oRs.close
				sSQL = "SELECT members.memID as memberID, members.* FROM members"
				sSQL = sSQL & " WHERE members.username = '" & Trim(Request.Form("email_address")) & "' and members.password = '" & Trim(Request.Form("password")) & "' and members.site='Funds4Me.co.uk'"
				oRs.Open sSQL, Conn, 3, 3
				if oRs.EOF and oRs.BOF then
					
					Response.Redirect("account.asp?email=known")
				
				else
					loggedIn = "yes"
					oRs.MoveFirst
					Session("memName") = oRs("firstName") & " " & oRs("lastName")
					Session("memID") = oRs("memberID")
					
  					response.cookies("details")("memberID") = oRs("memberID")
					response.cookies("details")("UserName") = Trim(Request.Form("email_address"))
					response.cookies("details")("Password") = Trim(Request.Form("password"))
					response.cookies("details").expires = "July 31,2008"
					response.cookies("details")("site") = "Funds4Me.co.uk"
					Response.Redirect("account.asp")
				end if	
				
			end if		
		
		oRs.Close
		Response.Redirect("account.asp")
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
	<body bgcolor="#ffffff" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0"><!--#INCLUDE FILE="header.inc"-->
		<table width="800" cellspacing="0" cellpadding="0">
			<tr>
				<td><IMG height=1 src="images/pix.gif" width=517></td>
				<td bgcolor="#99ccff"><IMG height=1 src="images/pix.gif" width=142></td>
				<td bgcolor="#cde4f6"><IMG height=1 src="images/pix.gif" width=141></td>
			</tr>
			<tr>
				<td valign="top" align="left" >
					<table width="498" cellpadding="0" cellspacing="0">
						<tr>
							<td>
								<IMG height=450 src="images/pix.gif" width=1>
							</td>
							<td align="middle" valign="top"><!-- Main Text area -->
								

<table cellpadding="0" cellspacing="0" border="0" width="517">
    <tr>
        <td bgcolor="#000000">
            <table cellpadding="0" cellspacing="0" border="0" align="left" width="517">
                <tr>
                    <td rowspan="4" class="black-text" bgcolor="#ffffff"><img src="/images/spacer.gif" width="25" height="1" border="0"></td>
                    <td colspan="2" class="black-text" bgcolor="#ffffff"><img src="/images/spacer.gif" width="1" height="25" border="0"></td>
                    <td rowspan="4" class="black-text" bgcolor="#ffffff"><img src="/images/spacer.gif" width="25" height="1" border="0"></td>
                </tr>
                <tr>
                    <td class="black-text" bgcolor="#ffffff"><h3>How does Funds4get-me-not Work?</h3>
                    Forget to visit Funds4Me first? Have no fear, Funds4get-me-not is here! Funds4get-me-not makes sure you don't miss out on the valuable Cash Back offered by Funds4Me.<br><br>
                    If you shop at any one of our 500 merchants, but forget to visit Funds4Me first, Funds4get-me-not will pop up to remind you to earn your Cash Back through our site.<br><br>
                    Remember to visit Funds4Me at www.funds4me.co.uk before shopping online to find the latest specials and exclusive offers guaranteed to save you money.
                    </td>
                    <td valign="top" bgcolor="#ffffff"><br></td>
                    
                </tr>
                <tr>
                    <td class="black-text" colspan="2" bgcolor="#ffffff"><br>*Funds4get-me-not is available for Windows Internet Explorer 4.0 and higher and AOL 5.0 and higher.</td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<%
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
		set oRs = Server.CreateObject("ADODB.recordset")
			oRs.CursorLocation = 3 
			oRs.CursorType = 1
		
			
			sSQL = "SELECT * FROM members WHERE email = '" & Trim(Request.Form("email_address")) & "' AND site = 'Funds4Me.co.uk'"
			oRs.open sSQL, conn, 1,3
				if oRs.EOF and oRs.BOF then
					oRs.AddNew
						oRs("firstName") = "Please complete"
						oRs("lastName") = "your account details"
						oRs("email") = Trim(Request.Form("email_address"))
						oRs("username") = Trim(Request.Form("email_address"))
						oRs("password") = Trim(Request.Form("password"))
						oRs("RegisterDate") = day(date) & "/" & month(date) & "/" & year(date)
						oRs("site") = "Funds4Me.co.uk"
						report = "yes"
						oRs("report") = report
						newsletter = "yes"					
						oRs("newsletter") = newsletter
						ipad = trim(request.servervariables("REMOTE_ADDR"))
						oRs("ipaddress") = ipad
					oRs.update()
	
  					oRs.close()
  					sSQL = "SELECT * FROM members WHERE email = '" & Trim(Request.Form("email_address")) & "' AND site = 'Funds4Me.co.uk'"
					oRs.open sSQL, conn, 1,3
					CookieUserID = oRs("memid")

					BodyText = "Dear Member<br><br>Welcome to Funds 4 Me"
					BodyText = BodyText & "<br><br>When you login as a member at Funds 4 Me and then follow the links on our site you are able to earn Cash Back from over 500 stores world wide."
					BodyText = BodyText & "<br><br>Your login details are:<br>Email: " & Trim(Request.Form("email_address")) & "<br>Password: " & Trim(Request.Form("password")) & "<br><br>Please remember to login to your account and complete your personal details, otherwise we will not know where to send your check!"
					BodyText = BodyText & "Do you have children and want to protect them when they are online, as a thank you for joining Funds4me you can purchase KidiSafe (see kidisafe.com) for just £14.95. Order by phone and quote funds4me.com offer <br><br>We have also just started your account with a £5 gift! now go shop and earn some cash back!<br><br><b>Funds 4 Me</b>"
						
			
			
					
					Set objCDO = Server.CreateObject("CDONTS.NewMail")
					objCDO.BodyFormat = 0 
					objCDO.MailFormat = 0 
					objCDO.To = email
					objCDO.Bcc = "rsm@funds4schools.co.uk"
					objCDO.From = "admin@Funds4Me.co.uk"
					objCDO.Subject = "Registration for Funds 4 Me"
					objCDO.Body = BodyText
	
					objCDO.Send
	
					Set objCDO = Nothing
					
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
			
			
					oRs.close()
					sSQL = "SELECT * FROM fundsraised WHERE memID = 999999999"
	
					oRs.open sSQL, conn, 1,3
								
					oRs.AddNew
				
					oRs("store") = "Funds4"
					oRs("memID") = CookieUserID
					oRs("schoolID") = "Funds4Me.co.uk"
					oRs("fundsraised") = "5.00"
					oRs("product") = "£5.00 Opening offer"
					oRs("site") = "Funds4Me.co.uk"
					oRs("currencypaid") = "£"
					oRs("shopDate") = day(date) & "/" & month(date) & "/" & year(date)
					oRs("invoicemonth") = invoiceuse
					oRs("invoicecalc") = invoicecalc
					oRs("DateEntered") = day(date) & "/" & month(date) & "/" & year(date)
					oRs("invoiced") = "No"
					oRs.update()
				
				else
					oRs.close
					sSQL = "SELECT members.memID as memberID, members.* FROM members"
					sSQL = sSQL & " WHERE members.username = '" & Trim(Request.Form("email_address")) & "' and members.password = '" & Trim(Request.Form("password")) & "' and members.site='Funds4Me.co.uk'"
					oRs.Open sSQL, Conn, 3, 3
					if oRs.EOF and oRs.BOF then
						
						Response.Redirect("account.asp?email=known")
				
					else
						loggedIn = "yes"
						oRs.MoveFirst
						Session("memName") = oRs("firstName") & " " & oRs("lastName")
						Session("memID") = oRs("memberID")
						
  						response.cookies("details")("memberID") = oRs("memberID")
						response.cookies("details")("UserName") = Trim(Request.Form("email_address"))
						response.cookies("details")("Password") = Trim(Request.Form("password"))
						response.cookies("details").expires = "July 31,2008"
						response.cookies("details")("site") = "Funds4Me.co.uk"
						Response.Redirect("account.asp")
					end if	
				
				end if		
				
				Session("memName") = "Complete your account"
  				Session("memID") = oRs("memid")
  				
				response.cookies("details")("memberID") = CookieUserID
				response.cookies("details")("UserName") = Trim(Request.Form("email_address"))
				response.cookies("details")("Password") = Trim(Request.Form("password"))
				response.cookies("details").expires = "July 31,2008"
				response.cookies("details")("site") = "Funds4Me.co.uk"
				
				
			oRs.Close


linkit = "http://www.tifps.com/cms/install.asp?memberid=mem" & cookieUserID & "F4MECOUK&client=funds4me.co.uk"

%>
<script LANGUAGE="JavaScript">
{
	<%
	Response.write("sTemp = """ & linkit & """")
	%>
	window.open(sTemp,'infowin','scrollbars=yes,width=570,height=570');
}
</script>

<!-- End of Main Text -->
					</td>
				</tr>
			</table></TD>
		<td bgcolor="#99ccff" valign="top" align="middle">
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
		<td bgcolor="#cde4f6" valign="top" align="middle"><!--#INCLUDE FILE="adverts.inc"-->
		</td></TR>
	<tr>		
		<td colspan="4"><!--#INCLUDE FILE="footer.inc"-->
		</td>
	</tr></TBODY></TABLE>	
</body>
</html>

