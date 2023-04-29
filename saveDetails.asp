<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	memberID = Trim(Request.QueryString("memID"))
	if memberID > "" then
		sUpdate = "yes"
	else
		sUpdate = "no"
	end if
	
	firstName = Trim(Request.Form("firstName"))
	lastName = Trim(Request.Form("lastName"))
	email = Trim(Request.Form("email"))
	password = Trim(Request.Form("password"))
	ipad = trim(request.servervariables("REMOTE_ADDR"))

		Dim conn, oRs, oRs2
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
		set oRs = Server.CreateObject("ADODB.recordset")
			oRs.CursorLocation = 3 
			oRs.CursorType = 1
		
			
		
		if sUpdate = "no" then		
			sSQL = "SELECT * FROM members WHERE email = '" & Trim(Request.Form("email")) & "' AND site = 'Funds4Me.co.uk'"
			oRs.open sSQL, conn, 1,3
				if oRs.EOF and oRs.BOF then
					oRs.AddNew
				else
					Response.Redirect("account.asp?email=known")
				end if		
		else
			sSQL = "SELECT * FROM members WHERE memID = " & cint(memberID)
			oRs.open sSQL, conn, 1,3
		
		end if
			
			oRs("firstName") = firstName
			oRs("lastName") = lastName
			oRs("address1") = Trim(Request.Form("address1"))
			oRs("address2") = Trim(Request.Form("address2"))
			oRs("city") = Trim(Request.Form("city"))
			oRs("county") = Trim(Request.Form("county"))
			oRs("country") = Trim(Request.Form("country"))
			oRs("postcode") = Trim(Request.Form("postcode"))
			oRs("email") = email
			oRs("phone") = Trim(Request.Form("phone"))
			oRs("username") = email
			oRs("password") = password
			oRs("RegisterDate") = day(date) & "/" & month(date) & "/" & year(date)
			oRs("site") = "Funds4Me.co.uk"
			report = Trim(Request.Form("report"))
			if report = "on" then
				report = "yes"
			else
				report = "no"
			end if
			oRs("report") = report
			newsletter = Trim(Request.Form("newsletter"))
			if newsletter = "on" then
				newsletter = "yes"
			else
				newsletter = "no"
			end if
			oRs("newsletter") = newsletter
			oRs("hearAbout") = Trim(Request.Form("hearAbout"))
			oRs("relation") = Trim(Request.Form("relation"))
			oRs("ipaddress") = ipad
			oRs.update()
	
  			oRs.close()
  		
		
		sSQL = "SELECT * FROM members WHERE email = '" & Trim(Request.Form("email")) & "' AND site = 'Funds4Me.co.uk'"
		oRs.open sSQL, conn, 1,3
		member = oRs("memID")					
		Session("memName") = firstName & " " & lastName
		Session("memID") = member
		temp1 = oRs("memid")
		
  		oRs.Close
  		if sUpdate = "no" then
  			
  			response.cookies("details")("UserName") = email
			response.cookies("details")("Password") = password
			response.cookies("details").expires = "July 31,2008"
			
  		
  			
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
		
		
			oRs.close()
		
		
		
		BodyText = "Dear " & firstName & " " & lastName & "<br><br>Welcome to Funds 4 Me"
  			BodyText = BodyText & "<br><br>When you login as a member at Funds 4 Me and then follow the links on our site you are able to earn Cash Back from over 500 stores world wide."
  			BodyText = BodyText & "<br><br>Your login details are:<br>Email: " & email & "<br>Password: " & password & "<br><br>"
			BodyText = BodyText & "We are currently redesigning the site so keep watching for the big launch!<br><br>As a funds4me member if you have your own blog why not run it from Zonchi.com for free, <b>YOU</b> choose the adverts displayed on the blog from the funds4me database and earn cash back from them being shown in your blog!"
  			BodyText = BodyText & "<br><br>Thanks for Joining<br><br><b>Funds 4 Me</b>"
  		
		
  			Dim objCDOSYSCon 

'Create the e-mail server object 
Set objCDOSYSMail = Server.CreateObject("CDO.Message") 
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 

'Out going SMTP server  
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.moorehouse.eu"   
 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")  = 25   
 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2   
 
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60   
 
objCDOSYSCon.Fields.Update  
  
'Update the CDOSYS Configuration 
Set objCDOSYSMail.Configuration = objCDOSYSCon 
  

objCDOSYSMail.From = "robert@Funds4Me.co.uk"   

objCDOSYSMail.To = toEmail

objCDOSYSMAIL.Bcc = "hullnew@hotmail.com"

objCDOSYSMail.Subject = "Registration for Funds 4 Me"

objCDOSYSMail.HTMLBody = BodyText

objCDOSYSMail.Send

'Close the server mail object  
Set objCDOSYSMail = Nothing 
Set objCDOSYSCon = Nothing 
			
		end if
  		
  		
  		

  		
  		
  		
  		
  		 		
  		Response.Redirect("account.asp")

%>
