<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	
		email = Trim(Request.Querystring("email"))

		Dim conn, oRs, oRs2
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
		
		set oRs = Server.CreateObject("ADODB.recordset")
		oRs.CursorLocation = 3 
		oRs.CursorType = 1
					
		sSQL = "SELECT * FROM members WHERE email = '" & email & "'"
		oRs.open sSQL, conn, 1,3
			if oRs.EOF and oRs.BOF then
				Response.Redirect("account.asp?email=" & email)
			else
				password = oRs("password")
				fullName = oRs("firstName") & " " & oRs("lastName")
			end if		
		oRs.close()
			
		  		
  		BodyText = "Dear " & fullName & "<br><br>Reminder of your Funds 4 Me login details"
  		BodyText = BodyText & "<br><br>Email: " & email & "<br>Password: " & password & "<br><br>"
  		BodyText = BodyText & "<b>Funds 4 Me</b>"
  		BodyText = BodyText & "<br><br>This information was delivered to you because we received<BR>a password reminder request through the Funds 4 Me web site <BR>that specified your email address."
  		BodyText = BodyText & "<BR><BR>If you did not make this request, just delete this message <BR>as another user likely mistyped their email address."
  		BodyText = BodyText & "<br>Your password is always secure and will only be sent to the email <BR>address you provided during registration."
  		
  		Dim objCDO
		Set objCDO = Server.CreateObject("CDONTS.NewMail")
		objCDO.BodyFormat = 0 
		objCDO.MailFormat = 0 
		objCDO.To = email
		objCDO.From = "admin@Funds4Me.co.uk"
		objCDO.Subject = "Password Reminder for Funds 4 Me"
		objCDO.Body = BodyText

		objCDO.Send

		Set objCDO = Nothing
  		
  		Response.Redirect("account.asp?text=password")

%>
