<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	
		memID = Session("memID")
		
		
		Dim conn, oRs, oRs2
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
		
		set oRs = Server.CreateObject("ADODB.recordset")
		oRs.CursorLocation = 3 
		oRs.CursorType = 1
					
		sSQL = "SELECT * FROM members WHERE memID = " & memID & ""
		oRs.open sSQL, conn, 1,3
			if oRs.EOF and oRs.BOF then
				Response.Redirect("account.asp")
			else
				email = oRs("email")
				fullName = oRs("firstName") & " " & oRs("lastName")
			end if		
		oRs.close()
			
		  		
  		BodyText = "Dear " & fullName & "<br><br>Your 25% Discount Voucher for KidiSafe"
  		BodyText = BodyText & "<br><br>Coupon: KSF4SM055894166<br><br>"
  		BodyText = BodyText & "Remember 60% of the purchase price of KidiSafe goes to the school of your choice via your Funds4Me members account. When you register the software please use the same email address as the one you registered with Funds4Me.co.uk<br><br><b>Funds 4 Me</b>"
  		
  		Dim objCDO
		Set objCDO = Server.CreateObject("CDONTS.NewMail")
		objCDO.BodyFormat = 0 
		objCDO.MailFormat = 0 
		objCDO.To = email
		objCDO.From = "admin@Funds4Me.co.uk"
		objCDO.Subject = "25% Discount Voucher for KidiSafe"
		objCDO.Body = BodyText

		objCDO.Send

		Set objCDO = Nothing
  		
  		Response.Redirect("account.asp")

%>
