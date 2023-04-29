<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	firstName = Trim(Request.Form("firstName"))
	lastName = Trim(Request.Form("lastName"))
	email = Trim(Request.Form("email"))
	SchoolID = Request.QueryString("sID")
	promoID = Request.QueryString("pID")
	ipad = trim(request.servervariables("REMOTE_ADDR"))
	
	Dim conn, oRs, oRs2
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
						
	set oRs = Server.CreateObject("ADODB.recordset")
	oRs.CursorLocation = 3 
	oRs.CursorType = 1
	
	sSQL = "SELECT * FROM promo where ID = " & cint(promoID)
	oRs.open sSQL, conn, 1,3
	
	toEmail = oRs("email")
	toCompany = oRs("Company")
	
	oRs.Close
		
	sSQL = "SELECT * FROM promocontacts Order by ID DESC"
	oRs.open sSQL, conn, 1,3
			

		
	oRs.AddNew
			
	oRs("firstName") = firstName
	oRs("lastName") = lastName
	oRs("address1") = Trim(Request.Form("address1"))
	oRs("address2") = Trim(Request.Form("address2"))
	oRs("city") = Trim(Request.Form("city"))
	oRs("county") = Trim(Request.Form("county"))
	oRs("postcode") = Trim(Request.Form("postcode"))
	oRs("email") = email
	oRs("phone") = Trim(Request.Form("phone"))
	oRs("fax") = Trim(Request.Form("fax"))
	oRs("comments") = Trim(Request.Form("comments"))
	oRs("RegisterDate") = day(date) & "/" & month(date) & "/" & year(date)
	oRs("promocode") = Request.QueryString("pID")
	oRs("site") = "Funds4Me.co.uk"	
	if SchoolID = "SchoolOfTheDay" then
		sDay=day(date)
		sMonth=month(date)
		if len(sDay)=1 then
			sDay = "0" & sDay
		end if
		if len(sMonth)=1 then
			sMonth="0" & sMonth
		end if
		SchoolID = "SOD" & sDay & sMonth & year(date)
	end if
	oRs("schoolID") = SchoolID
	oRs("ipaddress") = ipad
	oRs.update()
	
	oRs.close()
  		
  		
  		
				
	Set oRs = Nothing
  		
  		
  		
  		
	BodyText = "Dear " & firstName & " " & lastName & "<br><br>Your Details have been forwarded on to " & Request.QueryString("pID") 
	BodyText = BodyText & "<br><br>They should be in touch within the next 5 working days, please let us know if they do not contact you within this time."
	BodyText = BodyText & "<br><br>Finally would you please let us know if you make any purchase as a result of this contact, by replying to this email and telling us what you purchased so we may ensure your chosen school (" & SchoolID & ") receives any money due as a result.<br><br>"
	BodyText = BodyText & "<b>Funds 4 Me</b>"
  		
	Dim objCDO
	Set objCDO = Server.CreateObject("CDONTS.NewMail")
	objCDO.BodyFormat = 0 
	objCDO.MailFormat = 0 
	objCDO.To = email
	objCDO.From = "admin@Funds4Me.co.uk"
	objCDO.Subject = "Promotion on Funds 4 Me"
	objCDO.Body = BodyText

	objCDO.Send

	Set objCDO = Nothing
			
	BodyText = "Dear " & toCompany 
	BodyText = BodyText & "<br><br>The following details have been passed on from a Funds4Me.co.uk user, please contact this person within 5 working days."
	BodyText = BodyText & "<br><br>Please let us know if this results in a purchase so we may raise an invoice.<br><br>"
	BodyText = BodyText & "<b>Funds 4 Me</b>"
	BodyText = BodyText & "<br><br>First Name : " & firstName
	BodyText = BodyText & "<br><br>Last Name  : " & lastName
	BodyText = BodyText & "<br><br>Address    : " & Trim(Request.Form("address1"))
	BodyText = BodyText & "<br><br>" & Trim(Request.Form("address2"))
	BodyText = BodyText & "<br><br>" & Trim(Request.Form("city"))
	BodyText = BodyText & "<br><br>" & Trim(Request.Form("county"))
	BodyText = BodyText & "<br><br>" & Trim(Request.Form("postcode"))
	BodyText = BodyText & "<br><br>Email : " & email
	BodyText = BodyText & "<br><br>Phone : " & Trim(Request.Form("phone"))
	BodyText = BodyText & "<br><br>Fax   : " & Trim(Request.Form("fax"))
	BodyText = BodyText & "<br><br>Comments : " & Trim(Request.Form("comments"))
  				
	Set objCDO = Server.CreateObject("CDONTS.NewMail")
	objCDO.BodyFormat = 0 
	objCDO.MailFormat = 0 
	objCDO.To = toEmail
	objCDO.CC = "promo@Funds4Me.co.uk"
	objCDO.From = "promo@Funds4Me.co.uk"
	objCDO.Subject = "Promotion on Funds 4 Me"
	objCDO.Body = BodyText

	objCDO.Send

	Set objCDO = Nothing
  		
Response.Redirect("default.asp")

%>
