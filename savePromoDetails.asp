<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	firstName = Trim(Request.Form("firstName"))
	email = Trim(Request.Form("email"))
	promoID = Request.QueryString("promoID")
	password = Trim(Request.Form("password"))
	ipad = trim(request.servervariables("REMOTE_ADDR"))
	
	Dim conn, oRs, oRs2
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	set oRs = Server.CreateObject("ADODB.recordset")
			
	sSQL = "SELECT * FROM promo WHERE ID = " & cint(promoID)
	
	oRs.open sSQL, conn, 1,3
		
	oRs("Company") = firstName
	'oRs("lastName") = lastName
	oRs("address") = Trim(Request.Form("address1"))
	oRs("address1") = Trim(Request.Form("address2"))
	oRs("Town") = Trim(Request.Form("city"))
	oRs("county") = Trim(Request.Form("county"))
	oRs("postcode") = Trim(Request.Form("postcode"))
	oRs("email") = email
	oRs("Tel") = Trim(Request.Form("phone"))
	oRs("Fax") = Trim(Request.Form("fax"))
	oRs("site") = "Funds4Me.co.uk"
'	temp1 = trim(StripHTML("<html>","</head>",trim(Request.Form(sItem))))
	temp1 = trim(Request.Form("fullDetails"))
	temp2 = replace(temp1,chr(13),"",1,-1,1)
	temp3 = replace(temp2,chr(10),"",1,-1,1)
	temp4 = replace(temp3,chr(34),"'",1,-1,1)
	fullDetails = temp4
	oRs("formText") = fullDetails
	
'	temp1 = trim(StripHTML("<html>","</head>",trim(Request.Form(sItem))))
	temp1 = trim(Request.Form("fullDescription"))
	temp2 = replace(temp1,chr(13),"",1,-1,1)
	temp3 = replace(temp2,chr(10),"",1,-1,1)
	temp4 = replace(temp3,chr(34),"'",1,-1,1)
	fullDescription = temp4
	oRs("Description") = fullDescription
	
	oRs("password") = password
	oRs("LastUpdate") = day(date) & "/" & month(date) & "/" & year(date)
	oRs("ipaddress") = ipad
	oRs.update()
	
  	oRs.close()
  		
	
  		
  Function StripHTML(ByVal BeginPoint, ByVal EndPoint, ByVal HTMLCode)
	    If IsNull(BeginPoint) Or BeginPoint = "" Or Not Len(BeginPoint) > 0 Then StripHTML = HTMLCode : Exit Function
	    If IsNull(EndPoint) Or EndPoint = "" Or Not Len(EndPoint) > 0 Then StripHTML = HTMLCode : Exit Function
	    If IsNull(HTMLCode) Or HTMLCode = "" Or Not Len(HTMLCode) > 0 Then StripHTML = HTMLCode : Exit Function
	    
	    Dim xCount, tmpDiscard, retHTML	   
	    Dim intStart: intStart = InStr(LCase(HTMLCode), LCase(BeginPoint))
	    Dim intEnd: intEnd = (InStr(LCase(HTMLCode), LCase(EndPoint)) + Len(LCase(EndPoint))) '- 1

	    StripHTML = right(cstr(HTMLCode),len(HTMLCode) - IntEnd)
End Function		

  		
  		
  		
  		
  		 		
  		Response.Redirect("promoaccount.asp")

%>
