<%@ Language = VBScript%>
<%
	
	schoolID = Trim(request.form("ID"))
	if schoolID = "" then 
	schoolID = Trim(request.form("username"))
	end if
	if schoolID = "" then schoolID = 0
	password = Trim(Request.Form("password"))
	ipad = trim(request.servervariables("REMOTE_ADDR"))
	
	Dim conn, oRs, oRs2
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	set oRs = Server.CreateObject("ADODB.recordset")
			
	sSQL = "SELECT * FROM schools WHERE ID = " & schoolID
	
	oRs.open sSQL, conn, 1,3
	
	if oRs.BOF and oRs.EOF then
	oRs.AddNew
	oRs("pay_address") = Trim(Request.Form("address1"))
	'oRs("pay_address1") = Trim(Request.Form("address2")) 
	oRs("pay_Town") = Trim(Request.Form("city")) 
	oRs("pay_county") = Trim(Request.Form("county")) 
	oRs("pay_postcode") = Trim(Request.Form("postcode"))
	else
	oRs.MoveFirst
	end if
	
		
	oRs("school_name") = Trim(Request.Form("SchoolName")) & "<br>"
	oRs("school_address") = Trim(Request.Form("address1")) & "<br>"
	oRs("school_address1") = Trim(Request.Form("address2")) & "<br>"
	oRs("school_Town") = Trim(Request.Form("city")) & "<br>"
	oRs("school_county") = Trim(Request.Form("county")) & "<br>"
	oRs("school_postcode") = Trim(Request.Form("postcode"))
	
	oRs("pay_name") = Trim(Request.Form("payname"))
	
	
	
	oRs("email") = Trim(Request.Form("email"))
	oRs("Tel") = Trim(Request.Form("phone"))
	oRs("Fax") = Trim(Request.Form("fax"))
	oRs("website") = Trim(Request.Form("Schoolwebsite"))
	oRs("school_password") = Trim(Request.Form("Password"))
	oRs("contact") = Trim(Request.Form("contact"))
	
	if session("memID") = "29" then
	session("schoolID") = schoolID
	oRs("School_logo") = Trim(Request.Form("SchoolLogo"))
	oRs("School_pic") = Trim(Request.Form("Schoolpic"))
	oRs("ID") = schoolID
	oRs("start_date") = Trim(Request.Form("startdate"))
	oRs("cost_basis") = Trim(Request.Form("costbasis"))
	oRs("carriedforward") = Trim(Request.Form("carriedforward"))
	oRs("active") = Trim(Request.Form("active"))
	t1 = ""
	t1 = Trim(Request.Form("LEA"))
	if t1 = "" then
	t1 = 0
	end if
	oRs("LEA") = t1
	oRs("AmazonUK_ID") = Trim(Request.Form("AmazonUK_ID"))
	oRs("site") = "Funds4Me.co.uk"
	oRs("school_country") = "USA"
	oRs("pay_country") = "USA"
	end if
	 	
'	temp1 = trim(StripHTML("<html>","</head>",trim(Request.Form(sItem))))
	temp1 = trim(Request.Form("fullDetails"))
	temp2 = replace(temp1,chr(13),"",1,-1,1)
	temp3 = replace(temp2,chr(10),"",1,-1,1)
	temp4 = replace(temp3,chr(34),"'",1,-1,1)
	fullDetails = temp4
	oRs("school_short_description") = fullDetails
	
'	temp1 = trim(StripHTML("<html>","</head>",trim(Request.Form(sItem))))
	temp1 = trim(Request.Form("fullDescription"))
	temp2 = replace(temp1,chr(13),"",1,-1,1)
	temp3 = replace(temp2,chr(10),"",1,-1,1)
	temp4 = replace(temp3,chr(34),"'",1,-1,1)
	fullDescription = temp4
	oRs("school_description") = fullDescription
	
	oRs("LastUpdate") = now()
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

  		
  		
  		
  		
  		 		
  		Response.Redirect("schooladmin.asp")

%>
