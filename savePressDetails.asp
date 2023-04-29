<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	pressID = Request.QueryString("ID")
	
	Dim conn, oRs, oRs2
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	set oRs = Server.CreateObject("ADODB.recordset")
	
	if pressID = "new" then
		pressID = 0
		addnew = "yes"
	end if
		
	sSQL = "SELECT * FROM press WHERE ID = " & cint(pressID)
	
	oRs.open sSQL, conn, 1,3
	if addnew = "yes" then
		oRs.AddNew
	end if
	oRs("title") = request.Form("Presstitle")
	oRs("Image1") = request.Form("PressImage1")
	oRs("Image2") = request.Form("PressImage2")
	oRs("Image3") = request.Form("PressImage3")
	oRs("newsDate") = request.Form("pressDate")
	oRs("live") = request.Form("presslive")
	oRs("item") = request.Form("pressitem")
	oRs("site") = "Funds4Me.co.uk"
'	temp1 = trim(StripHTML("<html>","</head>",trim(Request.Form(sItem))))
	temp1 = trim(Request.Form("fullDetails"))
	temp2 = replace(temp1,chr(13),"",1,-1,1)
	temp3 = replace(temp2,chr(10),"",1,-1,1)
	temp4 = replace(temp3,chr(34),"'",1,-1,1)
	fullDetails = temp4
	oRs("story") = fullDetails
	
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

  		
  		
  		
  		
  		 		
  		Response.Redirect("pressadmin.asp")

%>
