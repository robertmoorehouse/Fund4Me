<%
'log user out
sLogOff = ""
sLogOff = Trim(Request.Querystring("logout"))
if sLogOff = "true" then
		response.cookies("details")("UserName") = ""
		response.cookies("details")("Password") = ""
		response.cookies("details")("UserName") = ""
		response.cookies("details")("Password") = ""
		session("memID") = ""
		Session("memName") = ""
end if

'See if is a member set memberID if they are
sUser = ""
sUser = Trim(Request.Form("username"))
sPassword = Trim(Request.Form("password"))
sEmail = Trim(Request.Querystring("email"))

if sUser = "" then
	sUser = Request.cookies("details")("UserName")
	sPassword = Request.cookies("details")("Password")			
end if

if (sEmail = "known") then
	sUser = ""
	sPassword = ""
	Session("memID") = ""
end if


if (sUser = "") and (sPassword = "") then
	if Session("memID") = "" then
		Session("memID") = "NotLoggedIn"
	end if
else
	if Session("memID") = "29" then
		if len(sUser) > 6 then
			sSQL = "SELECT members.memID as memberID, members.* FROM members"
			sSQL = sSQL & " WHERE members.username = '" & sUser & "' and members.site='Funds4Me.co.uk'"
		else
			sSQL = "SELECT members.memID as memberID, members.* FROM members"
			sSQL = sSQL & " WHERE members.memID = " & sUser & " and members.site='Funds4Me.co.uk'"
		end if
	else
		sSQL = "SELECT members.memID as memberID, members.* FROM members"
		sSQL = sSQL & " WHERE members.username = '" & sUser & "' and members.password = '" & sPassword & "' and members.site='Funds4Me.co.uk'"
		
	end if
		
	Set oRs = Server.CreateObject("ADODB.RecordSet")
		oRs.Open sSQL, objConn1, 3, 3
		'response.Write(sSQL)
	if oRs.EOF and oRs.BOF then
		oRs.close
		sSQL = "SELECT members.memID as memberID, members.* FROM members"
		sSQL = sSQL & " WHERE members.username = '" & sUser & "' and members.password = '" & sPassword & "' and members.site='Funds4Me.com'"
	
		oRs.Open sSQL, objConn, 3, 3
	
		if oRs.EOF and oRs.BOF then
		
			loggedIn = "no"
		
		else
			
			response.Redirect("http://www.funds4me.com")
		
		end if
	
	else
		loggedIn = "yes"
		oRs.MoveFirst
		Session("memName") = oRs("firstName") & " " & oRs("lastName")
		Session("memID") = oRs("memberID")
		
		response.cookies("details")("memberID") = memberID
		response.cookies("details")("UserName") = sUser
		response.cookies("details")("Password") = sPassword
		response.cookies("details").expires = "July 31,2008"
		response.cookies("details")("site") = "Funds4Me.co.uk"
		
	end if	
oRs.close
end if


memberID = session("memID")
memberName = Session("memName")

CookieUserID = session("memID")
AmazonUK = "dropthegrieffrom"


if CookieUserID = "NotLoggedIn" then
	register = "Register"
	cuid = "NotLoggedIn" & date()
	CookieUserID = replace(cuid,"/","-",1,-1,1)
	personalise = ""
else
	CookieUserID = "mem" & session("memID")
	personalise = "Your "
	register = "Shop"
	name = Request.cookies("details")("UserName")
	showval = "Your purchase will benefit " & name
end if

CookieUserID =  CookieUserID & "F4MECOUK"
%>