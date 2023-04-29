
<%
'for each x in Request.ServerVariables
'  response.write("<tr><td>" & x & "</td><td>" &_
'    Request.ServerVariables(x) & "</td></tr>")
'next
url = request.querystring


'response.Write(url)
'response.end

InitialString = url 

Set RegularExpressionObject = New RegExp

With RegularExpressionObject
.Pattern = "404;http://www.funds4me.co.uk:80/"
.IgnoreCase = True
.Global = True
End With

ReplacedString = RegularExpressionObject.Replace(InitialString, "")

MyNum = 0
MyNum = UBound(Split(ucase(ReplacedString), ucase("gocashbackshopping")))

ReplacedString = replace(ReplacedString,"gocashbackshopping/","",1,-1,1)

Set RegularExpressionObject = nothing

sql = "select * from stores where merchant = '" & ReplacedString & "';"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set InsertCursor = Server.CreateObject("ADODB.recordset")
InsertCursor.CursorLocation = 3 
InsertCursor.CursorType = 1
													
InsertCursor.open SQL, conn, 1,3

	if InsertCursor.bof and InsertCursor.eof then
	  redirecturl = "/default.asp"
	  MyNum = 0
	else
	
	insertcursor.MoveFirst
	
	redirecturl = "storedetails.asp" 
	session("shopid") = InsertCursor("id")
	
	
	store = InsertCursor("Merchant") & "(Funds4Me.co.uk)"
	cashback = InsertCursor("Charityrebatesval")
	shop = insertcursor("merchant")
	linkit = Replace(InsertCursor("Link"),"CookieUserID",CookieUserID,1, -1, 1)

	id = InsertCursor("ID")


	
	end if
		
Insertcursor.close

if MyNum > 0 then

	if Session("memID") > "" then
		CookieUserID  = "mem" & Session("memID")
	else
		CookieUserID  = request.cookies("details")("memberID")
	end if
	
	AmazonUK = request.cookies("details")("AmazonUK")
	
	if CookieUserID = "" then
		CookieUserID = session("customerID")
	end if

	if CookieUserID = "" then
		cuid = "NotLoggedIn" & date()
		CookieUserID = replace(cuid,"/","-",1,-1,1)
		AmazonUK = "dropthegrieffrom"
		showval = "Your purchase will benefit the School of the Day"	
	end if

	CookieUserID =  CookieUserID & "F4MECOUK"

	ipad = trim(request.servervariables("REMOTE_ADDR"))

	Dim objConn2
	Set objConn2 = Server.CreateObject("ADODB.Connection")
	objConn2.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
		
	set boRecordset = Server.CreateObject("ADODB.recordset")
	boRecordset.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
	boRecordset.CursorType = 1



	SQL = "SELECT * FROM clicks WHERE (member = '')"
	boRecordset.open SQL, objConn2, 1, 3
	boRecordset.AddNew
	boRecordset("member") = CookieUserID
	boRecordset("store") = shop
	boRecordset("whendate") = now() 
	boRecordset("max") = CashBack
	boRecordset("site") = "Funds4Me.co.uk"
	boRecordset("ipaddress") = ipad

	boRecordset.Update
	boRecordset.Close

	response.Redirect(linkit)
else
	Server.Transfer(redirecturl)
end if


%>