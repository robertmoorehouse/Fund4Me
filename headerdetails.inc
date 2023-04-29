<%
temp = request.QueryString("form")
shop = ""
if temp="yes" then
	shop = Request.Form("shopID")
else
	shop = Request("shop")
end if
if shop = "" then
	shop = session("shopid")
end if
session("shopid") = ""
if shop = "" then
	shop = "Funds4Me Cash Back Shopping"
	keyword = "Cash Back Shopping"
	keywordlist = "cash back, cashback, shopping for books, music, toys, videos, CDs, games &amp; electronics, computers, fashion, gifts, food wine, office &amp; business stores"
	titletext = "Cash Back Shopping with Funds4Me"
else


	ipad = trim(request.servervariables("REMOTE_ADDR"))
	
	sql = "select * from stores where id = " & shop & ";"
	
	
	set InsertCursor = Server.CreateObject("ADODB.recordset")
	InsertCursor.CursorLocation = 3 
	InsertCursor.CursorType = 1
													
	InsertCursor.open SQL, objConn3, 1,3
	
	insertcursor.MoveFirst
	
	store = shop & "(Funds4Me.co.uk)"
	failed = shop
	cashback = InsertCursor("shoppin")
	shop = insertcursor("merchant")
	linkit = Replace(InsertCursor("Link"),"CookieUserID",CookieUserID,1, -1, 1)
	c = linkit'Replace(linkit,"466441","1292226",1,-1,1)
	linkit  = c
	id = InsertCursor("ID")
	keyword = InsertCursor("keyword")
	keywordlist = InsertCursor("keywordlist")
	If Len(Trim(keyword)) = 0 Then keyword = "Cash Back, " & InsertCursor("Merchant")
	If Len(Trim(keywordlist)) = 0 Then keywordlist = InsertCursor("Merchant") & " " & InsertCursor("shoppin") & " Cash Back."
	notes = InsertCursor("Notes")
	banner = InsertCursor("banner")
	
	ipad = trim(request.servervariables("REMOTE_ADDR"))
	
	set boRecordset = Server.CreateObject("ADODB.recordset")
	boRecordset.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
	boRecordset.CursorType = 1
	
	
	
	SQL = "SELECT * FROM clicks WHERE (member = '')"
	boRecordset.open SQL, objConn1, 1, 3
	boRecordset.AddNew
	boRecordset("member") = CookieUserID
	boRecordset("store") = shop
	boRecordset("whendate") = now() 
	boRecordset("max") = CashBack
	boRecordset("site") = "Funds4Me.co.uk"
	boRecordset("ipaddress") = ipad
	
	boRecordset.Update
	boRecordset.Close
	
	SQL = "SELECT * FROM specials WHERE storeID = " & failed
	boRecordset.open SQL, objConn1, 1, 3
	if boRecordset.BOF and boRecordset.EOF then
		specof = "We have none listed, please let us know of any you are aware of any."
	else
		specof = boRecordset("description")
	end if
	boRecordset.Close
	titletext = shop & " " & CashBack & " Cash Back"
end if
%>