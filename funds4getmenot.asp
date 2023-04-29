<%
view = session("whichone") 
site = session("sitename")

CookieUserID = "mem" & Request.QueryString("memberID") & "F4MEUK"

shop = Request.QueryString("shop")
send = Request.QueryString("send")
sql = "select * from stores where id = " & shop & ";"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set InsertCursor = Server.CreateObject("ADODB.recordset")
InsertCursor.CursorLocation = 3 
InsertCursor.CursorType = 1
													
InsertCursor.open SQL, conn, 1,3

insertcursor.MoveFirst

b = Replace(InsertCursor("link"),"CookieUserID",CookieUserID,1, -1, 1)

session("page") = b
merchant = InsertCursor("merchant")
cashback = InsertCursor("shoppin")

InsertCursor.Close


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
	boRecordset("store") = merchant
	boRecordset("whendate") = now() 
	boRecordset("max") = CashBack
	boRecordset("site") = "REDIRECTOR"
	boRecordset("ipaddress") = ipad

	boRecordset.Update
	boRecordset.Close



Response.Redirect(send)
%>