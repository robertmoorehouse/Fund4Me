<%
view = session("whichone") 
site = session("sitename")
CookieUserID = ""
CookieUserID  = request.cookies("details")("memid")
if CookieUserID = "" then
CookieUserID = session("customerID")
end if
shop = Request.querystring("shop")


sql = "select * from stores where id = " & shop & ";"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set InsertCursor = Server.CreateObject("ADODB.recordset")
InsertCursor.CursorLocation = 3 
InsertCursor.CursorType = 1
													
InsertCursor.open SQL, conn, 1,3

insertcursor.MoveFirst

b = Replace(InsertCursor("Link"),"CookieUserID",CookieUserID,1, -1, 1)



InsertCursor.Close

ipad = trim(request.servervariables("REMOTE_ADDR"))

sql = "select * from stores where id = " & shop & ";"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set InsertCursor = Server.CreateObject("ADODB.recordset")
InsertCursor.CursorLocation = 3 
InsertCursor.CursorType = 1
													
InsertCursor.open SQL, conn, 1,3

insertcursor.MoveFirst

store = shop
failed = shop
cashback = InsertCursor("shoppin")
shop = insertcursor("productname")
linkit = Replace(InsertCursor("Link"),"CookieUserID",CookieUserID,1, -1, 1)
id = InsertCursor("ID")
merchant = insertcursor("merchant")

if cookieuserid = "" then
recordmem = ipad
else
recordmem = cookieuserid
end if

%>
 <table border="0" width="95%" ID="Table1">
	<tr>
		<td class="standart" valign="top" align="left" bgcolor="#ffffff" >
			<a href="http://www.funds4me.co.uk/shoprequired.asp?shop=<%=Request.querystring("shop")%>&passedmem=mem<%=cookieuserid%>F4MEUK" target=_parent><img src="http://www.funds4me.co.uk/images/logo.gif" border=0></a></td>
			<td class="standart" valign="top" align="left"><a href="http://www.funds4me.co.uk/shoprequired.asp?shop=<%=Request.querystring("shop")%>&passedmem=mem<%=cookieuserid%>F4MEUK" target=_parent><%=InsertCursor("Banner")%></a><br>
		<%=InsertCursor("Merchant")%> </td>

	</tr>
	
	<tr>
		<td class="standart" valign="top" align="left" bgcolor="#ffffff" colspan=2><B>STOP: Earn a maximum Cash Back payment of <%=InsertCursor("Charityrebatesval")%> by following the link on our site.</B></td>
	</tr>
	<tr>
			<td class="standart" valign="top" align="left" colspan=2><br><u><b>Coupons & Special Offers</b></u><br>
			
			
			
			<table cellspacing="1" Border="0">
													<tr>
														<%
									Dim objConn
									Set objConn = Server.CreateObject("ADODB.Connection")
									objConn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")'connection, user, ID	
																		
									sSQL = "SELECT * FROM specials"
									sSQL = sSQL & " WHERE (Live <> 'no' and ((country = 'uk') or (country = 'all')) and storeID = '" & Request.querystring("shop") & "') order by title"
									Set oRs = Server.CreateObject("ADODB.RecordSet")
									oRs.Open sSQL, objConn, 1, 3
									
									if oRs.BOF and oRs.EOF then
										
									else
										oRs.MoveFirst
										Do 
											If oRs.EOF Then Exit Do
											title = "<a href='" & oRs("url") & "'>" & oRs("title") & "</a>"
											banner = oRs("image") 
											text = oRs("description")
											
											CookieUserID = ""
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
											cuid = "sod" & date()
											CookieUserID = replace(cuid,"/","-",1,-1,1)
											AmazonUK = "dropthegrieffrom"
											else
											name = Request.cookies("details")("UserName")
											end if

											CookieUserID =  CookieUserID & "F4MECOUK"
											
											text1 = replace(text,"CookieUserID",CookieUserID,1,-1,1)
											text = text1
											
											startdate = oRs("startdate")
											enddate = oRs("enddate")
											ID = oRs("ID")
											Live = oRs("live")
											
											Set oRs1 = Server.CreateObject("ADODB.RecordSet")										
											sSQL = "SELECT * FROM specials WHERE ID = " & ID
											oRs1.Open sSQL, objConn, 1, 3
											if live = "yes" then
												if enddate < date() then
													oRs1("live") = "no"
													live = "no"
												end if
											else
												if startdate <= date() then
													oRs1("live") = "yes"
													live = "yes"
												end if
											end if
											oRs1.Update
											oRs1.Close
											if live = "yes" then
											%>
														<tr>
															<td align="left" class="standart"></td>
															<td align="left" colspan="3" class="LargeBlack"><b><%=title%></b></td>
															<td align="left" class="standart">&nbsp;&nbsp;</td>
														</tr>
														<tr>
															<td align="left" class="standart"></td>
															<td colspan="3">This offer Starts: <%=startdate%> and Ends: <%=enddate%></td>
															<td align="left" class="standart">&nbsp;&nbsp;</td>
														</tr>
														
														<tr>
															<td align="left" class="standart"></td>
															<td colspan="3"><%=text%></td>
															<td align="left" class="standart">&nbsp;&nbsp;</td>
														</tr>
														<tr>
															<td colspan="5"><hr>
															</td>
														</tr>
														<%
														end if
														oRs.MoveNext
														
										Loop    
							        end if
									  %>
												</table>
			
			
			
			
			
			</td>
	</tr>
	<tr>
		<td class="standart" valign="top" align="left" colspan="2"><br><br><u><b>About this Merchant</u></b><br><%=InsertCursor("notes")%><br><Br></td>
	</tr>

</table>

<%
InsertCursor.close
%>
</body>
</html>







