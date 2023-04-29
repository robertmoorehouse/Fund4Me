<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">



								<table border="0" cellpadding="0" cellspacing="0" width="490" summary="top">
								
									<tr>
										<td class="WhiteText" align="left" valign="middle" bgcolor="#0099cc">
						
											&nbsp;<b><i>Latest Coupons and Special Offers for Funds4Me users</i></b>
										</td>
									</tr>
									<tr>
										<td class="standart" align="left">
											
											<font size="2">
												
												
												<table cellspacing="1" Border="0">
													<tr>
														<%
									Dim objConn
									Set objConn = Server.CreateObject("ADODB.Connection")
									objConn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")'connection, user, ID	
																		
									sSQL = "SELECT * FROM specials"
									sSQL = sSQL & " WHERE (Live <> 'no' and ((country = 'uk') or (country = 'all'))) order by title"
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
															<td colspan="3"><%=banner%></td>
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
											</font>
										</td>
									</tr>
								</table>
							 
  
  
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->