<!--#INCLUDE FILE="newheader.inc"-->



<!--content -->
<div id="content">
<%
	usecountry = Request.QueryString("country")
	useID = Request.QueryString("catID")
	if (usecountry = "") or useID = "all" then 
	usecountry = "like '%,uk,%' OR ListCountry = ',all,'"
	usecountrytext = ""
	country = "UK"
	else
	if usecountry = "all" then
	usecountry = "> ''"
	else
	usecountry = "like '%," & Request.QueryString("country") & ",%' OR ListCountry = ',all,'"
	end if
	usecountrytext = " (from " & Request.QueryString("country") & ")"
	end if
	%>


	<table cellspacing="1"  Border="0" >
		<tr>
			<td >
				<b>Stores listed under the <%=Request.QueryString("catname")%> category</b><br />
				Filter by Category <br />
				<a href="stores.asp?catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>&country=<%=request.QueryString("country")%>"><%=Request.QueryString("catname")%></a>
				<%
				sql = ""
				usew = ""
				sql = "select * from category where Master = '" & useID & "' order by CategoryName;"
								
				set rs = Server.CreateObject("ADODB.recordset")
				rs.CursorLocation = 3 
				rs.CursorType = 1
				
				rs.open SQL, objConn3, 1,3
					do while not rs.eof
						response.write(" : <a href=""stores.asp?catID=" & request.QueryString("catID") & "&catname=" & request.QueryString("catname") & "&country=" & request.QueryString("country") & "&filterby=" & rs("Categoryid") & """>" & rs("CategoryName") & "</a>")
						rs.movenext
					loop
				rs.close
				%>
		<br>
			Filter by country - current country : <%=country%>
			<a href="stores.asp?country=uk&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_gb.gif" border=0 alt="England, Scotland, Wales & Northern Ireland"></a>
			<a href="stores.asp?country=usa&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_us.gif" border=0 alt="USA"></a></a>
			<a href="stores.asp?country=ca&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_CA.gif" border=0 alt="Canada"></a></a>
			<a href="stores.asp?country=aus&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_AUS.gif" border=0 alt="Australia"></a></a>
			<a href="stores.asp?country=ie&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_IE.gif" border=0 alt="Ireland"></a></a>
			<a href="stores.asp?country=at&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_AT.gif" border=0 alt="Austria"></a></a>
			<a href="stores.asp?country=be&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_BE.gif" border=0 alt="Belgium"></a></a>
			<a href="stores.asp?country=dk&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_DK.gif" border=0 alt="Denmark"></a></a>
			<a href="stores.asp?country=fi&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_FI.gif" border=0 alt="Finland"></a></a>
			<a href="stores.asp?country=fr&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_FR.gif" border=0 alt="France"></a></a>
			<a href="stores.asp?country=de&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_DE.gif" border=0 alt="Germany"></a></a>
			<a href="stores.asp?country=it&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_IT.gif" border=0 alt="Italy"></a></a>
			<a href="stores.asp?country=nl&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_NL.gif" border=0 alt="Netherlands"></a></a>
			<a href="stores.asp?country=no&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_NO.gif" border=0 alt="Norway"></a></a>
			<a href="stores.asp?country=es&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_ES.gif" border=0 alt="Spain"></a></a>
			<a href="stores.asp?country=se&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_SE.gif" border=0 alt="Sweden"></a></a>
			<a href="stores.asp?country=ch&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_CH.gif" border=0 alt="Switzerland"></a></a>
			<a href="stores.asp?country=all&catID=<%=request.QueryString("catID")%>&catname=<%=request.QueryString("catname")%>"><img src="siteimages/flag_WW.gif" border=0 alt="All stores world wide"></a>
			</td>
		</tr>
	</table>
	<table>
		<tr>
			<td colspan=5>
	<%
							
				filterby = ""
				filterby = request.QueryString("filterby")
				if filterby > "" then 
					useID = filterby
					sql = "select * from category where categoryID = '" & useID & "' order by CategoryName;"
				else
					sql = "select * from category where Master = '" & useID & "' order by CategoryName;"
				end if
			
				
				rs.open SQL, objConn3, 1,3
					do while not rs.eof
						Response.Write("</td></tr><tr><td colspan=5><hr><u>" & rs("CategoryName") & "</u>")
						useID = rs("categoryID")
												
						sql = "select * from stores WHERE ((category like '%," & useID & ",%') and ((ListCountry " & usecountry & "))) ORDER BY stores.merchant;" 
										
						set rs1 = Server.CreateObject("ADODB.recordset")
						rs1.CursorLocation = 3 
						rs1.CursorType = 1
						
						rs1.open SQL, objConn3, 1,3
						
						'response.Write(SQL)
						
						fFirstPass = true
							%>
						
								
			
		
		<tr>
			<td colspan=5>
				<hr>
			</td>
		</tr>
						
						<%
						
								Do while not rs1.eof    
									If Not fFirstPass Then        
										rs1.MoveNext    
									Else        
										fFirstPass = False    
									End If    
									If rs1.EOF Then Exit Do
									a = rs1("productname")
									'a = "<font size = '2'><a href='shoprequired.asp?shop=" & rs("ID") & "'>" & rs("Merchant") & "</a></font>"
									%> 
									
									<tr>
									
									<td align="left" class="standart"></td>
									<td align="left" class="LargeBlack"><b><%=a%>&nbsp;&nbsp;<%=rs1("shoppin")%></b></td>
									<td align="left" class="standart">&nbsp;&nbsp;</td>
									<td align="left" class="standart"></td>
									<td align="left" class="standart">&nbsp;&nbsp;</td>
									</tr>
									<tr>
									<td colspan=5>&nbsp;</td>
									</tr>
									<tr>
									<td align="left" class="standart"></td>
									<td colspan=3><a href="<%=replace(rs1("link"),"CookieUserID",cookieuserid,1,-1,1)%>" target="_blank">Click Here</a> to Shop at <%=rs1("merchant")%> and earn up to <%=rs1("shoppin")%> cash back.<br><%=rs1("Notes")%></td>
									<td align="left" class="standart">&nbsp;&nbsp;</td>
									</tr>
									<tr>
									<td colspan=5><hr></td>
									</tr>
									<%
								rs1.movenext
								Loop    
							rs1.close
												
						
						rs.movenext
					loop			
				rs.close
	

%>
		</td>
	</tr>
</table>

<%
function showstores(useID)


end function		
		%>






</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->