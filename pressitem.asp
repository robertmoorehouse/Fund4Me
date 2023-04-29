<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">



	<table border="0" cellpadding="0" cellspacing="0" width="498" valign="top">
      <tr>
        <td colspan="2" class="BlueText" width="408" valign="top" align="center">
			<img src="images/pix.gif" width="408" height="10" ><br>
			<b><a href="press.asp">Back to News & Press Items</a></b><br><br><br>
		</td>
	  </tr>
	</table>
			
		<%
		
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		pressID = request.QueryString("pid")
		sql = "select * from press WHERE id = " & pressID
		rs.Open SQL, objConn1, 1, 3
		presstext = rs("story")
		image1 = rs("image1")
		image2 = rs("image2")
		image3 = rs("image3")
		rs.Close
%>
			
<table border="0" cellpadding="0" cellspacing="0" width="490" valign="top" ID="Table1">
      
      
      <tr>
        <td width="490" valign="top" align="left">
          	<%=presstext%><br><br>
        </td>
      </tr>
      <%if image1 > "" then %>
      
      <tr>
        <td width="490" class="standart" valign="top" align="left">
          	<img src="images/press/<%=image1%>"><br><br>
        </td>
      </tr>
      <%end if%>
       <%if image2 > "" then %>
      
      <tr>
        <td width="490" class="standart" valign="top" align="left">
          	<img src="images/press/<%=image2%>"><br><br>
        </td>
      </tr>
      <%end if%>
       <%if image3 > "" then %>
      
      <tr>
        <td width="490" class="standart" valign="top" align="left">
          	<img src="images/press/<%=image3%>"><br><br>
        </td>
      </tr>
      <%end if%>
	</table>
	
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->
