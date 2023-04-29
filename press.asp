<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">



	<table border="0" cellpadding="0" cellspacing="0" width="498" valign="top">
      <tr>
        <td width="498" class="standart" valign="top" align="center">
                
			<table width="490" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ffffff">
				<td class="WhiteText" width="90" valign="top" align="left"></td>
				<td colspan="2" class="BlueText" width="408" valign="top" align="center">
					<img src="images/pix.gif" width="408" height="10" ><br>
					<b>News & Press Items</b><br><br><br>
				</td>
			</tr>
			</table>
			</table>
			
		<%
		
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		sql = "select * from press where site = 'Funds4Me.co.uk' Order BY id desc;"
		rs.Open SQL, objConn1, 1, 3
%>
			<table width="440" cellpadding="0" cellspacing="0" ID="Table1">


			
			<tr bgcolor="#ff9900">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="8" height="1">
				</td>
				<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
					<b><i>Press and News Items</i></b>
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="440" height="4">
				</td>
			</tr>
	<%
	fFirstPass = true		
			
	Do    
	If Not fFirstPass Then        
	rs.MoveNext    
	Else        
	fFirstPass = False    
	End If    
	If rs.EOF Then Exit Do
	
	%> 
	<tr bgcolor="#ffcc99">
	<td class="WhiteText" width="8" valign="top" align="left">
		<img src="images/pix.gif" width="15" height="15">
	</td>     
	<td class="WhiteText"  valign="top" align="left"><a href="pressitem.asp?pid=<%=rs("id")%>"><%= rs("title") %></a></td>
	<td class="WhiteText"  valign="top" align="left" nowrap><%=rs("newsdate")%>&nbsp;&nbsp;</td>
	</tr>
	<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" valign="top" align="left">&nbsp;
					
				</td>
				<td class="LargeBlack" valign="top" align="left">
					
				</td>
			</tr>
	<%          
	        
	Loop    
	        
	  %>
	  </tr>
	  			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" valign="top" align="left">&nbsp;
					
				</td>
				<td class="LargeBlack" valign="top" align="left">
					
				</td>
			</tr>
	  
	  			<tr bgcolor="#ff9900">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="8" height="1">
				</td>
				<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
					<b><i>Press Contact</i></b>
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="440" height="4">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td colspan=2 class="LargeBlack" valign="top" align="left">
					<b>Press Pack and further information available from: - </b><br>
Chris Shaw<br>
Pink Elephant PR<br>
Telephone: -	+44 (0) 1484 341 001<br>
Facsimile: -	+44 (0) 1484 341 003<br>
E-mail: -	<a href="mailto:chris@pinkelephantpr.com">chris@pinkelephantpr.com</a><br>
				</td>
			</tr>
	  			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" valign="top" align="left">&nbsp;
					
				</td>
				<td class="LargeBlack" valign="top" align="left">
					
				</td>
			</tr>
	  
</table>
		
       
      

						

	
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->