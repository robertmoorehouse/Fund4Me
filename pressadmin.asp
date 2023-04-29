<%@ Language = VBScript%>
<%	
	iMainMenu = 6
	iSchoolMenu = 0 
	
	sUser = Trim(Request.Form("username"))
	sPassword = Trim(Request.Form("password"))
		
	Dim objConn
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")'connection, user, ID	
	if session("memID") = "29" then
		temp = ""
		temp = request.QueryString("id")
		if temp = "new" then temp = 1
		if temp <> "" then Session("pressID") = temp
		if temp = "" then Session("pressID") = 0
		loggedIn = "yes"
		sSQL = "SELECT * FROM press"
		sSQL = sSQL & " WHERE ID = " & cint(Session("pressID"))
	else
		loggedIn = "never"
	end if

	IF loggedIn = "never" then
	else		
		Set oRs = Server.CreateObject("ADODB.RecordSet")
		oRs.Open sSQL, objConn, 1, 3
		
		
		if oRs.EOF and oRs.BOF then
			loggedIn = "yes"
		else
			loggedIn = "yes"
			addnew = request.QueryString("id")
			if (addnew = "new") or (addnew = "") then
			else
				oRs.MoveFirst
				pressID = oRs("ID")
				session("pressID") = oRs("ID")
				promotext = oRs("story")
				Presstitle = oRs("title")
				PressImage1 = oRs("Image1")
				PressImage2 = oRs("Image2")
				PressImage3 = oRs("Image3")
				pressDate = oRs("newsDate")
				presslive = oRs("live")
				pressitem = oRs("item")
			end if
  			oRs.close()
  			Set oRs = Nothing
  			
		end if	
	end if
	
	

%>
<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">




<SCRIPT language=Javascript src="js/postcodeutil.js" 
type=text/javascript></SCRIPT>

<SCRIPT language=Javascript src="js/stringutil.js" 
type=text/javascript></SCRIPT>

<SCRIPT language=Javascript src="js/numericutil.js" 
type=text/javascript></SCRIPT>

<SCRIPT language=Javascript src="js/errorutil.js" 
type=text/javascript></SCRIPT>

<SCRIPT language=Javascript src="js/generalutil.js" 
type=text/javascript></SCRIPT>
<title>Funds 4 Me</title>
</head>
<SCRIPT language=JavaScript> 
 

function verifyUploadDetails(submission){                                                                                 

	submission.fullDetails.value = submission.fullDetailsDHTMLEdit1.DOM.body.outerHTML;
	return true;
} 
</SCRIPT>
<script Language="JavaScript"><!--
	function setValues(checkValue) 
	{	
		var theForm = window.document.updateForm;
		theForm.fullDetails.focus();
	}	
	function RollOver(imgName,num)
	{ 
		document.images[imgName].src = "images/buttons/accMenu" + num + "On.gif"
	}
	function RollOut(imgName,num)
	{ 
		document.images[imgName].src = "images/buttons/accMenu" + num + ".gif"
	}
		
	
-->
</script>
<SCRIPT language=JavaScript src="js/dhtmled.js"></SCRIPT>
<SCRIPT language=JavaScript src="js/jscript_code.htm"></SCRIPT>


	<table border="0" cellpadding="0" cellspacing="0" width="498" valign="top">
      <tr>
        <td width="498" class="standart" valign="top" align="center">
        <%if loggedIn = "yes" then %>
        
        
			<table width="490" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ffffff">
				<td class="WhiteText" width="90" valign="top" align="left"></td>
				<td colspan="2" class="BlueText" width="408" valign="top" align="center">
					<img src="images/pix.gif" width="408" height="10" ><br>
					<b>News Items</b><br><br><br>
					<font size="2"><b><i>Your Profile</i></b></font>
				</td>
			</tr>
			</table>
			</table>
			
		<%
		
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		sql = "select * from press where site = 'Funds4Me.co.uk' Order BY id desc;"
		rs.Open SQL, objConn, 1, 3
%>
<a href="pressadmin.asp?id=new">Add New</a><br><br>
<table cellspacing="3" width="440" Border="0" ID="Table1">
	<tr>
		<td align="left" ><font size="2">Item</td>
		<td align="left"><font size="2">Title</td>
		<td align="left"><font size="2">Live</td>
		<td align="left" nowrap><font size="2">Listing-No</td>
	</tr>
	<tr>
		<td align="left" ><font size="2"></td>
		<td align="left"><font size="2">&nbsp;</td>
		<td align="left"><font size="2"></td>
		<td align="left"><font size="2"></td>
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
	       
	<tr>
		<td align="left" ><font size="2"><%= rs("id") %></td>
		<td align="left"><font size="2"><a href="pressadmin.asp?id=<%=rs("id")%>"><%= rs("title") %></a></td>
		<td align="left"><font size="2"><%=rs("live")%></td>
		<td align="left"><font size="2"><%=rs("item")%></td>
	</tr>
	<%          
	        
	Loop    
	        
	  %>
</table>
		
		
		<form name="updateForm" action="savePressDetails.asp?id=<%=request.QueryString("id")%>" method="post">

		<table width="440" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ff9900">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="8" height="1">
				</td>
				<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
					<b><i>Details</i></b>
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
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Title: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="Presstitle" size="30" value="<%=Presstitle%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>image1: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="Pressimage1" size="30" value="<%=Pressimage1%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>image2: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="Pressimage2" size="30" value="<%=Pressimage2%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>image3: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="Pressimage3" size="30" value="<%=Pressimage3%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Date: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="PressDate" size="30" value="<%=PressDate%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Live: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<SELECT tabIndex=1 name=pressLive ID="pressLive">
					<OPTION value=<%=pressLive%> selected><%=pressLive%></OPTION>
					<%if pressLive="Yes" then %>
					<OPTION value="No">No</OPTION>
					<%else%>
					<OPTION value="Yes">Yes</OPTION>
					<%end if%>
					</select>
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Item: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<SELECT tabIndex=1 name=pressitem >
					<%
					for n = 1 to 50
						if cint(pressitem) = n then
						%>
						<OPTION value=<%=pressitem%> selected><%=pressitem%></OPTION>
						<%
						else
						%>
						<OPTION value=<%=n%>><%=n%></OPTION>
						<%
						end if
                    next
                    %>
                    </select>
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">&nbsp;
					
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					
				</td>
			</tr>
			
			<tr bgcolor="#ff9900">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="8" height="1">
				</td>
				<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
					<b><i>Page Contents</i></b>
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">&nbsp;
					
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					
				</td>
			</tr>
			
			
			
			
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" colspan=2 valign="top" align="left">
					<b>Story:</font><br>
				
					
					
					
					
	<script LANGUAGE="javascript" FOR="fullDetailsDHTMLEdit1" EVENT="DisplayChanged">
	<!--
	return DHTMLEdit1_DisplayChanged(document.forms.updateForm.fullDetailsDHTMLEdit1);
	//-->
	</script>
	
	<script LANGUAGE="javascript" FOR="fullDetailsDHTMLEdit1" EVENT="DocumentComplete">
	<!--
	return DHTMLEdit1_DocumentComplete(document.forms.updateForm.fullDetailsDHTMLEdit1);
	//-->
	</script>
	
	<script LANGUAGE="javascript" FOR="fullDetailsDHTMLEdit1" EVENT="ShowContextMenu">
	<!--
	return DHTMLEdit1_ShowContextMenu(document.forms.updateForm.fullDetailsDHTMLEdit1);
	//-->
	</script>
	
	<script LANGUAGE="javascript" FOR="fullDetailsDHTMLEdit1" EVENT="ContextMenuAction(itemIndex)">
	<!--
	return DHTMLEdit1_ContextMenuAction('document.forms.updateForm.fullDetailsDHTMLEdit1',document.forms.updateForm.fullDetailsDHTMLEdit1,itemIndex);
	//-->
	</script>
	
	
	<DIV STYLE="border-width: thin; border-color: Orange; border-style: inset; text-align: left; background-color: buttonface; width: 450;">

	

	
	
	
<IMG SRC="js/Bar.gif" WIDTH=9 HEIGHT=24 ALT="" BORDER="0">
  <A NAME="DECMD_BOLD" onclick="return DECMD_BOLD_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsibold,0)"
  onMouseOut="outstyle(fullDetailsibold,0)"
  onMouseDown="downstyle(fullDetailsibold,0);"
  onMouseUp="upstyle(fullDetailsibold,0);"><img
  class="tbButton" src="js/bold.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsibold" ALT="Bold"></A><A
	NAME="DECMD_ITALIC" onclick="return DECMD_ITALIC_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiitalic,0)"
  onMouseOut="outstyle(fullDetailsiitalic,0)"
  onMouseDown="downstyle(fullDetailsiitalic,0);"
  onMouseUp="upstyle(fullDetailsiitalic,0);"><img
  class="tbButton" src="js/italic.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiitalic" ALT="Italic"></A><A
	NAME="DECMD_ITALIC" onclick="return DECMD_UNDERLINE_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiunderline,0)"
  onMouseOut="outstyle(fullDetailsiunderline,0)"
  onMouseDown="downstyle(fullDetailsiunderline,0);"
  onMouseUp="upstyle(fullDetailsiunderline,0);"><img
  class="tbButton" src="js/under.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiunderline" ALT="Underline"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">
  <A NAME="DECMD_FONT" onclick="return DECMD_FONT_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsifont,0)"
  onMouseOut="outstyle(fullDetailsifont,0)"
  onMouseDown="downstyle(fullDetailsifont,0);"
  onMouseUp="upstyle(fullDetailsifont,0);"><img
  class="tbButton" src="js/font.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsifont" ALT="Edit Font"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">
	<A
	NAME="DECMD_CUT" onclick="return DECMD_CUT_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsicut,0)"
  onMouseOut="outstyle(fullDetailsicut,0)"
  onMouseDown="downstyle(fullDetailsicut,0);"
  onMouseUp="outstyle(fullDetailsicut,0);"><img
  class="tbButton" src="js/cut.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsicut" ALT="Cut"></A><A
	NAME="DECMD_COPY" onclick="return DECMD_COPY_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsicopy,0)"
  onMouseOut="outstyle(fullDetailsicopy,0)"
  onMouseDown="downstyle(fullDetailsicopy,0);"
  onMouseUp="outstyle(fullDetailsicopy,0);"><img
  class="tbButton" src="js/copy.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsicopy" ALT="Copy"></A><A
	NAME="DECMD_PASTE" onclick="return DECMD_PASTE_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsipaste,0)"
  onMouseOut="outstyle(fullDetailsipaste,0)"
  onMouseDown="downstyle(fullDetailsipaste,0);"
  onMouseUp="outstyle(fullDetailsipaste,0);"><img
  class="tbButton" src="js/paste.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsipaste" ALT="Paste"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">	
<A
	NAME="DECMD_UNDO" onclick="return DECMD_UNDO_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiundo,0)"
  onMouseOut="outstyle(fullDetailsiundo,0)"
  onMouseDown="downstyle(fullDetailsiundo,0);"
  onMouseUp="outstyle(fullDetailsiundo,0);"><img
  class="tbButton" src="js/undo.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiundo" ALT="Undo"></A><A
	NAME="DECMD_REDO" onclick="return DECMD_REDO_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiredo,0)"
  onMouseOut="outstyle(fullDetailsiredo,0)"
  onMouseDown="downstyle(fullDetailsiredo,0);"
  onMouseUp="outstyle(fullDetailsiredo,0);"><img
  class="tbButton" src="js/redo.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiredo" ALT="Redo"></A>	
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">	
	<A
		NAME="DECMD_HYPERLINK" onclick="return DECMD_HYPERLINK_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1,'fullDetails');"
  onMouseOver="overstyle(fullDetailsilink,0)"
  onMouseOut="outstyle(fullDetailsilink,0)"
  onMouseDown="downstyle(fullDetailsilink,0);"
  onMouseUp="outstyle(fullDetailsilink,0);"><img
  class="tbButton" src="js/link.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsilink" ALT="Link"></A>
	<A
	NAME="DECMD_VISIBLEBORDERS" onclick="return DECMD_VISIBLEBORDERS_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiborders,0)"
  onMouseOut="outstyle(fullDetailsiborders,0)"
  onMouseDown="downstyle(fullDetailsiborders,0);"
  onMouseUp="outstyle(fullDetailsiborders,0);"><img
  class="tbButton" src="js/borders.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiborders" ALT="Show All"></A>


	

	
	<IMG SRC="js/seplong.gif" WIDTH=446 HEIGHT=2 ALT="" BORDER="0"><BR>
	
<IMG SRC="js/Bar.gif" WIDTH=9 HEIGHT=24 ALT="" BORDER="0">
  <A NAME="DECMD_JUSTIFYLEFT" onclick="return DECMD_JUSTIFYLEFT_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsileft,0)"
  onMouseOut="outstyle(fullDetailsileft,0)"
  onMouseDown="downstyle(fullDetailsileft,0);"
  onMouseUp="upstyle(fullDetailsileft,0);"><img
  class="tbButton" src="js/left.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsileft" ALT="Align Left"></A><A
	NAME="DECMD_JUSTIFYCENTER" onclick="return DECMD_JUSTIFYCENTER_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsicenter,0)"
  onMouseOut="outstyle(fullDetailsicenter,0)"
  onMouseDown="downstyle(fullDetailsicenter,0);"
  onMouseUp="upstyle(fullDetailsicenter,0);"><img
  class="tbButton" src="js/center.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsicenter" ALT="Center"></A><A
	NAME="DECMD_JUSTIFYRIGHT" onclick="return DECMD_JUSTIFYRIGHT_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiright,0)"
  onMouseOut="outstyle(fullDetailsiright,0)"
  onMouseDown="downstyle(fullDetailsiright,0);"
  onMouseUp="upstyle(fullDetailsiright,0);"><img
  class="tbButton" src="js/right.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiright" ALT="Align Right"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">
	<A
	NAME="DECMD_ORDERLIST" onclick="return DECMD_ORDERLIST_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiorder,0)"
  onMouseOut="outstyle(fullDetailsiorder,0)"
  onMouseDown="downstyle(fullDetailsiorder,0);"
  onMouseUp="upstyle(fullDetailsiorder,0);"><img
  class="tbButton" src="js/numlist.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiorder" ALT="Ordered List"></A><A
	NAME="DECMD_UNORDERLIST" onclick="return DECMD_UNORDERLIST_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiunorder,0)"
  onMouseOut="outstyle(fullDetailsiunorder,0)"
  onMouseDown="downstyle(fullDetailsiunorder,0);"
  onMouseUp="upstyle(fullDetailsiunorder,0);"><img
  class="tbButton" src="js/bullist.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiunorder" ALT="Unordered List"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">
	<A
	NAME="DECMD_OUTDENT" onclick="return DECMD_OUTDENT_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiout,0)"
  onMouseOut="outstyle(fullDetailsiout,0)"
  onMouseDown="downstyle(fullDetailsiout,0);"
  onMouseUp="upstyle(fullDetailsiout,0);"><img
  class="tbButton" src="js/deindent.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiout" ALT="Unindent"></A><A
	NAME="DECMD_INDENT" onclick="return DECMD_INDENT_onclick(document.forms.updateForm.fullDetailsDHTMLEdit1);"
  onMouseOver="overstyle(fullDetailsiin,0)"
  onMouseOut="outstyle(fullDetailsiin,0)"
  onMouseDown="downstyle(fullDetailsiin,0);"
  onMouseUp="upstyle(fullDetailsiin,0);"><img
  class="tbButton" src="js/inindent.gif" WIDTH="23" HEIGHT="22"
	NAME="fullDetailsiin" ALT="Indent"></A>


	
	</DIV>
	
	<OBJECT CLASSID="clsid:2D360201-FFF5-11D1-8D03-00A0C959BC0A"
        WIDTH="450"
        HEIGHT="150"
        BORDER="0"
        NAME="fullDetailsbox"
        ID="fullDetailsDHTMLEdit1"
        VIEWASTEXT="YES"></OBJECT>
		
		<textarea cols="0" rows="0" style="visibility:hidden" name="fullDetails"><%=PromoText%></textarea>
		
		<SCRIPT LANGUAGE=JAVASCRIPT>
				if ("<%=PromoText%>" == "") {
				} else {go('fullDetailsbox','fullDetails','home.asp');}
		</SCRIPT>
					
					
					
		
		
		
		
			
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">&nbsp;
					
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					
				</td>
			</tr>
			<tr bgcolor="#ffffff">
				<td class="WhiteText" width="15" valign="middle" align="left">
					<img src="images/pix.gif" width="15" height="20">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
				</td>
				<td class="LargeBlack" width="310" valign="top" align="right">
					<A href="javascript:if(verifyUploadDetails(window.document.updateForm)){window.document.updateForm.submit();};"><img src="images/buttons/update.gif" border="0"></a>
				</td>
			</tr>
			
  <%		else
  	
						If (loggedIn = "no") then
							Response.Write("<br><b>Your details have not been found, please try again </a></b>")
							Response.Write("<br><br><img src='images/blueSquare.gif' width='490' height='1' border='0'><br><br>")
						else %>
							
						<table width="360" cellpadding="0" cellspacing="0">
							<tr bgcolor="#ffffff">
								<td class="largeblack" valign="top" align="center"><br>
						<b>Press Admin Area</b>
						<br><br>
						</td></tr></table>
					<%	end if
	                  
	               
	                 
	%>				
					<span class="BlueText"><b>Please enter your User Name and Password.</b>
					</span><br><br>
					<form name="LoginForm" action="pressadmin.asp" method="post">
					<table width="360" cellpadding="0" cellspacing="0">
						<tr bgcolor="#0099CC">
							<td class="WhiteText" width="8" valign="top" align="left">
								<img src="images/pix.gif" width="8" height="1">
							</td>
							<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
								<b><i>User Details</i></b>
							</td>
						</tr>
						<tr bgcolor="#99CCFF">
							<td class="WhiteText" width="8" valign="top" align="left">
								<img src="images/pix.gif" width="8" height="1">
							</td>
							<td class="Standart" width="176" valign="top" align="left">
								User Name:<br>
								<input type="text" name="username" size="20">
							</td>
							<td class="Standart" width="176" valign="top" align="left">
								Password:<br>
								<input type="password" id="password" name="password" size="20">
							</td>
						</tr>
						<tr bgcolor="#99CCFF">
							<td class="WhiteText" width="8" valign="top" align="left">
								<img src="images/pix.gif" width="8" height="1">
							</td>
							<td class="WhiteText" width="176" valign="top" align="left">
								<a href="JavaScript:document.LoginForm.submit();"><img src="images/buttons/login.gif" border="0"></a>
							</td>
							<td class="WhiteText" width="176" valign="top" align="right">
							</td>
						</tr>
					</table><br>
					</form> 			
							
<%					
  				end if
  			
%>	
        </td>
      </tr>

      </tr>
    </table>   
      
      

						
	
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->
