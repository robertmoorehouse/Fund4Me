<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	iMainMenu = 6
	iSchoolMenu = 0 

	sUser = Trim(Request.Form("username"))
	sPassword = Trim(Request.Form("password"))
	sEmail = Trim(Request.Querystring("email"))
	
	if (sEmail = "known") then
		sUser = ""
		sPassword = ""
		Session("memID") = ""
	end if
	
	Dim objConn
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")'connection, user, ID	
	
	if (sUser = "") and (sPassword = "") then
		if Session("promoID") <> "" then
			loggedIn = "yes"
			sSQL = "SELECT * FROM promo"
			sSQL = sSQL & " WHERE ID = " & cint(Session("promoID"))
		else
			loggedIn = "never"
		end if
	else
		sSQL = "SELECT * FROM promo"
		sSQL = sSQL & " WHERE username = '" & sUser & "' and password = '" & sPassword & "'"
	end if

	IF loggedIn = "never" then
	else		
		Set oRs = Server.CreateObject("ADODB.RecordSet")
		oRs.Open sSQL, objConn, 1, 3
				
		if oRs.EOF and oRs.BOF then
			loggedIn = "no"
		else
			loggedIn = "yes"
			oRs.MoveFirst
			promoID = oRs("ID")
			session("promoID") = oRs("ID")
			promotext = oRs("FormText")
			PromoName = oRs("Company")
			PromoDescription = oRs("Description")
			PromoImage = oRs("Image")
			Promourl = oRs("url")
			Promoemail = oRs("email")
			PromoAddress = oRs("Address")
			PromoAddress1 = oRs("Address1")
			PromoTown = oRs("Town")
			PromoCounty = oRs("County")
			PromoCountry = oRs("Country")
			PromoPostCode = oRs("PostCode")
			PromoFax = oRs("Fax")
			PromoTel = oRs("Tel")
			PromoUserName = oRs("userName")
			PromoPassword = oRs("password")
			strfullDetails = oRs("FormText")
  			oRs.close()
  			Set oRs = Nothing
  			
		end if	
	end if
	
	

%>

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
	submission.FullDescription.value = submission.FullDescriptionDHTMLEdit1.DOM.body.outerHTML;
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
	function verify() {
		submit = true;
		if (document.updateForm.firstName.value == "") {
			alert("You must enter a First Name");
			document.updateForm.firstName.focus();
			submit = false;
		} else {
			if (document.updateForm.address1.value == "") {
				alert("You must enter a Address");
				document.updateForm.address1.focus();
				submit = false;
			}		else {
				if (document.updateForm.city.value == "") {
					alert("You must enter a City");
					document.updateForm.city.focus();
					submit = false;
				}	else {
					if (document.updateForm.county.value == "") {
						alert("You must enter a County");
						document.updateForm.county.focus();
						submit = false;
					}		else {	
						if (document.updateForm.postcode.value == "") {
							alert("You must enter a Postcode");
							document.updateForm.postcode.focus();
							submit = false;
						}	else {
							if (document.updateForm.email.value == "") {
								alert("You must enter an Email Address");
								document.updateForm.email.focus();
								submit = false;
							}	else {	
								if (document.updateForm.username.value == "") {
									alert("You must enter a User Name");
									document.updateForm.username.focus();
									submit = false;
								}	else {	
									if (document.updateForm.password.value == "") {
										alert("You must enter a Password");
										document.updateForm.password.focus();
										submit = false;
									}	else {	
										if (document.updateForm.confPassword.value != document.updateForm.password.value) {
											alert("The password confirmation is not the same as the password!");
											document.updateForm.password.value="";
											document.updateForm.confPassword.value="";
											document.updateForm.password.focus();
											submit = false;
										}	else {	
											pw = document.updateForm.password.value;
											if (pw.length < 6) {
												alert("Your password must have at least 6 characters.");
												document.updateForm.password.value="";
												document.updateForm.confPassword.value="";
												document.updateForm.password.focus();
												submit = false;
											}	else {
												emailaddress = document.updateForm.email.value;
												at = emailaddress.indexOf("@");
												dot = emailaddress.indexOf(".");
												if (at == -1){
													alert("You email address is not valid");
													submit = false;
												} else {
													if (dot == -1){
														alert("You email address is not valid");
														submit = false;
													}
												}
		
		
		}	}	}	}	}	}	}	}	}	}	
		
		
		if (submit == true) {
		document.updateForm.submit();
		}	
	
		
	}
	
	function fillValues() {
		document.updateForm.firstName.value = "<%=PromoName%>";
		document.updateForm.address1.value = "<%=PromoAddress%>";
		document.updateForm.address2.value = "<%=PromoAddress1%>";
		document.updateForm.city.value = "<%=PromoTown%>";
		document.updateForm.county.value = "<%=PromoCounty%>";
		document.updateForm.postcode.value = "<%=PromoPostCode%>";
		document.updateForm.email.value = "<%=Promoemail%>";
		document.updateForm.phone.value = "<%=PromoTel%>";
		document.updateForm.fax.value = "<%=PromoFax%>";
		document.updateForm.username.value = "<%=promousername%>";
		document.updateForm.password.value = "<%=promopassword%>";
		document.updateForm.confPassword.value = "<%=promopassword%>";
	}
-->
</script>

<SCRIPT language=JavaScript src="js/dhtmled.js"></SCRIPT>
<SCRIPT language=JavaScript src="js/jscript_code.htm"></SCRIPT>
<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">



	<table border="0" cellpadding="0" cellspacing="0" width="498" valign="top">
      <tr>
        <td width="498" class="standart" valign="top" align="center">
        <%if loggedIn = "yes" then %>
        
        
			<table width="490" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ffffff">
				<td class="WhiteText" width="90" valign="top" align="left"></td>
				<td colspan="2" class="BlueText" width="408" valign="top" align="center">
					<img src="images/pix.gif" width="408" height="10" ><br>
					<b>Some details can not be changed here.</b><br><br><br>
					<font size="2"><b><i>Your Profile</i></b></font>
				</td>
			</tr>
			</table>
			</table>
		<form name="updateForm" action="savePromoDetails.asp?promoid=<%=promoID%>" method="post" onSubmit="return verify(this);">

		<table width="440" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ff9900">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="8" height="1">
				</td>
				<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
					<b><i>Personal Details</i></b>
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
					<b>Company Name: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="firstName" size="30" value="<%=PromoName%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Address: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="address1" size="30" value="<%=PromoAddress%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="address2" size="30" value="<%=PromoAddress1%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>City: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="city" size="30" value="<%=PromoTown%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>County: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="county" size="30" value="<%=PromoCounty%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Postcode: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="postcode" size="15" value="<%=PromoPostCode%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Email: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="email" size="30"  value="<%=Promoemail%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Telephone:</b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="phone" size="30" value="<%=PromoTel%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Fax:</b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="fax" size="30" value="<%=PromoFax%>">
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
					<b><i>Account Details</i></b>
				</td>
			</tr>
			
			
			
			
				<tr bgcolor="#ffcc99">
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="400" height="4">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Username: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="username" size="30" disabled ID="Text1" value="<%=promousername%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Password: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="password" name="password" size="30" ID="Password1" value="<%=promopassword%>">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Confirm Password: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="password" name="confPassword" size="30" ID="Password2" value="<%=promopassword%>">
				</td>
			
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
					<b>Form Description: Top of Contact Forms.</font><br>
				
					
					
					
					
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
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" colspan=2 valign="top" align="left">
					<b>Full Company Description.</font><br>
				
					
					
					
					
	<script LANGUAGE="javascript" FOR="FullDescriptionDHTMLEdit1" EVENT="DisplayChanged">
	<!--
	return DHTMLEdit1_DisplayChanged(document.forms.updateForm.FullDescriptionDHTMLEdit1);
	//-->
	</script>
	
	<script LANGUAGE="javascript" FOR="FullDescriptionDHTMLEdit1" EVENT="DocumentComplete">
	<!--
	return DHTMLEdit1_DocumentComplete(document.forms.updateForm.FullDescriptionDHTMLEdit1);
	//-->
	</script>
	
	<script LANGUAGE="javascript" FOR="FullDescriptionDHTMLEdit1" EVENT="ShowContextMenu">
	<!--
	return DHTMLEdit1_ShowContextMenu(document.forms.updateForm.FullDescriptionDHTMLEdit1);
	//-->
	</script>
	
	<script LANGUAGE="javascript" FOR="FullDescriptionDHTMLEdit1" EVENT="ContextMenuAction(itemIndex)">
	<!--
	return DHTMLEdit1_ContextMenuAction('document.forms.updateForm.FullDescriptionDHTMLEdit1',document.forms.updateForm.FullDescriptionDHTMLEdit1,itemIndex);
	//-->
	</script>
	
	
	<DIV STYLE="border-width: thin; border-color: Orange; border-style: inset; text-align: left; background-color: buttonface; width: 450;">

	

	
	
	
<IMG SRC="js/Bar.gif" WIDTH=9 HEIGHT=24 ALT="" BORDER="0">
  <A NAME="DECMD_BOLD" onclick="return DECMD_BOLD_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptionibold,0)"
  onMouseOut="outstyle(FullDescriptionibold,0)"
  onMouseDown="downstyle(FullDescriptionibold,0);"
  onMouseUp="upstyle(FullDescriptionibold,0);"><img
  class="tbButton" src="js/bold.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptionibold" ALT="Bold"></A><A
	NAME="DECMD_ITALIC" onclick="return DECMD_ITALIC_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniitalic,0)"
  onMouseOut="outstyle(FullDescriptioniitalic,0)"
  onMouseDown="downstyle(FullDescriptioniitalic,0);"
  onMouseUp="upstyle(FullDescriptioniitalic,0);"><img
  class="tbButton" src="js/italic.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniitalic" ALT="Italic"></A><A
	NAME="DECMD_ITALIC" onclick="return DECMD_UNDERLINE_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniunderline,0)"
  onMouseOut="outstyle(FullDescriptioniunderline,0)"
  onMouseDown="downstyle(FullDescriptioniunderline,0);"
  onMouseUp="upstyle(FullDescriptioniunderline,0);"><img
  class="tbButton" src="js/under.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniunderline" ALT="Underline"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">
  <A NAME="DECMD_FONT" onclick="return DECMD_FONT_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptionifont,0)"
  onMouseOut="outstyle(FullDescriptionifont,0)"
  onMouseDown="downstyle(FullDescriptionifont,0);"
  onMouseUp="upstyle(FullDescriptionifont,0);"><img
  class="tbButton" src="js/font.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptionifont" ALT="Edit Font"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">
	<A
	NAME="DECMD_CUT" onclick="return DECMD_CUT_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptionicut,0)"
  onMouseOut="outstyle(FullDescriptionicut,0)"
  onMouseDown="downstyle(FullDescriptionicut,0);"
  onMouseUp="outstyle(FullDescriptionicut,0);"><img
  class="tbButton" src="js/cut.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptionicut" ALT="Cut"></A><A
	NAME="DECMD_COPY" onclick="return DECMD_COPY_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptionicopy,0)"
  onMouseOut="outstyle(FullDescriptionicopy,0)"
  onMouseDown="downstyle(FullDescriptionicopy,0);"
  onMouseUp="outstyle(FullDescriptionicopy,0);"><img
  class="tbButton" src="js/copy.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptionicopy" ALT="Copy"></A><A
	NAME="DECMD_PASTE" onclick="return DECMD_PASTE_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptionipaste,0)"
  onMouseOut="outstyle(FullDescriptionipaste,0)"
  onMouseDown="downstyle(FullDescriptionipaste,0);"
  onMouseUp="outstyle(FullDescriptionipaste,0);"><img
  class="tbButton" src="js/paste.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptionipaste" ALT="Paste"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">	
<A
	NAME="DECMD_UNDO" onclick="return DECMD_UNDO_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniundo,0)"
  onMouseOut="outstyle(FullDescriptioniundo,0)"
  onMouseDown="downstyle(FullDescriptioniundo,0);"
  onMouseUp="outstyle(FullDescriptioniundo,0);"><img
  class="tbButton" src="js/undo.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniundo" ALT="Undo"></A><A
	NAME="DECMD_REDO" onclick="return DECMD_REDO_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniredo,0)"
  onMouseOut="outstyle(FullDescriptioniredo,0)"
  onMouseDown="downstyle(FullDescriptioniredo,0);"
  onMouseUp="outstyle(FullDescriptioniredo,0);"><img
  class="tbButton" src="js/redo.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniredo" ALT="Redo"></A>	
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">	
	<A
		NAME="DECMD_HYPERLINK" onclick="return DECMD_HYPERLINK_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1,'FullDescription');"
  onMouseOver="overstyle(FullDescriptionilink,0)"
  onMouseOut="outstyle(FullDescriptionilink,0)"
  onMouseDown="downstyle(FullDescriptionilink,0);"
  onMouseUp="outstyle(FullDescriptionilink,0);"><img
  class="tbButton" src="js/link.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptionilink" ALT="Link"></A>
	<A
	NAME="DECMD_VISIBLEBORDERS" onclick="return DECMD_VISIBLEBORDERS_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniborders,0)"
  onMouseOut="outstyle(FullDescriptioniborders,0)"
  onMouseDown="downstyle(FullDescriptioniborders,0);"
  onMouseUp="outstyle(FullDescriptioniborders,0);"><img
  class="tbButton" src="js/borders.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniborders" ALT="Show All"></A>


	

	
	<IMG SRC="js/seplong.gif" WIDTH=446 HEIGHT=2 ALT="" BORDER="0"><BR>
	
<IMG SRC="js/Bar.gif" WIDTH=9 HEIGHT=24 ALT="" BORDER="0">
  <A NAME="DECMD_JUSTIFYLEFT" onclick="return DECMD_JUSTIFYLEFT_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptionileft,0)"
  onMouseOut="outstyle(FullDescriptionileft,0)"
  onMouseDown="downstyle(FullDescriptionileft,0);"
  onMouseUp="upstyle(FullDescriptionileft,0);"><img
  class="tbButton" src="js/left.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptionileft" ALT="Align Left"></A><A
	NAME="DECMD_JUSTIFYCENTER" onclick="return DECMD_JUSTIFYCENTER_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptionicenter,0)"
  onMouseOut="outstyle(FullDescriptionicenter,0)"
  onMouseDown="downstyle(FullDescriptionicenter,0);"
  onMouseUp="upstyle(FullDescriptionicenter,0);"><img
  class="tbButton" src="js/center.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptionicenter" ALT="Center"></A><A
	NAME="DECMD_JUSTIFYRIGHT" onclick="return DECMD_JUSTIFYRIGHT_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniright,0)"
  onMouseOut="outstyle(FullDescriptioniright,0)"
  onMouseDown="downstyle(FullDescriptioniright,0);"
  onMouseUp="upstyle(FullDescriptioniright,0);"><img
  class="tbButton" src="js/right.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniright" ALT="Align Right"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">
	<A
	NAME="DECMD_ORDERLIST" onclick="return DECMD_ORDERLIST_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniorder,0)"
  onMouseOut="outstyle(FullDescriptioniorder,0)"
  onMouseDown="downstyle(FullDescriptioniorder,0);"
  onMouseUp="upstyle(FullDescriptioniorder,0);"><img
  class="tbButton" src="js/numlist.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniorder" ALT="Ordered List"></A><A
	NAME="DECMD_UNORDERLIST" onclick="return DECMD_UNORDERLIST_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniunorder,0)"
  onMouseOut="outstyle(FullDescriptioniunorder,0)"
  onMouseDown="downstyle(FullDescriptioniunorder,0);"
  onMouseUp="upstyle(FullDescriptioniunorder,0);"><img
  class="tbButton" src="js/bullist.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniunorder" ALT="Unordered List"></A>
<IMG SRC="js/sep.gif" WIDTH=2 HEIGHT=22 ALT="" BORDER="0">
	<A
	NAME="DECMD_OUTDENT" onclick="return DECMD_OUTDENT_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniout,0)"
  onMouseOut="outstyle(FullDescriptioniout,0)"
  onMouseDown="downstyle(FullDescriptioniout,0);"
  onMouseUp="upstyle(FullDescriptioniout,0);"><img
  class="tbButton" src="js/deindent.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniout" ALT="Unindent"></A><A
	NAME="DECMD_INDENT" onclick="return DECMD_INDENT_onclick(document.forms.updateForm.FullDescriptionDHTMLEdit1);"
  onMouseOver="overstyle(FullDescriptioniin,0)"
  onMouseOut="outstyle(FullDescriptioniin,0)"
  onMouseDown="downstyle(FullDescriptioniin,0);"
  onMouseUp="upstyle(FullDescriptioniin,0);"><img
  class="tbButton" src="js/inindent.gif" WIDTH="23" HEIGHT="22"
	NAME="FullDescriptioniin" ALT="Indent"></A>


	
	</DIV>
	
	<OBJECT CLASSID="clsid:2D360201-FFF5-11D1-8D03-00A0C959BC0A"
        WIDTH="450"
        HEIGHT="150"
        BORDER="0"
        NAME="FullDescriptionbox"
        ID="FullDescriptionDHTMLEdit1"
        VIEWASTEXT="YES"></OBJECT>
		
		<textarea cols="0" rows="0" style="visibility:hidden" name="FullDescription"><%=PromoDescription%></textarea>
		
		<SCRIPT LANGUAGE=JAVASCRIPT>
				if ("<%=PromoDescription%>" == "") {
				} else {go('FullDescriptionbox','FullDescription','home.asp');}
		</SCRIPT>
					<!--
					<A href="javascript:alert(document.updateForm.fullDetailsDHTMLEdit1.DOM.body.parentElement.outerHTML);">
		Debug
		</A>-->
					
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
						<b>Sponsor and Promo Company Details</b>
						<br><br>
						</td></tr></table>
					<%	end if
	                  
	               
	                 
	%>				
					<span class="BlueText"><b>Please enter your User Name and password to get your account details.</b>
					</span><br><br>
					<form name="LoginForm" action="promoaccount.asp" method="post">
					<table width="360" cellpadding="0" cellspacing="0">
						<tr bgcolor="#0099CC">
							<td class="WhiteText" width="8" valign="top" align="left">
								<img src="images/pix.gif" width="8" height="1">
							</td>
							<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
								<b><i>Company Login</i></b>
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
