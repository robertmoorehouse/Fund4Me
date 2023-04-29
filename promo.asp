<%@ Language = VBScript%>
<!--#INCLUDE FILE="connection.inc"-->
<%
	iMainMenu = 0
	iSchoolMenu = 0 
	


%>
<html>
<HEAD>
	<title>CashBack Shopping - Funds4me (UK)</title>
	
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<META NAME="description" CONTENT="Join the internets oldest Cashback shopping site. Funds4me cash back was Britains first cashback site and rated as one of the top payers."><META NAME="keywords" CONTENT="cash back, cashback, shopping">
	<meta name="author" content="Inka Maier, TiF Pure Solutions Limited, + 44 (0) 1482 611021">
	
	<meta name="rating" content="General">
	<link rel="stylesheet" type="text/css" href="styles.css">

<script language="JavaScript"><!--
function verify() {
		submit = true;
		
		if (document.regiForm.firstName.value == "") {
			alert("You must enter a First Name");
			document.regiForm.firstName.focus();
			submit = false;
		} 
		if (document.regiForm.lastName.value == "") {
			alert("You must enter a Last Name");
			document.regiForm.lastName.focus();
			submit = false;
		}	
		if (document.regiForm.address1.value == "") {
			alert("You must enter a Address");
			document.regiForm.address1.focus();
			submit = false;
		}	
		if (document.regiForm.city.value == "") {
			alert("You must enter a City");
			document.regiForm.city.focus();
			submit = false;
		}
		if (document.regiForm.county.value == "") {
			alert("You must enter a County");
			document.regiForm.county.focus();
			submit = false;
		}		
		if (document.regiForm.postcode.value == "") {
			alert("You must enter a Postcode");
			document.regiForm.postcode.focus();
			submit = false;
		}
		if (document.regiForm.email.value == "") {
			alert("You must enter an Email Address");
			document.regiForm.email.focus();
			submit = false;
		}		
		emailaddress = document.regiForm.email.value;
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
		
		if (submit == true) {
		document.regiForm.submit();
		}	
	
		
	}
-->
</script>
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<!--#INCLUDE FILE="header.inc"-->
<%
SchoolID = id
PromoID = Request.QueryString("promoid")

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	
set oRs = Server.CreateObject("ADODB.recordset")
	
sSQL = "SELECT * FROM promo WHERE ID = " & PromoID
oRs.open sSQL, Conn, 1,3

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

oRs.close

if SchoolID > "" then 
else          
	SchoolID = SchoolOfDayID
end if %>

<table width="800" cellspacing="0" cellpadding="0">
	<tr>
		<td>
			<img src="images/pix.gif" width="517" height="10">
		</td>
		<td bgcolor="#99CCFF">
			<img src="images/pix.gif" width="142" height="10">
		</td>
		<td bgcolor="#CDE4F6">
			<img src="images/pix.gif" width="141" height="10">
		</td>
	</tr>
	<tr>
		<td valign="top" align="left" valign="top">
			<table width="517" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<img src="images/pix.gif" width="20" height="350">
					</td>
					<td align="center" valign="top">
<!-- Main Text area -->
	<table border="0" cellpadding="0" cellspacing="0" width="490" valign="top">
      <tr>
        <td width="490" class="standart" valign="top" align="left">
          	<%=promotext%>
        </td>
      </tr>
	</table>
	<form name="regiForm" action="SendPromoDetails.asp?sID=<%=SchoolID%>&amp;pID=<%=PromoID%>" method="post" onSubmit="return verify(this);">

		<table width="480" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ff9900">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="8" height="1">
				</td>
				<td colspan="2" class="WhiteText" width="352" valign="top" align="left">
					<b><i>Contact Details for <%=PromoName%></i></b>
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
					<b>First Name: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="firstName" size="30">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Last Name: <font color="#FFFFFF" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="lastName" size="30">
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
					<input type="text" name="address1" size="30">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<input type="text" name="address2" size="30">
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
					<input type="text" name="city" size="30">
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
					<input type="text" name="county" size="30">
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
					<input type="text" name="postcode" size="15">
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
					<input type="text" name="email" size="30" >
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
					<input type="text" name="phone" size="30">
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
					<input type="text" name="fax" size="30" >
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<b>Comments:</b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<textarea cols="30" rows="5" name="comments" ></textarea>
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="400" height="10">
				</td>
			</tr>
			
			<tr bgcolor="#ffcc99">
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="400" height="10">
				</td>
			</tr>
			<tr bgcolor="#ffffff">
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="400" height="10">
				</td>
			</tr>
			<tr bgcolor="#ffffff">
				<td class="WhiteText" width="15" valign="middle" align="left">
					<img src="images/pix.gif" width="15" height="20">
				</td>
				<td class="LargeBlack" width="150" valign="top" align="left">
					<a href="javascript:history.back(1);"><img src="images/buttons/back.gif" border="0"></a>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="right">
					<a href="javascript:verify();"><img src="images/buttons/finish.gif" border="0"></a>
				</td>
			</tr>
			
		</table>
		<br><br>	
	</form> 			

						

<!-- End of Main Text -->
					</td>
				</tr>
				
			</table>

		</td>
		<td bgcolor="#99CCFF" valign="top" align="center">
			<%  ' Supporters/Retailers/Specials
			
				if iRightMenu = 1 then
			%>	  <!--#INCLUDE FILE="coupons.inc"-->
			<%	end if
				if iRightMenu = 2 then
			%>	  <!--#INCLUDE FILE="retailer.inc"-->
			<%	end if
				if iRightMenu = 3 then
			%>	  <!--#INCLUDE FILE="specials.inc"-->
			<%	end if			
			%>
		</td>
		<td bgcolor="#CDE4F6" valign="top" align="center">
			<!--#INCLUDE FILE="adverts.inc"-->
		</td>
	</tr>
	<tr>		
		<td colspan="4">
			<!--#INCLUDE FILE="footer.inc"-->
		</td>
	</tr>
</table>	
</body>
</html>