
<script Language="JavaScript"><!--
	function pwreminder(){
		sEmail = document.LoginForm.username.value
		document.location.href="PWreminder.asp?email=" + sEmail	
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
			if (document.updateForm.lastName.value == "") {
				alert("You must enter a Last Name");
				document.updateForm.lastName.focus();
				submit = false;
			}	else {
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
		
		
		}	}	}	}	}	}	}	}	}	}	}
		
		
		if (submit == true) {
		document.updateForm.submit();
		}	
	
		
	}
	
	function fillValues() {
		document.updateForm.firstName.value = "<%=firstName%>";
		document.updateForm.lastName.value = "<%=lastName%>";
		document.updateForm.address1.value = "<%=address1%>";
		document.updateForm.address2.value = "<%=address2%>";
		document.updateForm.city.value = "<%=city%>";
		document.updateForm.county.value = "<%=county%>";
		document.updateForm.country.value = "<%=country%>";
		document.updateForm.postcode.value = "<%=postcode%>";
		document.updateForm.email.value = "<%=email%>";
		document.updateForm.phone.value = "<%=phone%>";
		document.updateForm.relation.value = "<%=relation%>";
		document.updateForm.username.value = "<%=username%>";
		document.updateForm.password.value = "<%=password%>";
		document.updateForm.confPassword.value = "<%=password%>";
		if ("<%=newsletter%>" == "yes") {
			document.updateForm.newsletter.checked = true;
		} else {
			document.updateForm.newsletter.checked = false;
		}
	}
-->
</script>
<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">
	<table border="0" cellpadding="0" cellspacing="0" width="498" valign="top">
      <tr>
        <td width="498" class="standart" valign="top" align="center">
			<table width="490" cellpadding="0" cellspacing="0">
			<tr bgcolor="#ffffff">
				<td class="WhiteText" width="90" valign="top" align="left"><!--<a href="AccSchools.asp"><img border="0" src="images/buttons/accMenu1.gif" name="men1" OnMouseOver="javascript:RollOver('men1',1);" OnMouseOut="javascript:RollOut('men1',1);" width="90"></a><a href="AccProfile.asp"><img border="0" src="images/buttons/accMenu2.gif" width="90" name="men2" OnMouseOver="javascript:RollOver('men2',2);" OnMouseOut="javascript:RollOut('men2',2);"></a><a href="AccReport.asp"><img border="0" src="images/buttons/accMenu3.gif" width="90" name="men3" OnMouseOver="javascript:RollOver('men3',3);" OnMouseOut="javascript:RollOut('men3',3);"></a>-->
</td>
				<td colspan="2" class="BlueText" width="408" valign="top" align="center">
					<img src="images/pix.gif" width="408" height="10"><br>
					<b>Welcome <%=Session("memName")%></b><br><br><br>
					<font size="2"><b><i>My Profile</i></b></font>
				</td>
			</tr>
			</table>
			</table>
	<form name="updateForm" action="saveDetails.asp?memID=<%=Session("memID")%>" method="post" onSubmit="return verify(this);">

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
					<b>State / County: <font color="#FFFFFF" size="2"><b>*</b></font></b>
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
					<b>Country: <font color="#ffffff" size="2"><b>*</b></font></b>
				</td>
				<td class="LargeBlack" width="310" valign="top" align="left">
					<select name="country" size="1">
					  <option value="USA" selected>United States of America</option>
					  <option value="Afghanistan">Afghanistan</option>
					  <option value="Albania">Albania</option>
					  <option value="Algeria">Algeria</option>
					  <option value="Angola">Angola</option>
					  <option value="Andorra">Andorra</option>
					  <option value="Anguilla">Anguilla</option>
					  <option value="Antigua">Antigua and Barbuda</option>
					  <option value="Argentina">Argentina</option>
					  <option value="Armenia">Armenia</option>
					  <option value="Aruba">Aruba</option>
					  <option value="Ascension Ils">Ascension Island</option>
					  <option value="Australia">Australia</option>
					  <option value="Austria">Austria</option>
					  <option value="Azerbaijan">Azerbaijan</option>
					  <option value="Azores">Azores</option>
					  <option value="Bahamas">Bahamas</option>
					  <option value="Bahrain">Bahrain</option>
					  <option value="Balearic Isles">Baleric Isles</option>
					  <option value="Bangladesh">Bangladesh</option>
					  <option value="Barbados">Barbados</option>
					  <option value="Belarus">Belarus</option>
					  <option value="Belgium">Belgium</option>
					  <option value="Belize">Belize</option>
					  <option value="Benin">Benin</option>
					  <option value="Bermuda">Bermuda</option>
					  <option value="Bhutan">Bhutan</option>
					  <option value="Bolivia">Bolivia</option>
					  <option value="Bosnia">Bosnia-Herzegovinia</option>
					  <option value="Botswana">Botswana</option>
					  <option value="Brazil">Brazil</option>
					  <option value="British Virgin">British Virgin Islands</option>
					  <option value="Brunei ">Brunei Darussalam</option>
					  <option value="Bulgaria">Bulgaria</option>
					  <option value="Burkina Faso">Burkina Faso</option>
					  <option value="Burundi">Burundi</option>
					  <option value="Cambodia">Cambodia</option>
					  <option value="Cameroon">Cameroon</option>
					  <option value="Canada">Canada</option>
					  <option value="Canary Islands">Canary Islands</option>
					  <option value="Cape Verde">Cape Verde</option>
					  <option value="Cayman Islands">Cayman Islands</option>
					  <option value="Central Africa">Central African Republic</option>
					  <option value="Chad">Chad</option>
					  <option value="Chille">Chille</option>
					  <option value="China">China</option>
					  <option value="Christmas Isla">Christmas Island</option>
					  <option value="Cocos Islands">Cocos Islands</option>
					  <option value="Columbia">Columbia</option>
					  <option value="Comoros">Comoros</option>
					  <option value="Congo">Congo</option>
					  <option value="Corsica">Corsica</option>
					  <option value="Costa Rica">Costa Rica</option>
					  <option value="Croatia">Croatia</option>
					  <option value="Cuba">Cuba</option>
					  <option value="Cuba Guantanam">Cuba Guantanamo Bay</option>
					  <option value="Cyprus">Cyprus</option>
					  <option value="Czech Republic">Czech Republic</option>
					  <option value="Denmark">Denmark</option>
					  <option value="Djibouti">Djibouti</option>
					  <option value="Dominica">Dominica</option>
					  <option value="Dominican Repu">Dominican Republic</option>
					  <option value="East Timor">East Timor</option>
					  <option value="Ecuador">Ecuador</option>
					  <option value="Egypt">Egypt</option>
					  <option value="El Salvador">El Salvador</option>
					  <option value="Equatorial Gui">Equatorial Guinea</option>
					  <option value="Eritrea">Eritrea</option>
					  <option value="Estonia">Estonia</option>
					  <option value="Ethiopia">Ethiopia</option>
					  <option value="Falkland Islan">Falkland Island</option>
					  <option value="Faroe Islands">Faroe Islands</option>
					  <option value="Fiji">Fiji</option>
					  <option value="Finland">Finland</option>
					  <option value="France">France</option>
					  <option value="French Guiana">French Guiana</option>
					  <option value="French Polynes">French Polynesia</option>
					  <option value="Gabon">Gabon</option>
					  <option value="The Gambia">The Gambia</option>
					  <option value="Gaza &amp; Kha">Gaza &amp; Khan Yunis</option>
					  <option value="Georgia">Georgia</option>
					  <option value="Germany">Germany</option>
					  <option value="Ghana">Ghana</option>
					  <option value="Gibraltar">Gibraltar</option>
					  <option value="Greece">Greece</option>
					  <option value="Greenland">Greenland</option>
					  <option value="Grenada">Grenada</option>
					  <option value="Guatemalia">Guatemalia</option>
					  <option value="Guernsey">Guernsey</option>
					  <option value="Guinea">Guinea</option>
					  <option value="Guinea-Bissau">Guinea-Bissau</option>
					  <option value="Guyana">Guyana</option>
					  <option value="Haiti">Haiti</option>
					  <option value="Honduras">Honduras</option>
					  <option value="Hong Kong">Hong Kong</option>
					  <option value="Hungary">Hungary</option>
					  <option value="Iceland">Iceland</option>
					  <option value="India">India</option>
					  <option value="Indonesia">Indonesia</option>
					  <option value="Iran">Iran</option>
					  <option value="Iraq">Iraq</option>
					  <option value="Ireland">Ireland</option>
					  <option value="Israel">Israel</option>
					  <option value="Italy">Italy</option>
					  <option value="Ivory Coast">Ivory Coast</option>
					  <option value="Jamaica">Jamaica</option>
					  <option value="Japan">Japan</option>
					  <option value="Jersey">Jersey</option>
					  <option value="Jordan">Jordan</option>
					  <option value="Kazakhstan">Kazakhstan</option>
					  <option value="Kenya">Kenya</option>
					  <option value="Kirghizstan">Kirghizstan</option>
					  <option value="Kiribati">Kiribati</option>
					  <option value="Korea">Korea</option>
					  <option value="Kuwaite">Kuwaite</option>
					  <option value="Laos">Laos</option>
					  <option value="Latvia">Latvia</option>
					  <option value="Lebanon">Lebanon</option>
					  <option value="Lesotho">Lesotho</option>
					  <option value="Liberia">Liberia</option>
					  <option value="Libya">Libya</option>
					  <option value="Luxembourg"></option>
					  <option value="Macao"></option>
					  <option value="Macedonia">Macedonia</option>
					  <option value="Maderia">Maderia</option>
					  <option value="Malawi">Malawi</option>
					  <option value="Malaysia">Malaysia</option>
					  <option value="Maldives">Maldives</option>
					  <option value="Mali">Mali</option>
					  <option value="Malta">Malta</option>
					  <option value="Marshall Islan">Marshall Islands</option>
					  <option value="Martinique">Martinique</option>
					  <option value="Mauritania">Mauritania</option>
					  <option value="Mauritius">Mauritius</option>
					  <option value="Mexico">Mexico</option>
					  <option value="Micronesia">Micronesia</option>
					  <option value="Moldova">Moldova</option>
					  <option value="Monaco">Monaco</option>
					  <option value="Mongolia">Mongolia</option>
					  <option value="Montserrat">Montserrat</option>
					  <option value="Morocco">Morocco</option>
					  <option value="Mozambique">Mozambique</option>
					  <option value="Nambia">Nambia</option>
					  <option value="Nauru">Nauru</option>
					  <option value="Nepal">Nepal</option>
					  <option value="Netherland"></option>
					  <option value="Netherlands An">Netherlands Antilles</option>
					  <option value="New Caledonia">New Caledonia</option>
					  <option value="New Zealand">New Zealand</option>
					  <option value="Nicaragua">Nicaragua</option>
					  <option value="Niger">Niger</option>
					  <option value="Nigeria">Nigeria</option>
					  <option value="Norfolk Island">Norflok Island</option>
					  <option value="Norway">Norway</option>
					  <option value="Oman">Oman</option>
					  <option value="Pakistan">Pakistan</option>
					  <option value="Panama">Panama</option>
					  <option value="Papua New Guin">Papua New Guinea</option>
					  <option value="Paraguay">Paraguay</option>
					  <option value="Peru">Peru</option>
					  <option value="Philippines">Philipines</option>
					  <option value="Pitcairn Islan">Pitcairn Islands</option>
					  <option value="Poland">Poland</option>
					  <option value="Portugal">Portugal</option>
					  <option value="Puerto Rico">Puerto Rico</option>
					  <option value="Qatar">Qatar</option>
					  <option value="Romania">Romania</option>
					  <option value="Russia">Russia</option>
					  <option value="Rwanda">Rwanda</option>
					  <option value="St Helena">St Helena</option>
					  <option value="St Lucia">St Lucia</option>
					  <option value="Samoa">Samoa</option>
					  <option value="San Marino">San Marino</option>
					  <option value="Sardinia">Sardinia</option>
					  <option value="Saudia Arabia">Saudia Arabia</option>
					  <option value="Senegal">Senegal</option>
					  <option value="Seychelles">Seychelles</option>
					  <option value="Sicily">Sicily</option>
					  <option value="Sierra Leone">Sierra Leone</option>
					  <option value="Singapore">Singapore</option>
					  <option value="Slovakia">Slovakia</option>
					  <option value="Slovenia">Slovenia</option>
					  <option value="Solomon Island">Solomon Islands</option>
					  <option value="Somalia">Somalia</option>
					  <option value="South Africa">South Africa</option>
					  <option value="Spain">Spain</option>
					  <option value="Sri Lanka">Sri Lanka</option>
					  <option value="Sudan">Sudan</option>
					  <option value="Swaziland">Swaziland</option>
					  <option value="Sweden">Sweden</option>
					  <option value="Switzerland">Switzerland</option>
					  <option value="Syria">Syria</option>
					  <option value="Taiwan">Taiwan</option>
					  <option value="Tanzania">Tanzania</option>
					  <option value="Thailand">Thailand</option>
					  <option value="Togo">Togo</option>
					  <option value="Tonga">Tonga</option>
					  <option value="Trinidad and T">Trinidad and Tobago</option>
					  <option value="Tunisia">Tunisia</option>
					  <option value="Turkey">Turkey</option>
					  <option value="Uganda">Uganda</option>
					  <option value="Ukraine">Ukraine</option>
					  <option value="United Arab Em">United Arab Emirates</option>
					  <option value="Uruguay">Uruguay</option>
					  <option value="Vatican City S">Vatican City State</option>
					  <option value="Venzuela">Venzuela</option>
					  <option value="Vietnam">Vietnam</option>
					  <option value="Western Samoa">Western Samoa</option>
					  <option value="Yemen">Yemen</option>
					  <option value="Yugoslavia">Yugoslavia</option>
					  <option value="Zambia">Zambia</option>
					  <option value="Zimbabwe">Zimbabwe</option>
					</select>
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
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="400" height="10">
				</td>
			</tr>
		</table>
		<br><br>
		<table width="440" cellpadding="0" cellspacing="0">
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
					<input type="text" name="username" size="30" disabled>
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
					<input type="password" name="password" size="30">
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
					<input type="password" name="confPassword" size="30">
				</td>
			</tr>
				<tr bgcolor="#ffcc99">
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="400" height="8">
				</td>
			</tr>
			
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td colspan="2" class="LargeBlack" width="460" valign="top" align="left">
					<input type="checkbox" name="newsletter">&nbsp;
					Would you like to get the Funds 4 Me email newsletter?
				</td>
			</tr>
			</tr>
				<tr bgcolor="#ffcc99">
				<td colspan="3" class="WhiteText" width="8" valign="top" align="left">
					<img src="images/pix.gif" width="400" height="8">
				</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="top" align="left">
					<img src="images/pix.gif" width="15" height="15">
				</td>
				<td colspan="2" class="LargeBlack" width="460" valign="top" align="left">
					How did you hear about Funds 4 Me?
					<select name="hearAbout">
						<option value="Flyer" <% if hearAbout="Flyer" then a="selected"%><%=a%>>Flyer/Brochure</option>
						<option value="Search" <% if hearAbout="Search" then a="selected"%><%=a%>>Search Engine</option>
						<option value="SchoolPerson" <% if hearAbout="member" then a="selected"%><%=a%>>Another member</option>
						<option value="Friend" <% if hearAbout="Friend" then a="selected"%><%=a%>>A friend or family member</option>
						<option value="Radio" <% if hearAbout="Radio" then a="selected"%><%=a%>>Radio Add</option>
						<option value="Newspaper" <% if hearAbout="Newspaper" then a="selected"%><%=a%>>Newspaper Add</option>
						<option value="Web" <% if hearAbout="Web" then a="selected"%><%=a%>>On the web</option>
						<option value="email" <% if hearAbout="email" then a="selected"%><%=a%>>An Email</option>
						<option value="Other" <% if hearAbout="Other" then a="selected"%><%=a%>>Other</option>
						<option value="Shoppin.net member" <% if hearAbout="Shoppin.net member" then a="selected"%><%=a%>>Shoppin.net member</option>
					</select>
					</td>
			</tr>
			<tr bgcolor="#ffcc99">
				<td class="WhiteText" width="15" valign="middle" align="left">
					<img src="images/pix.gif" width="15" height="25">
				</td>
				<td colspan="2" class="LargeBlack" width="460" valign="middle" align="left">
					Please be specific to help our marketing team:&nbsp;&nbsp;&nbsp;
					<input type="text" name="relation" size="50">
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
				</td>
				<td class="LargeBlack" width="310" valign="top" align="right">
					<a href="javascript:verify();"><img src="images/buttons/update.gif" border="0"></a>
				</td>
			</tr>
			
		</table>
		
	</form> 			
      
      

						

	
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->