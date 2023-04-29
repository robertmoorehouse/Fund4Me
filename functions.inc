<%
Function last10stores()

	set boRecordset = Server.CreateObject("ADODB.recordset")
	boRecordset.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
	boRecordset.CursorType = 1
	
	set Lookup5 = Server.CreateObject("ADODB.recordset")
	Lookup5.CursorLocation = 3 
	Lookup5.CursorType = 1				
	
	
	'if memberID = "NotLoggedIn" then
		SQL = "SELECT top 1 id FROM clicks order by ID desc;"
		Lookup5.open SQL, objConn1
		id = Lookup5("id")
		id = cLng(id) - 20
		Lookup5.close
		
		SQL = "SELECT distinct top 10 store FROM clicks where id > " & id & ";"
		
	'else
	'	SQL = "SELECT top 1 id FROM clicks order by ID desc;"
	'	Lookup5.open SQL, objConn1
	'	id = Lookup5("id")
	'	id = cLng(id) - 500
	'	Lookup5.close
	'	SQL = "SELECT distinct top 10 store FROM clicks WHERE (member like 'mem" & memberID & "%' and id > " & id & ");"
		
	'end if
	
	boRecordset.open SQL, objConn1, 1, 3
	
	
	
	
	do while not boRecordset.bof and not boRecordset.eof
		SQL = "SELECT * FROM stores where ((stores.merchant = '" &  boRecordset("store") & "')); "
		Lookup5.open SQL, objConn3	
		if Lookup5.eof and Lookup5.bof then
		
		else	
			link = replace(Lookup5("link"),"CookieUserID", CookieUserID,1,-1,1)
			merchant = Lookup5("Merchant")
			temp = trim(merchant) + " " + trim(Lookup5("shoppin"))	
			cashback =  trim(Lookup5("shoppin"))
			
			response.write("<li><a href=""gocashbackshopping/" & merchant & """ target=""_blank"" title=""Shop at " & merchant & " now and earn up to " & cashback & " cash back.""><img src=""siteimages/spacer.gif"" alt=""" & temp & """ border=""0""/>" & temp & "</a></li>" & chr(13))
			
		end if
		Lookup5.close
		boRecordset.movenext
	loop
	boRecordset.Close

end Function	




Function newstores()

	set Lookup5 = Server.CreateObject("ADODB.recordset")
	Lookup5.CursorLocation = 3 
	Lookup5.CursorType = 1				
	
	SQL = "SELECT top 10 * FROM stores where ((stores.ListCountry like '%,uk,%') and (shoppin <> 'Coupons Only')) order by ID Desc; "
	Lookup5.open SQL, objConn3	
	%>
	<% 
	do while not Lookup5.EOF
		link = replace(Lookup5("link"),"CookieUserID", CookieUserID,1,-1,1)
		merchant = Lookup5("Merchant")
		temp = trim(merchant) + " " + trim(Lookup5("shoppin"))			
		cashback =  trim(Lookup5("shoppin"))
		if len(temp) < 30 then
			'response.write("<li><a href=""" & link & """ target=""_blank"" title=""Shop at " & merchant & " now and earn up to " & cashback & " cash back.""><img src=""siteimages/spacer.gif"" alt=""" & temp & """ border=""0""/>" & temp & "</a></li>" & chr(13))
			
			response.write("<li><a href=""gocashbackshopping/" & merchant & """ target=""_blank"" title=""Shop at " & merchant & " now and earn up to " & cashback & " cash back.""><img src=""siteimages/spacer.gif"" alt=""" & temp & """ border=""0""/>" & temp & "</a></li>" & chr(13))
		end if
		Lookup5.MoveNext
	loop
	Lookup5.Close
	
end Function




Function Categories()

	set Lookup2 = Server.CreateObject("ADODB.recordset")
	Lookup2.CursorLocation = 3 
	Lookup2.CursorType = 1
	
	usecountry = Request.QueryString("country")
	if (usecountry = "") then 
	usecountry = "like '%,uk,%' OR ListCountry = ',all,'"
	usecountrytext = ""
	else
	if usecountry = "all" then
	usecountry = "> ''"
	else
	usecountry = "like '%," & Request.QueryString("country") & ",%' OR ListCountry = ',all,'"
	end if
	usecountrytext = " (from " & Request.QueryString("country") & ")"
	end if
	
	SQL = "SELECT category.* FROM category where Master = 'yes' order by categoryname"
	Lookup2.open SQL, objConn3, 1,3
	
	set Lookup3 = Server.CreateObject("ADODB.recordset")
	Lookup3.CursorLocation = 3 
	Lookup3.CursorType = 1
	
	a=0
	b=1
	
	do while not lookup2.EOF
	
		catval = ""
	
		SQL = "SELECT Count(Merchant) AS CountOfMerchant FROM stores where (category like '%," & lookup2("CategoryID") & ",%' and ((ListCountry " & usecountry & ")))"
		
		Lookup3.open SQL,objConn3 , 0, 1
			
			if Lookup3.bof and Lookup3.eof then
				temp = 0
			else
				temp = Lookup3("CountOfMerchant")
			end if
			
			response.write("<li><a href=""stores.asp?catID=" & lookup2("CategoryID") & "&catname=" & lookup2("CategoryName") & "&country=uk"" title=""Click to view the " & temp & " stores we have listed under the " & lookup2("CategoryName") & " category.""><img src=""siteimages/spacer.gif"" alt=""" & lookup2("CategoryName") & """ />" & lookup2("CategoryName") & "&nbsp;(" & temp & ")</a></li>" & chr(13) )
			
		lookup3.close
		
		lookup2.MoveNext
	loop				

end Function



function currentpage()
	url = Request.ServerVariables("url")
	response.write(url)
end function



Function totalearned()

	set oRsRebates = Server.CreateObject("ADODB.recordset")
	
	if session("memID") = "29" then
		sSQLRebates = "SELECT sum(fundsraised) as Com FROM fundsraised"
	else
		sSQLRebates = "SELECT sum(fundsraised) as Com FROM fundsraised WHERE memID = '" & Session("memID") & "'"
	end if
	oRsRebates.open sSQLRebates, objConn1, 1,3
	total = oRsRebates("com")
	totalearned = Num2Pounds(total)
	oRsRebates.close
	set oRsRebates=nothing

end function




function storesearch()

	set Lookup5 = Server.CreateObject("ADODB.recordset")
	Lookup5.CursorLocation = 3 
	Lookup5.CursorType = 1
		
	SQL = "SELECT id, merchant, ListCountry FROM stores order by merchant "
	Lookup5.open SQL, objConn3
		 
			do while not Lookup5.EOF
	
			
		response.write("<option value=""gocashbackshopping/" & trim(Lookup5("Merchant")) & """>" & trim(Lookup5("Merchant")) & "</option>")
			
				Lookup5.MoveNext
			loop
		
	Lookup5.Close

end function
%>	


<script language="JavaScript"><!--
function verify() {
		submit = true;
		
		if (document.regiForm.firstName.value == "") {
			alert("You must enter a First Name");
			document.regiForm.firstName.focus();
			submit = false;
		} else {
			if (document.regiForm.lastName.value == "") {
				alert("You must enter a Last Name");
				document.regiForm.lastName.focus();
				submit = false;
			}	else {
				if (document.regiForm.address1.value == "") {
					alert("You must enter a Address");
					document.regiForm.address1.focus();
					submit = false;
				}		else {
					if (document.regiForm.city.value == "") {
						alert("You must enter a City");
						document.regiForm.city.focus();
						submit = false;
					}	else {
						if (document.regiForm.county.value == "") {
							alert("You must enter a County");
							document.regiForm.county.focus();
							submit = false;
						}		else {	
							if (document.regiForm.postcode.value == "") {
								alert("You must enter a Postcode");
								document.regiForm.postcode.focus();
								submit = false;
							}	else {
								if (document.regiForm.email.value == "") {
									alert("You must enter an Email Address");
									document.regiForm.email.focus();
									submit = false;
								}	else {	
									if (document.regiForm.username.value == "") {
										alert("You must enter a User Name");
										document.regiForm.username.focus();
										submit = false;
									}	else {	
										if (document.regiForm.password.value == "") {
											alert("You must enter a Password");
											document.regiForm.password.focus();
											submit = false;
										}	else {	
											if (document.regiForm.confPassword.value != document.regiForm.password.value) {
												alert("The password confirmation is not the same as the password!");
												document.regiForm.password.value="";
												document.regiForm.confPassword.value="";
												document.regiForm.password.focus();
												submit = false;
											}	else {	
												pw = document.regiForm.password.value;
												if (pw.length < 6) {
													alert("Your password must have at least 6 characters.");
													document.regiForm.password.value="";
													document.regiForm.confPassword.value="";
													document.regiForm.password.focus();
													submit = false;
												}	else {
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
		
		
		}	}	}	}	}	}	}	}	}	}	}
		
		
		if (submit == true) {
		document.regiForm.submit();
		}	
	
		
	}
-->
			</script>  
			
			
			
			
<script type="text/javascript">
<!--//
function gotostore(theSelect)
{
var mySelect=document.getElementById(theSelect);
theIndx=mySelect.selectedIndex;

if (mySelect.options[theIndx].value != "none")
{
if (mySelect.options[theIndx].value.length > 5)
{
var a = "http://www.funds4me.co.uk/" + mySelect.options[theIndx].value;
}
else
{
var a = "http://www.funds4me.co.uk/shoprequired.asp?shop=" + mySelect.options[theIndx].value;
}

document.location.href = a
}
}

//-->

</script>
