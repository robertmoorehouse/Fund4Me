<%@ Language = VBScript%>
<%

Set conn3 = Server.CreateObject("ADODB.Connection")
conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set Lookup = Server.CreateObject("ADODB.recordset")
			
SQL = "SELECT category.* FROM category order by CategoryID"

Lookup.open SQL, conn3 , 1, 3
newcatstr = ""
a = 1
do while not lookup.EOF

if Lookup("Master") = "yes" then
	holdmaster = Lookup("categoryID")
end if
	tempname = Lookup("categoryID")
	temp = Request.Form("" & tempname)
	
	if temp = "on" then
		if right(newcatstr,1) = "," then
		else
		newcatstr = newcatstr & ","
		end if
		if trim(cstr(holdmaster)) = trim(Lookup("Master")) then
			newcatstr = newcatstr & Lookup("Master") & "," & Lookup("categoryID")
			holdmaster = "no"
		else
			if trim(cstr(holdmaster)) = trim(tempname) then
			else
				newcatstr = newcatstr & Lookup("categoryID")
			end if
		end if
	a = a + 1
	end if
	
	lookup.MoveNext
Loop
Lookup.close	
newcatstr = newcatstr & ","


SQL = "SELECT country.* FROM country order by country"

Lookup.open SQL, conn3 , 1, 3
newcountrystr = ""
a = 1
do while not lookup.EOF

	tempname = Lookup("country")
	temp = Request.Form("" & tempname)
	
	if temp = "on" then
		if right(newcountrystr,1) = "," then
		else
		newcountrystr = newcountrystr & ","
		end if
		newcountrystr = newcountrystr & Lookup("country")
		a = a + 1
	end if
	
	lookup.MoveNext
Loop
Lookup.close	
newcountrystr = newcountrystr & ","



if session("memID") = "29" then

wform = request.QueryString("add")
doform = request.QueryString("sub")

set Lookup2 = Server.CreateObject("ADODB.recordset")

if doform = "yes" then
'save form
	if wform = "add" then
	
		SQL = "SELECT * from stores order by id desc;"
		Lookup2.open SQL, conn3, 1, 3
		
		lookup2.movefirst
		id = cint(lookup2("id")) + 1
		lookup2.close
	else
	
		id = request.querystring("id")
	
	end if
	
	login = request.form("login")
	password = request.form("password")
	email = request.form("email")
	banner1 = request.form("banner")
	notes = request.form("notes")
	oldpassword = request.form("oldpassword")
	merchantURL = request.form("merchantURL")
	CommissionTerm = request.form("CommissionTerm")
	address = request.form("address")
	city = request.form("city")
	state = request.form("state")
	zip = request.form("zip")
	country = request.form("country")
	phone = request.form("phone")
	paymenttypes = request.form("paymenttypes")
	shippingtypes = request.form("ShippingTypes")
	SupplierID = request.form("SupplierID")
	Manufacturer = request.form("Manufacturer")
	amerchant = Request.Form("merchant")
	logo = Request.Form("logo")
	shortdescription = request.Form("shortdescription")
	fullmembertext = request.Form("fullmembertext")
	freemembertext = request.Form("freemembertext")
	
	
	ucountry = newcountrystr
if Manufacturer = "Coupons Only" or Manufacturer = "None" then
	ProductName = "<span id=""merchantlink"" title=""Find coupons and discounts for " & amerchant & " with funds4me""><a href=""" & amerchant & """>"& amerchant &"</a></span>"
else
	ProductName = "<span id=""merchantlink"" title=""" & Manufacturer & " Cashback from " & amerchant & """><a href=""" & amerchant & """>"& amerchant &"</a></span>"
end if
	banner = Replace(banner1,chr(34),chr(39),1, -1, 1)
	logo = Replace(logo,chr(34),chr(39),1, -1, 1)
	keyword = Request.Form("keyword")
	keywordlist = Request.Form("keywordlist")

	charityrebatesval = Request.Form("charityrebatesval")
	charityrebate = Request.Form("charityrebate")

	'
	' Make sure blank fields are not null
	'
	If Len(Trim(ucountry)) = 0 Then ListCountry = "all"
	If Len(Trim(charityrebate)) = 0 Then charityrebatesval = "yes"
	If Len(Trim(charityrebatesval)) = 0 Then charityrebatesval = "0"
	If Len(Trim(login)) < 3 Then login = "ID Required"
	If Len(Trim(password)) < 3 Then password = "abcdefg"
	If Len(Trim(email)) = 0 Then email = " "
	If Len(Trim(banner)) = 0 Then banner = " "
	If Len(Trim(notes)) = 0 Then notes = " "
	If Len(Trim(merchantURL)) = 0 Then merchantURL = " "
	If Len(Trim(CommissionTerm)) = 0 Then CommissionTerm = " "
	If Len(Trim(address)) = 0 Then address = "Address not on file"
	If Len(Trim(city)) = 0 Then city = " "
	If Len(Trim(state)) = 0 Then state = " "
	If Len(Trim(zip)) = 0 Then zip = " "
	If Len(Trim(country)) = 0 Then country = " "
	If Len(Trim(phone)) = 0 Then phone = " "
	If Len(Trim(paymenttypes)) = 0 Then paymenttypes = "Details not on file"
	If Len(Trim(ShippingTypes)) = 0 Then ShippingTypes = "Details not on file"
	If Len(Trim(coupons1)) = 0 Then coupons1 = " "
	If Len(Trim(SupplierID)) = 0 Then SupplierID = " "
	If Len(Trim(Manufacturer)) = 0 Then Manufacturer = " "
	If Len(Trim(keyword)) = 0 Then keyword = amerchant
	If Len(Trim(keywordlist)) = 0 Then keywordlist = amerchant & " " & Manufacturer & " Cash Back "
	If Len(Trim(logo)) = 0 Then logo = ""
	If Len(Trim(freemembertext)) = 0 Then freemembertext = ""
	If Len(Trim(fullmembertext)) = 0 Then fullmembertext = ""
	If Len(Trim(shortdescription)) = 0 Then shortdescription = "Click to see full description, coupon codes and discounts for " & amerchant
	'
	' Insert record into customers
	'
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
	set oRs = Server.CreateObject("ADODB.recordset")
			
	sSQL = "SELECT ProductName,Notes,ID,Merchant,shoppin,Link,banner,login,password,merchantURL,email,address,city,state,zip,country,phone,CommissionTerm,paymenttypes,ShippingTypes,ListCountry,Charityrebatesval,charityrebate,category,keyword,keywordlist,logo,shortdescription,freemembertext,fullmembertext FROM stores WHERE (ID = " & ID & ")"
	
	oRs.open sSQL, conn, 1,3
	if oRs.BOF and oRs.EOF then
		oRs.AddNew
		'oRs("ID") = ID
	else
	oRs.MoveFirst
	end if
	
	oRs("Merchant") = aMerchant
	oRs("ProductName") = ProductName
	oRs("Notes") = Notes
	oRs("shoppin") = Manufacturer
	oRs("Link") = SupplierID
	oRs("banner") = banner
	'oRs("login") = login
	'oRs("password") = password
	oRs("merchantURL") = merchantURL
	'oRs("email") = email
	'oRs("address") = address
	'oRs("city") = city
	'oRs("state") = state
	'oRs("zip") = zip
	'oRs("country") = country
	'oRs("phone") = phone
	oRs("CommissionTerm") = CommissionTerm
	'oRs("paymenttypes") = paymenttypes
	'oRs("ShippingTypes") = ShippingTypes
	oRs("ListCountry") = ucountry
	oRs("Charityrebatesval") = charityrebatesval
	oRs("charityrebate") = charityrebate
	oRs("category") =newcatstr
	oRs("keyword") =keyword
	oRs("keywordlist") =keywordlist
	oRs("logo") =logo
	oRs("shortdescription") =shortdescription
	oRs("freemembertext") =freemembertext
	oRs("fullmembertext") =fullmembertext
	oRs.Update
	
  	oRs.close


end if

'add new form

'response.End
'delete form
Response.Redirect("storeadmin.asp?id=" & id - 1 & "&add=")
else
Response.Redirect("default.asp")
end if
%>
