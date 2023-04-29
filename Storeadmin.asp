<%@ Language = VBScript%>
<%
	
if session("memID") = "29" then



wform = request.QueryString("add")
doform = request.QueryString("sub")




if wform = "delete" then
		id = request.querystring("id")
		a = 1
		

		Set conn3 = Server.CreateObject("ADODB.Connection")
		conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
		
		set Lookup = Server.CreateObject("ADODB.recordset")
		Lookup.CursorLocation = 3 
		Lookup.CursorType = 1
						
		do while a = 1

			SQL = "SELECT stores.* "
			SQL = SQL + " FROM stores "
			SQL = SQL + "WHERE (stores.ID = " & ID & ") order by stores.ID;"
			
			Lookup.open SQL, conn3,1,3
			
			if Lookup.BOF and lookup.EOF then
			a = 200
			else
			Lookup.delete
			end if
			Lookup.Close	

		Loop

		response.Redirect("storeadmin.asp?id=" & id)
		
end if



%>
<html>
<HEAD>
	<title>Funds 4 Me: </title>
	
	<link rel="stylesheet" type="text/css" href="styles.css">
</HEAD>
	<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

	
				<table cellpadding="0" cellspacing="0" ID="Table2">
					<tr>
						<td class=standart>
							<img src="images/pix.gif" width="20" height="450">
						</td>
						<td class=standart align="center" valign="top">
	<!-- Main Text area -->
					
<%

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set Lookup2 = Server.CreateObject("ADODB.recordset")

if wform = "add" then
	
		id = 0
		
		a = 1
		do while a = 1

		Set conn3 = Server.CreateObject("ADODB.Connection")
		conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
		
		set Lookup = Server.CreateObject("ADODB.recordset")
		Lookup.CursorLocation = 3 
		Lookup.CursorType = 1
						
		

		SQL = "SELECT stores.* "
		SQL = SQL + " FROM stores "
		SQL = SQL + "WHERE (stores.ID = " & ID & ") order by stores.ID;"
		
		Lookup.open SQL, conn3
		
		if Lookup.BOF and lookup.EOF then
		ID = ID + 1
		Lookup.Close	
		else
		a = 5
		end if
		
		loop
		
		
		Lookup.MoveFirst
		
		set Lookup5 = Server.CreateObject("ADODB.recordset")
		Lookup5.CursorLocation = 3 
		Lookup5.CursorType = 1
						
		

		SQL = "SELECT stores.id, stores.merchant "
		SQL = SQL + " FROM stores order by merchant"
		
		Lookup5.open SQL, conn3
		
%>
<center>

  <table border="0">
    <tr>
      <td class=standart colspan=2>Merchant Admin  - <a href="storeadmin.asp?id=<%=ID+1%>">Next Merchant</a> - <a href="storeadmin.asp?add=add">Add Merchant</a>-          <a href="storeadmin.asp?add=delete&id=<%=ID%>">Delete Merchant</a>&nbsp;&nbsp;
    
      <form method="POST" action="storeadmin.asp" id=form1 name=form1>
            <select size="1" name="nextID" ID="Select1"> 
				<% do while not Lookup5.EOF
				
				%>
				<option value="<%=Lookup5("ID")%>"><%=trim(Lookup5("Merchant"))%></option>
				<%
				Lookup5.MoveNext
				loop
				Lookup5.Close
				%>
			</select>
			<input type="submit" value="GoTo" name="B3"> </form></td>
      
    </tr>
	<form method="POST" action="admin_storecheck.asp" id=form2 name=form2 target="_blank">
	<tr>
		<td class=standart>Check</td>
		<td class=standart><input type="text" name="Merchantcheck" size="20" Value=""><input type="submit" value="Check" name="B4" ></td></tr>
	<tr>
	</form>
    <tr><form method="POST" action="storeadminsave.asp?id=<%=ID%>&sub=yes&add=<%=wform%>"  name="myform" >
		<td class=standart><b>Basic Site Details</b>/</td>
		<td class=standart></td></td>
		</tr>
		
		<tr>
		<td class=standart>CurrentID</td>
		<td class=standart>ADD NEW </td></tr>
		<tr>
		<td class=standart>Merchant</td>
        <td class=standart><input type="text" name="Merchant" size="80" Value=""></td></tr>
		<tr>
		<td class=standart>Site URL</td>
        <td class=standart><input type="text" name="MerchantUrl" size="80" Value="" ></td></tr>
		<tr>
		<td class=standart>Rebate URL</td>
        <td class=standart><input type="text" name="SupplierID" size="120" Value="" ></td></tr>
		<tr> 
		<td class=standart colspan=2>
		<input type="button" onClick="document.myform.SupplierID.value+='&clickref=CookieUserID'" value="AffWin" name="AffWin" title="Aff Window" >
		<input type="button" onClick="document.myform.SupplierID.value+='&epi=CookieUserID'" value="TD" name="TD" title="TD" >
		<input type="button" onClick="document.myform.SupplierID.value+='&nwk=CookieUserID'" value="UKAff" name="UKAff" title="UK Affiliates" >
		<input type="button" onClick="document.myform.SupplierID.value+='&u1=CookieUserID'" value="Linkshare" name="LinkShare" title="Linkshare" >
		<input type="button" onClick="document.myform.SupplierID.value+='&bfinfo=CookieUserID'" value="BeFree" name="BeFree" title="Be Free" >
		<input type="button" onClick="document.myform.SupplierID.value+='&SID=CookieUserID'" value="CJ" name="CJ" title="CJ" >
		
		<!--ls = &u1=CookieUserID&nbsp;&nbsp;bf = ?bfinfo=CookieUserID&<br>CJ = ?SID=CookieUserID&&nbsp;&nbsp;UKAff = &nwk=CookieUserID<br>tradedoub = &epi=CookieUserID&nbsp;&nbsp;AffWin = &clickref=CookieUserID--></td>
		</tr>
		<tr>
		<td class=standart>Rebate (Short)</td>
        <td class=standart><input type="text" name="Manufacturer" size="80" Value="" ></td></tr>
		<tr>
		<td class=standart>Rebate (long)</td>
        <td class=standart><input type="text" name="CommissionTerm" size="80" Value="" ></td></tr>
		<tr>
		<td class=standart>Cashback Text Full Member</td>
        <td class=standart><input type="text" name="fullmembertext" size="80" Value="" ></td></tr>
		<tr>
		<td class=standart>Cashback Text Free Member</td>
        <td class=standart><input type="text" name="freemembertext" size="80" Value="" ></td></tr>
		<tr>
		<td class=standart>Keyword</td>
        <td class=standart><input type="text" name="keyword" size="80" Value="" ></td></tr>
		<tr>
		<td class=standart>Keywords</td>
        <td class=standart><input type="text" name="keywordlist" size="80" Value="" ></td></tr>
        <tr>
		<td class=standart>Merchant Logo no html - url only</td>
        <td class=standart><input type="text" name="logo" size="80" Value="" ><input type="button" onClick="document.myform.logo.value+='/logos/' + document.myform.Merchant.value + '.gif'" value="Add Local" name="Add Local" title="Add Local" ></td></tr>
		<tr>
		<td class=standart>Short description</td>
		<td class=standart><textarea rows="3" name="shortdescription" cols="50" ID="Textarea1"></textarea></td></tr>
		<tr>
		<td class=standart>Landing Page Description</td>
        <td class=standart><textarea rows="5" name="notes" cols="50" ID="Textarea1"></textarea></td></tr>
		<tr>
		<td colspan=2 class=standart>Image URL must be a HTML link no hyperlink max size 250 * 250 
        (e.g &lt;img src='...')"></td></tr>
		<tr>
		<td class=standart></td>
		<td>
          <input type="text" name="banner" size="80" Value="" ID="Text12"></td>
        </tr>
		<tr>
		  <td class=standart valign=top>
          List Country</td>
        <td class=standart> 
          
			<table >
				<tr>
       
							  <%
					set Lookup3 = Server.CreateObject("ADODB.recordset")		  
					SQL = "SELECT country.* FROM country order by country"
							
					Lookup3.open SQL, conn3 , 1, 3
					b=0
					e=1
					do while not lookup3.EOF
						if e=7 then
							%>
							</tr><tr>
							<% 
							e=1
						end if
						b=b+1		          
						%>
						<td class=standart>
						<INPUT id=<%=lookup3("country")%> name=<%=lookup3("country")%> type=checkbox ><%=lookup3("country")%>&nbsp;&nbsp;&nbsp;
						</td>
						<%
						e=e+1
							
							
						lookup3.MoveNext
					loop
					lookup3.close
						
						%>
					</td>
				</tr></table>
					
					
			
			</td></tr>
			
			
			
		<tr>
		<td class=standart>
        Category</td>
        <td class=standart></td>
        </tr>
        
        <tr><td class=standart colspan=2>
        <table >
        <tr>
       
					  <%
					  
			set Lookup2 = Server.CreateObject("ADODB.recordset")
			
			SQL = "SELECT category.* FROM category where Master = 'yes' order by categoryname"
					
			Lookup2.open SQL, conn3 , 1, 3
					
			set Lookup3 = Server.CreateObject("ADODB.recordset")
			
			tempsplit = " "
			'tempsplit = lookup("category")
			
			if isnull(tempsplit) then
			tempsplit=" "
			end if
			if len(tempsplit) < 2 then
			tempsplit=" "
			end if
			iscat = split(tempsplit,",")
			'keycnt = ubound(iscat)
			

			cnt = 1
			For Each word in iscat
				'Response.Write(word & "-" & cnt & "<BR>")
				cnt = cnt + 1
			next 
			'lookup1.MoveFirst
			a=0
			b=1
			do while not lookup2.EOF

				catval = ""
				for c = 0 to ubound(iscat)
					if iscat(c) = trim(cstr(lookup2("CategoryID"))) then
						catval = "checked" 
					else
						if catval = "checked" then
						else
						catval = ""
						end if
					end if
									
					'Response.write(cstr(iscat(c)) & " - " & trim(cstr(lookup1("CategoryID"))) & " - " & cnt & " - " & catval & "<br>")
				next  
					
				
						          
				%>
				<td colspan = 3 class=standart>
				<INPUT  id=<%=lookup2("CategoryID")%>  name=<%=lookup2("CategoryID")%> type=checkbox <%=catval%>><b><%=lookup2("CategoryName")%></b>
				</td></tr><tr>
				<%
				b=b+1
				
				SQL = "SELECT category.* FROM category where Master = '" & Lookup2("categoryID") & "' order by categoryname"
					
				Lookup3.open SQL, conn3 , 0, 1
				
				e=1
				do while not lookup3.EOF
					if e=6 then
						%>
						</tr><tr>
						<% 
						e=1
					end if
					catval = ""
					for c = 0 to ubound(iscat)
						if iscat(c) = trim(cstr(lookup3("CategoryID"))) then
							catval = "checked" 
						else
							if catval = "checked" then
							else
							catval = ""
							end if
						end if
										
						'Response.write(cstr(iscat(c)) & " - " & trim(cstr(lookup1("CategoryID"))) & " - " & cnt & " - " & catval & "<br>")
					next  
						
					
							          
					%>
					<td class=standart>
					<INPUT   id=<%=lookup3("CategoryID")%>   name=<%=lookup3("CategoryID")%> type=checkbox <%=catval%>><%=lookup3("CategoryName")%>&nbsp;&nbsp;&nbsp;
					</td>
					<%
					e=e+1
					
					
					lookup3.MoveNext
				loop
				lookup3.close
				
				%>
				</tr><tr>
				<%
				
				
				
				
				lookup2.MoveNext
			loop
		  
		       
          %>
        </tr></table></td>
        </tr>
		
		<!--          <tr>
		<td class=standart>Email Address</td>
        <td class=standart><input type="text" name="email" size="80" Value="" ID="Text1"></td></tr>
		<tr>
		<td class=standart>Address</td>
        <td class=standart><input type="text" name="Address" size="80" Value="" ID="Text2"></td></tr>
		<tr>
		<td class=standart>City</td>
        <td class=standart><input type="text" name="City" size="80" Value="" ID="Text3"></td></tr>
		<tr>
		<td class=standart>State</td>
        <td class=standart><input type="text" name="State" size="80" Value="" ID="Text4"></td></tr>
		<tr>
		<td class=standart>Zip</td>
        <td class=standart><input type="text" name="Zip" size="80" Value="" ID="Text5"></td></tr>
		<tr>
		<td class=standart>Country</td>
        <td class=standart><input type="text" name="Country" size="80" Value="" ID="Text6"></td></tr>
		<tr>
		<td class=standart>Phone</td>
        <td class=standart><input type="text" name="Phone" size="80" Value="" ID="Text7"></td></tr>
		<tr>
		<td class=standart>Payment Types</td>
        <td class=standart><input type="text" name="PaymentTypes" size="80" Value="" ID="Text8"></td></tr>
		<tr>
		<td class=standart>Shipping Types</td>
        <td class=standart><input type="text" name="ShippingTypes" size="80" Value="" ID="Text9"></td></tr>
		<tr>
		<td class=standart>Your Login ID</td>
        <td class=standart><input type="text" name="login" size="80" Value="" ID="Text10"> </td></tr>
		<tr>
		<td class=standart>Your Password</td>
        <td class=standart><input type="text" name="password" size="80" Value="" ID="Text11"></td></tr>-->
		
         
		 
		 
		<tr><td>
          <input type="hidden" name="oldpassword" size="20" Value=""></td><td>
          <input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></td>
		
        </form>
        <p>&nbsp;</td>
    </tr>
    <tr>
      <td class=standart colspan="2"></td>
    </tr>
  












<%		
		

		
else
	
		id = request.querystring("id")
	
	
	

'add new form


'delete form






	%>
	

<%
'
' Get the order number and password
'

ipad = trim(request.servervariables("REMOTE_ADDR"))
Password = Trim(Request("Password"))
customerID = Request("OrderNumber")
nextid = request.form("nextid")
viewid = request.QueryString("vid")
if viewid = "" then

	if nextid = "" then
		id = request.querystring("id")
		if id = "" then
			id = 1
		else 
			id = id + 1
		end if
	else
		id = nextid
	end if
else
id = viewid
end if

'
' Look up order and get customer information
'
%> 



<%		
		a = 1
		do while a = 1

		Set conn3 = Server.CreateObject("ADODB.Connection")
		conn3.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
		
		set Lookup = Server.CreateObject("ADODB.recordset")
		Lookup.CursorLocation = 3 
		Lookup.CursorType = 1
						
		

		SQL = "SELECT stores.* "
		SQL = SQL + " FROM stores "
		SQL = SQL + "WHERE (stores.ID = " & ID & ") order by stores.ID;"
		
		Lookup.open SQL, conn3
		
		if Lookup.BOF and lookup.EOF then
		ID = ID + 1
		Lookup.Close	
		else
		a = 5
		end if
		
		loop
		
		
		Lookup.MoveFirst
		
		set Lookup5 = Server.CreateObject("ADODB.recordset")
		Lookup5.CursorLocation = 3 
		Lookup5.CursorType = 1
						
		

		SQL = "SELECT stores.id, stores.merchant "
		SQL = SQL + " FROM stores order by merchant"
		
		Lookup5.open SQL, conn3
		
%>
<center>

  <table border="0">
    <tr>
      <td class=standart colspan=2>Merchant Admin  - <a href="storeadmin.asp?id=<%=ID+1%>">Next Merchant</a> - <a href="storeadmin.asp?add=add">Add Merchant</a>-&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="storeadmin.asp?add=delete&id=<%=ID%>">Delete Merchant</a>&nbsp;&nbsp;
    
      <form method="POST" action="storeadmin.asp" id=form1 name=form1>
            <select size="1" name="nextID" ID="Select1"> 
				<% do while not Lookup5.EOF
				
				%>
				<option value="<%=Lookup5("ID")%>"><%=trim(Lookup5("Merchant"))%></option>
				<%
				Lookup5.MoveNext
				loop
				Lookup5.Close
				%>
			</select>
			<input type="submit" value="GoTo" name="B3"> </form></td>
      
    </tr>
    <tr><form method="POST" action="storeadminsave.asp?id=<%=ID%>&sub=yes&add=<%=wform%>" name="myform" >
		<td class=standart><b>Basic Site Details</b>/</td>
		<td class=standart><input type="submit" value="Submit" name="B1"></td></td>
		</tr>
		<tr>
		<td class=standart>CurrentID</td>
		<td class=standart><%=ID%> </td></tr>
		<tr>
		<td class=standart>Merchant</td>
        <td class=standart><input type="text" name="Merchant" size="80" Value="<%= lookup("Merchant") %>"></td></tr>
		<!--<tr>
		<td class=standart>Site URL</td>
        <td class=standart><input type="text" name="MerchantUrl" size="80" Value="<%= lookup("MerchantUrl") %>" ></td></tr>
		<tr>-->
		

		<td class=standart>Rebate URL</td>
        <td class=standart><input type="text" name="SupplierID" size="120" Value="<%= lookup("Link") %>" ></td></tr>
		<tr> 
		<td class=standart colspan=2>
		<input type="button" onClick="document.myform.SupplierID.value+='&clickref=CookieUserID'" value="AffWin" name="AffWin" title="Aff Window" >
		<input type="button" onClick="document.myform.SupplierID.value+='&epi=CookieUserID'" value="TD" name="TD" title="TD" >
		<input type="button" onClick="document.myform.SupplierID.value+='&nwk=CookieUserID'" value="UKAff" name="UKAff" title="UK Affiliates" >
		<input type="button" onClick="document.myform.SupplierID.value+='&u1=CookieUserID'" value="Linkshare" name="LinkShare" title="Linkshare" >
		<input type="button" onClick="document.myform.SupplierID.value+='&bfinfo=CookieUserID'" value="BeFree" name="BeFree" title="Be Free" >
		<input type="button" onClick="document.myform.SupplierID.value+='&SID=CookieUserID'" value="CJ" name="CJ" title="CJ" >
		
		<!--ls = &u1=CookieUserID&nbsp;&nbsp;bf = ?bfinfo=CookieUserID&<br>CJ = ?SID=CookieUserID&&nbsp;&nbsp;UKAff = &nwk=CookieUserID<br>tradedoub = &epi=CookieUserID&nbsp;&nbsp;AffWin = &clickref=CookieUserID--></td>
		</tr>
		<tr>
		<td class=standart>Rebate (Short)</td>
        <td class=standart><input type="text" name="Manufacturer" size="80" Value="<%= lookup("Shoppin") %>" ></td></tr>
		<tr>
		<td class=standart>Rebate (long)</td>
        <td class=standart><input type="text" name="CommissionTerm" size="80" Value="<%= lookup("CommissionTerm") %>" ></td></tr>
		<tr>
		<td class=standart>Cashback Text Full Member</td>
        <td class=standart><input type="text" name="fullmembertext" size="80" Value="<%= lookup("fullmembertext") %>" ></td></tr>
		<tr>
		<td class=standart>Cashback Text Free Member</td>
        <td class=standart><input type="text" name="freemembertext" size="80" Value="<%= lookup("freemembertext") %>" ></td></tr>
		
		<tr>
		<td class=standart>Keyword</td>
        <td class=standart><input type="text" name="keyword" size="80" Value="<%= lookup("keyword") %>" ></td></tr>
		<tr>
		<td class=standart>Keywords</td>
        <td class=standart><input type="text" name="keywordlist" size="80" Value="<%= lookup("keywordlist") %>" ></td></tr>
		<tr>
		<td class=standart>Merchant Logo no html - url only</td>
        <td class=standart><input type="text" name="logo" size="80" Value="<%= lookup("logo") %>" ><input type="button" onClick="document.myform.logo.value+='/logos/' + document.myform.Merchant.value + '.gif'" value="Add Local" name="Add Local" title="Add Local" ></td></tr>
		<tr>
		<td class=standart>Short description</td>
		<td class=standart><textarea rows="3" name="shortdescription" cols="50" ID="Textarea1"><%= lookup("shortdescription") %></textarea></td></tr>
		
		<tr>
		<td class=standart>Landing Page Description</td>
        <td class=standart><textarea rows="5" name="notes" cols="50" ID="Textarea1"><%= lookup("Notes") %></textarea></td></tr>
		<tr>
		<td colspan=2 class=standart>Image URL must be a HTML link no hyperlink max size 250 * 250 
        (e.g &lt;img src='...')"></td></tr>
		<tr>
		<td class=standart></td>
		<td>
          <input type="text" name="banner" size="80" Value="<%= lookup("Banner") %>" ID="Text12">
		</td>
        </tr>	
		<tr>
      	<td class=standart colspan="2">
			<img src="<%=lookup("logo")%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=lookup("Banner")%><br><br>
			<input type="submit" value="Submit" name="B1"><br><br>
	  		<a href="http://www.funds4me.co.uk/shoprequired.asp?shop=<%=lookup("ID")%>" target="_new">test ID link</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  		<a href="http://www.funds4me.co.uk/<%=lookup("merchant")%>" target="_new">test Merchant landing page</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<a href="<%=lookup("link")%>" target="_new">test Merchant link</a>
			
	  	</td>
		</tr>
		<tr>
		  <td class=standart valign=top>
          List Country</td>
        <td class=standart> 
          	
			 <table>
        <tr>
       
					  <%
					  
			set Lookup3 = Server.CreateObject("ADODB.recordset")
			
			tempsplit = " "
			tempsplit = lookup("listcountry")
			
			if isnull(tempsplit) then
			tempsplit=" "
			end if
			if len(tempsplit) < 2 then
			tempsplit=" "
			end if
			iscat = split(tempsplit,",")
			'keycnt = ubound(iscat)
			

			cnt = 1
			For Each word in iscat
				'Response.Write(word & "-" & cnt & "<BR>")
				cnt = cnt + 1
			next 
			'lookup1.MoveFirst
			a=0
			b=1
							
				SQL = "SELECT country.* FROM country order by country"
					
				Lookup3.open SQL, conn3 , 0, 1
				
				e=1
				do while not lookup3.EOF
					if e=7 then
						%>
						</tr><tr>
						<% 
						e=1
					end if
					catval = ""
					for c = 0 to ubound(iscat)
						if iscat(c) = trim(cstr(lookup3("country"))) then
							catval = "checked" 
						else
							if catval = "checked" then
							else
							catval = ""
							end if
						end if
										
						'Response.write(cstr(iscat(c)) & " - " & trim(cstr(lookup1("CategoryID"))) & " - " & cnt & " - " & catval & "<br>")
					next  
						
					
							          
					%>
					<td class=standart>
					<INPUT   id=<%=lookup3("country")%>   name=<%=lookup3("country")%> type=checkbox <%=catval%>><%=lookup3("country")%>&nbsp;&nbsp;&nbsp;
					</td>
					<%
					e=e+1
					
					
					lookup3.MoveNext
				loop
				lookup3.close
				
          %>
        </tr></table>
			
			</td></tr>
		<tr>
		<td class=standart>
        Category</td>
        <td class=standart></td>
        </tr>
        
        <tr><td class=standart colspan=2>
        <table>
        <tr>
       
					  <%
					  
			set Lookup2 = Server.CreateObject("ADODB.recordset")
			
			SQL = "SELECT category.* FROM category where Master = 'yes' order by categoryname"
					
			Lookup2.open SQL, conn3 , 1, 3
					
			set Lookup3 = Server.CreateObject("ADODB.recordset")
			
			tempsplit = " "
			tempsplit = lookup("category")
			
			if isnull(tempsplit) then
			tempsplit=" "
			end if
			if len(tempsplit) < 2 then
			tempsplit=" "
			end if
			iscat = split(tempsplit,",")
			'keycnt = ubound(iscat)
			

			cnt = 1
			For Each word in iscat
				'Response.Write(word & "-" & cnt & "<BR>")
				cnt = cnt + 1
			next 
			'lookup1.MoveFirst
			a=0
			b=1
			do while not lookup2.EOF

				catval = ""
				for c = 0 to ubound(iscat)
					if iscat(c) = trim(cstr(lookup2("CategoryID"))) then
						catval = "checked" 
					else
						if catval = "checked" then
						else
						catval = ""
						end if
					end if
									
					'Response.write(cstr(iscat(c)) & " - " & trim(cstr(lookup1("CategoryID"))) & " - " & cnt & " - " & catval & "<br>")
				next  
					
				
						          
				%>
				<td colspan = 3 class=standart>
				<INPUT id=<%=lookup2("CategoryID")%> name=<%=lookup2("CategoryID")%> type=checkbox <%=catval%>><b><%=lookup2("CategoryName")%></b>
				</td></tr><tr>
				<%
				b=b+1
				
				SQL = "SELECT category.* FROM category where Master = '" & Lookup2("categoryID") & "' order by categoryname"
					
				Lookup3.open SQL, conn3 , 0, 1
				
				e=1
				do while not lookup3.EOF
					if e=6 then
						%>
						</tr><tr>
						<% 
						e=1
					end if
					catval = ""
					for c = 0 to ubound(iscat)
						if iscat(c) = trim(cstr(lookup3("CategoryID"))) then
							catval = "checked" 
						else
							if catval = "checked" then
							else
							catval = ""
							end if
						end if
										
						'Response.write(cstr(iscat(c)) & " - " & trim(cstr(lookup1("CategoryID"))) & " - " & cnt & " - " & catval & "<br>")
					next  
						
					
							          
					%>
					<td class=standart>
					<INPUT  id=<%=lookup3("CategoryID")%>  name=<%=lookup3("CategoryID")%> type=checkbox <%=catval%>><%=lookup3("CategoryName")%>&nbsp;&nbsp;&nbsp;
					</td>
					<%
					e=e+1
					
					
					lookup3.MoveNext
				loop
				lookup3.close
				
				%>
				</tr><tr>
				<%
				
				
				
				
				lookup2.MoveNext
			loop
		  
		       
          %>
        </tr></table></td>
        </tr>
		  <!--        <tr>
		<td class=standart>Email Address</td>
        <td class=standart><input type="text" name="email" size="80" Value="<%= lookup("email") %>" ID="Text1"></td></tr>
		<tr>
		<td class=standart>Address</td>
        <td class=standart><input type="text" name="Address" size="80" Value="<%= lookup("Address") %>" ID="Text2"></td></tr>
		<tr>
		<td class=standart>City</td>
        <td class=standart><input type="text" name="City" size="80" Value="<%= lookup("City") %>" ID="Text3"></td></tr>
		<tr>
		<td class=standart>State</td>
        <td class=standart><input type="text" name="State" size="80" Value="<%= lookup("State") %>" ID="Text4"></td></tr>
		<tr>
		<td class=standart>Zip</td>
        <td class=standart><input type="text" name="Zip" size="80" Value="<%= lookup("Zip") %>" ID="Text5"></td></tr>
		<tr>
		<td class=standart>Country</td>
        <td class=standart><input type="text" name="Country" size="80" Value="<%= lookup("Country") %>" ID="Text6"></td></tr>
		<tr>
		<td class=standart>Phone</td>
        <td class=standart><input type="text" name="Phone" size="80" Value="<%= lookup("Phone") %>" ID="Text7"></td></tr>
		<tr>
		<td class=standart>Payment Types</td>
        <td class=standart><input type="text" name="PaymentTypes" size="80" Value="<%= lookup("PaymentTypes") %>" ID="Text8"></td></tr>
		<tr>
		<td class=standart>Shipping Types</td>
        <td class=standart><input type="text" name="ShippingTypes" size="80" Value="<%= lookup("ShippingTypes") %>" ID="Text9"></td></tr>
		<tr>
		<td class=standart>Your Login ID</td>
        <td class=standart><input type="text" name="login" size="80" Value="<%= lookup("Login") %>" ID="Text10"> </td></tr>
		<tr>
		<td class=standart>Your Password</td>
        <td class=standart><input type="text" name="password" size="80" Value="<%= lookup("password") %>" ID="Text11"></td></tr>
		-->
		<tr><td>
          <input type="hidden" name="oldpassword" size="20" Value=""></td><td>
          <input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></td>
		
        </form>
        <p>&nbsp;</td>
    </tr>
    

  <%end if
%>

</table>

<!-- End of Main Text -->
					</td>
				</tr>
			</table>
<%
else
Response.Redirect("default.asp")
end if
%>		
</body>
</html>
