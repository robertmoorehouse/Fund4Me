<%
temp = request.QueryString("form")
shop = ""
if temp="yes" then
	shop = Request.Form("shopID")
else
	shop = Request("shop")
end if
if shop = "" then
shop = session("shopid")
end if
session("shopid") = ""
if shop = "" then

end if
CookieUserID = ""
if Session("memID") > "" then
	CookieUserID  = "mem" & Session("memID")
else
	CookieUserID  = request.cookies("details")("memberID")
end if

AmazonUK = "dropthegrieffrom"

if CookieUserID = "" then
	CookieUserID = session("customerID")
end if

if CookieUserID = "" then
	register = "Register"
	cuid = "NoneMember" & date()
	CookieUserID = replace(cuid,"/","-",1,-1,1)
else
	register = "Shop"
	name = Request.cookies("details")("UserName")
	showval = "Your purchase will benefit " & name
end if

CookieUserID =  CookieUserID & "F4MECOUK"

ipad = trim(request.servervariables("REMOTE_ADDR"))

sql = "select * from stores where id = " & shop & ";"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open Application("Connection3_ConnectionString"), Application("Connection3_RuntimeUserName"), Application("Connection3_RuntimePassword")
set InsertCursor = Server.CreateObject("ADODB.recordset")
InsertCursor.CursorLocation = 3 
InsertCursor.CursorType = 1
													
InsertCursor.open SQL, conn, 1,3

insertcursor.MoveFirst

store = shop & "(Funds4Me.co.uk)"
failed = shop
cashback = InsertCursor("shoppin")
shop = insertcursor("merchant")
linkit = Replace(InsertCursor("Link"),"CookieUserID",CookieUserID,1, -1, 1)
c = linkit'Replace(linkit,"466441","1292226",1,-1,1)
linkit  = c
id = InsertCursor("ID")
keyword = InsertCursor("keyword")
keywordlist = InsertCursor("keywordlist")
If Len(Trim(keyword)) = 0 Then keyword = "Cash Back, " & InsertCursor("Merchant")
If Len(Trim(keywordlist)) = 0 Then keywordlist = InsertCursor("Merchant") & " " & InsertCursor("shoppin") & " Cash Back."
notes = InsertCursor("Notes")
banner = InsertCursor("banner")




ipad = trim(request.servervariables("REMOTE_ADDR"))
Dim objConn2
Set objConn2 = Server.CreateObject("ADODB.Connection")
objConn2.Open Application("Connection1_ConnectionString"), Application("Connection1_RuntimeUserName"), Application("Connection1_RuntimePassword")
	
set boRecordset = Server.CreateObject("ADODB.recordset")
boRecordset.CursorLocation = 3 'adUseServer = 2, adUseClient = 3
boRecordset.CursorType = 1



	SQL = "SELECT * FROM clicks WHERE (member = '')"
	boRecordset.open SQL, objConn2, 1, 3
	boRecordset.AddNew
	boRecordset("member") = CookieUserID
	boRecordset("store") = shop
	boRecordset("whendate") = now() 
	boRecordset("max") = CashBack
	boRecordset("site") = "Funds4Me.co.uk"
	boRecordset("ipaddress") = ipad

	boRecordset.Update
	boRecordset.Close
	
	SQL = "SELECT * FROM specials WHERE storeID = " & failed
	boRecordset.open SQL, objConn2, 1, 3
	if boRecordset.BOF and boRecordset.EOF then
		specof = "We have none listed, please let us know of any you are aware of any."
	else
		specof = Replace(boRecordset("description"),"CookieUserID",CookieUserID,1, -1, 1)
	end if
	boRecordset.Close
	%>
	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title><%=shop%>&nbsp;<%=CashBack%>&nbsp;Cash Back</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="PRAGMA" content="NO-CACHE">
    <meta content="<%=shop%>,&nbsp;cashback, funds4me" name="keywords">
    <meta content="Get <%=CashBack%>&nbsp;Cash Back on&nbsp;<%=shop%>. Sign up to funds4me. The fastest growing Cash Back shopping site on the net." name="description">
    <link href="se_layoutcss.css" type="text/css" rel="stylesheet">
    <link href="se_presentationcss.css" type="text/css" rel="stylesheet">

    <script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
    </script>

    <meta content="MSHTML 6.00.2900.2963" name="GENERATOR">
</head>
<body>
    <!-- Header -->
    <div align="center">
        <div id="hdr">
            <h1>
                <%=shop%>
                <br>
                Cash Back & Coupons
            </h1>
            <h2 class="h2">
                Earn <%=CashBack%> Cashback on <%=shop%> through Funds4Me.co.uk</h2>
        </div>
    </div>
    <!-- end of header -->
    <div align="center">
	
	<%
if cashback = "None" then
%>
<div id="lh-col">
            <h3 class="h3" align="center">
                <%=register%> today to get CashBack from 100's of shops.</h3>
            <div class="red" align="center">
                <strong><%=shop%> coupons and discounts.</strong></div>
            <div class="text">
                <p>
                    Shop at <%=shop%> using the coupons and discounts codes on the right                    make on <strong><%=shop%></strong>. <br>
                    <br>
                    Remember to join funds4me and get £5 free for opening your account and recieve a £10 cheque for refering new members when they get their first cheque.
                </p>
            </div>
            <div class="signup">
			<p>
               <strong>Coupons / Discounts</strong><br><br>
			   <%=specof%>
			   </p>
            </div>
        </div>
		<!-- end of left column -->
		
		
		<!-- right column -->
        <div id="rh-col">
            <div align="justify">
			
                <h3 class="h3">
                    <a href="<%=linkit%>" title="Shop today to get <%=CashBack%> CashBack on <%=shop%>" target="_blank">About <%=shop%></a></h3>
                <span class="text2"><%=notes%></span><br>
                <br>
                
                <br>
                <br><a href="<%=linkit%>" title="Shop today to get <%=CashBack%> CashBack on <%=shop%>" target="_blank">
                <%=banner%></a>
            </div>
            <div class="footer" align="center">
      <a href="Contact.asp">Contact</a> | <a href="Privacy.asp">Privacy Policy</a> | <a href="Technical.asp">Technical Requirements</a> | <a href="press.asp">News & Press</a> | <a href="useragreement.asp">EULA</a>	
			<br><br>
			&copy; 1998 - 2006 Robert Moorehouse
                </div>
        </div>
		
		</a>
		
		
		<%
else


	'if is a member
	if Session("memID") > "" then
	%>
	
		<div id="lh-col">
            <h3 class="h3" align="center">
                <%=register%> today to get <%=CashBack%> CashBack on <%=shop%></h3>
            <div class="red" align="center">
                <strong>Earn <%=CashBack%> on <%=shop%></strong></div>
            <div class="text">
                <p>
                    <%'=name%> Shop now and as a Funds4Me member you will earn <strong><%=CashBack%></strong> Cash Back on any purchase you
                    make on <strong><%=shop%></strong>. <br>
                    <br>
                    Remember to use any of the coupon codes on the left to get further discounts from this store.
					<br><br>
					Recieve a £10 cheque for refering a new member when they get their first cheque.
                </p>
            </div>
            <div class="signup">
			<p>
               <strong>Coupons / Discounts</strong><br><br>
			   <%=specof%>
			   </p>
            </div>
        </div>
		<!-- end of left column -->
		
		
		<!-- right column -->
        <div id="rh-col">
            <div align="justify">
			
                <h3 class="h3">
                    <a href="<%=linkit%>" title="Shop today to get <%=CashBack%> CashBack on <%=shop%>" target="_blank">About <%=shop%></a></h3>
                <span class="text2"><%=notes%></span><br>
                <br>
                
                <br>
                <br><a href="<%=linkit%>" title="Shop today to get <%=CashBack%> CashBack on <%=shop%>" target="_blank">
                <%=banner%></a>
            </div>
            <div class="footer" align="center">
      <a href="Contact.asp">Contact</a> | <a href="Privacy.asp">Privacy Policy</a> | <a href="Technical.asp">Technical Requirements</a> | <a href="press.asp">News & Press</a> | <a href="useragreement.asp">EULA</a>	
			<br><br>
			&copy; 1998 - 2006 Robert Moorehouse
                </div>
        </div>
		
		</a>
		
	<%
	else
	%>
	<!-- left column -->
        <div id="lh-col">
            <h3 class="h3" align="center">
                Register today to get <%=CashBack%> CashBack on <%=shop%></h3>
            <div class="red" align="center">
                <strong>Earn <%=CashBack%> on <%=shop%></strong></div>
            <div class="text">
                <p>
                    By <a title="signing up" href="http://www.funds4me.co.uk/register2.asp">signing up</a>
                    to Funds4Me.co.uk you will earn <strong><%=CashBack%></strong> Cash Back on any purchase you
                    make on <strong><%=shop%></strong> and we will give you £5.00 for becoming
                    a member. It doesn't stop there! Not only will you earn Cash Back from <%=shop%>
                    but we have over a thousand retailers, including large high street brands for you
                    to choose from and earn even more Cash Back. You will even benefit from any Sales
                    and Discounts they have to offer.<br>
                    <br>
                    <a title="Sign Up now" href="http://www.funds4me.co.uk/register2.asp">Sign Up now</a>
                    and start saving together with thousands of other shoppers using this site. There
                    is No Catch and we dont not sell or give out your details to third parties. <a title="Read more and Sign up here"
                        href="http://www.funds4me.co.uk/register2.asp">Read more and Sign up here</a>
					<br><br>
					<a title="Sign Up now" href="http://www.funds4me.co.uk/register2.asp">Sign Up now</a> and 
					recieve a £10 cheque for refering new members when they get their first cheque, you don't have to have any balance on your own account!
                </p>
            </div>
            <div class="signup">
                <a href="http://www.funds4me.co.uk/register2.asp">
                    <img alt="sign up for <%=shop%> savings. Click Here" src="images/signup.jpg"
                        border="0"></a>
            </div>
        </div>
        <!-- end of left column -->
		
		<!-- right column -->
        <div id="rh-col">
           <div class="text">
               
				<h3 class="h3">
                     <a href="<%=linkit%>" title="Shop today to get <%=CashBack%> CashBack on <%=shop%>" target="_blank">About <%=shop%></a></h3>
                <span class="text2"><%=notes%></span><br>
                <br>
                <a title="Sign Up" href="http://www.funds4me.co.uk/register2.asp"><strong>Yes, please
                    I want to start saving and want to sign up - click here</strong></a>
                <br>
                <br>
                <a title="No Thanks" href="<%=linkit%>"
                    target="_blank" rel="nofollow">No Thanks - I'm a bit crazy and dont want to get
                    cash back. Take me to <%=shop%><br><br>
					<%=banner%></a>
            </div>
			<div class="signup">
                <div class="signup">
			<p>
               <strong>Coupons / Discounts</strong><br><br>
			   <%=specof%>
			   </p>
            </div>
            </div>
            
        </div>
		<div class="footer" align="center">
      <a href="Contact.asp">Contact</a> | <a href="Privacy.asp">Privacy Policy</a> | <a href="Technical.asp">Technical Requirements</a> | <a href="press.asp">News & Press</a> | <a href="useragreement.asp">EULA</a>	
			<br><br>
			&copy; 1998 - 2006 Robert Moorehouse
                </div>
		
		
	<%
	end if

end if

	%>
	
        
        
    </div>
<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-220799-7";
urchinTracker();
</script>
</body>
</html>
