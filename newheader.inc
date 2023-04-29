<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<!--#INCLUDE FILE="connectionstrings.inc"-->
<!--#INCLUDE FILE="userdetails.inc"-->
<!--#INCLUDE FILE="headerdetails.inc"-->
<!--#INCLUDE FILE="functions.inc"-->
<!--#INCLUDE FILE="displaypounds.inc"-->

<head>
    <title><%=titletext%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta http-equiv="PRAGMA" content="NO-CACHE">
    <meta content="Get <%=CashBack%>&nbsp;Cash Back on&nbsp;<%=shop%>. Sign up to funds4me.co.uk. The fastest growing Cash Back shopping site on the net." name="description">
	<META NAME="keywords" CONTENT="<%=keywordlist%>">
	<link rel="stylesheet" href="funds4me.css" type="text/css" />
</head>
<body>


<div id="contentHolder">

<div id="masthead">
  <h1 id="siteName">The Cash Back shopping portal</h1>
  <div class="whiteline"></div>
  <div id="globalNav"><form method="POST" action="javascript:gotostore('shopID')" id="form1" name="form1"><select size="1" id="shopID" name="shopID"><% storesearch() %></select><input type="submit" value="Go Shop" name="B3" ID="Submit1"><img src="siteimages/spacer.gif" width="225px" height="1px" /><a href="index.asp" title="Return to Home Page">Home</a> | <a href="fastcash.asp" title="Earn fast cash without spending a penny">Fast Cash</a> | <a href="prizes.asp" title="Visit price sites">Prizes</a> | <a href="coupons.asp" title="Latest coupons and vouchers">Coupons</a> | <a href="help.asp" title="read our help pages">Help Desk</a> | <a href="register.asp" title="Join now and start earning cashback while you shop">Sign Up</a> | <a href="login.asp">Login</a> </form></div>
</div>
  
  <div class="whiteline"></div>
<div id="navBar">
  <div id="search">
    <form action="#">
      <label>search</label>
      <input name="searchFor" type="text" size="10" />
      <input name="goButton" type="submit" value="go" />
    </form>
  </div>
  <div class="whiteline"></div>
  <div class="sectionLinks">
    <ul>
      <li><a href="index.asp" title="Return to Home Page"><img src="siteimages/spacer.gif" alt="Home" />Home</a></li>
      <li><a href="about.asp" title="Visit the About us page"><img src="siteimages/spacer.gif" alt="About us" />About us</a></li>
      <li><a href="help.asp" title="Visit the Help Center"><img src="siteimages/spacer.gif" alt="Help Center" />Help Center</a></li>
      <li><a href="faq.asp" title="Read the FAQ's"><img src="siteimages/spacer.gif" alt="FAQ's" />FAQ's</a></li>
      <li><a href="advertise.asp" title="Advertise on our site"><img src="siteimages/spacer.gif" alt="Advertise" />Advertise</a></li>
    </ul>
  </div>
  <div class="whiteline"></div>
  <div class="relatedLinks">
    <ul>
	<% if memberID <> "NotLoggedIn" then %>
		<li><a href="upgrade.asp" title="Double commission for full members - upgrade now!"><img src="siteimages/spacer.gif" alt="Upgrade Account" />Upgrade Account</a></li>
		<li><a href="refer.asp" title="Earn CASH for refering new users!"><img src="siteimages/spacer.gif" alt="Refer a friend" />Refer a friend</a></li>
		<li><a href="account.asp" title="View your account"><img src="siteimages/spacer.gif" alt="Your Account" />Your Account</a></li>
		<li><a href="favourites.asp" title="View the list of your favourite stores"><img src="siteimages/spacer.gif" alt="Favourites" />Favourites</a></li>
		<li><a href="recent.asp" title="View all your recent stores"><img src="siteimages/spacer.gif" alt="Recently Visited" />Recently Visited</a></li>
	<% else %>
		<li><a href="register.asp" title="Join now and start earning cashback while you shop"><img src="siteimages/spacer.gif" alt="Sign up" />Sign up</a></li>
		<li><a href="refer.asp" title="Earn CASH for refering new users!"><img src="siteimages/spacer.gif" alt="Refer a friend" />Refer a friend</a></li>
	<% end if %>
    </ul>
  </div>
  <div class="whiteline"></div>
  <% if memberID <> "NotLoggedIn" then %>
  <div id="search">
    <form action="<%=currentpage() %>?logout=true" method="post">
      <label>You have earned
      <% response.write(totalearned()) %> to date.</label>
	  <input name="goButton" type="submit" value="Login Out" />
    </form>
  </div>
  <div class="clearLinks">
    <h3>Your 5 favourite sites</h3>
	shows only if logged in
    <ul>
      <li><a href="#"><%=memberID%></a></li>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
    </ul>
  </div>
  <% else %>
  <div id="search">
    <form action="<%=currentpage() %>" method="post">
      <label>Log In</label>
      <input name="UserName" type="text" size="15" /><br />
	  <input name="Password" type="text" size="15" /><br />
      <input name="goButton" type="submit" value="Login" />
    </form>
  </div>
  <% end if %>
  <div class="whiteline"></div>
  <div class="Last10">
    <h3><%'=personalise%>Latest Stores Visited</h3>
    <ul>
<% last10stores() %>
    </ul>
  </div>
</div>
<!--end navBar div -->


<!--Right Bar -->

<div id="headlines">

   <div class="topstores">
    <ul>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
      <li><a href="#">Related Link</a></li>
    </ul>
  </div>
  <div class="whiteline"></div>
   <div class="newstores">
   <ul>
<% newstores() %>
    </ul>
  </div>
<div class="whiteline"></div>
 <div class="categories">
    <ul>
<% Categories() %>
	</ul>
  </div>

</div>
<!--end Right Bar -->

