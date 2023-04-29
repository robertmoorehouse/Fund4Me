<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">



<%
if cashback = "Coupons Only" or cashback = "None" then
%>
<!-- Header -->
    <div align="center">
        <div id="hdr">
            <h1>
                <%=shop%>
                <br>
                Coupons and Discounts
            </h1>
            <h2 class="h2">
                Find coupons and discounts for <%=shop%> with Funds4Me.co.uk</h2>
        </div>
    </div>
    <!-- end of header -->
    <div align="center">
	<div id="lh-col">
            <h3 class="h3" align="center">
                <%=register%> today to get <%=CashBack%> CashBack with funds4me.co.uk</h3>
            <div class="red" align="center">
                <strong>Coupons and Discounts for <%=shop%></strong></div>
            <div class="text">
                <p>
                    Here is a list of the latest coupons, discount and offers available from <%=shop%>.<br><br>
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
			<a href="<%=linkit%>" title="Shop today at <%=shop%> using our coupons and discount codes" target="_blank">
                <h3 class="h3">
                    About <%=shop%></h3></a>
                <span class="text2"><%=notes%></span><br>
                <br>
                
                <br>
                <br><a href="<%=linkit%>" title="Shop today at <%=shop%> using our coupons and discount codes" target="_blank">
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
    <!-- Header -->
    <div align="center">
        <div id="hdr">
            <h1>
                <%=shop%>
                <br>
                Cash Back
            </h1>
            <h2 class="h2">
                Earn <%=CashBack%> Cashback on <%=shop%> through Funds4Me.co.uk</h2>
        </div>
    </div>
    <!-- end of header -->
    <div align="center">
	
		
	<%
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
			<a href="<%=linkit%>" title="Shop today to get <%=CashBack%> CashBack on <%=shop%>" target="_blank">
                <h3 class="h3">
                    About <%=shop%></h3></a>
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
                <a href="<%=linkit%>" title="Shop today to get <%=CashBack%> CashBack on <%=shop%>" target="_blank">
				<h3 class="h3">
                    About <%=shop%></h3></a>
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
		
		
	<%
	end if
	%>
	
        
        
    </div>
<%
end if
%>
  
  
  
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->
<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
</script>
<script type="text/javascript">
_uacct = "UA-220799-7";
urchinTracker();
</script>
</body>
</html>
