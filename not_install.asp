
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
-->
</script>
<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">



								



<table cellpadding="0" cellspacing="0" border="0" width="517">
    <tr>
        <td bgcolor="#000000">
            <table cellpadding="0" cellspacing="0" border="0" align="left" width="517">
                <tr>
                    <td rowspan="4" class="black-text" bgcolor="#ffffff"><img src="/images/spacer.gif" width="25" height="1" border="0"></td>
                    <td colspan="2" class="black-text" bgcolor="#ffffff"><img src="/images/spacer.gif" width="1" height="25" border="0"></td>
                    <td rowspan="4" class="black-text" bgcolor="#ffffff"><img src="/images/spacer.gif" width="25" height="1" border="0"></td>
                </tr>
                <tr>
                    <td class="black-text" bgcolor="#ffffff"><h3>How does Funds4get-me-not Work?</h3>
                    Forget to visit Funds4Me first? Have no fear, Funds4get-me-not is here! Funds4get-me-not makes sure you don't miss out on the valuable Cash Back offered by Funds4Me.<br><br>
                    If you shop at any one of our 500 merchants, but forget to visit Funds4Me first, Funds4get-me-not will pop up to remind you to earn your Cash Back through our site.<br><br>
                    Remember to visit Funds4Me at www.funds4me.co.uk before shopping online to find the latest specials and exclusive offers guaranteed to save you money.
                    </td>
                    <td valign="top" bgcolor="#ffffff"><br></td>
                    
                </tr>
                <tr>
                    <td class="black-text" colspan="2" bgcolor="#ffffff"><br>*Funds4get-me-not is available for Windows Internet Explorer 4.0 and higher and AOL 5.0 and higher.</td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<%
		

cookieUserID = "funds4metest"
linkit = "http://www.tifps.com/cms/install.asp?memberid=funds4metest&client=funds4me.co.uk"

%>
<script LANGUAGE="JavaScript">
{
	<%
	Response.write("sTemp = """ & linkit & """")
	%>
	window.open(sTemp,'infowin','scrollbars=yes,width=570,height=570');
}
</script>
	
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->
