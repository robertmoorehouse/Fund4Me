<!--#INCLUDE FILE="newheader.inc"-->


  
<!--content -->
<div id="content">



						
							    <script LANGUAGE="JavaScript">
								{
									<%
									Response.write("sTemp = """ & linkit & """")
									%>
									window.open(sTemp,'infowin');
								}
								</script>

								  <table border="0" width="490">
								    <tr>
								      <td class="standart" width="26%" valign="top" align="left" bgcolor="#FFCC99" colspan="2">
										
											Funds4Me.co.uk in association with <%=shop%> <%= Cashtext %>. Enjoy your shopping.
										<br><b><%=showval%></b></td>
								      
								      
								    </tr>
								    <tr>
								    <td width="26%" valign="top" align="left" bgcolor="#ffffff" colspan="2"><font size="2"></td></tr>
								    <tr>
								    <td width="26%" valign="top" align="left" bgcolor="#ffffff" colspan="2">
								   </td></tr>
								    <tr>
								      <td class="standart" width="26%" valign="top" align="left"><%=InsertCursor("Banner")%><br>
								        <%=InsertCursor("Merchant")%> <br><br>
								        <b>Maximum Cash Back of <%=InsertCursor("Charityrebatesval")%></td>
								      <td class="standart" width="14%" valign="top" align="left"><b>Email&nbsp;:&nbsp;</b><br><%=InsertCursor("email")%><br><br>
								        <b>Address&nbsp;:&nbsp;</b><br><%=InsertCursor("address")%><br>
								        <%=InsertCursor("city")%>&nbsp;
								        <%=InsertCursor("state")%>&nbsp;
								        <%=InsertCursor("zip")%>&nbsp;
								        <%=InsertCursor("country")%><br><br>
								        <b>Phone&nbsp;:&nbsp;</b><br><%=InsertCursor("phone")%><br><Br></td>
								    </tr>
								    <tr>
								      <td class="standart" width="40%" valign="top" align="left" colspan="2"><b>Payment Options.</b><br><%=InsertCursor("PaymentTypes")%><br></td>
								    </tr>
								    <tr>
								      <td class="standart" width="40%" valign="top" align="left" colspan="2"><b>Shipping Policy</b><br><%=InsertCursor("ShippingTypes")%><br></td>
								    </tr>
								    <tr>
								      <td class="standart" width="40%" valign="top" align="left" colspan="2"><b>About this Merchant</b><br><%=InsertCursor("notes")%><br><Br></td>
								    </tr>
								    <tr>
								      <td class="standart" width="40%" valign="top" align="left" colspan="2"><b>Coupons & Special Offers</b><br><%=specof%><br><Br></td>
								    </tr>
								    <tr>
								      <td class="standart" width="26%" valign="top" align="left" colspan="2">

									  </td>
									</tr>
								</table>
					
						

  
  
  
</div>
<!--end content -->



<!--#INCLUDE FILE="newfooter.inc"-->