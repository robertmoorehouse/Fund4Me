<SCRIPT LANGUAGE=VBScript RUNAT=Server>
 FUNCTION Num2Dollars (ByVal Num)

	if isnull(num) then
	num = 0
	end if
	
   intDollars = int(Num)
   intCents = CInt((Num - intDollars) * 100)

 ' Formatting the cents
   If intCents = 0 Then
     strCents = "00"
   ElseIf intCents < 10 Then
     strCents = "0" & intCents
   Else
     strCents = CStr(intCents)
   End If

 ' Initializing Dollars
   strDollars = ""

 ' Determine the dollars
   While len(CStr(intDollars)) > 3
     intTemp = intDollars mod 1000

     If intTemp = 0 Then
	strDollars = "000" & strDollars
     ElseIf intTemp < 10 Then
	strDollars = "00" & CStr(intTemp) & strDollars
     ElseIf intTemp < 100 Then
	strDollars = "0" & CStr(intTemp) & strDollars
     Else
	strDollars = CStr(intTemp) & strDollars
     End If

     strDollars = "," & strDollars
     intDollars = (intDollars - intTemp) \ 1000
   Wend

   strDollars = CStr(intDollars) & strDollars

   Num2Dollars = strDollars & "." & strCents

 END FUNCTION
</SCRIPT>
