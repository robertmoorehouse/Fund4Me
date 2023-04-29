<SCRIPT LANGUAGE=VBScript RUNAT=Server>
 FUNCTION Num2Pounds (ByVal Num)
if num < 0 then
	Num2Pounds = "£" & CStr(Num)
else

   intPounds = int(Num)
   intPence = CInt((Num - intPounds) * 100)

 ' Formatting the Pence
   If intPence = 0 Then
     strPence = "00"
   ElseIf intPence < 10 Then
     strPence = "0" & intPence
   Else
     strPence = CStr(intPence)
   End If

 ' Initializing Pounds
   strPounds = ""

 ' Determine the Pounds
   While len(CStr(intPounds)) > 3
     intTemp = intPounds mod 1000

     If intTemp = 0 Then
	strPounds = "000" & strPounds
     ElseIf intTemp < 10 Then
	strPounds = "00" & CStr(intTemp) & strPounds
     ElseIf intTemp < 100 Then
	strPounds = "0" & CStr(intTemp) & strPounds
     Else
	strPounds = CStr(intTemp) & strPounds
     End If

     strPounds = "," & strPounds
     intPounds = (intPounds - intTemp) \ 1000
   Wend

   strPounds = CStr(intPounds) & strPounds

   Num2Pounds = "£" & strPounds & "." & strPence
end if
 END FUNCTION
</SCRIPT>
<SCRIPT LANGUAGE=VBScript RUNAT=Server>
 FUNCTION noNum2Pounds (ByVal Num)

   intPounds = int(Num)
   intPence = CInt((Num - intPounds) * 100)

 ' Formatting the Pence
   If intPence = 0 Then
     strPence = "00"
   ElseIf intPence < 10 Then
     strPence = "0" & intPence
   Else
     strPence = CStr(intPence)
   End If

 ' Initializing Pounds
   strPounds = ""

 ' Determine the Pounds
   While len(CStr(intPounds)) > 3
     intTemp = intPounds mod 1000

     If intTemp = 0 Then
	strPounds = "000" & strPounds
     ElseIf intTemp < 10 Then
	strPounds = "00" & CStr(intTemp) & strPounds
     ElseIf intTemp < 100 Then
	strPounds = "0" & CStr(intTemp) & strPounds
     Else
	strPounds = CStr(intTemp) & strPounds
     End If

     strPounds = "," & strPounds
     intPounds = (intPounds - intTemp) \ 1000
   Wend

   strPounds = CStr(intPounds) & strPounds

   noNum2Pounds = strPounds & "." & strPence

 END FUNCTION
</SCRIPT>