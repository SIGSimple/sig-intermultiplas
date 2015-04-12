<%
Function CaptalizeText(Text)
	Temp = Split(Text, " ")
	For i = LBound(Temp) to UBound(Temp)
		TextTemp = TextTemp & " " & UCase(Left(Temp(i), 1)) & Mid(Temp(i), 2)
	Next
	CaptalizeText = TextTemp
End Function
%>