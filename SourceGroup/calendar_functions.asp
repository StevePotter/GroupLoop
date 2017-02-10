<%
'-------------------------------------------------------------------------
' These print the individual table divisions for faded days of the month
'-------------------------------------------------------------------------
Sub Write_TD_Faded()
'------------------------End Code-----------------------------
%>
	<td valign=top width="14%" class="<% PrintTDMain %>">&nbsp;</td>
<%	
'-----------------------Begin Code----------------------------	
End Sub

'------------------------------------------------------------
' This function finds the last date of the given month
'------------------------------------------------------------
Function GetLastDay(intMonthNum, intYearNum)
	if intMonthNum = 4 or intMonthNum = 6 or intMonthNum = 9 or intMonthNum = 11 then
		GetLastDay = 30
	elseif intMonthNum = 2 then
		if intYearNum Mod 4 = 0 then
			GetLastDay = 29
		else
			GetLastDay = 28
		end if
	else
		GetLastDay = 31
	end if
End Function

%>