<p class="Heading" align="<%=HeadingAlignment%>">Rate</p>

<%
'-----------------------Begin Code----------------------------
if Request("Rating") <> "" and RateAnnouncements = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Item</a></p>
<%
	AddRating rsItem("ID"), "Announcements"
end if

Sub AddRating( intID, strTable )
	Query = "SELECT TimesRated, TotalRating, RatingScore FROM " & strTable & " WHERE (ID = " & intID & " AND CustomerID = " & CustomerID & ")"
	Set rsTempRating = Server.CreateObject("ADODB.Recordset")
	rsTempRating.Open Query, Connect, adOpenStatic, adLockOptimistic

	'Make sure they aren't fucking around
	if CInt(Request("Rating")) > RatingMax OR CInt(Request("Rating")) < 1 or rsTempRating.EOF then
		Set rsTempRating = Nothing
		Redirect("error.asp")
	end if

	rsTempRating("TimesRated") = rsTempRating("TimesRated") + 1
	rsTempRating("TotalRating") = rsTempRating("TotalRating") + Request("Rating")
	rsTempRating.Update
	rsTempRating("RatingScore") = GetRatingScore( rsTempRating("TotalRating"), rsTempRating("TimesRated") )
	rsTempRating.Update

	if rsTempRating("TimesRated") = 1 then
'------------------------End Code-----------------------------
%>
		This has now been rated once, with a rating of <%=rsTempRating("TotalRating")%>. <br>
<%
'-----------------------Begin Code----------------------------
	else
'------------------------End Code-----------------------------
%>
		This has now been rated <%=rsTempRating("TimesRated")%> times, with 
		an average rating of <%=GetRating( rsTempRating("TotalRating"), rsTempRating("TimesRated") ) %>. <br>
<%
'-----------------------Begin Code----------------------------
	end if
	rsTempRating.Close
	set rsTempRating = Nothing
End Sub

'------------------------End Code-----------------------------
%>