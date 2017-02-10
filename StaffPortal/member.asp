<%
'
'-----------------------Begin Code----------------------------
'Check for valid info passed
if Request("ID") = "" then Redirect("error.asp")

'Open up the member's record
intID = Int(Request("ID"))
Query = "SELECT * FROM Members WHERE CustomerID = " & intCustomerID & " AND ID = " & intID
Set rsMember = Server.CreateObject("ADODB.Recordset")
rsMember.Open Query, Connect, adOpenStatic, adLockReadOnly
'Check for valid record
if rsMember.EOF then Redirect("error.asp")
'------------------------End Code-----------------------------
%>

<p class=Heading align=<%=rsSite("HeadingAlignment")%>>Member Information</p>
<p class=LinkText align=<%=rsSite("HeadingAlignment")%>><a href="javascript:history.back(1)">Back</a></p>
<br>
<%
if Request("Rating") <> "" and rsSite("RateMembers") = 1 then
	AddRating rsMember("ID"), "Members"
	%><a href="javascript:history.back(1)">Click here</a> to go back.<%
else
%>

	NickName: <%=rsMember("NickName")%><br>
<%
	if rsMember("PrivateName") = 0 OR LoggedMember then
%>
		Name: <%=rsMember("FirstName")%>&nbsp;<%=rsMember("LastName")%><br>
<%
	end if
	if (rsMember("PrivateEMail") = 0 OR LoggedMember) AND (rsMember("EMail1") <> "" AND rsMember("EMail2") <> "") then
		if rsMember("EMail1") <> "" AND rsMember("EMail2") = "" then
%>
			<br>E-Mail: <a href=mailto:<%=rsMember("EMail1")%>><%=rsMember("EMail1")%></a><br>
<%
		else
%>
			<br>E-Mail: <a href=mailto:<%=rsMember("EMail1")%>><%=rsMember("EMail1")%></a>,&nbsp;
					<a href=mailto:<%=rsMember("EMail2")%>><%=rsMember("EMail2")%></a><br>
<%
		end if
	end if
	if (rsMember("PrivateBeeper") = 0 OR LoggedMember) AND rsMember("Beeper") <> "" then
%>
		<br>Beeper Number: <%=rsMember("Beeper")%><br>
<%
	end if
	if (rsMember("PrivateCellPhone") = 0 OR LoggedMember) AND rsMember("CellPhone") <> "" then
%>
		<br>Cell Phone Number: <%=rsMember("CellPhone")%><br>
<%
	end if
	if (rsMember("PrivateHome") = 0 OR LoggedMember) AND rsMember("HomeStreet") <> "" then
%>
		<br>Home Address:<br>
		&nbsp;&nbsp;&nbsp;<%=rsMember("HomeStreet")%><br>
		&nbsp;&nbsp;&nbsp;<%=rsMember("HomeCity")%>,&nbsp;<%=rsMember("HomeState")%>&nbsp;<%=rsMember("HomeZip")%><br>
		&nbsp;&nbsp;&nbsp;<%=rsMember("HomePhone")%><br>
<%
	end if
	if (rsMember("PrivateSecondary") = 0 OR LoggedMember) AND rsMember("SecondaryDescription") <> "" AND rsMember("SecondaryStreet") <> "" then
%>
		<br><%=rsMember("SecondaryDescription")%> Address:<br>
		&nbsp;&nbsp;&nbsp;<%=rsMember("SecondaryStreet")%><br>
		&nbsp;&nbsp;&nbsp;<%=rsMember("SecondaryCity")%>,&nbsp;<%=rsMember("SecondaryState")%>&nbsp;<%=rsMember("SecondaryZip")%><br>
		&nbsp;&nbsp;&nbsp;<%=rsMember("SecondaryPhone")%><br>
<%
	end if

	if rsSite("RateMembers") = 1 then
		PrintRatingPulldown rsMember("ID"), "", "Members", "member.asp", "member"
	end if

	if rsSite("ReviewMembers") = 1 then
%>
		<a href="review.asp?Source=member.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Members">Add a review</a><br>
<%
		if ReviewsExist( "Members", rsMember("ID") ) then
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "member.asp", "Members", rsMember("ID")
			Set rsPage = Nothing
		end if
	end if
%>	<!-- #include file ="member_stats.asp" --> <%
end if

%>



<%
rsMember.Close
set rsMember = Nothing
%>