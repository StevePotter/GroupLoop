<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeCalendar ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))

'If no date is specified, make it today
if IsEmpty(Request("ID")) OR NOT IsDate(Request("ID")) Then
	EventDate = Date
else
	EventDate = CDate(Request("ID"))
end if
intEventDay = Day(EventDate)
intEventMonth = Month(EventDate)
intEventYear = Year(EventDate)
strEventMonthName = MonthName(intEventMonth)
strDateTitle = strEventMonthName & " " & intEventDay & ", " & intEventYear
'------------------------End Code-----------------------------
%>
<p class="Heading" align="<%=HeadingAlignment%>">Events For <%=strDateTitle%></p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back</a></p>

<%
'-----------------------Begin Code----------------------------
'Log them in if they requested it
if Request("Action") = "Login" and not LoggedMember then Redirect( "login.asp?Source=calendar_read.asp&ID=" & EventDate & "&Submit=Read" )

if Request("Rating") <> "" and RateCalendar = 1 then
	'Get the ID of the item
	if Request("ID") <> "" then
		intID = CInt(Request("ID"))
	else
		Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
	end if
	AddRating intID, "Calendar"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Event</a></p>
<%
else
	Query = "SELECT DisplayDateItemCalendar, DisplayAuthorItemCalendar, DisplaySubjectItemCalendar  FROM Look WHERE CustomerID = " & CustomerID
	Set rsItems = Server.CreateObject("ADODB.Recordset")
	rsItems.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
		DisplayDate = CBool(rsItems("DisplayDateItemCalendar"))
		DisplayAuthor = CBool(rsItems("DisplayAuthorItemCalendar"))
		DisplaySubject = CBool(rsItems("DisplaySubjectItemCalendar"))
	rsItems.Close

	Query = "SELECT ID, Date, CustomerID, MemberID, Subject, Body, Private, StartDate, EndDate, TotalRating, TimesRated FROM Calendar WHERE (CustomerID = " & CustomerID & " AND StartDate <='" & EventDate & " 11:59:59 PM' AND EndDate >='" & EventDate & " 12:00:00 AM')"
	rsItems.CacheSize = 40
	rsItems.Open Query, Connect, adOpenStatic, adLockReadOnly
	Set ID = rsItems("ID")
	Set ItemDate = rsItems("Date")
	Set MemberID = rsItems("MemberID")
	Set Subject = rsItems("Subject")
	Set Body = rsItems("Body")
	Set StartDate = rsItems("StartDate")
	Set EndDate = rsItems("EndDate")
	Set IsPrivate = rsItems("Private")
	Set TotalRating = rsItems("TotalRating")
	Set TimesRated = rsItems("TimesRated")

	if CalendarShowBirthdays = 1 then
		Query = "Select ID FROM Members WHERE ( CustomerID = " & CustomerID & " AND " & _
			"( Day(Birthdate) ='" & intEventDay & "' AND Month(Birthdate) = '" & intEventMonth & "') )"
	else
		'force an empty recordset.  this seems to avoid a real complicated if statement thing.  Pretty ghetto though...
		Query = "SELECT NickName FROM Members WHERE ID = 0"
	end if

	Set rsBDays = Server.CreateObject("ADODB.Recordset")
	rsItems.CacheSize = 40
	rsBDays.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsItems.EOF AND rsBDays.EOF then
'------------------------End Code-----------------------------
%>
		<p>Sorry, but there are no events for <%=strDateTitle%> at the moment.</p>
<%
'-----------------------Begin Code----------------------------
		rsBDays.Close
		set rsBDays = Nothing
		rsItems.Close
	else
		if CalendarShowBirthdays = 1 then
			do until rsBDays.EOF
'------------------------End Code-----------------------------
%>
				<p class="Heading">Happy birthday, <%=GetNickName(rsBDays("ID"))%>! &nbsp;<%=CalendarBirthdayMessage %></p>
<%
'-----------------------Begin Code----------------------------
				rsBDays.MoveNext
			loop
		end if

		do until rsItems.EOF
			IncrementStat "CalendarEventsRead"
			IncrementHits ID, "Calendar"
			if LoggedAdmin or (LoggedMember and Session("MemberID") = MemberID) then
%>
				<table align=<%=HeadingAlignment%>>
				<tr>
				<td align=right width="50%" class="LinkText"><a href="members_calendar_modify.asp?Submit=Edit&ID=<%=ID%>">Edit</a>&nbsp;&nbsp;</td>
				<td align=left width="50%" class="LinkText">&nbsp;&nbsp;<a href="javascript:DeleteBox('If you delete this event, there is no way to get it back.  Are you sure?', 'members_calendar_modify.asp?Submit=Delete&ID=<%=ID%>')">Delete</a></td>
				</tr>
				</table>
<%
			end if
			PrintTableHeader 100
			if DisplayAuthor then
'------------------------End Code-----------------------------
%>
			<tr>
				<td class="<% PrintTDMainSwitch %>">Author: <%=PrintTDLink(GetNickNameLink(rsItems("MemberID")))%></td>
			</tr>
<%
'-----------------------Begin Code----------------------------
			end if
			'Different days
			if FormatDateTime(rsItems("StartDate"), 2) <> FormatDateTime(rsItems("EndDate"), 2) then
'------------------------End Code-----------------------------
%>
				<tr>
					<td class="<% PrintTDMainSwitch %>">Dates: <%=FormatDateTime(rsItems("StartDate"), 2)%>&nbsp;<%=FormatDateTime(rsItems("StartDate"), 3)%> - <%=FormatDateTime(rsItems("EndDate"), 2)%>&nbsp;<%=FormatDateTime(rsItems("EndDate"), 3)%></td>
				</tr>
<%
'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<tr>
					<td class="<% PrintTDMainSwitch %>">Times: <%=FormatDateTime(rsItems("StartDate"), 3)%> - <%=FormatDateTime(rsItems("EndDate"), 3)%></td>
				</tr>
<%
'-----------------------Begin Code----------------------------
			end if
			if DisplaySubject then
'------------------------End Code-----------------------------
%>
			<tr>
				<td class="<% PrintTDMainSwitch %>">Subject: <%=Subject%></td>
			</tr>
<%
			end if
%>
			</table>
			<br>
<%
			if IsPrivate = 1 AND not LoggedMember then
%>
				This is a private event.  If you are a member, <a href="calendar_read.asp?ID=<%=Server.URLEncode(EventDate)%>&Action=Login">click here</a> to log in and view the event.
<%			else
				Response.Write Body
%>
			<br>
			<br>
<%
'-----------------------Begin Code----------------------------
				strRating = ""
				if TimesRated > 0 then strRating = "Rating: " & GetRating( TotalRating, TimesRated ) & "<br>"
				Response.Write strRating

				if RateCalendar = 1 and ReviewCalendar = 0 then
					%><a href="calendar_event_read.asp?ID=<%=ID%>">Rate This Event</a><br><%
				elseif RateCalendar = 0 and ReviewCalendar = 1 then
					if ReviewsExist( "Calendar", ID ) then
						%><a href="calendar_event_read.asp?ID=<%=ID%>">Read/Add Reviews</a><br><%
					else
						%><a href="calendar_event_read.asp?ID=<%=ID%>">Add A Review</a><br><%
					end if
				elseif RateCalendar = 1 and ReviewCalendar = 1 then
					if ReviewsExist( "Calendar", ID ) then
						%><a href="calendar_event_read.asp?ID=<%=ID%>">Rate This Event and Read/Add Reviews</a><br><%
					else
						%><a href="calendar_event_read.asp?ID=<%=ID%>">Rate/Review This Event</a><br><%
					end if
				end if
				Response.Write "<br>"
			end if
			rsItems.MoveNext
		loop
		rsItems.Close
		set rsItems = Nothing
		rsBDays.Close
		set rsBDays = Nothing
	end if

end if
'------------------------End Code-----------------------------
%>