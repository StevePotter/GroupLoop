<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeCalendar ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class="Heading" align="<%=HeadingAlignment%>"><%=CalendarTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if

Public DisplayDate, DisplayAuthor, DisplaySubject

Query = "SELECT DisplayDateItemCalendar, DisplayAuthorItemCalendar, DisplaySubjectItemCalendar  FROM Look WHERE CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplayDate = CBool(rsItem("DisplayDateItemCalendar"))
	DisplayAuthor = CBool(rsItem("DisplayAuthorItemCalendar"))
	DisplaySubject = CBool(rsItem("DisplaySubjectItemCalendar"))
rsItem.Close

'Open up the item
Query = "SELECT Date, MemberID, Subject, Body, Private, StartDate, EndDate FROM Calendar WHERE ID = " & intID & " AND CustomerID = " & CustomerID
rsItem.Open Query, Connect, adOpenStatic, adLockReadOnly

if rsItem("Private") = 1 AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=calendar_event_read.asp&ID=" & intID & "&Submit=Read" )
end if

if Request("Rating") <> "" and RateCalendar = 1 then
	AddRating intID, "Calendar"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Event</a></p>
<%
else
	IncrementStat "CalendarEventsRead"
	IncrementHits intID, "Calendar"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.back(1)">Back</a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID")) then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_calendar_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;<a href="javascript:DeleteBox('If you delete this event, there is no way to get it back.  Are you sure?', 'members_calendar_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a></td>
		</tr>
		</table>
<%
	end if
'------------------------End Code-----------------------------
%>
	</p>
	<% PrintTableHeader 100 %>
	<tr>
		<td colspan="2" class="<% PrintTDMain %>">
		<table width=100% cellspacing=0 cellpadding=0>
		<tr>
		<% if DisplayAuthor then
			strAlign = "align='left'"
			if not DisplayDate then strAlign = "colspan=2 align='left'"
		%>
		<td class="<% PrintTDMain %>" <%=strAlign%>>Author: <%=PrintTDLink(GetNickNameLink(rsItem("MemberID")))%></td>
		<% end if %>	
		<% if DisplayDate then
			strAlign = "align='right'"
			if not DisplayAuthor then strAlign = "colspan=2 align='left'"
		%>
		<td class="<% PrintTDMain %>" <%=strAlign%>>Date Written: <%=FormatDateTime(rsItem("Date"), 2)%></td>
		<% end if %>	
		</tr>
		<tr>
		<%
		if FormatDateTime(rsItem("StartDate"), 2) = FormatDateTime(rsItem("EndDate"), 2) then
		%>
		<td class="<% PrintTDMainSwitch %>" align="left" colspan=2>Date of Event: <%=FormatDateTime(rsItem("StartDate"), 2)%><br>Time: <%=FormatDateTime(rsItem("StartDate"), 3)%> - <%=FormatDateTime(rsItem("EndDate"), 3)%></td>
		<%
		else
		%>
		<td class="<% PrintTDMainSwitch %>" align="left" colspan=2>Dates of Event: <%=FormatDateTime(rsItem("StartDate"), 2)%>&nbsp;<%=FormatDateTime(rsItem("StartDate"), 3)%> - <%=FormatDateTime(rsItem("EndDate"), 2)%>&nbsp;<%=FormatDateTime(rsItem("EndDate"), 3)%></td>
		<%
		end if
		%>
		</tr>
		</table>

		</td>
	</tr>
<%	if DisplaySubject then	%>		
	<tr>
		<td class="<% PrintTDMainSwitch %>" align="left" colspan="2">Subject: <%=rsItem("Subject")%></td>
	</tr>
	<% end if %>
	</table>
	<br>
	<%=rsItem("Body")%>

	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RateCalendar = 1 then
		PrintRatingPulldown intID, "", "Calendar", "calendar_event_read.asp", "event"
	end if
	if ReviewCalendar = 1 then
%>
		<a href="review.asp?Source=calendar_event_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Calendar">Add a review</a><br>
<%
		if ReviewsExist( "Calendar", intID ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=calendar_event_read.asp?ID=<%=intID%>&TargetTable=Calendar&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "calendar_event_read.asp", "Calendar", intID
			Set rsPage = Nothing
		end if
	end if
end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>