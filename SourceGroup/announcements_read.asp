<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeAnnouncements ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=AnnouncementsTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if



Public DisplayDate, DisplayAuthor, DisplaySubject

Query = "SELECT DisplayDateItemAnnouncements, DisplayAuthorItemAnnouncements, DisplaySubjectItemAnnouncements  FROM Look WHERE CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplayDate = CBool(rsItem("DisplayDateItemAnnouncements"))
	DisplayAuthor = CBool(rsItem("DisplayAuthorItemAnnouncements"))
	DisplaySubject = CBool(rsItem("DisplaySubjectItemAnnouncements"))
rsItem.Close



'Open up the item
Query = "SELECT ID, Date, MemberID, Subject, Body, Private FROM Announcements WHERE ID = " & intID & " AND CustomerID = " & CustomerID
rsItem.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The announcement does not exist.  If you pasted a link, there may be a typo, or the announcement may have been deleted.  Please refer to the announcement list to find the desired announcement, if it still exists."))
end if

if rsItem("Private") = 1 AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=announcements_read.asp&ID=" & intID & "&Submit=Read" )
end if

if Request("Rating") <> "" and RateAnnouncements = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Announcement</a></p>
<%
	AddRating intID, "Announcements"
else
	IncrementStat "AnnouncementsRead"

	IncrementHits intID, "Announcements"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_announcements_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
<a href="javascript:DeleteBox('If you delete this announcement, there is no way to get it back.  Are you sure?', 'members_announcements_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a>
</td>
		</tr>
		</table>
<%
	end if

'------------------------End Code-----------------------------
%>
	</p>
<%
	if DisplayDate or DisplayAuthor or DisplaySubject  then
		PrintTableHeader 100
		if DisplayDate or DisplayAuthor then %>

		<tr>
			<td colspan="2" class="<% PrintTDMain %>">
			<table width=100% cellspacing=0 cellpadding=0>
			<tr>
	
			<% if DisplayAuthor then %>
			<td class="<% PrintTDMain %>" align="left">Author: <%=PrintTDLink(GetNickNameLink(rsItem("MemberID")))%></td>
			<% end if %>	
			<% if DisplayDate then
				strAlign = "left"
				if DisplayAuthor then strAlign = "right"
			%>
			<td class="<% PrintTDMainSwitch %>" align="<%=strAlign%>">Date Written: <%=FormatDateTime(rsItem("Date"), 2)%></td>
			<% end if %>		
			</tr>	
			</table>
			</td>
		</tr>
<%
		end if
		if DisplaySubject then
%>	
		<tr>
			<td class="<% PrintTDMainSwitch %>" align="left" colspan="2">Subject: <%=rsItem("Subject")%></td>
		</tr>
<%
		end if
		Response.Write "</table>"
	end if
%>
	<br>
	<%=rsItem("Body")%>

	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RateAnnouncements = 1 then
		PrintRatingPulldown intID, "", "Announcements", "announcements_read.asp", "announcement"
	end if
	if ReviewAnnouncements = 1 then
%>
		<a href="review.asp?Source=announcements_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Announcements">Add A Review</a><br>
<%
		if ReviewsExist( "Announcements", intID ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=announcements_read.asp?ID=<%=intID%>&TargetTable=Announcements&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "announcements_read.asp", "Announcements", intID
			Set rsPage = Nothing
		end if
	end if
end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>