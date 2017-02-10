<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeStories ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=StoriesTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if



Public ListType, DisplayDate, DisplayAuthor, DisplaySubject

Query = "SELECT DisplayDateItemStories, DisplayAuthorItemStories, DisplaySubjectItemStories  FROM Look WHERE CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplayDate = CBool(rsItem("DisplayDateItemStories"))
	DisplayAuthor = CBool(rsItem("DisplayAuthorItemStories"))
	DisplaySubject = CBool(rsItem("DisplaySubjectItemStories"))
rsItem.Close



'Open up the item
Query = "SELECT ID, Date, MemberID, Subject, Body, Private FROM Stories WHERE ID = " & intID & " AND CustomerID = " & CustomerID
rsItem.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The story does not exist.  If you pasted a link, there may be a typo, or the story may have been deleted.  Please refer to the story list to find the desired story, if it still exists."))
end if

if rsItem("Private") = 1 AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=stories_read.asp&ID=" & intID & "&Submit=Read" )
end if

if Request("Rating") <> "" and RateStories = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Story</a></p>
<%
	AddRating intID, "Stories"
else
	IncrementStat "StoriesRead"
	IncrementHits intID, "Stories"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_stories_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
<a href="javascript:DeleteBox('If you delete this story, there is no way to get it back.  Are you sure?', 'members_stories_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a>
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
	if RateStories = 1 then
		PrintRatingPulldown intID, "", "Stories", "stories_read.asp", "story"
	end if
	if ReviewStories = 1 then
%>
		<a href="review.asp?Source=stories_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Stories">Add A Review</a><br>
<%
		if ReviewsExist( "Stories", intID ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=stories_read.asp?ID=<%=intID%>&TargetTable=Stories&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "stories_read.asp", "Stories", intID
			Set rsPage = Nothing
		end if
	end if
end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>