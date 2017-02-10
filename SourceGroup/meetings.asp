<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeMeetings ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))

if SectionViewMeetings = "Members" and not LoggedMember() then Redirect("login.asp?Source=meetings.asp&Message=" & Server.URLEncode("Only members can view this section.  If you are a member, please log in with your information below.  Otherwise, sorry, but you may not view this section."))
if SectionViewMeetings = "Administrators" and not LoggedAdmin() then
	if LoggedMember() then
		Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section can only be viewed by an administrator."))
	else
		Redirect("login.asp?Source=meetings.asp&Message=" & Server.URLEncode("Only <b>sitre administrators</b> can view this section.  If you are an administrator, please log in with your information below.  If you are a regular member or a non-member, sorry, but you may not view this section."))
	end if
end if

'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading><%=MeetingsTitle%></span><br>
<%
if (IncludeAddButtons = 1 or LoggedMember()) and (LoggedAdmin() or CBool( MeetingsMembers )) then
%>
	<span class=LinkText><a href="members_meetings_add.asp">Add A Meeting</a></span>
<%
end if
%>
</p>
<%
'-----------------------Begin Code----------------------------
'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")

intRateMeetings = RateMeetings
intReviewMeetings = ReviewMeetings

'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if Request("Keywords") <> "" then
	Query = "SELECT ID, Date, MemberID, Subject, Body FROM Meetings WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	Set rsList = Server.CreateObject("ADODB.Recordset")
	rsList.CacheSize = 100
	rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set MemberID = rsList("MemberID")
		Set Body = rsList("Body")
		Set Subject = rsList("Subject")
	intSearchID = SingleSearch()
	Session("SearchID") = intSearchID
	rsList.Close
end if


Set FileSystem = CreateObject("Scripting.FileSystemObject")
Public PostPath
PostPath = GetPath("posts")

Set rsList = Server.CreateObject("ADODB.Recordset")


Public ListType, DisplayDate, DisplayAuthor, DisplayPrivacy, blBulletImg, ItemNumber
	strImagePath = GetPath("images")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	ItemNumber = 0	'This will be set by the PrintPagesHeader sub

Query = "SELECT IncludePrivacyMeetings, DisplaySearchMeetings, DisplayDaysOldMeetings, InfoTextMeetings, ListTypeMeetings, DisplayDateListMeetings, DisplayAuthorListMeetings, DisplayPrivacyListMeetings  FROM Look WHERE CustomerID = " & CustomerID
rsList.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplaySearch = CBool(rsList("DisplaySearchMeetings"))
	DisplayDaysOld = CBool(rsList("DisplayDaysOldMeetings"))
	InfoText = rsList("InfoTextMeetings")
	ListType = rsList("ListTypeMeetings")
	DisplayDate = CBool(rsList("DisplayDateListMeetings"))
	DisplayAuthor = CBool(rsList("DisplayAuthorListMeetings"))
	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = (CBool(rsList("DisplayPrivacyListMeetings")) and CBool(rsList("IncludePrivacyMeetings"))) and not cBool(SiteMembersOnly)
rsList.Close

if DisplaySearch or DisplayDaysOld then
%>
	<form METHOD="POST" ACTION="meetings.asp">
<%	if DisplayDaysOld then	%>
	View Meeting Minutes In The Last <% PrintDaysOld %>
	<br>
<%		if DisplaySearch then Response.Write "Or "
	end if
	if DisplaySearch then	%>
	Search For <input type="text" name="Keywords" size="25">
	<input type="submit" name="Submit" value="Go"><br>
<%	end if	%>	
	</form>
<%
end if

if intSearchID <> "" then
	'Their search came up empty
	if intSearchID = 0 then
		if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
			<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
			Try again, or <a href="meetings.asp">click here</a> to view all meetings.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			Try again, or <a href="meetings.asp">click here</a> to view all meetings.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	else
		'They have search results, so lets list their results
		Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
		rsPage.CacheSize = PageSize
		Set TargetID = rsPage("TargetID")
%>
		<form METHOD="POST" ACTION="meetings.asp">
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
		PrintPagesHeader
		PrintListHeader

		'Instantiate the recordset for the output
		Query = "SELECT ID, Date, MemberID, Subject, TotalRating, TimesRated, Private, FileName, FileLinkDirect FROM Meetings WHERE CustomerID = " & CustomerID
		rsList.CacheSize = PageSize
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set MemberID = rsList("MemberID")
		Set TotalRating = rsList("TotalRating")
		Set TimesRated = rsList("TimesRated")
		Set Subject = rsList("Subject")
		Set IsPrivate = rsList("Private")
		Set FileName = rsPage("FileName")
		Set FileLinkDirect = rsPage("FileLinkDirect")

		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				rsList.Filter = "ID = " & TargetID

				PrintTableData

				rsPage.MoveNext
			end if
		next
		PrintListFooter
		rsPage.Close
		set rsPage = Nothing
	end if
'They are just cycling through the Meetings.  No searching.
else
	if InfoText <> " " and InfoText <> "" then 	Response.Write "<p>" & InfoText & "</p>"
	if Request("DaysOld") <> "" then
		CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
		Query = "SELECT ID, Date, MemberID, Subject, TotalRating, TimesRated, Private, CommitteeID, FileName, FileLinkDirect FROM Meetings WHERE (CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	else
		Query = "SELECT ID, Date, MemberID, Subject, TotalRating, TimesRated, Private, CommitteeID, FileName, FileLinkDirect FROM Meetings WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	end if
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then
%>
		<form METHOD="POST" ACTION="meetings.asp">
		<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set MemberID = rsPage("MemberID")
		Set TotalRating = rsPage("TotalRating")
		Set TimesRated = rsPage("TimesRated")
		Set CommitteeID = rsPage("CommitteeID")
		Set Subject = rsPage("Subject")
		Set IsPrivate = rsPage("Private")
		Set FileName = rsPage("FileName")
		Set FileLinkDirect = rsPage("FileLinkDirect")

		PrintPagesHeader
		PrintListHeader
		for j = 1 to rsPage.PageSize
			if not rsPage.EOF then
				PrintTableData
				rsPage.MoveNext
			end if
		next
		PrintListFooter
	else
		if Request("DaysOld") <> "" then
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there have been no meetings added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
		else
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there are no meetings at the moment.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	end if
	rsPage.Close
	set rsPage = Nothing
end if

set rsList = Nothing
Set FileSystem = Nothing

'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(Subject & Body & ItemDate & GetNickName(MemberID) )
End Function



'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintListHeader
	if ListType = "Table" then
		PrintTableHeader 0
%>		
	<tr>
		<% if DisplayDate then %>
		<td class="TDHeader">Date</td>
		<% end if %>	
		<% if DisplayAuthor then %>
		<td class="TDHeader">Author</td>
		<% end if %>	
		<td class="TDHeader">Subject</td>
		<% if intRateMeetings = 1  and intReviewMeetings = 0 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% elseif intRateMeetings = 0  and intReviewMeetings = 1 then %>
			<td class="TDHeader" align=center>Review</td>
		<% elseif intRateMeetings = 1  and intReviewMeetings = 1 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% end if %>	
		<% if DisplayPrivacy then %>
		<td class="TDHeader">Public?</td>
		<% end if %>	
	</tr>
<%
	elseif ListType = "Bulleted" and not blBulletImg then
		Response.Write "<ul>"
	else
		Response.Write "<p>"
	end if
End Sub

'-------------------------------------------------------------
'This prints the closing for the list
'-------------------------------------------------------------
Sub PrintListFooter
	if ListType = "Table" then
		Response.Write("</table>")

	elseif ListType = "Bulleted" and not blBulletImg then
		Response.Write "</ul>"
	else
		Response.Write "</p>"
	end if

	'Give them the link to change the section's properties
	if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
		Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Meetings&Source=meetings.asp'>Change Section Options</a></p>"
	end if
End Sub


'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
	if ListType = "Table" then
%>
	<tr>
		<% if DisplayDate then %>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<% end if %>	
		<% if DisplayAuthor then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
		<% end if %>
<%
		'If there is a file, we will include it heres
		if FileName <> "" and FileLinkDirect = 1 and FileSystem.FileExists( PostPath & FileName ) then
			strLink = NonSecurePath & "posts/" & FileName
%>
		<td class="<% PrintTDMain %>"><a href="<%=strLink%>"><%=PrintTDLink(Subject)%></a></td>
<%
		else
%>
		<td class="<% PrintTDMain %>"><a href="meetings_read.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
<%
		end if
		if intRateMeetings = 1 and intReviewMeetings = 0 then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="meetings_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate")%></a></font></td>
<%		elseif intRateMeetings = 0 and intReviewMeetings = 1 then
			if ReviewsExist( "Meetings", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="meetings_read.asp?ID=<%=ID%>"><%=PrintTDLink("Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="meetings_read.asp?ID=<%=ID%>"><%=PrintTDLink("Add Review")%></a></font></td>
<%			end if
		elseif intRateMeetings = 1 and intReviewMeetings = 1 then
			if ReviewsExist( "Meetings", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="meetings_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate and Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="meetings_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate/Add Review")%></a></font></td>
<%			end if
		end if%>
		<% if DisplayPrivacy then %>
		<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
		<% end if %>	
	</tr>
<%
	else
		strHeader = ""
		strFooter = "<br>"
		if ListType = "Bulleted" then
			if blBulletImg then
				strHeader = "<img src='images/BulletImage." & strBulletExt & "'>"
			else
				strHeader = "<li>"
				strFooter = "</li>"
			end if
		elseif ListType = "Numbered" then
				ItemNumber = ItemNumber + 1
				strHeader = ItemNumber & ".&nbsp;"
		end if

		Response.Write strHeader

		'If there is a file, we will include it heres
		if FileName <> "" and FileLinkDirect = 1 and FileSystem.FileExists( PostPath & FileName ) then
			strLink = NonSecurePath & "posts/" & FileName
%>
		<a href="<%=strLink%>"><%=Subject%></a>&nbsp;&nbsp;&nbsp;&nbsp;
<%
		else
%>
		<a href="meetings_read.asp?ID=<%=ID%>"><%=Subject%></a>&nbsp;&nbsp;&nbsp;&nbsp;
<%
		end if
%>
		<% if DisplayDate then %>
		<% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayAuthor then %>
		By: <%=GetNickNameLink(MemberID)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayPrivacy and IsPrivate = 1 then Response.Write "Private&nbsp;&nbsp;"
		if intRateMeetings = 1 and intReviewMeetings = 0 then
%>			<%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="meetings_read.asp?ID=<%=ID%>">Rate</a></font>&nbsp;&nbsp;
<%		elseif intRateMeetings = 0 and intReviewMeetings = 1 then
			if ReviewsExist( "Meetings", ID ) then
%>				<font size="-2"><a href="meetings_read.asp?ID=<%=ID%>">Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<font size="-2"><a href="meetings_read.asp?ID=<%=ID%>">Add Review</a></font>&nbsp;&nbsp;
<%			end if
		elseif intRateMeetings = 1 and intReviewMeetings = 1 then
			if ReviewsExist( "Meetings", ID ) then
%>				<%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="meetings_read.asp?ID=<%=ID%>">Rate and Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="meetings_read.asp?ID=<%=ID%>">Rate/Add Review</a></font>&nbsp;&nbsp;
<%			end if
		end if
		Response.Write strFooter
	end if
End Sub

'------------------------End Code-----------------------------
%>
