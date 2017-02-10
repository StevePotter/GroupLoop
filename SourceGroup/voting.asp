<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeVoting ) then Redirect("error.asp")
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading><%=VotingTitle%></span><br>
<%
if (IncludeAddButtons = 1 or LoggedMember()) and (LoggedAdmin() or CBool( VotingMembers )) then
%>
<span class=LinkText><a href="members_polls_add.asp">Add a Poll</a></span>
<%
end if
%>
</p>

<%
'-----------------------Begin Code----------------------------
'Link to the results of closed polls if there are any
if ClosedExist then
	%><a href="voting_results.asp?Action=ViewClosed">Click here</a> to view results from closed polls.<br><%
end if

blLoggedMember = LoggedMember
blLoggedAdmin = LoggedAdmin



intRateVoting = RateVoting
intReviewVoting = ReviewVoting

Set rsPage = Server.CreateObject("ADODB.Recordset")


Public InfoText, ListType, DisplayDate, DisplayAuthor, DisplayPrivacy, blBulletImg, ItemNumber
	strImagePath = GetPath("images")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	ItemNumber = 0	'This will be set by the PrintPagesHeader sub


Query = "SELECT IncludePrivacyVoting,  InfoTextVoting, ListTypeVoting, DisplayDateListVoting, DisplayAuthorListVoting, DisplayPrivacyListVoting  FROM Look WHERE CustomerID = " & CustomerID
rsPage.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	InfoText = rsPage("InfoTextVoting")
	ListType = rsPage("ListTypeVoting")
	DisplayDate = CBool(rsPage("DisplayDateListVoting"))
	DisplayAuthor = CBool(rsPage("DisplayAuthorListVoting"))
	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = (CBool(rsPage("DisplayPrivacyListVoting")) and CBool(rsPage("IncludePrivacyVoting"))) and not cBool(SiteMembersOnly)
rsPage.Close


if InfoText <> " " and InfoText <> "" then Response.Write "<p>" & InfoText & "</p>"

Query = "SELECT ID, Date, MemberID, Subject, TotalRating, TimesRated, Private, ResultsSecurity FROM VotingPolls WHERE (OpenToVote = 1 AND CustomerID = " & CustomerID & ") ORDER BY Date DESC"
rsPage.CacheSize = PageSize
rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
if not rsPage.EOF then

%>
	<form METHOD="POST" ACTION="voting.asp">
<%
	Set ID = rsPage("ID")
	Set ItemDate = rsPage("Date")
	Set MemberID = rsPage("MemberID")
	Set TotalRating = rsPage("TotalRating")
	Set TimesRated = rsPage("TimesRated")
	Set Subject = rsPage("Subject")
	Set IsPrivate = rsPage("Private")
	Set ResultsSecurity = rsPage("ResultsSecurity")
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
'------------------------End Code-----------------------------
%>
	<p>Sorry, but there are no open voting polls at the moment.</p>
<%
'-----------------------Begin Code----------------------------
end if
rsPage.Close
set rsPage = Nothing


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintListHeader
	if ListType = "Table" then
		PrintTableHeader 0
%>		
	<tr>
		<td class="TDHeader">&nbsp;</td>
		<% if DisplayDate then %>
		<td class="TDHeader">Date</td>
		<% end if %>	
		<% if DisplayAuthor then %>
		<td class="TDHeader">Author</td>
		<% end if %>	
		<td class="TDHeader">Subject</td>
		<% if intRateVoting = 1  and intReviewVoting = 0 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% elseif intRateVoting = 0  and intReviewVoting = 1 then %>
			<td class="TDHeader" align=center>Review</td>
		<% elseif intRateVoting = 1  and intReviewVoting = 1 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% end if %>	
		<% if DisplayPrivacy then %>
		<td class="TDHeader">Public?</td>
		<% end if %>	
		<td class="TDHeader">&nbsp;</td>
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
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Voting&Source=voting.asp'>Change Section Options</a></p>"
		end if
End Sub



'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
	if ListType = "Table" then
%>
	<tr>
		<td class="<% PrintTDMain %>"><a href="voting_cast.asp?ID=<%=ID%>"><%=PrintTDLink("Vote")%></a></td>
		<% if DisplayDate then %>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<% end if %>	
		<% if DisplayAuthor then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
		<% end if %>	
		<td class="<% PrintTDMain %>"><a href="voting_cast.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
<%		if intRateVoting = 1 and intReviewVoting = 0 then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="voting_results.asp?ID=<%=ID%>"><%=PrintTDLink("Rate")%></a></font></td>
<%		elseif intRateVoting = 0 and intReviewVoting = 1 then
			if ReviewsExist( "VotingPolls", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="voting_results.asp?ID=<%=ID%>"><%=PrintTDLink("Read/Add Reviews")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="voting_results.asp?ID=<%=ID%>"><%=PrintTDLink("Add Review")%></a></font></td>
<%			end if
		elseif intRateVoting = 1 and intReviewVoting = 1 then
			if ReviewsExist( "Voting", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="voting_results.asp?ID=<%=ID%>"><%=PrintTDLink("Rate and Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="voting_results.asp?ID=<%=ID%>"><%=PrintTDLink("Rate/Add Review")%></a></font></td>
<%			end if
		end if%>
		<% if DisplayPrivacy  then %>
		<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
		<% end if %>	
		<td class="<% PrintTDMain %>">
		
<%
		if not (( ResultsSecurity = 1 AND not blLoggedMember ) or (ResultsSecurity = 2 AND not blLoggedAdmin)) then
%>
		<a href="voting_results.asp?ID=<%=ID%>"><%=PrintTDLink("View Results")%></a>
<%
		end if
%>		
		</td>
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
%>
		<%=strHeader%>
		<a href="voting_cast.asp?ID=<%=ID%>">Vote</a>&nbsp;&nbsp;&nbsp;&nbsp;
		<a href="voting_cast.asp?ID=<%=ID%>"><%=Subject%></a>&nbsp;&nbsp;&nbsp;&nbsp;
		<% if DisplayDate then %>
		<% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayAuthor then %>
		By: <%=GetNickNameLink(MemberID)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayPrivacy and IsPrivate = 1 then Response.Write "Private&nbsp;&nbsp;" %>
		<a href="voting_results.asp?ID=<%=ID%>">View Results</a>&nbsp;&nbsp;
<%
		if intRateVoting = 1 and intReviewVoting = 0 then
%>			<%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="voting_results.asp?ID=<%=ID%>">Rate</a></font>&nbsp;&nbsp;
<%		elseif intRateVoting = 0 and intReviewVoting = 1 then
			if ReviewsExist( "Voting", ID ) then
%>				<font size="-2"><a href="voting_results.asp?ID=<%=ID%>">Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<font size="-2"><a href="voting_results.asp?ID=<%=ID%>">Add Review</a></font>&nbsp;&nbsp;
<%			end if
		elseif intRateVoting = 1 and intReviewVoting = 1 then
			if ReviewsExist( "Voting", ID ) then
%>				<%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="voting_results.asp?ID=<%=ID%>">Rate and Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="voting_results.asp?ID=<%=ID%>">Rate/Add Review</a></font>&nbsp;&nbsp;
<%			end if
		end if
		Response.Write strFooter
	end if
End Sub

Function ClosedExist()
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "PollsClosedExist"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	ClosedExist = CBool(blExists)
End Function
'------------------------End Code-----------------------------
%>
