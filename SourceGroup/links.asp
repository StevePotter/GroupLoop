<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeLinks ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading><%=LinksTitle%></span><br>
<%
if IncludeAddButtons = 1 or LoggedMember() then
%>
<span class=LinkText><a href="members_links_add.asp">Add A Link</a></span>
<%
end if
%>
</p>
<%
'-----------------------Begin Code----------------------------
'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")

intRateLinks = RateLinks
intReviewLinks = ReviewLinks

Set rsList = Server.CreateObject("ADODB.Recordset")

'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if Request("Keywords") <> "" then
	Query = "SELECT ID, Date, MemberID, URL, Name, Description FROM Links WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	rsList.CacheSize = 100
	rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set MemberID = rsList("MemberID")
		Set URL = rsList("URL")
		Set Name = rsList("Name")
		Set Description = rsList("Description")
	intSearchID = SingleSearch()
	Session("SearchID") = intSearchID
	rsList.Close
end if




Public ListType, DisplayDate, DisplayAuthor, DisplayPrivacy, blBulletImg, ItemNumber
	strImagePath = GetPath("images")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	ItemNumber = 0	'This will be set by the PrintPagesHeader sub

Query = "SELECT IncludePrivacyLinks, DisplaySearchLinks, DisplayDaysOldLinks, InfoTextLinks, ListTypeLinks, DisplayDateListLinks, DisplayAuthorListLinks, DisplayPrivacyListLinks  FROM Look WHERE CustomerID = " & CustomerID
rsList.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplaySearch = CBool(rsList("DisplaySearchLinks"))
	DisplayDaysOld = CBool(rsList("DisplayDaysOldLinks"))
	InfoText = rsList("InfoTextLinks")
	ListType = rsList("ListTypeLinks")
	DisplayDate = CBool(rsList("DisplayDateListLinks"))
	DisplayAuthor = CBool(rsList("DisplayAuthorListLinks"))
	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = (CBool(rsList("DisplayPrivacyListLinks")) and CBool(rsList("IncludePrivacyLinks"))) and not cBool(SiteMembersOnly)
rsList.Close


if DisplaySearch or DisplayDaysOld then
%>
	<form METHOD="POST" ACTION="links.asp">
<%	if DisplayDaysOld then	%>
	View Links In The Last <% PrintDaysOld %>
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
			Try again, or <a href="links.asp">click here</a> to view all links.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			Try again, or <a href="links.asp">click here</a> to view all links.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	else
		'They have search results, so lets list their results
		Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
		rsPage.CacheSize = PageSize
%>
		<form METHOD="POST" ACTION="links.asp">
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
		PrintPagesHeader
		PrintListHeader

		'Instantiate the recordset for the output
		Set rsList = Server.CreateObject("ADODB.Recordset")
		Query = "SELECT ID, Date, MemberID, URL, Name, Description, TotalRating, TimesRated, Private FROM Links WHERE CustomerID = " & CustomerID
		rsList.CacheSize = PageSize
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set MemberID = rsList("MemberID")
		Set TotalRating = rsList("TotalRating")
		Set TimesRated = rsList("TimesRated")
		Set URL = rsList("URL")
		Set Name = rsList("Name")
		Set Description = rsList("Description")
		Set IsPrivate = rsList("Private")

		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				rsList.Filter = "ID = " & rsPage("TargetID")

				PrintTableData

				rsPage.MoveNext
			end if
		next
		PrintListFooter

		rsPage.Close
		set rsPage = Nothing
		set rsList = Nothing
	end if
'They are just cycling through the links.  No searching.
else
	if InfoText <> " " and InfoText <> "" then Response.Write "<p>" & InfoText & "</p>"
	'This is if they requested links written in a time period
	if Request("DaysOld") <> "" then
		CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
		Query = "SELECT ID, Date, MemberID, URL, Name, Description, TotalRating, TimesRated, Private FROM Links WHERE (CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	else
		Query = "SELECT ID, Date, MemberID, URL, Name, Description, TotalRating, TimesRated, Private FROM Links WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	end if
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then
%>
		<form METHOD="POST" ACTION="links.asp">
		<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set MemberID = rsPage("MemberID")
		Set TotalRating = rsPage("TotalRating")
		Set TimesRated = rsPage("TimesRated")
		Set URL = rsPage("URL")
		Set Name = rsPage("Name")
		Set Description = rsPage("Description")
		Set IsPrivate = rsPage("Private")

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
			<p>Sorry, but there have been no links added in that time period. <a href="javascript:hilink.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
		else
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there are no links at the moment.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	end if
	rsPage.Close
	set rsPage = Nothing
end if


'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(URL & Name & Description & ItemDate & GetNickName(MemberID) )
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
		<td class="TDHeader">Link</td>
		<td class="TDHeader">Description</td>
		<% if intRateLinks = 1  and intReviewLinks = 0 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% elseif intRateLinks = 0  and intReviewLinks = 1 then %>
			<td class="TDHeader" align=center>Review</td>
		<% elseif intRateLinks = 1  and intReviewLinks = 1 then %>
			<td class="TDHeader" align=center>Rating</td>
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
		Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Links&Source=links.asp'>Change Section Options</a></p>"
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
		if IsPrivate = 1 and not LoggedMember then %>
			<td class="<% PrintTDMain %>" colspan="2">This is a private link. &nbsp;<a href="login.asp?Source=links_read.asp&ID=<%=ID%>&Submit=Read"><%=PrintTDLink("Click here")%></a> to log in and view it.</td>
<%		else
			strName = Name
			if Name = "" then strName = URL
%>
			<td class="<% PrintTDMain %>"><a href="<%=URL%>" target="_blank"><%=PrintTDLink(strName)%></a></td>
			<td class="<% PrintTDMain %>"><%=Description%></td>
<%		end if
		if intRateLinks = 1 and intReviewLinks = 0 then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="links_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate")%></a></font></td>
<%		elseif intRateLinks = 0 and intReviewLinks = 1 then
			if ReviewsExist( "Links", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="links_read.asp?ID=<%=ID%>"><%=PrintTDLink("Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="links_read.asp?ID=<%=ID%>"><%=PrintTDLink("Add Review")%></a></font></td>
<%			end if
		elseif intRateLinks = 1 and intReviewLinks = 1 then
			if ReviewsExist( "Links", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="links_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate and Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="links_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate/Add Review")%></a></font></td>
<%			end if
		end if%>

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
<%
		if IsPrivate = 1 and not LoggedMember then %>
			This is a private link. &nbsp;<a href="login.asp?Source=links_read.asp&ID=<%=ID%>&Submit=Read">Click here</a> to log in and view it.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%		else
			strName = Name
			if Name = "" then strName = URL
%>
			<a href="<%=URL%>" target="_blank"><%=strName%></a>&nbsp;&nbsp;&nbsp;&nbsp;
			<%=Description%>&nbsp;&nbsp;&nbsp;&nbsp;
<%		end if %>
		<% if DisplayDate then %>
		<% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayAuthor then %>
		By: <%=GetNickNameLink(MemberID)%>&nbsp;&nbsp;
		<% end if %>	
<%
		if intRateLinks = 1 and intReviewLinks = 0 then
%>			<%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="links_read.asp?ID=<%=ID%>">Rate</a></font>&nbsp;&nbsp;
<%		elseif intRateLinks = 0 and intReviewLinks = 1 then
			if ReviewsExist( "Links", ID ) then
%>				<font size="-2"><a href="links_read.asp?ID=<%=ID%>">Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<font size="-2"><a href="links_read.asp?ID=<%=ID%>">Add Review</a></font>&nbsp;&nbsp;
<%			end if
		elseif intRateLinks = 1 and intReviewLinks = 1 then
			if ReviewsExist( "Links", ID ) then
%>				<%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="links_read.asp?ID=<%=ID%>">Rate and Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="links_read.asp?ID=<%=ID%>">Rate/Add Review</a></font>&nbsp;&nbsp;
<%			end if
		end if
		Response.Write strFooter
	end if
End Sub
'------------------------End Code-----------------------------
%>
