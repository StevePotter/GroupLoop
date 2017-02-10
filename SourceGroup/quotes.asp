<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeQuotes ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading><%=QuotesTitle%></span><br>
<%
if IncludeAddButtons = 1 or LoggedMember() then
%>
<span class=LinkText><a href="members_quotes_add.asp">Add A Quote</a></span>
<%
end if
%>
</p>

<form METHOD="POST" ACTION="quotes.asp">
	View Quotes In The Last <% PrintDaysOld %>
	<br>
	Or Search For <input type="text" name="Keywords" size="25">
	<input type="submit" name="Submit" value="Go"><br>
</form>
<%
'-----------------------Begin Code----------------------------
'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")

intRateQuotes = RateQuotes
intReviewQuotes = ReviewQuotes

Set rsList = Server.CreateObject("ADODB.Recordset")


Public ListType, DisplayDate, DisplayAuthor, DisplayPrivacy, blBulletImg, ItemNumber
	strImagePath = GetPath("images")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	ItemNumber = 0	'This will be set by the PrintPagesHeader sub

Query = "SELECT ListTypeQuotes, DisplayDateListQuotes, DisplayAuthorListQuotes, DisplayPrivacyListQuotes  FROM Look WHERE CustomerID = " & CustomerID
rsList.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	ListType = rsList("ListTypeQuotes")
	DisplayDate = CBool(rsList("DisplayDateListQuotes"))
	DisplayAuthor = CBool(rsList("DisplayAuthorListQuotes"))
	DisplayPrivacy = CBool(rsList("DisplayPrivacyListQuotes"))
rsList.Close



'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if Request("Keywords") <> "" then
	Query = "SELECT ID, Date, MemberID, Author, Quote, Description FROM Quotes WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	Set rsList = Server.CreateObject("ADODB.Recordset")
	rsList.CacheSize = 100
	rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set MemberID = rsList("MemberID")
		Set Author = rsList("Author")
		Set Description = rsList("Description")
		Set Quote = rsList("Quote")
	intSearchID = SingleSearch()
	Session("SearchID") = intSearchID
	rsList.Close
end if

if intSearchID <> "" then
	'Their search came up empty
	if intSearchID = 0 then
		if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
			<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
			Try again, or <a href="quotes.asp">click here</a> to view all quotes.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			Try again, or <a href="quotes.asp">click here</a> to view all quotes.</p>
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
		<form METHOD="POST" ACTION="quotes.asp">
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
		PrintPagesHeader
		PrintListHeader

		'Instantiate the recordset for the output
		Set rsList = Server.CreateObject("ADODB.Recordset")
		Query = "SELECT ID, Date, MemberID, Author, Quote, TotalRating, TimesRated, Private FROM Quotes WHERE CustomerID = " & CustomerID
		rsList.CacheSize = PageSize
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set MemberID = rsList("MemberID")
		Set TotalRating = rsList("TotalRating")
		Set TimesRated = rsList("TimesRated")
		Set Author = rsList("Author")
		Set Quote = rsList("Quote")
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
'They are just cycling through the quotes.  No searching.
else
	'This is if they requested quotes written in a time period
	if Request("DaysOld") <> "" then
		CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
		Query = "SELECT ID, Date, MemberID, Author, Quote, TotalRating, TimesRated, Private FROM Quotes WHERE (CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	else
		Query = "SELECT ID, Date, MemberID, Author, Quote, TotalRating, TimesRated, Private FROM Quotes WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	end if
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then
%>
		<form METHOD="POST" ACTION="quotes.asp">
		<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set MemberID = rsPage("MemberID")
		Set TotalRating = rsPage("TotalRating")
		Set TimesRated = rsPage("TimesRated")
		Set Author = rsPage("Author")
		Set Quote = rsPage("Quote")
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
			<p>Sorry, but there have been no quotes added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
		else
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there are no quotes at the moment.</p>
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
	GetDesc = UCASE(Author & Quote & ItemDate & GetNickName(MemberID) )
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
		<td class="TDHeader">Quote</td>
		<% if intRateQuotes = 1  and intReviewQuotes = 0 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% elseif intRateQuotes = 0  and intReviewQuotes = 1 then %>
			<td class="TDHeader" align=center>Review</td>
		<% elseif intRateQuotes = 1  and intReviewQuotes = 1 then %>
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
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Quotes&Source=quotes.asp'>Change Section Options</a></p>"
		end if
End Sub


'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
%>
	<tr>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
<%		if IsPrivate = 1 and not LoggedMember then %>
			<td class="<% PrintTDMain %>">This is a private quote. &nbsp;<a href="login.asp?Source=quotes_read.asp&ID=<%=ID%>&Submit=Read"><%=PrintTDLink("Click here")%></a> to log in and view it.</td>
<%		else	%>

		<td class="<% PrintTDMain %>"><a href="quotes_read.asp?ID=<%=ID%>"><%=PrintTDLink("&quot;" & Quote & "&quot; - " & Author)%></a></td>
<%		end if
		if intRateQuotes = 1 and intReviewQuotes = 0 then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="quotes_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate")%></a></font></td>
<%		elseif intRateQuotes = 0 and intReviewQuotes = 1 then
			if ReviewsExist( "Quotes", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="quotes_read.asp?ID=<%=ID%>"><%=PrintTDLink("Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="quotes_read.asp?ID=<%=ID%>"><%=PrintTDLink("Add Review")%></a></font></td>
<%			end if
		elseif intRateQuotes = 1 and intReviewQuotes = 1 then
			if ReviewsExist( "Quotes", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="quotes_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate and Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="quotes_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate/Add Review")%></a></font></td>
<%			end if
		end if%>
		<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
	</tr>
<%
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
		<td class="<% PrintTDMain %>"><%=GetNickNameLink(MemberID)%></td>
		<% end if %>	
<%		if IsPrivate = 1 and not LoggedMember then %>
			<td class="<% PrintTDMain %>">This is a private quote. &nbsp;<a href="login.asp?Source=quotes_read.asp&ID=<%=ID%>&Submit=Read">Click here</a> to log in and view it.</td>
<%		else	%>

		<td class="<% PrintTDMain %>"><a href="quotes_read.asp?ID=<%=ID%>">&quot;<%=Quote%>&quot; - <%=Author%></a></td>
<%		end if
		if intRateQuotes = 1 and intReviewQuotes = 0 then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Rate</a></font></td>
<%		elseif intRateQuotes = 0 and intReviewQuotes = 1 then
			if ReviewsExist( "Quotes", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Read/Add Review</a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Add Review</a></font></td>
<%			end if
		elseif intRateQuotes = 1 and intReviewQuotes = 1 then
			if ReviewsExist( "Quotes", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Rate and Read/Add Review</a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Rate/Add Review</a></font></td>
<%			end if
		end if%>
		<% if DisplayPrivacy  then %>
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
%>
		<%=strHeader%>
<%		if IsPrivate = 1 and not LoggedMember then %>
			This is a private quote. &nbsp;<a href="login.asp?Source=quotes_read.asp&ID=<%=ID%>&Submit=Read">Click here</a> to log in and view it.
<%		else	%>
		<a href="quotes_read.asp?ID=<%=ID%>">&quot;<%=Quote%>&quot; - <%=Author%></a>
<%		end if%>
		&nbsp;&nbsp;&nbsp;&nbsp;

		<% if DisplayDate then %>
		<% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayAuthor then %>
		By: <%=GetNickNameLink(MemberID)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayPrivacy and IsPrivate = 1 then Response.Write "Private&nbsp;&nbsp;"
		if intRateQuotes = 1 and intReviewQuotes = 0 then
%>			<%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Rate</a></font>&nbsp;&nbsp;
<%		elseif intRateQuotes = 0 and intReviewQuotes = 1 then
			if ReviewsExist( "Quotes", ID ) then
%>				<font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Add Review</a></font>&nbsp;&nbsp;
<%			end if
		elseif intRateQuotes = 1 and intReviewQuotes = 1 then
			if ReviewsExist( "Quotes", ID ) then
%>				<%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Rate and Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="Quotes_read.asp?ID=<%=ID%>">Rate/Add Review</a></font>&nbsp;&nbsp;
<%			end if
		end if
		Response.Write strFooter
	end if
End Sub
'------------------------End Code-----------------------------
%>
