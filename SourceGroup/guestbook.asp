<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeGuestbook ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading><%=GuestbookTitle%></span><br>
<span class=LinkText><a href="guestbook_post.asp">Add An Entry</a></span></p>

<%
'-----------------------Begin Code----------------------------
'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")

intRateGuestbook = RateGuestbook
intReviewGuestbook = ReviewGuestbook

Set rsList = Server.CreateObject("ADODB.Recordset")


Public ListType, DisplayDate, DisplayAuthor, DisplayPrivacy, blBulletImg, ItemNumber
	strImagePath = GetPath("images")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	ItemNumber = 0	'This will be set by the PrintPagesHeader sub

Query = "SELECT DisplaySearchGuestbook, DisplayDaysOldGuestbook, InfoTextGuestbook FROM Look WHERE CustomerID = " & CustomerID
rsList.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplaySearch = CBool(rsList("DisplaySearchGuestbook"))
	DisplayDaysOld = CBool(rsList("DisplayDaysOldGuestbook"))
	InfoText = rsList("InfoTextGuestbook")
rsList.Close

if DisplaySearch or DisplayDaysOld then
%>
	<form METHOD="POST" ACTION="guestbook.asp">
<%	if DisplayDaysOld then	%>
	View Entries In The Last <% PrintDaysOld %>
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

'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if Request("Keywords") <> "" then
	Query = "SELECT ID, Date, Email, Author, Body FROM Guestbook WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	rsList.CacheSize = 100
	rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set Email = rsList("Email")
		Set Author = rsList("Author")
		Set Body = rsList("Body")
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
			Try again, or <a href="guestbook.asp">click here</a> to view all entries.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			Try again, or <a href="guestbook.asp">click here</a> to view all entries.</p>
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
		<form METHOD="POST" ACTION="guestbook.asp">
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
		PrintPagesHeader
		PrintTableHeader 0
		PrintTableTitle

		'Instantiate the recordset for the output
		Set rsList = Server.CreateObject("ADODB.Recordset")
		Query = "SELECT ID, Date, Email, Author, Body, TotalRating, TimesRated FROM Guestbook WHERE CustomerID = " & CustomerID
		rsList.CacheSize = PageSize
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set ID = rsList("ID")
		Set ItemDate = rsList("Date")
		Set Email = rsList("Email")
		Set Author = rsList("Author")
		Set Body = rsList("Body")
		Set TotalRating = rsList("TotalRating")
		Set TimesRated = rsList("TimesRated")

		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				rsList.Filter = "ID = " & rsPage("TargetID")

				PrintTableData

				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
		set rsPage = Nothing
		set rsList = Nothing
	end if
'They are just cycling through the guestbook.  No searching.
else
	if InfoText <> " " and InfoText <> "" then Response.Write "<p>" & InfoText & "</p>"
	'This is if they requested guestbook written in a time period
	if Request("DaysOld") <> "" then
		CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
		Query = "SELECT ID, Date, Email, Author, Body, TotalRating, TimesRated FROM Guestbook WHERE (CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	else
		Query = "SELECT ID, Date, Email, Author, Body, TotalRating, TimesRated FROM Guestbook WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	end if
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then
%>
		<form METHOD="POST" ACTION="guestbook.asp">
		<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Email = rsPage("Email")
		Set Author = rsPage("Author")
		Set Body = rsPage("Body")
		Set TotalRating = rsPage("TotalRating")
		Set TimesRated = rsPage("TimesRated")

		PrintPagesHeader
		PrintTableHeader 0
		PrintTableTitle
		for j = 1 to rsPage.PageSize
			if not rsPage.EOF then
				PrintTableData
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		'Give them the link to change the section's properties
		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Guestbook&Source=guestbook.asp'>Change Section Options</a></p>"
		end if
	else
		if Request("DaysOld") <> "" then
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there have been no entries added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
		else
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there are no entries at the moment.</p>
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
	GetDesc = UCASE(Email & Author & Body )
End Function

'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">Author</td>
		<td class="TDHeader">Entry</td>
		<% if intRateGuestbook = 1  and intReviewGuestbook = 0 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% elseif intRateGuestbook = 0  and intReviewGuestbook = 1 then %>
			<td class="TDHeader" align=center>Review</td>
		<% elseif intRateGuestbook = 1  and intReviewGuestbook = 1 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% end if %>	
	</tr>
<%
End Sub

'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
	if InStr( Email, "@" ) then
		strAuthor = "<a href='mailto:" & Email & "'>" & Author & "</a>"
	else
		strAuthor = Author
	end if
%>
	<tr>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<td class="<% PrintTDMain %>"><%=strAuthor%></td>
		<td class="<% PrintTDMain %>"><%=Body%></td>
<%		if intRateGuestbook = 1 and intReviewGuestbook = 0 then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="guestbook_read.asp?ID=<%=ID%>">Rate</a></font></td>
<%		elseif intRateGuestbook = 0 and intReviewGuestbook = 1 then
			if ReviewsExist( "Guestbook", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="guestbook_read.asp?ID=<%=ID%>">Read/Add Review</a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="guestbook_read.asp?ID=<%=ID%>">Add Review</a></font></td>
<%			end if
		elseif intRateGuestbook = 1 and intReviewGuestbook = 1 then
			if ReviewsExist( "Guestbook", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="guestbook_read.asp?ID=<%=ID%>">Rate and Read/Add Review</a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="guestbook_read.asp?ID=<%=ID%>">Rate/Add Review</a></font></td>
<%			end if
		end if%>
	</tr>
<%
End Sub
'------------------------End Code-----------------------------
%>
