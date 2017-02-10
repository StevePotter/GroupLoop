<!-- #include file="media_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeMedia ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading><%=MediaTitle%></span><br>
<%
if (IncludeAddButtons = 1 or LoggedMember()) and (LoggedAdmin() or CBool( MediaMembers )) then
%>
<span class=LinkText><a href="members_media_add.asp?ID=<%=Request("ID")%>">Add a File</a></span>
<%
end if
%>
</p>



<%
'-----------------------Begin Code----------------------------
intCategoryID = Request("ID")
if intCategoryID <> "" then intCategoryID = CInt(intCategoryID)

'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")
if intSearchID <> "" then intSearchID = CInt(intSearchID)

strPath = GetPath("media")

'Check for a default category
if intCategoryID = "" then intCategoryID = GetDefaultCat()

'Start them off in the first category if none is specified
if intCategoryID = 0 then

	'They entered text to search for, so we are going to get matches and put them into the SectionSearch
	if Request("Keywords") <> "" AND intSearchID = "" then
		Query = "SELECT ID, Date, FileName, Description, MemberID FROM Media WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set FileName = rsList("FileName")
			Set Description = rsList("Description")
			Set MemberID = rsList("MemberID")
		intSearchID = SingleSearch()
		rsList.Close
		set rsList = Nothing
	end if

	if intSearchID = "" then
%>
		<form METHOD="POST" ACTION="media.asp">
			Search For <input type="text" name="Keywords" size="15">
			<input type="submit" name="Submit" value="Go"><br>
		</form>
<%
		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then Response.Write "<p><a href='admin_mediacategories_modify.asp'>Modify Categories</a></p>"

		Set rsPage = Server.CreateObject("ADODB.Recordset")
		PrintCategoryMenu "media.asp"
		Set rsPage = Nothing
	else
		'Their search came up empty
		if intSearchID = 0 then
			if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
				<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
				Try again, or <a href="media.asp">click here</a> to go back to the category list.</p>
<%
'-----------------------Begin Code----------------------------
			else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="media.asp">click here</a> to go back to the category list.</p>
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
			<form METHOD="POST" ACTION="media.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
	<%
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle

			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Date, MemberID, FileName, Description, TotalRating, TimesRated FROM Media WHERE CustomerID = " & CustomerID
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set MemberID = rsList("MemberID")
			Set TotalRating = rsList("TotalRating")
			Set TimesRated = rsList("TimesRated")
			Set FileName = rsList("FileName")
			Set Description = rsList("Description")

			Set FileSystem = CreateObject("Scripting.FileSystemObject")

			for p = 1 to rsPage.PageSize
				if not rsPage.EOF then
					rsList.Filter = "ID = " & TargetID

					PrintTableData

					rsPage.MoveNext
				end if
			next
			Response.Write("</table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing

			Set FileSystem = Nothing
		end if
	end if

'They have a category selected
else
	if not ValidCategory(intCategoryID) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but that is not a valid category."))

	GetCategoryInfo intCategoryID, strName, blPrivate

	if blPrivate AND not LoggedMember then Redirect( "login.asp?Source=media.asp&ID=" & intCategoryID & "&Submit=Go" )

	'Keep track of shit
	IncrementHits intCategoryID, "MediaCategories"

	'------------------------End Code-----------------------------
	%>

	<form METHOD="POST" ACTION="media.asp">
		<input type="hidden" name="ID" value="<%=intCategoryID%>">
		Search <%=strName%> For <input type="text" name="Keywords" size="15">
		<input type="submit" name="Submit" value="Go"><br>
	</form>

	<table width="100%">
		<tr>
			<td align="left">
				<span class="Heading">Category: <%=strName%></span>
			</td>

<%			if NeedCategoryMenu("MediaCategories") then	%>
			<td align="right">
				<font size="-1">Change Category To:</font><br>
				<form action="media.asp" method="post">
					<% PrintCategoryPullDown intCategoryID, 0, 0 %>
					<input type="Submit" value="Switch">
				</form>
			</td>
<%			end if %>
		</tr>
	</table>

<%
'-----------------------Begin Code----------------------------
	'They entered text to search for, so we are going to get matches and put them into the SectionSearch
	if Request("Keywords") <> "" AND intSearchID = "" then
		Query = "SELECT ID, Date, FileName, Description, MemberID FROM Media WHERE (CategoryID = " & intCategoryID & " AND CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set FileName = rsList("FileName")
			Set Description = rsList("Description")
			Set MemberID = rsList("MemberID")
		intSearchID = SingleSearch()
		rsList.Close
		set rsList = Nothing
	end if

	if intSearchID <> "" then
		'Their search came up empty
		if intSearchID = 0 then
			if Session("MemberID") <> "" then
	'-----------------------End Code----------------------------
	%>
				<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
				Try again, or <a href="media.asp?ID=<%=intCategoryID%>">click here</a> to view all files in this category.</p>
	<%
	'-----------------------Begin Code----------------------------
			else
	'-----------------------End Code----------------------------
	%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="media.asp?ID=<%=intCategoryID%>">click here</a> to view all files in this category.</p>
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
			<form METHOD="POST" ACTION="media.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
			<input type="hidden" name="ID" value="<%=intCategoryID%>">
	<%
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle

			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Date, MemberID, FileName, Description, TotalRating, TimesRated FROM Media WHERE CategoryID = " & intCategoryID & " AND CustomerID = " & CustomerID
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set MemberID = rsList("MemberID")
			Set TotalRating = rsList("TotalRating")
			Set TimesRated = rsList("TimesRated")
			Set FileName = rsList("FileName")
			Set Description = rsList("Description")

			Set FileSystem = CreateObject("Scripting.FileSystemObject")

			for p = 1 to rsPage.PageSize
				if not rsPage.EOF then
					rsList.Filter = "ID = " & TargetID

					PrintTableData

					rsPage.MoveNext
				end if
			next
			Response.Write("</table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing

			Set FileSystem = Nothing
		end if

	'They are just cycling through the Media.  No searching.
	else
		'Instantiate the recordset for the output
		Query = "SELECT ID, Date, MemberID, FileName, Description, TotalRating, TimesRated FROM Media WHERE CategoryID = " & intCategoryID & " AND CustomerID = " & CustomerID & " ORDER BY Date DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		'Don't navigate if it's empty
		if not rsPage.EOF then
%>
			<form METHOD="POST" ACTION="media.asp">
			<input type="hidden" name="ID" value="<%=intCategoryID%>">
<%
			Set ID = rsPage("ID")
			Set ItemDate = rsPage("Date")
			Set MemberID = rsPage("MemberID")
			Set TotalRating = rsPage("TotalRating")
			Set TimesRated = rsPage("TimesRated")
			Set FileName = rsPage("FileName")
			Set Description = rsPage("Description")

			PrintPagesHeader
			PrintTableHeader 100
			PrintTableTitle

			strPath = GetPath ("media")
			Set FileSystem = CreateObject("Scripting.FileSystemObject")

			for j = 1 to rsPage.PageSize
				if not rsPage.EOF then
					PrintTableData

					rsPage.MoveNext
				end if
			next

			Response.Write("</table>")



			Set FileSystem = Nothing
			set rsPage = Nothing
		else
			'If there are no available Media
			Response.Write "<p>Sorry, but there are no files in this category.</p>"
		end if

		'Give them the link to change the section's properties
		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Media&Source=media.asp'>Change Section Options</a></p>"
		end if

		set rsPage = Nothing
	end if

end if

'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(FileName & Description )
End Function


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<% if IncludeDate = 1 then %>
		<td class="TDHeader" align=center>Date</td>
		<% end if %>	
		<% if IncludeAuthor = 1 then %>
		<td class="TDHeader">Author</td>
		<% end if %>	
		<td class="TDHeader">File</td>
		<td class="TDHeader">Description</td>
		<% if RateMedia = 1  and ReviewMedia = 0 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% elseif RateMedia = 0  and ReviewMedia = 1 then %>
			<td class="TDHeader" align=center>Review</td>
		<% elseif RateMedia = 1  and ReviewMedia = 1 then %>
			<td class="TDHeader" align=center>Rating</td>
		<% end if %>	
	</tr>
<%
End Sub


'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
	strFileName = strPath & "/" & FileName
	if FileSystem.FileExists (strFileName) then
		Set TestFile = FileSystem.GetFile( strFileName )
		dblSize = Round((TestFile.Size / 1000000), 2 )
		strLink = "<a href='media/" & FileName & "'>" & FileName & "</a> &nbsp;<font size=-2>(" & dblSize & " Megs)</font>"
		Set TestFile = Nothing
	else
		strLink = "File Does Not Exist"
	end if
	'------------------------End Code-----------------------------
	%>
	<tr>
		<% if IncludeDate = 1 then %>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<% end if %>	
		<% if IncludeAuthor = 1 then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
		<% end if %>	
		<td class="<% PrintTDMain %>"><%=strLink%></td>
		<td class="<% PrintTDMain %>"><a href="media_read.asp?ID=<%=ID%>"><%=PrintTDLink(PrintStart(Description))%></a></td>
<%		if RateMedia = 1 and ReviewMedia = 0 then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="media_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate")%></a></font></td>
<%		elseif RateMedia = 0 and ReviewMedia = 1 then
			if ReviewsExist( "Media", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="media_read.asp?ID=<%=ID%>"><%=PrintTDLink("Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="media_read.asp?ID=<%=ID%>"><%=PrintTDLink("Add Review")%></a></font></td>
<%			end if
		elseif RateMedia = 1 and ReviewMedia = 1 then
			if ReviewsExist( "Media", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="media_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate and Read/Add Review")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="media_read.asp?ID=<%=ID%>"><%=PrintTDLink("Rate/Add Review")%></a></font></td>
<%			end if
		end if%>
	</tr>
<%
End Sub
'------------------------End Code-----------------------------
%>
