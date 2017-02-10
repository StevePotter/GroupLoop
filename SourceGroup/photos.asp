<!-- #include file="photos_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludePhotos ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading><%=PhotosTitle%></span><br>
<%
if (IncludeAddButtons = 1 or LoggedMember()) and (LoggedAdmin() or CBool( PhotosMembers )) then
%>
<span class=LinkText><a href="members_photos_add.asp?ID=<%=Request("ID")%>">Add a Photo</a></span>
<%
end if
%>
</p>
<%
'-----------------------Begin Code----------------------------
Table = "PhotoCategories"


Query = "SELECT DisplaySearchPhotos, InfoTextPhotos  FROM Look WHERE CustomerID = " & CustomerID
Set rsList = Server.CreateObject("ADODB.Recordset")
rsList.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	InfoText = rsList("InfoTextPhotos")
	DisplaySearch = CBool(rsList("DisplaySearchPhotos"))
rsList.Close


intCategoryID = Request("ID")
if intCategoryID <> "" then intCategoryID = CInt(intCategoryID)

'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")
if intSearchID <> "" then intSearchID = CInt(intSearchID)

intPhotosPerRow = PhotosPerRow

'Start them off in the first category if none is specified
if intCategoryID = "" then

	'They entered text to search for, so we are going to get matches and put them into the SectionSearch
	if Request("Keywords") <> "" AND intSearchID = "" then
		Query = "SELECT ID, Date, Name, MemberID FROM Photos WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set Name = rsList("Name")
			Set MemberID = rsList("MemberID")

		Query = "SELECT PhotoID, Caption FROM PhotoCaptions WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		Set rsCapList = Server.CreateObject("ADODB.Recordset")
		rsCapList.CacheSize = 100
		rsCapList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set Caption = rsCapList("Caption")

		intSearchID = SingleSearch()
		rsList.Close
		set rsList = Nothing
		rsCapList.Close
		set rsCapList = Nothing
	end if

	if intSearchID = "" then

		if DisplaySearch then
%>
		<form METHOD="POST" ACTION="photos.asp">
			Search For <input type="text" name="Keywords" size="15">
			<input type="submit" name="Submit" value="Go"><br>
		</form>
<%
		end if
		if InfoText <> " " and InfoText <> "" then Response.Write "<p>" & InfoText & "</p>"

		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then Response.Write "<p><a href='admin_photocategories_modify.asp'>Modify Categories</a></p>"

		Set rsPage = Server.CreateObject("ADODB.Recordset")
		PrintCategoryMenu "photos.asp", 0, Table
		Set rsPage = Nothing

		'Give them the link to change the section's properties
		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Photos&Source=photos.asp'>Change Section Options</a></p>"
		end if

	else
		'Their search came up empty
		if intSearchID = 0 then
			if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
				<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
				Try again, or <a href="photos.asp">click here</a> to go back to the category list.</p>
<%
'-----------------------Begin Code----------------------------
			else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="photos.asp">click here</a> to go back to the category list.</p>
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
			<p><a href="photos.asp">Click here</a> to go back to the category list.</p>

			<form METHOD="POST" ACTION="photos.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
			PrintPagesHeader
			PrintTableHeader 100

			Response.Write "<tr>"
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, CategoryID, Name, Ext, Thumbnail, ThumbnailExt FROM Photos WHERE CustomerID = " & CustomerID
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set CategoryID = rsList("CategoryID")
			Set Name = rsList("Name")
			Set Ext = rsList("Ext")
			Set Thumbnail = rsList("Thumbnail")
			Set ThumbnailExt = rsList("ThumbnailExt")

			ChangeTDMain
			for p = 0 to (rsPage.PageSize - 1)
				if not rsPage.EOF then
					if p mod intPhotosPerRow = 0 then
						Response.Write "</tr><tr>"
						ChangeTDMain
					end if

					rsList.Filter = "ID = " & TargetID

					strCommDisp = ""
					intNumCaptions = GetNumCaptions(ID)
					if intNumCaptions = 1 then
						strCommDisp = "<br>1 Caption"
					elseif intNumCaptions > 1 then
						strCommDisp = "<br>" & intNumCaptions & " Captions"
					end if

	'------------------------End Code-----------------------------
	%>
					<td class="<% PrintTDMain %>" align="center" valign="middle">
					<a href="photos.asp?ID=<%=CategoryID%>"><%=GetCategoryName( CategoryID, Table ) %></a><br>
	<%
						if Thumbnail = 1 then
	%>
							<a href="photos_view.asp?ID=<%=ID%>"><img src="photos/<%=ID%>t.<%=ThumbnailExt%>" border=0 alt="<%=Name%>">
	<%
						else
	%>
							<a href="photos_view.asp?ID=<%=ID%>"><%=PrintTDLink(Name)%>
	<%
						end if
	%>
						<%=PrintTDLink(strCommDisp)%></a>
					</td>
	<%
	'-----------------------Begin Code----------------------------
					rsPage.MoveNext
				else
					exit for
				end if
			next

			'Print the empty cells if we've had more than one row
			if p > intPhotosPerRow then
				do until p mod intPhotosPerRow = 0
	%>
						<td class="<% PrintTDMain %>" align="center" valign="middle">
							&nbsp;
						</td>
	<%
					p = p + 1
				loop
			end if

			Response.Write("</tr></table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing

		end if
	end if

'They have a category selected
else
	if not ValidCategory(intCategoryID, Table) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but that is not a valid category."))

	GetCategoryInfo intCategoryID, strName, blPrivate, strBody

	Response.Write "<p>" & GetCatHeiarchy(intCategoryID, "photos.asp", Table, PhotosTitle) & "</p>"

	if blPrivate AND not LoggedMember then Redirect( "login.asp?Source=photos.asp&ID=" & intCategoryID & "&Submit=Go" )

	'Keep track of shit
	IncrementHits intCategoryID, Table


'	<form METHOD="POST" ACTION="photos.asp">
'		<input type="hidden" name="ID" value="< %=intCategoryID% >">
'		Search < %=strName% > For <input type="text" name="Keywords" size="15">
'		<input type="submit" name="Submit" value="Go"><br>
'	</form>

'------------------------End Code-----------------------------
%>


	<table width="100%">
		<tr>
			<td align="left">
				<span class="Heading">Category: <%=strName%></span>
			</td>
<%			if NeedCategoryMenu(Table) then %>

			<td align="right">
				<form action="photos.asp" method="post">
					<font size="-1">Change Category To:</font><br>
					<% PrintCategoryPullDown intCategoryID, 1, 1, 0, 1, Table, "ID", "" %>
					<input type="Submit" value="Switch">
				</form>
			</td>
<%			end if %>
		</tr>
	</table>

<%
'-----------------------Begin Code----------------------------

	if not IsNull(strBody) then
		if strBody <> "" then
%>
		<p><%=strBody%></p>
<%	
		end if
	end if

	if CategoryHasChild( intCategoryID, Table ) then
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		PrintCategoryMenu "photos.asp", intCategoryID, Table
		Set rsPage = Nothing
	end if

	'They entered text to search for, so we are going to get matches and put them into the SectionSearch
	if Request("Keywords") <> "" AND intSearchID = "" then
		Query = "SELECT ID, Date, Name, MemberID FROM Photos WHERE (CategoryID = " & intCategoryID & " AND CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set Name = rsList("Name")
			Set MemberID = rsList("MemberID")

		Query = "SELECT PhotoID, Caption FROM PhotoCaptions WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		Set rsCapList = Server.CreateObject("ADODB.Recordset")
		rsCapList.CacheSize = 100
		rsCapList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set Caption = rsCapList("Caption")

		intSearchID = SingleSearch()
		rsList.Close
		set rsList = Nothing
		rsCapList.Close
		set rsCapList = Nothing
	end if

	if intSearchID <> "" then
		'Their search came up empty
		if intSearchID = 0 then
			if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
				<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
				Try again, or <a href="photos.asp?ID=<%=intCategoryID%>">click here</a> to go back to the photos in <%=strName%>.</p>
<%
'-----------------------Begin Code----------------------------
			else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="photos.asp?ID=<%=intCategoryID%>">click here</a> to go back to the photos in <%=strName%>.</p>
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
			<p><a href="photos.asp?ID=<%=intCategoryID%>">Click here</a> to go back to the photos in <%=strName%>.</p>

			<form METHOD="POST" ACTION="photos.asp">
			<input type="hidden" name="ID" value="<%=intCategoryID%>">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
			PrintPagesHeader
			PrintTableHeader 100

			Response.Write "<tr>"
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Name, Ext, Thumbnail, ThumbnailExt FROM Photos WHERE CategoryID = " & intCategoryID & " AND CustomerID = " & CustomerID
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set Name = rsList("Name")
			Set Ext = rsList("Ext")
			Set Thumbnail = rsList("Thumbnail")
			Set ThumbnailExt = rsList("ThumbnailExt")

			ChangeTDMain
			for p = 0 to (rsPage.PageSize - 1)
				if not rsPage.EOF then
					if p mod intPhotosPerRow = 0 then
						Response.Write "</tr><tr>"
						ChangeTDMain
					end if

					rsList.Filter = "ID = " & TargetID

					strCommDisp = ""
					intNumCaptions = GetNumCaptions(ID)
					if intNumCaptions = 1 then
						strCommDisp = "<br>1 Caption"
					elseif intNumCaptions > 1 then
						strCommDisp = "<br>" & intNumCaptions & " Captions"
					end if

	'------------------------End Code-----------------------------
	%>
					<td class="<% PrintTDMain %>" align="center" valign="middle">
	<%
						if Thumbnail = 1 then
	%>
							<a href="photos_view.asp?ID=<%=ID%>"><img src="photos/<%=ID%>t.<%=ThumbnailExt%>" border=0 alt="<%=Name%>">
	<%
						else
	%>
							<a href="photos_view.asp?ID=<%=ID%>"><%=PrintTDLink(Name)%>
	<%
						end if
	%>
						<%=PrintTDLink(strCommDisp)%></a>
					</td>
	<%
	'-----------------------Begin Code----------------------------
					rsPage.MoveNext
				else
					exit for
				end if
			next

			'Print the empty cells if we've had more than one row
			if p > intPhotosPerRow then
				do until p mod intPhotosPerRow = 0
	%>
						<td class="<% PrintTDMain %>" align="center" valign="middle">
							&nbsp;
						</td>
	<%
					p = p + 1
				loop
			end if

			Response.Write("</tr></table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing

		end if

	else
		Query = "SELECT ID, Name, Ext, Thumbnail, ThumbnailExt FROM Photos WHERE CustomerID = " & CustomerID & " AND CategoryID = " & intCategoryID & " ORDER BY Date DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsPage("ID")
			Set Name = rsPage("Name")
			Set Ext = rsPage("Ext")
			Set Thumbnail = rsPage("Thumbnail")
			Set ThumbnailExt = rsPage("ThumbnailExt")

		'Don't navigate if it's empty
		if not rsPage.EOF then
	%>
			<form METHOD="POST" ACTION="photos.asp">
			<input type="hidden" name="ID" value="<%=intCategoryID%>">
	<%
			PrintPagesHeader
			PrintTableHeader 100

			Response.Write "<tr>"

			ChangeTDMain
			for p = 0 to (rsPage.PageSize - 1)
				if not rsPage.EOF then
					if p mod intPhotosPerRow = 0 then
						Response.Write "</tr><tr>"
						ChangeTDMain
					end if

					strCommDisp = ""
					intNumCaptions = GetNumCaptions(ID)
					if intNumCaptions = 1 then
						strCommDisp = "<br>1 Caption"
					elseif intNumCaptions > 1 then
						strCommDisp = "<br>" & intNumCaptions & " Captions"
					end if

	%>
					<td class="<% PrintTDMain %>" align="center" valign="middle">
	<%
						if Thumbnail = 1 then
	%>
							<a href="photos_view.asp?ID=<%=ID%>"><img src="photos/<%=ID%>t.<%=ThumbnailExt%>" border=0 alt="<%=Name%>">
	<%
						else
	%>
							<a href="photos_view.asp?ID=<%=ID%>"><%=PrintTDLink(Name)%>
	<%
						end if
	%>
						<%=PrintTDLink(strCommDisp)%></a>
					</td>
	<%
					rsPage.MoveNext
				else
					exit for
				end if
			next

			'Print the empty cells if we've had more than one row
			if p > intPhotosPerRow then
				do until p mod intPhotosPerRow = 0
	%>
						<td class="<% PrintTDMain %>" align="center" valign="middle">
							&nbsp;
						</td>
	<%
					p = p + 1
				loop
			end if

			Response.Write("</tr></table>")
			rsPage.Close
		else
			Response.Write "<p>Sorry, but there are no photos in this category.</p>"
		end if

		'Give them the link to change the section's properties
		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Photos&Source=photos.asp'>Change Section Options</a></p>"
		end if
		set rsPage = Nothing
	end if
end if

'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	rsCapList.Filter = "PhotoID = " & ID
	strCaps = ""
	do until rsCapList.EOF
		strCaps = strCaps & Caption
		rsCapList.MoveNext
	loop

	GetDesc = UCASE( strCaps & ItemDate & Name & GetNickName(MemberID) )
End Function


%>
