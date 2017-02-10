<!-- #include file="photos_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not CatsExist then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but no photos can be modified until the administrator creates a category."))
if not CBool( IncludePhotos ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_photos_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Photos</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
blLoggedAdmin = LoggedAdmin

if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

strSubmit = Request("Submit")

Table = "PhotoCategories"

'Deleting a file
if strSubmit = "Delete" and Request("PhotoID") <> "" then
	intID = CInt(Request("PhotoID"))

	Query = "SELECT Ext, Thumbnail, ThumbnailExt FROM Photos WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	strPath = GetPath ("photos")
	strFileName = strPath & "/" & intID & "." & rsUpdate("Ext")
	strThumbName =  strPath & "/" & intID & "t." & rsUpdate("ThumbnailExt")

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE PhotoCaptions WHERE PhotoID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	'Delete the files
	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	if FileSystem.FileExists(strFileName) then FileSystem.DeleteFile(strFileName)
	if FileSystem.FileExists(strThumbName) then FileSystem.DeleteFile(strThumbName)

	Set FileSystem = Nothing
'------------------------End Code-----------------------------
%>
	<p>The photo has been deleted. &nbsp;<a href="members_photos_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

'Editing a photo
elseif strSubmit = "Edit" and Request("PhotoID") <> "" then
	intID = CInt(Request("PhotoID"))

	Query = "SELECT Date, CategoryID, Name, Ext  FROM Photos WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
	<%	if blLoggedAdmin then	%>
			if (form.Date.value == ""){
				strError += "Sorry, but you forgot to enter a date. \n";
				alert (strError);
				return false;
			}
	<%	end if	%>
			if(form.PhotoFile.value == "") {
				return true;
			}
			else{
				alert ('Uploading your file may take some time, so please be patient and dont constantly click the Update button, because that wont speed anything up.');
				return true;
			}
		}

	//-->
	</SCRIPT>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

<%	if blLoggedAdmin then	%>
	* indicates required information<br>
<%	end if	%>

	<form enctype="multipart/form-data" method="post" action="<%=SecurePath%>members_photos_modify_process.asp" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted" name="MyForm">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="PhotoID" value="<%=intID%>">
<%	PrintTableHeader 0
	if blLoggedAdmin then	%>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
<%	end if	%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Category</td>
		<td class="<% PrintTDMain %>"> 
			<% PrintCategoryPullDown rsEdit("CategoryID"), 1, 0, 0, 1, "PhotoCategories", "ID", "" %>
     	</td>
   	</tr>

	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Description of Photo</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Name" size="55" value="<%=FormatEdit( rsEdit("Name") )%>">
     	</td>
    </tr>
<%
		if LCase(rsEdit("Ext")) = "jpg" then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				Edit photo.  Click the button here to resize, rotate, sharpen, etc. your photo.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="button" value="Edit Photo" onClick="Redirect('members_images_modify.asp?Path=photos&FileName=<%=intID%>.jpg&Thumbnail=true')" >
			</td>
		</tr>
<%
		end if
%>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Photo File.  If you leave this blank, the original photo will be kept.  Otherwise, the new photo will erase the old one.
		</td>
		<td class="<% PrintTDMain %>">
			<input type="file" name="PhotoFile">
		</td>
	</tr>

	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
    </tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing

'Updating a caption
elseif strSubmit = "Update" and Request("CaptionID") <> "" then
	intID = CInt(Request("CaptionID"))

	if (blLoggedAdmin and Request("Date") = "") or Request("Caption") = "" then Redirect("incomplete.asp")

	Query = "SELECT Private, Caption, Date, IP, ModifiedID, PhotoID FROM PhotoCaptions WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	intPhotoID = rsUpdate("PhotoID")

	if Request("Private") = "1" then 
		rsUpdate("Private") = 1
	else
		rsUpdate("Private") = 0
	end if
	if blLoggedAdmin then rsUpdate("Date") = Request("Date")
	rsUpdate("Caption") = GetTextArea( Request("Caption") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The caption has been edited. &nbsp;<a href="members_photos_modify.asp?Submit=Modify+Captions&PhotoID=<%=intPhotoID%>">Click here</a> to modify another caption for this photo.<br>
		<a href="members_photos_modify.asp">Click here</a> to modify another photo.<br>
		<a href="photos_view.asp?ID=<%=intPhotoID%>">Click here</a> to view the photo with the edited caption.
	</p>
<%
'-----------------------Begin Code----------------------------
'Delete a caption
elseif strSubmit = "Delete" and Request("CaptionID") <> "" then
	intID = CInt(Request("CaptionID"))

	Query = "SELECT PhotoID FROM PhotoCaptions WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	intPhotoID = rsUpdate("PhotoID")

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'PhotoCaptions' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The caption has been deleted. &nbsp;
<%	if PhotoCaptionsExist( intPhotoID ) then %>	
		<a href="members_photos_modify.asp?Submit=Modify+Captions&PhotoID=<%=intPhotoID%>">Click here</a> to modify another caption for this photo.<br>
<%	end if	%>
		<a href="members_photos_modify.asp">Click here</a> to modify another photo.<br>
		<a href="photos_view.asp?ID=<%=intPhotoID%>">Click here</a> to view the photo with the deleted caption.
	</p>
<%
'Edit a caption
elseif strSubmit = "Edit" and Request("CaptionID") <> "" then
	intID = CInt(Request("CaptionID"))

	Query = "SELECT Private, Caption, Date FROM PhotoCaptions WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsEdit.EOF then
		set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if rsEdit("Private") = 1 then 
		strChecked = "checked"
	else
		strChecked = ""
	end if
'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	* indicates required information<br>

	<form method="post" action="<%=SecurePath%>members_photos_modify.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="CaptionID" value="<%=intID%>">
	<%PrintTableHeader 0%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>
<%	if blLoggedAdmin then	%>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
<%	end if	%>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">* Caption</td>
    	<td class="<% PrintTDMain %>"> 
			<% TextArea "Caption", 55, 4, True, rsEdit("Caption") %>
    	</td>
    </tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
    </tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing

'Listing the captions
elseif strSubmit = "Modify Captions" and Request("PhotoID") <> "" then
	intID = CInt(Request("PhotoID"))

	'They have search results, so lets list their results
	Query = "SELECT Name, Thumbnail, ThumbnailExt FROM Photos WHERE ID = " & intID
	Set rsPhoto = Server.CreateObject("ADODB.Recordset")
	rsPhoto.Open Query, Connect, adOpenStatic, adLockReadOnly
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>
<%
	if rsPhoto("Thumbnail") = 1 then	%>
		<p align=center><a href="photos_view.asp?ID=<%=intID%>"><img src="photos/<%=intID%>t.<%=rsPhoto("ThumbnailExt")%>" border=0 alt="<%=rsPhoto("Name")%>"></a></p>
<%	else	%>
		<p align=center>Photo description: <a href="photos_view.asp?ID=<%=intID%>"><%=rsPhoto("Name")%></a></p>
<%	end if
	rsPhoto.Close
	set rsPhoto = Nothing
	PrintTableHeader 0
%>
	<tr>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">Author</td>
		<td class="TDHeader">Caption</td>
		<td class="TDHeader">Public?</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
	Query = "SELECT ID, Date, MemberID, Caption, Private FROM PhotoCaptions WHERE PhotoID = " & intID & " AND " & strMatch & " ORDER BY ID DESC"
	Set rsList = Server.CreateObject("ADODB.Recordset")
	rsList.CacheSize = 20
	rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
	do until rsList.EOF

'------------------------End Code-----------------------------
%>
		<form METHOD="POST" ACTION="members_photos_modify.asp">
		<input type="hidden" name="CaptionID" value="<%=rsList("ID")%>">
		<tr>
			<td class="<% PrintTDMain %>"><%=FormatDateTime(rsList("Date"), 2)%></td>
			<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink( rsList("MemberID") ))%></td>
			<td class="<% PrintTDMain %>"><%=rsList("Caption")%></td>
			<td class="<% PrintTDMain %>"><%=PrintPublic(rsList("Private"))%> 
			<input type="submit" name="Submit" value="Edit">
			<td class="<% PrintTDMainSwitch %>"><input type="button" value="Delete" onClick="DeleteBox('If you delete this caption, there is no way to get it back.  Are you sure?', 'members_photos_modify.asp?Submit=Delete&CaptionID=<%=rsList("ID")%>')">
			<%if ReviewsExist( "PhotoCaptions", rsList("ID") ) AND blLoggedAdmin then%>
				<input type="button" value="Modify Reviews" onClick="Redirect('admin_reviews_modify.asp?Source=members_photos_modify.asp&TargetTable=PhotoCaptions&TargetID=<%=rsList("ID")%>')">
			<%end if%>	
			</td>

		</tr>
		</form>
<%
'-----------------------Begin Code----------------------------
		rsList.MoveNext
	loop
	Response.Write("</table>")

	rsList.Close
	set rsList = Nothing

'Listing the photos
'this is tough because members may have captions in someone else's photos
'so we just can't list their photos, we have list every one that has one of their captions too
else
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
			Set rsList = Server.CreateObject("ADODB.Recordset")
			rsList.CacheSize = 100
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
				Set ID = rsList("ID")
				Set ItemDate = rsList("Date")
				Set Name = rsList("Name")
				Set MemberID = rsList("MemberID")

			Query = "SELECT PhotoID, Caption FROM PhotoCaptions WHERE (" & strMatch & ") ORDER BY Date DESC"
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
	%>
			<form METHOD="POST" ACTION="members_photos_modify.asp">
				Search For <input type="text" name="Keywords" size="15">
				<input type="submit" name="Submit" value="Go"><br>
			</form>
	<%
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintCategoryMenu "members_photos_modify.asp", 0, Table
			Set rsPage = Nothing
		else
			'Their search came up empty
			if intSearchID = 0 then
				if Session("MemberID") <> "" then
	'-----------------------End Code----------------------------
	%>
					<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
					Try again, or <a href="members_photos_modify.asp">click here</a> to go back to the category list.</p>
	<%
	'-----------------------Begin Code----------------------------
				else
	'-----------------------End Code----------------------------
	%>
					<p>Sorry, but your search came up empty.<br>
					Try again, or <a href="members_photos_modify.asp">click here</a> to go back to the category list.</p>
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

				Set rsList = Server.CreateObject("ADODB.Recordset")
				Query = "SELECT ID, Name, Ext, Thumbnail, ThumbnailExt, MemberID FROM Photos WHERE CustomerID = " & CustomerID
				rsList.CacheSize = PageSize
				rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

				Set ID = rsList("ID")
				Set Name = rsList("Name")
				Set Ext = rsList("Ext")
				Set Thumbnail = rsList("Thumbnail")
				Set ThumbnailExt = rsList("ThumbnailExt")
				Set MemberID = rsList("MemberID")

				if blLoggedAdmin then
		%>
					<p><a href="members_photos_modify.asp">Click here</a> to go back to the category listing.</p>
					<form METHOD="POST" ACTION="members_photos_modify.asp">
					<input type="hidden" name="SearchID" value="<%=intSearchID%>">
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
							rsList.Filter = "ID = " & TargetID
							PrintTableData
							rsPage.MoveNext
						else
							exit for
						end if
					next
				else
					%><p><a href="members_photos_modify.asp">Click here</a> to go back to the category listing.</p><%
					p = 0
					PrintTableHeader 100
					Response.Write "<tr>"
					do until rsPage.EOF
						if MemberID = Session("MemberID") or PhotoCaptionsMemberExist( TargetID, Session("MemberID") ) then
							if p mod intPhotosPerRow = 0 then
								Response.Write "</tr><tr>"
								ChangeTDMain
							end if
							rsList.Filter = "ID = " & TargetID
							PrintTableData
							p = p + 1
						end if
						rsPage.MoveNext
					loop

					if p = 0 then Response.Write "<td class=BodyText>Sorry, but there is nothing for you to edit in this category.</td>"


				end if

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

		if blPrivate AND not LoggedMember then Redirect( "login.asp?Source=members_photos_modify.asp&ID=" & intCategoryID & "&Submit=Go" )

		'Keep track of shit
		IncrementHits intCategoryID, "PhotoCategories"

	'------------------------End Code-----------------------------
	%>
		<form METHOD="POST" ACTION="members_photos_modify.asp">
			<input type="hidden" name="ID" value="<%=intCategoryID%>">
			Search <%=strName%> For <input type="text" name="Keywords" size="15">
			<input type="submit" name="Submit" value="Go"><br>
		</form>
		<table width="100%">
		<tr>
			<td align="left">
				<span class="Heading">Category: <%=strName%></span>
			</td>
<%			if NeedCategoryMenu(Table) then %>

			<td align="right">
				<form action="members_photos_modify.asp" method="post">
					<font size="-1">Change Category To:</font><br>
					<% PrintCategoryPullDown intCategoryID, 0, 1, 0, 1, Table, "ID", "" %>
					<input type="Submit" value="Switch">
				</form>
			</td>
<%			end if %>
		</tr>
		</table>

	<%
	'-----------------------Begin Code----------------------------

		if CategoryHasChild( intCategoryID, Table ) then
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintCategoryMenu "members_photos_modify.asp", intCategoryID, Table
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

			Query = "SELECT PhotoID, Caption FROM PhotoCaptions WHERE (" & strMatch & ") ORDER BY Date DESC"
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
					Try again, or <a href="members_photos_modify.asp?ID=<%=intCategoryID%>">click here</a> to go back to the photos in <%=strName%>.</p>
	<%
	'-----------------------Begin Code----------------------------
				else
	'-----------------------End Code----------------------------
	%>
					<p>Sorry, but your search came up empty.<br>
					Try again, or <a href="members_photos_modify.asp?ID=<%=intCategoryID%>">click here</a> to go back to the photos in <%=strName%>.</p>
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

				Set rsList = Server.CreateObject("ADODB.Recordset")
				Query = "SELECT ID, Name, Ext, Thumbnail, ThumbnailExt, MemberID FROM Photos WHERE CategoryID = " & intCategoryID & " AND CustomerID = " & CustomerID
				rsList.CacheSize = PageSize
				rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

				Set ID = rsList("ID")
				Set Name = rsList("Name")
				Set Ext = rsList("Ext")
				Set Thumbnail = rsList("Thumbnail")
				Set ThumbnailExt = rsList("ThumbnailExt")
				Set MemberID = rsList("MemberID")

				if blLoggedAdmin then
		%>
					<p><a href="members_photos_modify.asp?ID=<%=intCategoryID%>">Click here</a> to go back to the photos in <%=strName%>.</p>

					<form METHOD="POST" ACTION="members_photos_modify.asp">
					<input type="hidden" name="ID" value="<%=intCategoryID%>">
					<input type="hidden" name="SearchID" value="<%=intSearchID%>">
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
							rsList.Filter = "ID = " & TargetID
							PrintTableData
							rsPage.MoveNext
						else
							exit for
						end if
					next
				else
					%><p><a href="members_photos_modify.asp?ID=<%=intCategoryID%>">Click here</a> to go back to the photos in <%=strName%>.</p><%
					p = 0
					PrintTableHeader 100
					Response.Write "<tr>"
					do until rsPage.EOF
						if MemberID = Session("MemberID") or PhotoCaptionsMemberExist( TargetID, Session("MemberID") ) then
							if p mod intPhotosPerRow = 0 then
								Response.Write "</tr><tr>"
								ChangeTDMain
							end if
							rsList.Filter = "ID = " & TargetID
							PrintTableData
							p = p + 1
						end if
						rsPage.MoveNext
					loop

					if p = 0 then Response.Write "<td class=BodyText>Sorry, but there is nothing for you to edit in this category.</td>"


				end if

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
			Query = "SELECT ID, Name, Ext, Thumbnail, ThumbnailExt, MemberID FROM Photos WHERE CustomerID = " & CustomerID & " AND CategoryID = " & intCategoryID & " ORDER BY Date DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.CacheSize = PageSize
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
				Set ID = rsPage("ID")
				Set Name = rsPage("Name")
				Set Ext = rsPage("Ext")
				Set Thumbnail = rsPage("Thumbnail")
				Set ThumbnailExt = rsPage("ThumbnailExt")
				Set MemberID = rsPage("MemberID")

			'Don't navigate if it's empty
			if not rsPage.EOF then
				if blLoggedAdmin then
		%>
					<form METHOD="POST" ACTION="members_photos_modify.asp">
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
							PrintTableData
							rsPage.MoveNext
						else
							exit for
						end if
					next
				else
					p = 0
					PrintTableHeader 100
					Response.Write "<tr>"
					do until rsPage.EOF
						if MemberID = Session("MemberID") or PhotoCaptionsMemberExist( ID, Session("MemberID") ) then
							if p mod intPhotosPerRow = 0 then
								Response.Write "</tr><tr>"
								ChangeTDMain
							end if
							PrintTableData
							p = p + 1
						end if
						rsPage.MoveNext
					loop

					if p = 0 then Response.Write "<td class=BodyText>Sorry, but there is nothing for you to edit in this category.</td>"


				end if

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
			set rsPage = Nothing
		end if
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


Sub PrintTableData
%>
		<form METHOD="POST" ACTION="members_photos_modify.asp">
		<input type="hidden" name="PhotoID" value="<%=ID%>">
		<td class="<% PrintTDMain %>" align="center" valign="middle">
	<%	if Thumbnail = 1 then	%>
			<a href="photos_view.asp?ID=<%=ID%>"><img src="photos/<%=ID%>t.<%=ThumbnailExt%>" border=0 alt="<%=Name%>"></a>
	<%	else	%>
			<a href="photos_view.asp?ID=<%=ID%>"><%=PrintTDLink(Name)%></a>
	<%	end if	%>

	<%	if blLoggedAdmin or MemberID = Session("MemberID") then	%>
			<br>
			<input type="submit" name="Submit" value="Edit">  
			<input type="button" value="Delete" onClick="DeleteBox('If you delete this photo, there is no way to get it or its captions back.  Are you sure?', 'members_photos_modify.asp?Submit=Delete&PhotoID=<%=ID%>')">
	<%	end if
		if (blLoggedAdmin and PhotoCaptionsExist(ID)) or PhotoCaptionsMemberExist( ID, Session("MemberID") ) then	%>
			<br><input type="submit" name="Submit" value="Modify Captions">
<%		end if				%>
	</td>
	</form>
<%
End Sub


Function PhotoCaptionsExist( intPhotoID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "PhotoCaptionsExist"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@ItemID") = intPhotoID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	PhotoCaptionsExist = CBool(blExists)
End Function

Function PhotoCaptionsMemberExist( intPhotoID, intMemberID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "PhotoCaptionsMemberExist"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@ItemID") = intPhotoID
		.Parameters("@MemberID") = intMemberID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	PhotoCaptionsMemberExist = CBool(blExists)
End Function
'------------------------End Code-----------------------------
%>