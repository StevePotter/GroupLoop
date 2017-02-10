<%
Sub MakeLongNames(strTable)
	Query = "SELECT Name, LongName, ID FROM " & strTable & " WHERE (CustomerID = " & CustomerID & ") ORDER BY Name"

	Set rsTempCats = Server.CreateObject("ADODB.Recordset")

	rsTempCats.CacheSize = 30
	rsTempCats.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect
	do until rsTempCats.EOF
		rsTempCats("LongName") = GetCatHeiarchy( rsTempCats("ID"), "", strTable, "" )
		rsTempCats.Update
		rsTempCats.MoveNext
	loop
	rsTempCats.Close

	Set rsTempCats = Nothing
End Sub

'-------------------------------------------------------------
'This prints out a heiarchy of categories (Store > Guns > Pistols > Ammo...)
'-------------------------------------------------------------
Function GetCatHeiarchy( intCategoryID, strLink, strTable, strFirstTitle )
	intTempCat = intCategoryID

	'Print out the beginning of the heiarchy
	if strLink <> "" then
		strHeiarchy = "<a href=" & strLink & "?ID=" & intTempCat & ">" & GetCategoryName(intTempCat, strTable) & "</a>"
	else
		strHeiarchy = GetCategoryName(intTempCat, strTable)
	end if

	'If we have a parent, rewind to the base category
	if CategoryHasParent(intTempCat, strTable) then
		do until not CategoryHasParent(intTempCat, strTable)
			intTempCat = GetCatParentID( intTempCat, strTable )
			if strLink <> "" then
				strHeiarchy = "<a href=" & strLink & "?ID=" & intTempCat & ">" & GetCategoryName(intTempCat, strTable) & "</a>&nbsp;&gt;&nbsp;" & strHeiarchy
			else
				strHeiarchy = GetCategoryName(intTempCat, strTable) & "&nbsp;&gt;&nbsp;" & strHeiarchy
			end if
		loop
	end if

	'Put the front and end tags on this bitch
	if strFirstTitle <> "" then
		if strLink <> "" then
			strHeiarchy = "<a href=" & strLink & ">" & strFirstTitle & "</a>&nbsp;&gt;&nbsp;" & strHeiarchy
		else
			strHeiarchy = strFirstTitle & "&nbsp;&gt;&nbsp;" & strHeiarchy
		end if
	end if

	GetCatHeiarchy = strHeiarchy
End Function


'-------------------------------------------------------------
'This function sees if a category has siblings
'-------------------------------------------------------------
Function CategoryHasParent( intCategoryID, strTable )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ParentCategoriesExist"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CategoryID") = intCategoryID
		.Parameters("@Table") = strTable

		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing

	CategoryHasParent = blExists
End Function


'-------------------------------------------------------------
'This function gets the parent's ID
'-------------------------------------------------------------
Function GetCatParentID( intCategoryID, strTable )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetCatParentID"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CategoryID") = intCategoryID
		.Parameters("@Table") = strTable

		.Execute , , adExecuteNoRecords

		intTempParentID = .Parameters("@ParentID")
	End With
	Set cmdTemp = Nothing

	GetCatParentID = intTempParentID
End Function


'-------------------------------------------------------------
'This function sees if a category has a child
'-------------------------------------------------------------
Function CategoryHasChild( intCategoryID, strTable )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ChildCategoriesExist"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CategoryID") = intCategoryID
		.Parameters("@Table") = strTable

		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing

	CategoryHasChild = blExists
End Function

'-------------------------------------------------------------
'This function prints a category name
'-------------------------------------------------------------
Function GetCategoryName( intCategoryID, strTable )
	if intCategoryID = "" then Redirect("error.asp")
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetCategoryName"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@ItemID") = intCategoryID
		.Parameters("@Table") = strTable

		.Execute , , adExecuteNoRecords
		strTempName = .Parameters("@Name")
	End With
	Set cmdTemp = Nothing

	GetCategoryName = strTempName
End Function


'-------------------------------------------------------------
'This function sees if a category exists
'-------------------------------------------------------------
Function ValidCategory( intCategoryID, strTable )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ValidCategory"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@Table") = strTable
		.Parameters("@ItemID") = intCategoryID
		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blResult = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	ValidCategory = CBool(blResult)
End Function

'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the other categories
'-------------------------------------------------------------
Sub PrintCategoryMenu( strSource, intParentID, strTable )
	Query = "SELECT ID, Name, Date FROM " & strTable & " WHERE (ParentID = " & intParentID & " AND CustomerID = " & CustomerID & ") ORDER BY Name"

	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if rsPage.EOF then
		if intParentID = 0 then Response.Write "<p>Sorry, but there are no categories yet.  Once the administrator creates one, the store can be used.</p>"
	else
		if intParentID = 0 then
			Response.Write "<p><span class='Heading'>Please choose a category:</span><br>"
		else
			Response.Write "<p><span class='Heading'>Sub-Categories:</span><br>"
		end if
		Set ID = rsPage("ID")
		Set Name = rsPage("Name")
		Set ItemDate = rsPage("Date")

'		if rsPage.RecordCount > PageSize then

'			<form METHOD="POST" ACTION="strSource  ">
'			PrintPagesHeader
'		end if
'		for p = 1 to rsPage.PageSize
'			if not rsPage.EOF then
		do until rsPage.EOF
'------------------------End Code-----------------------------
%>
				&nbsp;&nbsp;&nbsp;&nbsp;<% PrintNew(ItemDate) %><a href="<%=strSource%>?ID=<%=ID%>"><%=Name%></a> &nbsp;<br>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
		loop
'			end if
'		next
		Response.Write "</p>"
	end if

End Sub



Sub GetCategoryInfo( intCategoryID, strName, blPrivate, strBody )
	'Get the category info
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetPhotoCategoryInfo"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh
		.Parameters("@ItemID") = CInt(intCategoryID)

		.Execute , , adExecuteNoRecords
		strName = .Parameters("@Name")
		strBody = .Parameters("@Body")

		blPrivate = CBool(.Parameters("@Private"))
	End With
	Set cmdTemp = Nothing
End Sub


Function CatsExist()
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "PhotoCategoriesExist"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing
	CatsExist = CBool(blExists)
End Function


'-------------------------------------------------------------
'This function prints a category name
'-------------------------------------------------------------
Function GetNumCaptions( intPhotoID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetNumCaptions"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Count", adInteger, adParamOutput )

		.Parameters("@ItemID") = intPhotoID

		.Execute , , adExecuteNoRecords
		intCount = .Parameters("@Count")
	End With
	Set cmdTemp = Nothing

	GetNumCaptions = intCount
End Function


'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the categories
'if intHighLight = 0, then it prints a list excluding the intCategoryID
'if intHighLight = 1, then it prints a list, highlighting the intCategoryID
'if intChangeSubmit = 1, when they change the option it submits the form
'strname = name of pulldown menu
'strTable = table name from sql
'-------------------------------------------------------------
Sub PrintCategoryPullDown( intCategoryID, intHighLightID, intChangeSubmit, intPrintBlank, intLongName, strTable, strName, strSpecialQuery )

	if intCategoryID <> "" then
		intCategoryID = CInt(intCategoryID)
	else
		intCategoryID = 0
	end if
	'exclude/include the category
	if strSpecialQuery <> "" then
			Query = strSpecialQuery
	elseif intHighLightID = 0 then
			Query = "SELECT Name, LongName, Private, ID FROM " & strTable & " WHERE (ID <> " & intCategoryID & " AND CustomerID = " & CustomerID & ") ORDER BY LongName"
	else
			Query = "SELECT Name, LongName, Private, ID FROM " & strTable & " WHERE (CustomerID = " & CustomerID & ") ORDER BY LongName"
	end if 
	Set rsTempCats = Server.CreateObject("ADODB.Recordset")
	rsTempCats.CacheSize = 30
	rsTempCats.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	if rsTempCats.EOF then
		Set rsTempCats = Nothing
		Exit Sub
	end if

	'They use the short name, not the long one
	if intLongName = 0 then
		Set Name = rsTempCats("Name")
	else
		Set Name = rsTempCats("LongName")
	end if
	Set ID = rsTempCats("ID")
	Set IsPrivate = rsTempCats("Private")

	'Make the size 3 if there are many members
		%><select name="<%=strName%>" size="1" <%


	if intChangeSubmit = 1 then
		%>onChange="this.form.submit();"<%
	end if

	Response.Write ">"
	if intPrintBlank = 1 then Response.Write "<option value = ''>None</option>" & vbCrlf

	do until rsTempCats.EOF
		strPrivate = ""
		if IsPrivate = 1 then strPrivate = " (private)"
		'Highlight the current category
		if intHighLightID = 1 and ID = intCategoryID then
			Response.Write "<option value = '" & ID & "' SELECTED>" & Name & strPrivate & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & ID & "'>" & Name & strPrivate & "</option>" & vbCrlf
		end if

		rsTempCats.MoveNext
	loop
	rsTempCats.Close

	set rsTempCats = Nothing
	Response.Write("</select>")

End Sub



'-------------------------------------------------------------
'This sub creates a thumbnail for an image
'strFileName - name of photo
'strThumbFileName - name of thumbnail to save as
'-------------------------------------------------------------
Sub CreateThumbnail( strFileName, strThumbFileName )
	'Create the object and load up the big image
	Set Image = Server.CreateObject("AspImage.Image")
	Image.LoadImage(strFileName)


	if Image.MaxX = 0 or Image.MaxY = 0 then
		Set Image = Nothing
		Response.Write "Cannot create thumbnail because your image is too large for memory or is an unacceptable type.  If you scanned it, use a lower DPI, or resize it smaller.  Also try saving it as a JPG.<br>"
		Exit Sub
	end if

	'now lets get the ratio of the image size so we can set the thumbnail image size
	dblRatio = Image.MaxX / Image.MaxY	'X/Y

	'Horizontal image, resize to 100 pixels wide
	if dblRatio > 1 then
		intNewX = 100
	'Vertical image, resize to 70 pixels wide
	else
		intNewX = 70
	end if

	'Get the height and resize
	intNewY = Round( intNewX / dblRatio )
	Image.ResizeR intNewX, intNewY

	'Now set the filetype, name, and quality and save it
	Image.FileName = strThumbFileName
	Image.ImageFormat = 1
	Image.JPEGQuality = 80
	Image.SaveImage

	'DIE MOTHERFUCKER DIE MOTHERFUCKER DIE
	Set Image = Nothing
End Sub


'-------------------------------------------------------------
'This sub optimizes an image for oversizing and makes it a JPEG
'strFile - name of image, omitting the extension
'strExt - extension of image.  It gets changed to jpg
'-------------------------------------------------------------
Sub OptimizeImage( strFile, strExt )

	strFileName = strFile & "." & strExt
	Set TempFileSystem = CreateObject("Scripting.FileSystemObject")
	if not TempFileSystem.FileExists(strFileName) then
		Response.Write ("Cannot optimize because photo does not exist - " & strFileName)
		Exit Sub
	end if

	'Create the object and load up the big image
	Set Image = Server.CreateObject("AspImage.Image")
	Image.LoadImage(strFileName)

	if Image.MaxX = 0 or Image.MaxY = 0 then
		Set Image = Nothing
		Response.Write "Cannot optimize image because your image is too large for memory or is an unacceptable type.  If you scanned it, use a lower DPI, or resize it smaller.  Also try saving it as a JPG.<br>"
		Exit Sub
	end if

	intOrgX = Image.MaxX
	intOrgY = Image.MaxY

	'now lets get the ratio of the image size so we can set the thumbnail image size
	dblRatio = intOrgX / intOrgY	'X/Y

	'Horizontal image, resize to 600 pixels wide
	if dblRatio >= 1 and intOrgX > 600 then
		intNewX = 600
		intNewY = Round( intNewX / dblRatio )
		Image.ResizeR intNewX, intNewY
	'Vertical image, resize to 450 pixels wide
	elseif dblRatio < 1 and intOrgX > 600 then
		intNewX = 450
		intNewY = Round( intNewX / dblRatio )
		Image.ResizeR intNewX, intNewY
	end if

	'Delete the old file
	TempFileSystem.DeleteFile(strFileName)
	Set TempFileSystem = Nothing

	strFileName = strFile & ".jpg"

	'Now set the filetype, name, and quality and save it
	Image.FileName = strFileName
	Image.ImageFormat = 1
	Image.JPEGQuality = 75
	Image.SaveImage

	'DIE MOTHERFUCKER DIE MOTHERFUCKER DIE
	Set Image = Nothing
End Sub
%>