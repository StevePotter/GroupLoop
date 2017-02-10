<%
Sub MakeLongNames(strSection)
	Query = "SELECT Name, LongName, ID FROM Categories WHERE (CustomerID = " & CustomerID & " AND Section = '" & strSection & "') ORDER BY Name"

	Set rsTempCats = Server.CreateObject("ADODB.Recordset")

	rsTempCats.CacheSize = 30
	rsTempCats.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect
	do until rsTempCats.EOF
		rsTempCats("LongName") = GetCatHeiarchy( rsTempCats("ID"), "", strSection, "" )
		rsTempCats.Update
		rsTempCats.MoveNext
	loop
	rsTempCats.Close

	Set rsTempCats = Nothing
End Sub



'-------------------------------------------------------------
'This prints out a heiarchy of categories (Store > Guns > Pistols > Ammo...)
'-------------------------------------------------------------
Function GetCatHeiarchy( intCategoryID, strLink, strSection, strFirstTitle )
	intTempCat = intCategoryID

	'Print out the beginning of the heiarchy
	if strLink <> "" then
		strHeiarchy = "<a href=" & strLink & "?ID=" & intTempCat & ">" & GetCategoryName(intTempCat, strSection) & "</a>"
	else
		strHeiarchy = GetCategoryName(intTempCat, strSection)
	end if

	'If we have a parent, rewind to the base category
	if CategoryHasParent(intTempCat, strSection) then
		do until not CategoryHasParent(intTempCat, strSection)
			intTempCat = GetCatParentID( intTempCat, strSection )
			if strLink <> "" then
				strHeiarchy = "<a href=" & strLink & "?ID=" & intTempCat & ">" & GetCategoryName(intTempCat, strSection) & "</a>&nbsp;&gt;&nbsp;" & strHeiarchy
			else
				strHeiarchy = GetCategoryName(intTempCat, strSection) & "&nbsp;&gt;&nbsp;" & strHeiarchy
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
Function CategoryHasParent( intCategoryID, strSection )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ParentCategoriesExist"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CategoryID") = intCategoryID
		.Parameters("@Section") = strSection

		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing

	CategoryHasParent = blExists

End Function



'-------------------------------------------------------------
'This function gets the parent's ID
'-------------------------------------------------------------
Function GetCatParentID( intCategoryID, strSection )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetCatParentID"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CategoryID") = intCategoryID
		.Parameters("@Section") = strSection

		.Execute , , adExecuteNoRecords

		intTempParentID = .Parameters("@ParentID")
	End With
	Set cmdTemp = Nothing

	GetCatParentID = intTempParentID
End Function


'-------------------------------------------------------------
'This function sees if a category has a child
'-------------------------------------------------------------
Function CategoryHasChild( intCategoryID, strSection )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ChildCategoriesExist"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CategoryID") = intCategoryID
		.Parameters("@Section") = strSection

		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing

	CategoryHasChild = blExists
End Function

'-------------------------------------------------------------
'This function gets a category's main info
'-------------------------------------------------------------
Sub GetCategoryInfo( intCategoryID, strSection, strTempName, strLongName, strBody )
	'Get the category info
	Set cmdName = Server.CreateObject("ADODB.Command")
	With cmdName
		.ActiveConnection = Connect
		.CommandText = "GetCategoryInfo"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh
		.Parameters("@CategoryID") = intCategoryID
		.Parameters("@Section") = strSection

		.Execute , , adExecuteNoRecords
		strTempName = .Parameters("@Name")
		strLongName = .Parameters("@LongName")
		strBody = .Parameters("@Body")
	End With
	Set cmdName = Nothing
End Sub

'-------------------------------------------------------------
'This function prints a category name
'-------------------------------------------------------------
Function GetCategoryName( intCategoryID, strSection )
	if intCategoryID = "" then Redirect("error.asp")

	Set cmdName = Server.CreateObject("ADODB.Command")
	With cmdName
		.ActiveConnection = Connect
		.CommandText = "GetCategoryName"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh
		.Parameters("@ItemID") = intCategoryID
		.Parameters("@Section") = strSection

		.Execute , , adExecuteNoRecords
		strTempName = .Parameters("@Name")
	End With
	Set cmdName = Nothing

	GetCategoryName = strTempName
End Function


'-------------------------------------------------------------
'This function sees if a category exists
'-------------------------------------------------------------
Function ValidCategory( intCategoryID, strSection )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ValidCategory"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@Section") = strSection
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
Sub PrintCategoryMenu( strSource, intParentID, strSection )
	Query = "SELECT ID, Name, Date FROM Categories WHERE (Section = '" & strSection & "' AND ParentID = " & intParentID & " AND CustomerID = " & CustomerID & ") ORDER BY Name"
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if rsPage.EOF then
		if intParentID = 0 then Response.Write "<p>Sorry, but there are no categories yet.</p>"
	else
		if intParentID = 0 then
			Response.Write "<p><span class='Heading'>Please choose a category:</span><br>"
		else
			Response.Write "<p><span class='Heading'>Sub-Categories:</span><br>"
		end if
		Set ID = rsPage("ID")
		Set Name = rsPage("Name")
		Set ItemDate = rsPage("Date")


		do until rsPage.EOF
'------------------------End Code-----------------------------
%>
				&nbsp;&nbsp;&nbsp;&nbsp;<% PrintNew(ItemDate) %><a href="<%=strSource%>?ID=<%=ID%>"><%=Name%></a> &nbsp;<br>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
		loop
		Response.Write "</p>"
	end if

End Sub


'-------------------------------------------------------------
'This function sees if a category exists
'-------------------------------------------------------------
Function CategoriesExist(strSection)
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "CategoriesExist"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@Section") = strSection

		.Execute , , adExecuteNoRecords
		blResult = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	CategoriesExist = CBool(blResult)
End Function


'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the categories
'if intHighLight = 0, then it prints a list excluding the intCategoryID
'if intHighLight = 1, then it prints a list, highlighting the intCategoryID
'if intChangeSubmit = 1, when they change the option it submits the form
'strTempName = name of pulldown menu
'strSection = section for cat
'-------------------------------------------------------------
Sub PrintCategoryPullDown( intCategoryID, intHighLightID, intChangeSubmit, intPrintBlank, intLongName, strSection, strTempName, strSpecialQuery )

	if intCategoryID <> "" then
		intCategoryID = CInt(intCategoryID)
	else
		intCategoryID = 0
	end if
	'exclude/include the category
	if strSpecialQuery <> "" then
			Query = strSpecialQuery
	elseif intHighLightID = 0 then
			Query = "SELECT Name, LongName, Private, ID FROM Categories WHERE (Section = '" & strSection & "' AND ID <> " & intCategoryID & " AND CustomerID = " & CustomerID & ") ORDER BY LongName"
	else
			Query = "SELECT Name, LongName, Private, ID FROM Categories WHERE (Section = '" & strSection & "' AND CustomerID = " & CustomerID & ") ORDER BY LongName"
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

	%><select name="<%=strTempName%>" size="1" <%


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


%>
