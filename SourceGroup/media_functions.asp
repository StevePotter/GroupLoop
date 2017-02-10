<%
Function GetDefaultCat()
	'Get the category info
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetDefaultMediaCategory"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CategoryID", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )

		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		intCatID = CInt(.Parameters("@CategoryID"))
	End With
	Set cmdTemp = Nothing
	GetDefaultCat = intCatID
End Function

Function MediaFileExists( strFileName )
	'Get the category info
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "MediaFileExists"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@FileName", adVarWChar, adParamInput, 400 )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@FileName") = strFileName

		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
	End With
	Set cmdTemp = Nothing
	MediaFileExists = blExists
End Function

Sub GetCategoryInfo( intCategoryID, strName, blPrivate )
	'Get the category info
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetMediaCategoryInfo"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamOutput, 100 )
		.Parameters.Append .CreateParameter ("@Private", adInteger, adParamOutput )

		.Parameters("@ItemID") = intCategoryID

		.Execute , , adExecuteNoRecords
		strName = .Parameters("@Name")
		blPrivate = CBool(.Parameters("@Private"))
	End With
	Set cmdTemp = Nothing
End Sub

Function CatsExist()
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "MediaCategoriesExist"
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
'This function sees if a category exists
'-------------------------------------------------------------
Function ValidCategory( intCategoryID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "MediaValidCategory"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@ItemID") = intCategoryID
		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	ValidCategory = CBool(blExists)
End Function


'-------------------------------------------------------------
'This function prints a category name
'-------------------------------------------------------------
Function GetCategory( intCategoryID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetMediaCategory"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamOutput, 200 )

		.Parameters("@ItemID") = intCategoryID

		.Execute , , adExecuteNoRecords
		strName = .Parameters("@Name")
	End With
	Set cmdTemp = Nothing

	GetCategory = strName
End Function



'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the categories
'if intHighLight = 0, then it prints a list excluding the intCategoryID
'if intHighLight = 1, then it prints a list, highlighting the intCategoryID
'-------------------------------------------------------------
Sub PrintCategoryPullDown( intCategoryID, intMemberID, intHighLightID )

	if intCategoryID <> "" then intCategoryID = CInt(intCategoryID)
	'Now we are going to get the group names to list in the pull-down menu
	if intHighLightID = 0 then
		if intMemberID = 0 then
			Query = "SELECT Name, Private, ID FROM MediaCategories WHERE (ID <> " & intCategoryID & " AND CustomerID = " & CustomerID & ") ORDER BY Name"
		else
			Query = "SELECT Name, Private, ID FROM MediaCategories WHERE (ID <> " & intCategoryID & " AND MemberID = " & intMemberID & ") ORDER BY  Name"
		end if
	else
		if intMemberID = 0 then
			Query = "SELECT Name, Private, ID FROM MediaCategories WHERE (CustomerID = " & CustomerID & ") ORDER BY Name"
		else
			Query = "SELECT Name, Private, ID FROM MediaCategories WHERE (MemberID = " & intMemberID & ") ORDER BY Name"
		end if
	end if 
	Set rsTempCats = Server.CreateObject("ADODB.Recordset")
	rsTempCats.CacheSize = 30
	rsTempCats.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsTempCats.EOF then
		Set rsTempCats = Nothing
		Exit Sub
	end if

	Set Name = rsTempCats("Name")
	Set ID = rsTempCats("ID")
	Set IsPrivate = rsTempCats("Private")

	'Make the size 3 if there are many members
	if rsTempCats.RecordCount <= 30 then
		%><select name="ID" size="1"><%
	else
		%><select name="ID" size="3"><%
	end if

	do until rsTempCats.EOF
		strPrivate = ""
		if IsPrivate = 1 then strPrivate = " (private)"
		'Highlight the current category
		if intHighLightID = 0 then
			Response.Write "<option value = '" & ID & "'>" & Name & strPrivate & "</option>" & vbCrlf
		else
			if ID = intCategoryID then
				Response.Write "<option value = '" & ID & "' SELECTED>" & Name & strPrivate & "</option>" & vbCrlf
			else
				Response.Write "<option value = '" & ID & "'>" & Name & strPrivate & "</option>" & vbCrlf
			end if
		end if 

		rsTempCats.MoveNext
	loop
	rsTempCats.Close
	set rsTempCats = Nothing
	Response.Write("</select>")

End Sub

'-------------------------------------------------------------
'This subroutine prints a pulldown menu for the other categories
'-------------------------------------------------------------
Sub PrintCategoryMenu( strSource )
	Query = "SELECT Private, ID, Name, Date FROM MediaCategories WHERE (CustomerID = " & CustomerID & ") ORDER BY Name"
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsPage.EOF then
		Response.Write "<p>Sorry, but there are no categories yet.  Once the administrator creates one, the photos section can be used.</p>"
	else
		Response.Write "<p><span class='Heading'>Please choose a category:</span><br>"
		Set IsPrivate = rsPage("Private")
		Set ID = rsPage("ID")
		Set Name = rsPage("Name")
		Set ItemDate = rsPage("Date")

%>
		<form METHOD="POST" ACTION="<%=strSource%>">
<%
		PrintPagesHeader
		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				strPrivate = ""
				if IsPrivate = 1 then strPrivate = "(private)"
'------------------------End Code-----------------------------
%>
				&nbsp;&nbsp;&nbsp;&nbsp;<% PrintNew(ItemDate) %><a href="<%=strSource%>?ID=<%=ID%>"><%=Name%></a> &nbsp;<font size="-2"><%=strPrivate%></font><br>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write "</p>"
	end if

End Sub
%>