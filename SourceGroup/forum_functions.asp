<%
'-------------------------------------------------------------
'This function tells whether or not a message has replies
'-------------------------------------------------------------
Function HasReplies( intMessageID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "HasReplies"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@ItemID") = intMessageID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing
	HasReplies = CBool(blExists)
End Function


Sub GetCategory( intCategoryID, strName, blPrivate, blMembersOnly )
	'Get the category info
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetForumCategoryInfo"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamOutput, 200 )
		.Parameters.Append .CreateParameter ("@Private", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@MembersOnly", adInteger, adParamOutput )

		.Parameters("@ItemID") = intCategoryID

		.Execute , , adExecuteNoRecords
		strName = .Parameters("@Name")
		blPrivate = CBool(.Parameters("@Private"))
		blMembersOnly = CBool(.Parameters("@MembersOnly"))
	End With
	Set cmdTemp = Nothing
End Sub


Function IsExpanded( intSessionID, intMessageID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ForumIsExpanded"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@SessionID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@MessageID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@SessionID") = intSessionID
		.Parameters("@MessageID") = intMessageID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	IsExpanded = CBool(blExists)
End Function

'-------------------------------------------------------------
'This function prints a category name
'-------------------------------------------------------------
Function GetForumCategory( intCategoryID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetForumCategory"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamOutput, 200 )

		.Parameters("@ItemID") = intCategoryID

		.Execute , , adExecuteNoRecords
		strName = .Parameters("@Name")
	End With
	Set cmdTemp = Nothing

	GetForumCategory = strName
End Function


'-------------------------------------------------------------
'This function sees if a category exists
'-------------------------------------------------------------
Function ValidCategory( intCategoryID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "ForumValidCategory"
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
'This subroutine prints a pulldown menu for the other categories
'-------------------------------------------------------------
Sub PrintCategoryPullDown( intCategoryID )
	intCategoryID = CInt(intCategoryID)
	'Now we are going to get the group names to list in the pull-down menu
	Query = "SELECT ID, Private, Name FROM ForumCategories WHERE (CustomerID = " & CustomerID & ") ORDER BY Name"
	Set rsTempCats = Server.CreateObject("ADODB.Recordset")
	rsTempCats.CacheSize = 30
	rsTempCats.Open Query, Connect, adOpenStatic, adLockReadOnly

	'Make the size 3 if there are many members
	if rsTempCats.RecordCount <= 30 then
		%><select name="ID" size="1" onChange="this.form.submit();"><%
	else
		%><select name="ID" size="3" onChange="this.form.submit();"><%
	end if

	do While not rsTempCats.EOF
		'Highlight the current section
		if rsTempCats("Private") = 1 then
			strPrivate = " (private)"
		else
			strPrivate = ""
		end if
		strSelected = ""
		if rsTempCats("ID") = intCategoryID then strSelected = " SELECTED"
		Response.Write "<option value = '" & rsTempCats("ID") & "' " & strSelected & ">" & rsTempCats("Name") & strPrivate & "</option>" & vbCrlf
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
	Query = "SELECT Private, MembersOnly, ID, Name, Date FROM ForumCategories WHERE (CustomerID = " & CustomerID & ") ORDER BY Name"
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsPage.EOF then
		Response.Write "<p>Sorry, but there are no topics in the forum yet.  Once the administrator creates one, the forum can be used.</p>"
	else
		Response.Write "<p><span class='Heading'>Please choose a topic:</span><br>"
%>
		<form METHOD="POST" ACTION="<%=strSource%>">
<%
		PrintPagesHeader
		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				strPrivate = ""
				if rsPage("Private") = 1 then strPrivate = "(private)"
				if rsPage("MembersOnly") = 1 then strPrivate = "(only members can post)"
'------------------------End Code-----------------------------
%>
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=strSource%>?ID=<%=rsPage("ID")%>"><%=rsPage("Name")%></a> &nbsp;<font size="-2"><%=strPrivate%></font><br>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write "</p>"
	end if

End Sub


'-------------------------------------------------------------
'This subroutine prints the plus for explanding threads
'-------------------------------------------------------------
Sub PrintPlus()

	if ForumPlusImage = 1 then	%>
		<img src="images/ForumPlusImage.<%=ForumPlusImageExt%>" border="0"><%
	else
		Response.Write "+"
	end if

End Sub


'-------------------------------------------------------------
'This subroutine prints the minus for explanding threads
'-------------------------------------------------------------
Sub PrintMinus()

	if ForumMinusImage = 1 then	
		%><img src="images/ForumMinusImage.<%=ForumMinusImageExt%>" border="0"><%
	else
		Response.Write "-"
	end if

End Sub

%>