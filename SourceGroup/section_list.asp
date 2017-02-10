<%

Sub GoList()
	intSearchID = GetSearchID()

	if intSearchID = 0 then
%>
		<p>Sorry, but your search came up empty.<br>Try again, or <a href="<%=strListSource%>">click here</a> to view all <%=strPluralNoun%>.</p>
<%
	else
		if InfoText <> " " and InfoText <> "" then Response.Write "<p>" & InfoText & "</p>"

		if strOrderBy <> "" then strOrderBy = " ORDER BY " & strOrderBy
		if strFields = "" then strFields = "*"

		'This is if they requested items written in a time period
		if Request("DaysOld") <> "" then
			CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
			Query = "SELECT " & strFields & " FROM " & strTable & " WHERE (CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "')" & strOrderBy
		else
			Query = "SELECT " & strFields & " FROM " & strTable & " WHERE (CustomerID = " & CustomerID & ")" & strOrderBy
		end if
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
		if not rsPage.EOF then
%>
			<form METHOD="POST" ACTION="<%=strListSource%>">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
			<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%

			blShowModify = ShowModify( blShowModify, rsPage )

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
				Response.Write "<p>Sorry, but there have been no " & LCase(strPluralNoun) & " added in that time period. <a href='javascript:history.back(1)'>Click here to go back</a></p>"
			else
				Response.Write "<p>Sorry, but there are no " & LCase(strPluralNoun) & " at the moment.  Please try again later.  <a href='index.asp?Action=Old'>Click here to go back to the home page</a></p>"
			end if
		end if
		rsPage.Close
		set rsPage = Nothing
	end if
End Sub


Sub PrintTitle( strTitle, intMembersCanAdd, strAddSource, strNoun )
%>
	<p align="<%=HeadingAlignment%>"><span class=Heading><%=strTitle%></span><br>
<%
	if (IncludeAddButtons = 1 or LoggedMember()) and (LoggedAdmin() or CBool( intMembersCanAdd )) then
%>
	<span class=LinkText><a href="<%=strAddSource%>">Add <%=PrintAn(strNoun)%>&nbsp;<%=strNoun%></a></span>
<%
	end if
	Response.Write "</p>"

End Sub


Sub CheckSection( intInclude, strSectionView, strSource )
	if not CBool( intInclude ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. Only an administrator can reactivate it."))

	if strSectionView = "Members" and not LoggedMember() then Redirect("login.asp?Source=" & strSource & "&Message=" & Server.URLEncode("Only members can view this section.  If you are a member, please log in below.  Otherwise, sorry, but you may not view this section."))
	if strSectionView = "Administrators" and not LoggedAdmin() then
		if LoggedMember() then
			Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section can only be viewed by an administrator."))
		else
			Redirect("login.asp?Source=" & strSource & "&Message=" & Server.URLEncode("Only <b>site administrators</b> can view this section.  If you are an administrator, please log in below.  If you are a regular member or a non-member, sorry, but you may not view this section."))
		end if
	end if
End Sub



Function PrintHeaderCol( strContents, strAlign )
	if strAlign <> "" then strAlign = " align=" & strAlign
	if strContents = "" then strContents = "&nbsp;"

	PrintHeaderCol = "<td class='TDHeader'" & strAlign & ">" & strContents & "</td>"
End Function

'Set the modify values
Function DisplayModifyCol( strTable )
	blShowModify = False

	if Request("Modify") = "Yes" then
		Session("ModifyItems") = "Yes"
	elseif Request("Modify") = "No" then
		Session("ModifyItems") = "No"
	elseif blLoggedMember and Session("ModifyItems") = "" then	'Their first time, set the modify to yes
		Session("ModifyItems") = "Yes"
	end if

	if blLoggedMember then
		if blLoggedAdmin or GetNumMemberItems( strTable, Session("MemberID") ) > 0 then blShowModify = True
	end if

	DisplayModifyCol = blShowModify
End Sub


Function GetNumMemberItems( strTable, intMemberID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetNumMemberItems"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@Table", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Count", adInteger, adParamOutput )
		.Parameters("@Table") = strTable
		.Parameters("@MemberID") = intMemberID
		.Execute , , adExecuteNoRecords
		intCount = .Parameters("@Count")
	End With
	Set cmdTemp = Nothing

	GetNumMemberItems = intCount
End Function



'If there is a member going through, see if they need the buttons on this page
Function ShowModify( blAlreadyShow, rsObject )
	'True was passed
	if blAlreadyShow then
		'Check if they are just a regular member
		if LoggedMember() and not LoggedAdmin() then
			rsObject.Filter = "MemberID = " & Session("MemberID")
			'if it's empty, don't show
			blAlreadyShow = not rsObject.EOF
			rsObject.Filter = ""
		end if

		'Make sure they haven't set this to no modify
		if Session("ModifyItems") <> "Yes" then blAlreadyShow = False

	end if

	ShowModify = blAlreadyShow
End Function



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

	'They can modify announcements, and aren't already showing the edit/delete buttons
	if blShowModify and Session("ModifyItems") = "No" then
		if blLoggedAdmin then
			strYour = ""
		else
			strYour = "Your"
		end if
		Response.Write "<br><br><p align=left><a href='" & strListSource & "?Modify=Yes'>Show Edit/Delete Buttons For " & strYour & "&nbsp;" & strPluralNoun & "</a></p>"
	elseif blShowModify and Session("ModifyItems") = "Yes" then
		Response.Write "<br><br><p align=left><a href='" & strListSource & "?Modify=No'>Hide Edit/Delete Buttons</a></p>"
	end if

	'Give them the link to change the section's properties
	if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
		Response.Write "<div align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=" & strTable & "&Source=" & strListSource & "'>Change Section Options</a></div>"
	end if

End Sub

Function PrintAn(strFollowWord)
	if IsNull( strFollowWord) then
		PrintAn = "a"
	elseif strFollowWord = "" then
		PrintAn = "a"
	else
		testchar = Lcase( Left( strFollowWord,1 ) )

		if Instr( testchar, "aeiou" ) then
			PrintAn = "an"
		else
			PrintAn = "a"
		end if
	end if

End Function

Sub PrintSearch( DisplaySearch, DisplayDaysOld, strListSource, strPluralNoun )
	if DisplaySearch or DisplayDaysOld then
%>
		<form METHOD="POST" ACTION="<%=strListSource%>">
<%		if DisplayDaysOld then	%>
		View <%=strPluralNoun%> In The Last <% PrintDaysOld %>
		<br>
<%			if DisplaySearch then Response.Write "Or "
		end if
		if DisplaySearch then	%>
		Search For <input type="text" name="Keywords" size="25">
		<input type="submit" name="Submit" value="Go"><br>
<%		end if	%>	
		</form>
<%
	end if
End Sub
%>
