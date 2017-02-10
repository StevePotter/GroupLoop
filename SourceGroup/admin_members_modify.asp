<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_members_edit.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Members</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
strMatch = "CustomerID = " & CustomerID
strSubmit = Request("Submit")


'-------------------------------------------------------------
'This function returns the nickname of a given member
'-------------------------------------------------------------
Function GetNickNameOnly( intMemberID )
	Set cmdTempFunction = Server.CreateObject("ADODB.Command")
	With cmdTempFunction
		.ActiveConnection = Connect
		.CommandText = "GetNickName"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Display", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@Nick", adVarWChar, adParamOutput, 1000 )

		.Parameters("@ItemID") = intMemberID
		.Parameters("@Display") = "NickName"

		.Execute , , adExecuteNoRecords
		strNick = .Parameters("@Nick")
	End With
	Set cmdTempFunction = Nothing

	GetNickNameOnly = strNick
End Function


'We are going to check for errors if they are updating the profile
if strSubmit = "Update" then
	strNickName = Format(Request("NickName"))

	if Request("ID") = "" or Request("FirstName") = "" or Request("LastName") = "" or Request("NickName") = "" or Request("Password") = ""	then Redirect("incomplete.asp")

	if GetNickNameOnly(Request("ID")) <> strNickName and NickNameTaken( strNickName, CustomerID ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but the " & UsernameLabel & "&nbsp;" & strNickName & " is already taken."))

	Query = "SELECT * FROM Members WHERE ID = " & Request("ID")
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
	if not rsMember("CustomerID") = CustomerID then Redirect("error.asp")

	rsMember("Admin") = Request("Admin")
	rsMember("PrivateName") = Request("PrivateName")
	rsMember("PrivateBirthdate") = Request("PrivateBirthdate")
	rsMember("PrivateEMail") = Request("PrivateEMail")
	rsMember("PrivateHome") = Request("PrivateHome")
	rsMember("PrivateSecondary") = Request("PrivateSecondary")
	rsMember("PrivateBeeper") = Request("PrivateBeeper")
	rsMember("PrivateCellPhone") = Request("PrivateCellPhone")


	rsMember("FirstName") = Format(Request("FirstName"))
	rsMember("LastName") = Format(Request("LastName"))
	rsMember("NickName") = strNickName
	rsMember("Password") = Request("Password")
	rsMember("EMail1") = Request("EMail1")
	rsMember("EMail2") = Request("EMail2")
	rsMember("WebSite") = Request("WebSite")
	rsMember("Beeper") = Request("Beeper")
	rsMember("CellPhone") = Request("CellPhone")
	rsMember("Birthdate") = AssembleDate("Birthdate")
	rsMember("HomeStreet") = Request("HomeStreet")
	rsMember("HomeCity") = Request("HomeCity")
	rsMember("HomeState") = Request("HomeState")
	rsMember("HomeZip") = Request("HomeZip")
	rsMember("HomeCountry") = Request("HomeCountry")
	rsMember("HomePhone") = Request("HomePhone")
	rsMember("SecondaryDescription") = Format( Request("SecondaryDescription") )
	rsMember("SecondaryStreet") = Request("SecondaryStreet")
	rsMember("SecondaryCity") = Request("SecondaryCity")
	rsMember("SecondaryState") = Request("SecondaryState")
	rsMember("SecondaryZip") = Request("SecondaryZip")
	rsMember("SecondaryCountry") = Request("SecondaryCountry")
	rsMember("SecondaryPhone") = Request("SecondaryPhone")
	rsMember("SecondaryPExt") = Request("SecondaryPExt")
	rsMember("CustomerID") = CustomerID

	rsMember.Update
	rsMember.Close
	set rsMember = Nothing
'------------------------End Code-----------------------------
%>
	<p>The member has been edited. &nbsp;<a href="admin_members_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" or strSubmit = "Delete Member And Their Additions" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT MemberID FROM Customers WHERE ID = " & CustomerID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenForwardOnly, adLockReadOnly
		intOwnerID = rsUpdate("MemberID")
	rsUpdate.Close
	'We are trying to delete the site owner
	if intID = intOwnerID then
		set rsUpdate = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("You are trying to delete the owner of the site.  Sorry, but the owner cannot be deleted.  If you really want to delete the owner, e-mail support@GroupLoop.com, and we can do it."))
	end if

	Query = "SELECT ID FROM Members WHERE ID = " & intID & " AND " & strMatch
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close

	Query = "DELETE Reviews WHERE TargetTable = 'Members' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	if strSubmit = "Delete Member And Their Additions" then
		Query = "DELETE Announcements WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
		Query = "DELETE Calendar WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
		Query = "DELETE Stories WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
		Query = "DELETE VotingOptions WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
		Query = "DELETE VotingPolls WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
		Query = "DELETE PhotoCaptions WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
		Query = "DELETE QuizQuestions WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
		Query = "DELETE Quizzes WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
		Query = "DELETE Links WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		'Delete all their categories and messages in them
		Query = "SELECT ID FROM ForumCategories WHERE MemberID = " & intID
		rsUpdate.CacheSize = 100
		rsUpdate.Open Query, Connect, adOpenStatic, adLockReadOnly
		do until rsUpdate.EOF
			Query = "DELETE ForumMessages WHERE CategoryID = " & rsUpdate("ID")
			Connect.Execute Query, , adCmdText + adExecuteNoRecords
			rsUpdate.MoveNext
		loop
		rsUpdate.Close

		Query = "DELETE ForumCategories WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		Set rsSecondary = Server.CreateObject("ADODB.Recordset")
		rsSecondary.CacheSize = 100

		'Delete all their messages, making sure the replies are kept
		Query = "SELECT ID, BaseID FROM ForumMessages WHERE MemberID = " & intID & " ORDER BY ID"
		rsUpdate.Open Query, Connect, adOpenStatic, adLockReadOnly
		do until rsUpdate.EOF
			'If this is a base messages with replies, then get rid of this base and make the first reply a base
			if rsUpdate("BaseID") = 0 then
				Query = "SELECT ID, BaseID, IP, ModifiedID FROM ForumMessages WHERE MemberID <> " & intID & " AND BaseID = " & rsUpdate("ID") & " ORDER BY ID"
				rsSecondary.Open Query, Connect, adOpenStatic, adLockOptimistic
				if not rsSecondary.EOF then
					intNewBaseID = rsSecondary("ID")
					rsSecondary("BaseID") = 0
					rsSecondary("IP") = Request.ServerVariables("REMOTE_HOST")
					rsSecondary("ModifiedID") = Session("MemberID")
					rsSecondary.Update
					if rsSecondary.RecordCount > 1 then
						rsSecondary.MoveNext
						do until rsSecondary.EOF
							rsSecondary("IP") = Request.ServerVariables("REMOTE_HOST")
							rsSecondary("ModifiedID") = Session("MemberID")
							rsSecondary("BaseID") = intNewBaseID
							rsSecondary.Update
							rsSecondary.MoveNext
						loop
					end if
				end if
				rsSecondary.Close
			end if
			rsUpdate.MoveNext
		loop
		rsUpdate.Close

		Query = "DELETE ForumMessages WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		Set FileSystem = CreateObject("Scripting.FileSystemObject")

		strPath = GetPath ("photos")

		'Delete all their categories and photos in them
		Query = "SELECT ID FROM PhotoCategories WHERE MemberID = " & intID
		rsUpdate.Open Query, Connect, adOpenStatic, adLockReadOnly
		do until rsUpdate.EOF
			Query = "SELECT ID, Ext, ThumbnailExt FROM Photos WHERE CategoryID = " & rsUpdate("ID")
			rsSecondary.Open Query, Connect, adOpenStatic, adLockReadOnly
			do until rsSecondary.EOF
				strPhotoName = rsSecondary("ID") & rsSecondary("Ext")
				strThumbName = rsSecondary("ID") & "t." & rsSecondary("ThumbnailExt")
				if FileSystem.FileExists (strPath & "/" & strPhotoName) then FileSystem.DeleteFile (strPath & "/" & strPhotoName)
				if FileSystem.FileExists (strPath & "/" & strThumbName) then FileSystem.DeleteFile (strPath & "/" & strThumbName)
				rsSecondary.MoveNext
			loop
			rsSecondary.Close

			Query = "DELETE Photos WHERE CategoryID = " & rsUpdate("ID")
			Connect.Execute Query, , adCmdText + adExecuteNoRecords

			rsUpdate.MoveNext
		loop
		rsUpdate.Close

		Query = "DELETE PhotoCategories WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		Query = "SELECT ID, Ext, ThumbnailExt FROM Photos WHERE MemberID = " & intID
		rsSecondary.Open Query, Connect, adOpenStatic, adLockReadOnly
		do until rsSecondary.EOF
			Query = "DELETE PhotoCaptions WHERE PhotoID = " & rsSecondary("ID")
			Connect.Execute Query, , adCmdText + adExecuteNoRecords

			strPhotoName = rsSecondary("ID") & rsSecondary("Ext")
			strThumbName = rsSecondary("ID") & "t." & rsSecondary("ThumbnailExt")
			if FileSystem.FileExists (strPath & "/" & strPhotoName) then FileSystem.DeleteFile (strPath & "/" & strPhotoName)
			if FileSystem.FileExists (strPath & "/" & strThumbName) then FileSystem.DeleteFile (strPath & "/" & strThumbName)
			rsSecondary.MoveNext
		loop
		rsSecondary.Close

		Query = "DELETE Photos WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		strPath = GetPath ("media")

		'Delete all their categories and photos in them
		Query = "SELECT ID FROM MediaCategories WHERE MemberID = " & intID
		rsUpdate.Open Query, Connect, adOpenStatic, adLockReadOnly
		do until rsUpdate.EOF
			Query = "SELECT FileName FROM Media WHERE CategoryID = " & rsUpdate("ID")
			rsSecondary.Open Query, Connect, adOpenStatic, adLockReadOnly
			do until rsSecondary.EOF
				if FileSystem.FileExists (strPath & "/" & rsSecondary("FileName")) then FileSystem.DeleteFile (strPath & "/" & rsSecondary("FileName"))
				rsSecondary.MoveNext
			loop
			rsSecondary.Close

			Query = "DELETE Media WHERE CategoryID = " & rsUpdate("ID")
			Connect.Execute Query, , adCmdText + adExecuteNoRecords

			rsUpdate.MoveNext
		loop
		rsUpdate.Close

		Query = "DELETE MediaCategories WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		Query = "SELECT FileName FROM Media WHERE MemberID = " & intID
		rsSecondary.Open Query, Connect, adOpenStatic, adLockReadOnly
		do until rsSecondary.EOF
			if FileSystem.FileExists (strPath & "/" & rsSecondary("FileName")) then FileSystem.DeleteFile (strPath & "/" & rsSecondary("FileName"))
			rsSecondary.MoveNext
		loop
		rsSecondary.Close

		Query = "DELETE Media WHERE MemberID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		set rsSecondary = Nothing
		set rsUpdate = Nothing
		Set FileSystem = Nothing
	end if
'------------------------End Code-----------------------------
%>
	<p>The member has been deleted. &nbsp;<a href="admin_members_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete Member" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT MemberID FROM Customers WHERE ID = " & CustomerID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenForwardOnly, adLockReadOnly
		intOwnerID = rsUpdate("MemberID")
	rsUpdate.Close
	'We are trying to delete the site owner
	if intID = intOwnerID then
		set rsUpdate = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("You are trying to delete the owner of the site.  Sorry, but the owner cannot be deleted.  If you really want to delete the owner, e-mail support@GroupLoop.com, and we can do it."))
	end if

	Query = "SELECT ID FROM Members WHERE ID = " & intID & " AND " & strMatch
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'Members' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The member has been deleted. &nbsp;<a href="admin_members_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Suspend" or strSubmit = "Unsuspend" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT Suspended FROM Members WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	if rsUpdate("Suspended") = 0 then
		rsUpdate("Suspended") = 1
	else
		rsUpdate("Suspended") = 0
	end if
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	if strSubmit = "Suspend" then
		strDisplay = "suspended"
	else
		strDisplay = "unsuspended"
	end if
'------------------------End Code-----------------------------
%>
	<p>The member has been <%=strDisplay%>. &nbsp;<a href="admin_members_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM Members WHERE ID = " & intID & " AND " & strMatch
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsMember.EOF then
		set rsMember = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
%>
	* indicates required information<br>
	<a href="admin_members_modify.asp?ID=<%=intID%>&Submit=EMail">Click here</a> to e-mail them their name/password.<br>

	<form METHOD="post" ACTION="admin_members_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=intID%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What is their access level?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "Admin", 0, "Regular member access<br>", rsMember("Admin")
				PrintRadioOption "Admin", 1, "Administrator access<br>", rsMember("Admin")
%>
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Name & Such
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can see their name?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "PrivateName", 0, "Anybody<br>", rsMember("PrivateName")
				PrintRadioOption "PrivateName", 1, "Just Members<br>", rsMember("PrivateName")
				PrintRadioOption "PrivateName", 2, "Nobody<br>", rsMember("PrivateName")
%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* First Name
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="FirstName" value="<%=rsMember("FirstName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Last Name
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="LastName" value="<%=rsMember("LastName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* <%=UsernameLabel%>
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="NickName" value="<%=rsMember("NickName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* Password
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Password" value="<%=rsMember("Password")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can see their Birthday?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "PrivateBirthDate", 0, "Anybody<br>", rsMember("PrivateBirthDate")
				PrintRadioOption "PrivateBirthDate", 1, "Just Members<br>", rsMember("PrivateBirthDate")
				PrintRadioOption "PrivateBirthDate", 2, "Nobody<br>", rsMember("PrivateBirthDate")
%>			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Birthday
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "Birthdate", rsMember("Birthdate"), 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can see their E-Mail address?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "PrivateEMail", 0, "Anybody<br>", rsMember("PrivateEMail")
				PrintRadioOption "PrivateEMail", 1, "Just Members<br>", rsMember("PrivateEMail")
				PrintRadioOption "PrivateEMail", 2, "Nobody<br>", rsMember("PrivateEMail")
%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Primary E-Mail
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail1" value="<%=rsMember("EMail1")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Secondary E-Mail
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail2" value="<%=rsMember("EMail2")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Home Page (http://www.yourpage.com)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="WebSite" value="<%=rsMember("WebSite")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can see their beeper number?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "PrivateBeeper", 0, "Anybody<br>", rsMember("PrivateBeeper")
				PrintRadioOption "PrivateBeeper", 1, "Just Members<br>", rsMember("PrivateBeeper")
				PrintRadioOption "PrivateBeeper", 2, "Nobody<br>", rsMember("PrivateBeeper")
%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Beeper (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="Beeper" value="<%=rsMember("Beeper")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can see their cell phone number?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "PrivateCellPhone", 0, "Anybody<br>", rsMember("PrivateCellPhone")
				PrintRadioOption "PrivateCellPhone", 1, "Just Members<br>", rsMember("PrivateCellPhone")
				PrintRadioOption "PrivateCellPhone", 2, "Nobody<br>", rsMember("PrivateCellPhone")
%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Cell Phone (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="CellPhone" value="<%=rsMember("CellPhone")%>">
			</td>
		</tr>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Home Address
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can see their home address?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "PrivateHome", 0, "Anybody<br>", rsMember("PrivateHome")
				PrintRadioOption "PrivateHome", 1, "Just Members<br>", rsMember("PrivateHome")
				PrintRadioOption "PrivateHome", 2, "Nobody<br>", rsMember("PrivateHome")
%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeStreet" value="<%=rsMember("HomeStreet")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeCity" value="<%=rsMember("HomeCity")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "HomeState", rsMember("HomeState") %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeZip" value="<%=rsMember("HomeZip")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="HomePhone" value="<%=rsMember("HomePhone")%>">

			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Country
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="HomeCountry" value="<%=rsMember("HomeCountry")%>">
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Secondary Address (optional - school, current residence, etc)
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can see their secondary address?
			</td>
			<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "PrivateSecondary", 0, "Anybody<br>", rsMember("PrivateSecondary")
				PrintRadioOption "PrivateSecondary", 1, "Just Members<br>", rsMember("PrivateSecondary")
				PrintRadioOption "PrivateSecondary", 2, "Nobody<br>", rsMember("PrivateSecondary")
%>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Description (school, work, beach, etc.)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryDescription" value="<%=rsMember("SecondaryDescription")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryStreet" value="<%=rsMember("SecondaryStreet")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryCity" value="<%=rsMember("SecondaryCity")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "SecondaryState", rsMember("SecondaryState") %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryZip" value="<%=rsMember("SecondaryZip")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone (xxx.xxx.xxxx)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="SecondaryPhone" value="<%=rsMember("SecondaryPhone")%>">
				&nbsp;&nbsp;&nbsp;ext. <input type="text" size="4" name="SecondaryPExt" value="<%=rsMember("SecondaryPExt")%>">				 
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Country
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="SecondaryCountry" value="<%=rsMember("SecondaryCountry")%>">
			</td>
		</tr>
<%
	if IncludeNewsletter then
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<%=NewsletterTitle%> Subscription
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Subscribe to the newsletter?
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintRadio rsMember("SubscribeSiteNewsletter"), "SubscribeSiteNewsletter" %>
			</td>
		</tr>
		
<%
	end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Update">
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsMember.Close
	Set rsMember = Nothing
elseif strSubmit = "EMail" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT Password, NickName, FirstName, LastName, EMail1 FROM Members WHERE ID = " & intID & " AND " & strMatch
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsMember.EOF then
		set rsMember = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The member you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	strRecipName = rsMember("FirstName") & " " & rsMember("LastName")
	strNickName = rsMember("NickName")
	strEMail = rsMember("EMail1")
	strPassword = rsMember("Password")

	rsMember.Close


	if strEMail = "" then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but the member does not have an e-mail address entered, so it couldn't be sent."))


	Query = "SELECT Subdirectory, UseDomain, DomainName FROM Customers WHERE ID = " & CustomerID
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsMember.EOF then
		set rsMember = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The customer you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	blUseDomain = CBool(rsMember("UseDomain"))
	if blUseDomain then
		strURL = rsMember("DomainName")
	else
		strURL = "http://www.GroupLoop.com/" & rsMember("Subdirectory")
	end if

	rsMember.Close
	set rsMember = Nothing

	strSubject = "Your GroupLoop.com site membership info"

	strBody = "Dear " & strRecipName & "," & VbCrLf & VbCrLf & _
	"Here is your membership info for '" & Title & "' in case you lost it: " & VbCrLf & _
	"Your site address: " & strURL & VbCrLf & _
	"Your " & UsernameLabel & ": " & strNickName & VbCrLf & _
	"Your password: " & strPassword & VbCrLf & VbCrLf & _
	"Please do not respond to this e-mail.  Your questions may be answered by your site administrator(s) " & _
	"or the GroupLoop.com staff." & VbCrLf & _
	"You may e-mail the GroupLoop.com staff at: support@grouploop.com" & VbCrLf & VbCrLf & _
	"                                     Thank you and enjoy," & VbCrLf & _
	"                                         Stephen Potter" & VbCrLf & _
	"                                         President, GroupLoop.com" & VbCrLf & VbCrLf & _
	"Please read GroupLoop.com's Terms Of Service at http://www.GroupLoop.com/homegroup/tos.asp .  Your signing into your site verifies that you have read and accept the Terms Of Service, so please read it carefully." & VbCrLf

	'Set the rest of the mailing info and send it
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.IgnoreMalformedAddress = true
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = MailerFromName
	Mailer.FromAddress = "support@grouploop.com"
	Mailer.AddRecipient strRecipName, strEMail
	Mailer.Subject    = FormatEdit(strSubject)
	Mailer.BodyText   = strBody

	if not Mailer.SendMail then 
%>		<p>There has been an error, and the email has not been sent to the new member.  Please make sure you had a valid e-mail address entered.  Try again, and if the problem
		persists, e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a>.  Please include the error below.<br>
		Error was '<%=Mailer.Response%>'</p>
<%	else	%>
		<p>The e-mail has been sent. &nbsp;<a href="admin_members_modify.asp">Click here</a> to modify another member.</p>
<%
	end if
	Set Mailer = Nothing
else
	Query = "SELECT ID, FirstName, LastName, NickName, Suspended FROM Members WHERE CustomerID = " & CustomerID & " AND ID <> " & Session("MemberID") & " ORDER BY LastName"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsPage.EOF then
		Set rsPage = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("You have to have members before you can modify them, " & GetNickNameSession))
	end if

	Set ID = rsPage("ID")
	Set FirstName = rsPage("FirstName")
	Set LastName = rsPage("LastName")
	Set NickName = rsPage("NickName")
	Set Suspended = rsPage("Suspended")
'-----------------------End Code----------------------------
%>
	<form METHOD="POST" ACTION="admin_members_modify.asp">
<%
'-----------------------Begin Code----------------------------
	PrintPagesHeader
'-----------------------End Code----------------------------
%>
	<p>If you click 'Delete All Additions Also', everything the member ever created will be deleted.</p>
	<%PrintTableHeader 0%>
	<tr>
		<td class="TDHeader">Name</td>
		<td class="TDHeader">Nickname</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
	for p = 1 to rsPage.PageSize
		if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
		<form METHOD="post" ACTION="admin_members_modify.asp">
		<input type="hidden" name="ID" value="<%=ID%>">
			<tr>
				<td class="<% PrintTDMain %>"><%=FirstName%>&nbsp;<%=LastName%></td>
				<td class="<% PrintTDMain %>"><%=NickName%></td>
				<td class="<% PrintTDMainSwitch %>">
				<input type="submit" name="Submit" value="Edit"> 
				<input type="button" value="Delete Member And Their Additions" onClick="DeleteBox('If you delete this member, there is no way to get them or any of their additions back.  Are you sure?', 'admin_members_modify.asp?Submit=Delete+Member+And+Their+Additions&ID=<%=ID%>')"> 
				<input type="button" value="Delete Member" onClick="DeleteBox('If you delete this member, there is no way to get them back.  Are you sure?', 'admin_members_modify.asp?Submit=Delete+Member&ID=<%=ID%>')">
<%
				if Suspended = 0 then
%>
					<input type="submit" name="Submit" value="Suspend">  
<%
				else
%>
					<input type="submit" name="Submit" value="Unsuspend"> 
<%
				end if
				if ReviewsExist( "Members", ID ) then
%>
					<input type="button" value="Modify Reviews" onClick="Redirect('admin_reviews_modify.asp?Source=admin_members_modify.asp&TargetTable=Members&TargetID=<%=ID%>')">
<%
				end if
%>					
				</td>
			</tr>
		</form>
<%
'-----------------------Begin Code----------------------------
		rsPage.MoveNext
		end if
	next
	Response.Write("</table>")
	rsPage.Close
	set rsPage = Nothing
end if


Function GetReverseRadio( intbool )
	if intbool = 0 then GetReverseRadio = 1
	if intbool = 1 then GetReverseRadio = 0
End Function

'------------------------End Code-----------------------------
%>