<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember then Redirect("members.asp?Source=members_info_edit.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p class="Heading" align="<%=HeadingAlignment%>">Change Your Info</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------

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
strSubmit = Request("Submit")
if strSubmit = "Update" or strSubmit = "Update All My Membership Records" or strSubmit = "Update Just This Membership" then
	strNickName = Format(Request("NickName"))

	if Request("FirstName") = "" or Request("LastName") = "" or strNickName = "" or Request("Password") = "" then Redirect("incomplete.asp")

	if UCase(GetNickNameOnly(Session("MemberID"))) <> UCase(strNickName) AND NickNameTaken( strNickName, CustomerID ) then
		Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but that " & UsernameLabel & " is already taken by another member.  Try another nickname."))
	end if

	Query = "SELECT * FROM Members WHERE ID = " & Session("MemberID") & " AND CustomerID = " & CustomerID
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsMember.EOF then
		set rsMember = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("We can't find your member record."))
	end if

	'We are updating multiple memberships (linked by the CommonID, which must be >0)
	if strSubmit = "Update All My Membership Records" then
		if rsMember("CommonID") = 0 then
			set rsMember = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("You don't have any linking membership records."))
		end if
		'Get the linking ID
		intCommonID = rsMember("CommonID")

		'Close then reopen the recordset with multiple recs
		rsMember.Close
		Query = "SELECT * FROM Members WHERE CommonID = " & intCommonID
		rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
		do until rsMember.EOF

			SetData

			rsMember.MoveNext
		loop

	else
		'Set the data, you fucking retard
		SetData
	end if

	'Set the new PW
	Session("Password") = Request("Password")

	if Request.Cookies("SiteNum"&CustomerID)("AutoLogin") = "1" then
		Response.Cookies("SiteNum"&CustomerID)("NickName") = GetJustNickNameSession()
		Response.Cookies("SiteNum"&CustomerID)("Password") = Session("Password")
	end if

	'Close dis bitch
	rsMember.Close
	set rsMember = Nothing
'------------------------End Code-----------------------------
%>
	<!-- #include file="write_index.asp" -->
	<p>Your info has been edited. &nbsp;<a href="members_info_edit.asp">Click here</a> to edit it again.</p>
<%
'-----------------------Begin Code----------------------------

else
	Query = "SELECT * FROM Members WHERE ID = " & Session("MemberID")
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic
'------------------------End Code-----------------------------
%>

	<p>Here you can change all your information.  You may allow non-members to see your information by checking the various boxes throughout this page.</p>

	* indicates required information<br>

	<form METHOD="post" ACTION="members_info_edit.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=Request("ID")%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Name & Such
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who can see your name?
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
				* Nickname
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
				Who can see your Birthday?
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
				Who can see your E-Mail address?
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
				Who can see your beeper number?
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
				Who can see your cell phone number?
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
				Who can see your home address?
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
				Who can see your secondary address?
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

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">

<%
	'They have linked records, so allow them to update all their records
	if rsMember("CommonID") > 0 and CommonMember( rsMember("CommonID") ) > 1 then
%>
		<input type="submit" name="Submit" value="Update All My Membership Records"> 
		<input type="submit" name="Submit" value="Update Just This Membership">
<%
	else
%>
		<input type="submit" name="Submit" value="Update">
<%
	end if
%>
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	set rsMember = Nothing
end if



'This member has more than one membership
'The common ID is the ID of their first member record.  Each one after that gets the same commonid
'if the first record is lost, the commonID is still unique, so no big deal
Function CommonMember( intCommonID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetNumCommonMembers"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CommonID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Count", adInteger, adParamOutput )

		.Parameters("@CommonID") = intCommonID

		.Execute , , adExecuteNoRecords
		intCount = .Parameters("@Count")
	End With
	Set cmdTemp = Nothing
	CommonMember = intCount

End Function


Sub SetData
	rsMember("PrivateName") = Request("PrivateName")
	rsMember("PrivateBirthdate") = Request("PrivateBirthdate")
	rsMember("PrivateEMail") = Request("PrivateEMail")
	rsMember("PrivateHome") = Request("PrivateHome")
	rsMember("PrivateSecondary") = Request("PrivateSecondary")
	rsMember("PrivateBeeper") = Request("PrivateBeeper")
	rsMember("PrivateCellPhone") = Request("PrivateCellPhone")

	rsMember("FirstName") = Request("FirstName")
	rsMember("LastName") = Request("LastName")
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

	rsMember.Update
End Sub
'------------------------End Code-----------------------------
%>
