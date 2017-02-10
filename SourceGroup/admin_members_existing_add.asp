<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_members_existing_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add A Member</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
strMatch = "CustomerID = " & CustomerID

strSubmit = Request("Submit")

if strSubmit = "Add" then
	if Request("ID") = "" or Request("CommonID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intCopyID = CInt(Request("ID"))
	intCommonID = CInt(Request("CommonID"))

	Set cmdNick = Server.CreateObject("ADODB.Command")
	With cmdNick
		.ActiveConnection = Connect
		.CommandText = "AddMemberExisting"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CopyID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CommonID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Error", adInteger, adParamOutput )

		.Parameters("@CopyID") = intCopyID
		.Parameters("@CommonID") = intCommonID
		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		intError = .Parameters("@Error")
	End With
	Set cmdNick = Nothing

	if intError = 1 then Redirect("error.asp?Message=" & Server.URLEncode("The copy member doesn't exist."))

'------------------------End Code-----------------------------
%>
	<p>The member has been added. &nbsp;<a href="admin_members_existing_add.asp">Click here</a> to add another.</p>
<%
'-----------------------Begin Code----------------------------
else


	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetMultiSiteMembers"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )

		.Parameters("@CustomerID") = CustomerID

		rsPage.Open cmdTemp, , adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

		'We have all the distinct commonIDs, so check if the member is a member of this subsite also
		'if so, we don't include them
		.CommandText = "MemberCommonCustIDExists"

		.Parameters.Append .CreateParameter ("@CommonID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )
		.Parameters("@CustomerID") = CustomerID
	End With


	Function ThisSiteMemberExists( intCommonID )
		intCommonID = Int( intCommonID )
		With cmdTemp
			.Parameters("@CommonID") = intCommonID

			.Execute , , adExecuteNoRecords

			intExists = .Parameters("@Exists")

		End With

		ThisSiteMemberExists = CBool( intExists )
	End Function


	if not rsPage.EOF then
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">Name</td>
			<td class="TDHeader">NickName</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		strNickName = ""

		do until rsPage.EOF
			if rsPage("NickName") <> strNickName then
				strNickName = rsPage("NickName")
				if not ThisSiteMemberExists( rsPage("CommonID") ) then
'------------------------End Code-----------------------------
%>
					<form METHOD="post" ACTION="admin_members_existing_add.asp">
					<input type="hidden" name="ID" value="<%=rsPage("ID")%>">
					<input type="hidden" name="CommonID" value="<%=rsPage("CommonID")%>">
						<tr>
							<td class="<% PrintTDMain %>"><%=rsPage("FirstName")%>&nbsp;<%=rsPage("LastName")%></td>
							<td class="<% PrintTDMain %>"><%=rsPage("NickName")%></td>
							<td class="<% PrintTDMainSwitch %>">
							<input type="Submit" name="Submit" value="Add"> 
							</td>
						</tr>
					</form>
<%
'-----------------------Begin Code----------------------------
				end if
			end if
			rsPage.MoveNext
		loop
		Response.Write("</table>")
		rsPage.Close
	else
'------------------------End Code-----------------------------
%>
		<p>You have no existing members to add.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	if IsObject(cmdTemp) then Set cmdTemp = Nothing

	set rsPage = Nothing
end if


'------------------------End Code-----------------------------
%>