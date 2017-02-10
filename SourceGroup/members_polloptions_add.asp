<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeVoting ) or CBool( VotingMembers ) ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_polloptions_add.asp")
blLoggedAdmin = LoggedAdmin
if not (LoggedAdmin or CBool( VotingMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add Voting Poll Answers</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
intID = CInt(Request("ID"))

if Request("Submit") = "Add" then
	Query = "SELECT ID FROM VotingPolls WHERE ID = " & intID & " AND " & strMatch
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	rsNew.Open Query, Connect, adOpenForwardOnly, adLockOptimistic
	if rsNew.EOF then
		Set rsNew = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsNew.Close

	Query = "SELECT Name, MemberID, PollID, CustomerID, IP, ModifiedID FROM VotingOptions"
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic

	for i = 1 to 6
		if Request(i) <> "" then
			rsNew.AddNew
				rsNew("Name") = Format( Request(i) )
				rsNew("MemberID") = Session("MemberID")
				rsNew("PollID") = Request("ID")
				rsNew("CustomerID") = CustomerID
				rsNew("IP") = Request.ServerVariables("REMOTE_HOST")
				rsNew("ModifiedID") = Session("MemberID")
			rsNew.Update
		end if
	next
	rsNew.Close
	Set rsNew = Nothing
'------------------------End Code-----------------------------
%>
	<p>The answer(s) have been added. &nbsp;<a href="members_polloptions_add.asp?ID=<%=Request("ID")%>">Click here</a> to add more.  &nbsp;<a href="members_polls_modify.asp">Click here</a> to modify another poll.</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	<p>You may add up to six answers at a time.  Whatever answers you don't need, leave blank.</p>
	<form METHOD="post" ACTION="<%=SecurePath%>members_polloptions_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Poll
		</td>
		<td class="<% PrintTDMain %>">
<%
'-----------------------Begin Code----------------------------
		PrintPollPullDown intID, Session("MemberID")
'------------------------End Code-----------------------------
%>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Answer 1
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="1">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Answer 2
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="2">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Answer 3
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="3">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Answer 4
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="4">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Answer 5
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="5">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Answer 6
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="6">
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Add">
		</td>
	</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if


'-------------------------------------------------------------
'This function writes a pulldown menu for voting polls
'-------------------------------------------------------------
Sub PrintPollPullDown( intPollID, intMemberID )
	intPollID = CInt(intPollID)
	'Now we are going to get the group names to list in the pull-down menu
	if intMemberID = 0 then
		Query = "SELECT ID, Subject FROM VotingPolls WHERE (CustomerID = " & CustomerID & ")"
	else
		Query = "SELECT ID, Subject FROM VotingPolls WHERE (MemberID = " & intMemberID & " AND CustomerID = " & CustomerID & ")"
	end if
	Set rsTempPolls = Server.CreateObject("ADODB.Recordset")
	rsTempPolls.CacheSize = 20
	rsTempPolls.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsTempPolls.EOF then
		set rsTempPolls = Nothing
		exit sub
	end if

	'Make the size 3 if there are many members
	if rsTempPolls.RecordCount <= 30 then
		%><select name="ID" size="1"><%
	else
		%><select name="ID" size="3"><%
	end if

	Set ID = rsTempPolls("ID")
	Set Subject = rsTempPolls("Subject")

	do until rsTempPolls.EOF
		'Highlight the current section
		if ID = intPollID then
			Response.Write "<option value = '" & ID & "' SELECTED>" & Subject & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & ID & "'>" & Subject & "</option>" & vbCrlf
		end if
		rsTempPolls.MoveNext
	loop
	rsTempPolls.Close
	set rsTempPolls = Nothing
	Response.Write("</select>")
End Sub
'------------------------End Code-----------------------------
%>
