<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeVoting ) or CBool( VotingMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember then Redirect("members.asp?Source=members_polls_close.asp")
blLoggedAdmin = LoggedAdmin
if not (blLoggedAdmin or CBool( VotingMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Close A Poll</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if


'update info
if Request("Submit") = "Close" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT OpenToVote, IP, ModifiedID FROM VotingPolls WHERE ID = " & intID & " AND " & strMatch
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("OpenToVote") = 0
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The poll has been closed. &nbsp;If you want to re-open it or delete it, just goto Modify Voting Polls, and change it there. &nbsp;<a href="members_polls_close.asp">Click here</a> to close another.</p>
<%
'-----------------------Begin Code----------------------------
else

	Query = "SELECT ID, Date, Subject FROM VotingPolls WHERE (OpenToVote = 1 AND " & strMatch & ") ORDER BY Date DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Subject = rsPage("Subject")
%>
		<form METHOD="POST" ACTION="members_polls_close.asp">
<%
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">&nbsp;</td>
			<td class="TDHeader">Subject</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<form METHOD="post" ACTION="members_polls_close.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="voting_results.asp?ID=<%=ID%>">View</a></td>
						<td class="<% PrintTDMain %>"><%=Subject%></td>
						<td class="<% PrintTDMainSwitch %>"><input type="Submit" name="Submit" value="Close"></td>
					</tr>
				</form>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
	else
'------------------------End Code-----------------------------
%>
		<p>You have to open a poll before you can close it, <%=GetNickNameSession()%>.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if


'------------------------End Code-----------------------------
%>