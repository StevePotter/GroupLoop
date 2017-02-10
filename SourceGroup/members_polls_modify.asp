<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeVoting ) or CBool( VotingMembers ) ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_polls_modify.asp")
blLoggedAdmin = LoggedAdmin
if not (blLoggedAdmin or CBool( VotingMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Voting Polls</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------

if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

strSubmit = Request("Submit")

'update info
if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM VotingPolls WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	if Request("VoteLimit") = "Yes" then
		if Request("MaxType") = "Daily" then
			rsUpdate("MaxDailyVotes") = CInt(Request("MaxVotes"))
			rsUpdate("MaxTotalVotes") = 0
		else
			rsUpdate("MaxTotalVotes") = CInt(Request("MaxVotes"))
			rsUpdate("MaxDailyVotes") = 0
		end if
	else
		rsUpdate("MaxDailyVotes") = 0
		rsUpdate("MaxTotalVotes") = 0
	end if
	if LoggedAdmin then rsUpdate("Date") = AssembleDate("Date")
	
	rsUpdate("OpenToVote") = Request("OpenToVote")
	rsUpdate("Private") = Request("Private")
	rsUpdate("ResultsSecurity") = Request("ResultsSecurity")
	rsUpdate("Subject") = Format( Request("Subject") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close

	Query = "SELECT ID, Name FROM VotingOptions WHERE PollID = " & Request("ID")
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	do until rsUpdate.EOF
		if Request("Delete"&rsUpdate("ID")) = "Delete" or Request(rsUpdate("ID")) = "" then
			rsUpdate.Delete
		else
			rsUpdate("Name") = Format( Request(rsUpdate("ID")) )
		end if
		rsUpdate.Update
		rsUpdate.MoveNext
	loop

	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The poll has been edited.  &nbsp;<a href="members_polls_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM VotingPolls WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'VotingPolls' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	Query = "DELETE VotingOptions WHERE PollID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The poll has been deleted.  &nbsp;<a href="members_polls_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Subject, Private, MaxDailyVotes, MaxTotalVotes, OpenToVote, ResultsSecurity FROM VotingPolls WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsEdit.EOF then
		set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if rsEdit("Private") = 1 then 
		strPrivateChecked = "checked"
	else
		strPrivateChecked = ""
	end if


	strRChecked1 = strRChecked2 = strRChecked3 = ""


	if rsEdit("ResultsSecurity") = 0 then strRChecked1 = "checked"
	if rsEdit("ResultsSecurity") = 1 then strRChecked2 = "checked"
	if rsEdit("ResultsSecurity") = 2 then strRChecked3 = "checked"

	strLimitChecked1 = strLimitChecked2 = ""
	strMaxTypeChecked1 = strMaxTypeChecked2 = ""
	intLimit = ""
	if rsEdit("MaxDailyVotes") = 0 AND rsEdit("MaxTotalVotes") = 0 then 
		strLimitChecked2 = "checked"
	else
		strLimitChecked1 = "checked"
		if rsEdit("MaxDailyVotes") > 0 then
			intLimit = rsEdit("MaxDailyVotes")
			strMaxTypeChecked1 = "selected"
		else
			intLimit = rsEdit("MaxTotalVotes")
			strMaxTypeChecked2 = "selected"
		end if
	end if

'------------------------End Code-----------------------------
%>
	<p><a href="members_polloptions_add.asp?ID=<%=rsEdit("ID")%>">Click here</a> to add answers to this poll.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_polls_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<% PrintTableHeader 0%>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			* Poll Name
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="Subject" value="<%=FormatEdit( rsEdit("Subject") )%>">
		</td>
	</tr>
<%	if blLoggedAdmin then %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<%DatePulldown "Date", rsEdit("Date"), 1 %>
     	</td>
    </tr>
<%	end if %>

	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Is the poll open to voting?
		</td>
		<td class="<% PrintTDMain %>">
				<% PrintRadio rsEdit("OpenToVote"), "OpenToVote" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Can only members vote on it?
		</td>
		<td class="<% PrintTDMain %>">
				<% PrintRadio rsEdit("Private"), "Private" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Who can view the results?
		</td>
		<td class="<% PrintTDMain %>">
			<input type="radio" name="ResultsSecurity" value="0" <%=strRChecked1%>>Anyone 
			<input type="radio" name="ResultsSecurity" value="1" <%=strRChecked2%>>Members Only 
			<input type="radio" name="ResultsSecurity" value="2" <%=strRChecked3%>>Just Administrators
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Is there a vote limit?
		</td>
		<td class="<% PrintTDMain %>">
			<input type="radio" name="VoteLimit" value="Yes" <%=strLimitChecked1%>>Yes 
			<input type="radio" name="VoteLimit" value="No" <%=strLimitChecked2%>>No 
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			If there is a limit, 
		</td>
		<td class="<% PrintTDMain %>" valign="middle">
			allow <input type="text" name="MaxVotes" size="2" value="<%=intLimit%>"> votes
			<select size="1" name="MaxType">
				<option value="Daily" <%=strMaxTypeChecked1%>>per day</option>
				<option value="Total" <%=strMaxTypeChecked2%>>total</option>
			</select>.
		</td>
	</tr>
<%
'-----------------------Begin Code----------------------------
	Query = "SELECT ID, Name FROM VotingOptions WHERE PollID = " & Request("ID") & " ORDER BY ID"
	Set rsOptions = Server.CreateObject("ADODB.Recordset")
	rsOptions.CacheSize = 40
	rsOptions.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsOptions.EOF then
		Set ID = rsOptions("ID")
		Set Name = rsOptions("Name")
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" align=center colspan=2>
				You may modify the answers to the poll below.
			</td>
		</tr>
		<tr>
			<td class="TDHeader" align=center>
				Delete?
			</td>
			<td class="TDHeader" align=center>
				Answer
			</td>
		</tr>
<%
'-----------------------Begin Code----------------------------
		do until rsOptions.EOF
'------------------------End Code-----------------------------
%>
			<tr>
				<td class="<% PrintTDMain %>"  align="center">
					<input type=checkbox name="Delete<%=ID%>" value="Delete">
				</td>
				<td class="<% PrintTDMain %>">
					<input type="text" size="50" name="<%=ID%>" value="<%=FormatEdit( rsOptions("Name") )%>">
				</td>
			</tr>
<%
'-----------------------Begin Code----------------------------
			rsOptions.MoveNext
		loop
		rsOptions.Close
	end if
	set rsOptions = Nothing
'------------------------End Code-----------------------------
%>

	<tr>
		<td colspan="2" class="<% PrintTDMain %>" align="center">
			<input type="submit" name="Submit" value="Update">
		</td>
	</tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing

else
	Query = "SELECT ID, Date, Subject FROM VotingPolls WHERE (CustomerID = " & CustomerID & " AND " & strMatch & ") ORDER BY Date DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Subject = rsPage("Subject")
%>
		<form METHOD="POST" ACTION="members_polls_modify.asp">
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
				<form METHOD="post" ACTION="members_polls_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="voting_results.asp?ID=<%=ID%>">View</a></td>
						<td class="<% PrintTDMain %>"><%=Subject%></td>
						<td class="<% PrintTDMainSwitch %>">
						<input type="submit" name="Submit" value="Edit"> 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this voting poll, there is no way to get it back.  Are you sure?', 'members_polls_modify.asp?Submit=Delete&ID=<%=ID%>')">
						<%if ReviewsExist( "VotingPolls", ID ) AND blLoggedAdmin then%>
							<input type="button" value="Modify Reviews" onClick="Redirect('admin_reviews_modify.asp?Source=members_polls_modify.asp&TargetTable=VotingPolls&TargetID=<%=ID%>')">
						<%end if%>
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
	else
'------------------------End Code-----------------------------
%>
		<p>Sorry, but there are no voting polls to modify.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if


'------------------------End Code-----------------------------
%>