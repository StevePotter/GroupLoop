<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeVoting ) or CBool( VotingMembers ) ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_polls_add.asp")
if not (LoggedAdmin or CBool( VotingMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add A New Voting Poll</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Add" then
	if Request("Subject") = "" then Redirect("incomplete.asp")
	'Create our connection object and open a connection to our database
	Query = "SELECT * FROM VotingPolls"
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	'Get all the records
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
	rsNew.AddNew
		if Request("VoteLimit") = "Yes" then
			if Request("MaxType") = "Daily" then
				rsNew("MaxDailyVotes") = CInt(Request("MaxVotes"))
			else
				rsNew("MaxTotalVotes") = CInt(Request("MaxVotes"))
			end if
		end if

		rsNew("MemberID") = Session("MemberID")
		rsNew("ModifiedID") = Session("MemberID")
		rsNew("Private") = Request("Private")
		rsNew("ResultsSecurity") = Request("ResultsSecurity")
		rsNew("Subject") = Format( Request("Subject") )
		rsNew("CustomerID") = CustomerID
		rsNew("IP") = Request.ServerVariables("REMOTE_HOST")
	rsNew.Update
	rsNew.MovePrevious
	rsNew.MoveNext
	intID = rsNew("ID")
	rsNew.Close

	'Now add the voting options
	Query = "SELECT Name, MemberID, PollID, CustomerID, IP, ModifiedID FROM VotingOptions"
	'Get all the records
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic

	for i = 1 to 6
		if Request(i) <> "" then
			rsNew.AddNew
				rsNew("Name") = Format( Request(i) )
				rsNew("MemberID") = Session("MemberID")
				rsNew("PollID") = intID
				rsNew("CustomerID") = CustomerID
				rsNew("IP") = Request.ServerVariables("REMOTE_HOST")
				rsNew("ModifiedID") = Session("MemberID")
			rsNew.Update
		end if
	next
	Set rsNew = Nothing

'------------------------End Code-----------------------------
%>
	<p>The poll has been added. &nbsp;<a href="members_polloptions_add.asp?ID=<%=intID%>">Click here</a> to add more options to it.<br>
	<a href="members_polls_add.asp">Click here</a> to add another poll.
	</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	* indicates required information<br>
	<form METHOD="post" ACTION="<%=SecurePath%>members_polls_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			* Poll Name
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="Subject">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Can only members vote on it?
		</td>
		<td class="<% PrintTDMain %>">
			<input type="radio" name="Private" value="1">Yes 
			<input type="radio" name="Private" value="0" checked>No 
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Who can view the results?
		</td>
		<td class="<% PrintTDMain %>">
			<input type="radio" name="ResultsSecurity" value="0" checked>Anyone 
			<input type="radio" name="ResultsSecurity" value="1">Members Only 
			<input type="radio" name="ResultsSecurity" value="2">Just Administrators
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Is there a vote limit?
		</td>
		<td class="<% PrintTDMain %>">
			<input type="radio" name="VoteLimit" value="Yes" checked>Yes 
			<input type="radio" name="VoteLimit" value="No">No 
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			If there is a limit, 
		</td>
		<td class="<% PrintTDMain %>" valign="middle">
			allow <input type="text" name="MaxVotes" size="2" value="2"> votes
			<select size="1" name="MaxType">
				<option value="Daily">per day</option>
				<option value="Total">total</option>
			</select>.
		</td>
	</tr>
	<tr>
		<td class="TDHeader" colspan=2  align=center>
			Poll Answers - You may add up to 6 answers right now, but you can add more right after this.
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Option 1
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="1">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Option 2
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="2">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Option 3
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="3">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Option 4
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="4">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Option 5
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="5">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Option 6
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="6">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="center" colspan="2">
			<input type="submit" name="Submit" value="Add">
		</td>
	</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if
'------------------------End Code-----------------------------
%>
