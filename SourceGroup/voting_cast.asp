<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeVoting ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading><%=VotingTitle%></span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span><br>

<%
intID = Request("ID")
if intID = "" then Redirect("error.asp?Disable=yes&Message=" & Server.URLEncode("There was a problem with your vote.  You probably did not select an option to vote on.  <a href='javascript:history.back(1)'>Click here to go back and try again</a>."))
intID = CInt(intID)

'Open up the item
if Request("PollID") <> "" then
	intPollID = CInt(Request("PollID"))
	Query = "SELECT MaxDailyVotes, Private, Subject, ResultsSecurity FROM VotingPolls WHERE OpenToVote = 1 AND ID = " & intPollID & " AND CustomerID = " & CustomerID
else
	Query = "SELECT MaxDailyVotes, Private, Subject, ResultsSecurity FROM VotingPolls WHERE OpenToVote = 1 AND ID = " & intID & " AND CustomerID = " & CustomerID
end if
Set rsPoll = Server.CreateObject("ADODB.Recordset")
rsPoll.Open Query, Connect, adOpenStatic, adLockReadOnly

ResultsSecurity = rsPoll("ResultsSecurity")
blLoggedMember = LoggedMember
blLoggedAdmin = LoggedAdmin

if not (( ResultsSecurity = 1 AND not blLoggedMember ) or (ResultsSecurity = 2 AND not blLoggedAdmin)) then
%>
<span class=LinkText><a href="voting_results.asp?ID=<%=Request("ID")%>">View Current Results</a></span>
<%
end if
%>

</p>

<%
'-----------------------Begin Code----------------------------


if rsPoll.EOF then
	Set rsPoll = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The poll does not exist.  If you pasted a link, there may be a typo, or the poll may have been deleted.  Please refer to the poll list to find the desired poll, if it still exists."))
end if

if rsPoll("Private") = 1 AND not LoggedMember then
	rsPoll.Close
	Set rsPoll = Nothing
	Redirect( "login.asp?Source=voting_cast.asp&ID=" & intID & "&Submit=Vote" )
end if

if Request("Submit") = "Cast My Vote" then
	rsPoll.Close
	set rsPoll = Nothing

	if not AllowVote( intPollID ) then Redirect("voting_cast.asp?ID=" & intPollID)

	IncrementStat "VotesCast"
	IncrementHits intID, "VotingOptions"

	'Add this to the list of responses
	Query = "SELECT OptionID, CustomerID, PollID, IP, MemberID FROM VotingResponses"
	Set rsResponses = Server.CreateObject("ADODB.Recordset")
	rsResponses.Open Query, Connect, adOpenStatic, adLockOptimistic
	rsResponses.AddNew
		rsResponses("OptionID") = intID
		rsResponses("CustomerID") = CustomerID
		rsResponses("PollID") = intPollID
		rsResponses("IP") = Request.ServerVariables("REMOTE_HOST")
		if LoggedMember then rsResponses("MemberID") = Session("MemberID")
	rsResponses.Update
	rsResponses.Close
	set rsResponses = Nothing

	Redirect("voting_results.asp?ID=" & intPollID)
else
	Query = "SELECT ID, Name FROM VotingOptions WHERE PollID = " & intID & " ORDER BY ID"
	Set rsOptions = Server.CreateObject("ADODB.Recordset")
	rsOptions.CacheSize = 20
	rsOptions.Open Query, Connect, adOpenStatic, adLockReadOnly

	'They can't vote - no options
	if rsOptions.EOF then
'------------------------End Code-----------------------------
%>
		<p>Sorry, but there are no options in this poll, so you can't vote on it.</p>
<%
'-----------------------Begin Code----------------------------
	'They have already voted enough
	elseif not AllowVote( intID ) then
		if rsPoll("MaxDailyVotes") > 0 then
'------------------------End Code-----------------------------
%>
			<p>Sorry, you have voted the maximum number of times today.  You may vote again tomorrow.</p>
<%
'-----------------------Begin Code----------------------------
		else
'------------------------End Code-----------------------------
%>
			<p>Sorry, you have voted the maximum number of times.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	else
		IncrementHits intID, "VotingPolls"
'------------------------End Code-----------------------------
%>
		<form method="post" action="voting_cast.asp">
		<input type="hidden" name="PollID" value="<%=intID%>">
		<p><span class=Heading><%=rsPoll("Subject")%></span><br>
<%
'-----------------------Begin Code----------------------------
		do until rsOptions.EOF
'------------------------End Code-----------------------------
%>
				&nbsp;&nbsp;&nbsp; <input type="radio" name="ID" value="<%=rsOptions("ID")%>"><%=rsOptions("Name")%><br>
<%
'-----------------------Begin Code----------------------------	
			rsOptions.MoveNext
		loop
		rsOptions.Close
'------------------------End Code-----------------------------
%>
		<br><input type="submit" name="Submit" value="Cast My Vote"></p>
		</form>
<%
'-----------------------Begin Code----------------------------
	end if
	set rsOptions = Nothing

	rsPoll.Close
	set rsPoll = Nothing
end if


'-------------------------------------------------------------
'This function tells if the person can vote or not
'-------------------------------------------------------------
Function AllowVote( intID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "PollsAllowVote"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@PollID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@IP", adVarWChar, adParamInput, 50 )
		.Parameters.Append .CreateParameter ("@Allow", adInteger, adParamOutput )

		.Parameters("@PollID") = intID
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@MemberID") = CInt(Session("MemberID"))
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")

		.Execute , , adExecuteNoRecords
		blAllow = .Parameters("@Allow")
	End With
	Set cmdTemp = Nothing
	AllowVote = CBool(blAllow)
End Function

'------------------------End Code-----------------------------
%>