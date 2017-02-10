<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeVoting ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading><%=VotingTitle%></span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span></p>

<%
'-----------------------Begin Code----------------------------
intID = Request("ID")
if intID <> "" then intID = CInt(intID)
if intID = "" and not Request("Action") = "ViewClosed" then Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))


strSubject = ""

'Get the ID of the next poll
Query = "SELECT ID, Date, MemberID, Subject, TotalRating, TimesRated, Private FROM VotingPolls WHERE (OpenToVote = 1 AND CustomerID = " & CustomerID & ") ORDER BY Date DESC"
Set rsPage = Server.CreateObject("ADODB.Recordset")
rsPage.CacheSize = PageSize
rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly

intLastID = 0
intNextID = 0
do until rsPage.EOF
	intLastID = rsPage("ID")

	rsPage.MoveNext

	'We are on the next record
	if intLastID = intID and not rsPage.EOF then intNextID = rsPage("ID")
loop

rsPage.Close

'Open up the item
if Request("Action") = "ViewClosed" then
	Query = "SELECT ID, ResultsSecurity, Subject, MemberID, OpenToVote FROM VotingPolls WHERE (CustomerID = " & CustomerID & " AND OpenToVote = 0)"
else
	Query = "SELECT ID, ResultsSecurity, Subject, MemberID, OpenToVote FROM VotingPolls WHERE ID = " & intID
end if

rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly

if rsPage.EOF then
	Set rsPage = Nothing
	Redirect ("error.asp?Message=" & Server.URLEncode("The poll you requested to see doesn't exist.  Go back to the voting section and make sure the poll exists.") )
end if

if Request("Rating") <> "" and RateVoting = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To Polls</a></p>
<%
	AddRating intID, "VotingPolls"
else
	Set MemberID = rsPage("MemberID")
	Set ID = rsPage("ID")
	Set ResultsSecurity = rsPage("ResultsSecurity")
	Set Subject = rsPage("Subject")
	Set OpenToVote = rsPage("OpenToVote")

	strSubject = Subject

	PrintPagesHeader

	blLoggedMember = LoggedMember
	blLoggedAdmin = LoggedAdmin

	Set rsOptions = Server.CreateObject("ADODB.Recordset")
	rsOptions.CacheSize = 20

	do until rsPage.EOF
		intID = ID

		'If we can't see the results, log in
		if ResultsSecurity = 1 AND not blLoggedMember then
	'------------------------End Code-----------------------------
	%>
			<p>Sorry, but for the poll <b><%=Subject%></b>, only members may view the results.  If you are a member, <a href="login.asp?Source=voting_results.asp&ID=<%=intID%>&Submit=View+Results">Click here</a> to log in.</p>
	<%
	'-----------------------Begin Code----------------------------
								if intNextID > 0 then Response.Write "<p><a href=voting_cast.asp?ID=" & intNextID & ">Vote on the next poll</a></p>"
		elseif ResultsSecurity = 2 AND not blLoggedAdmin then
	'------------------------End Code-----------------------------

'			<p>Sorry, but for the poll <b>< %=Subject% ></b>, only administrators may view the results.</p>

	'-----------------------Begin Code----------------------------
			if intNextID > 0 then Response.Write "<p><a href=voting_cast.asp?ID=" & intNextID & ">Vote on the next poll</a></p>"

					'Print out the rating and reviews link
					if RateVoting = 1 then
						PrintRatingPulldown intID, "", "VotingPolls", "voting_results.asp", "poll"
					end if
					if ReviewVoting = 1 then
				%>
						<a href="review.asp?Source=voting_results.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=VotingPolls&Title=<%=Server.URLEncode(Subject)%>">Add A Review</a><br>
				<%
					end if


		else
			IncrementHits intID, "VotingPolls"


'Query = "SELECT MemberID, COUNT( MemberID ) As MemberVotes FROM VotingResponses " & _
'"WHERE PollID = " & intID & " GROUP BY MemberID ORDER BY MemberVotes DESC, MemberID"
', COUNT( VotingResponses.MemberID ) As Hits FROM
 Query = "SELECT Name, COUNT(VotingResponses.MemberID ) As Hits FROM VotingOptions " &_
 "INNER JOIN VotingResponses  " &_
 "ON VotingOptions.ID = VotingResponses.OptionID " &_
 "WHERE VotingOptions.PollID = " & intID & " AND VotingOptions.ID > 0 GROUP BY VotingOptions.Name"
 



	'		Query = "SELECT Hits, Name FROM VotingOptions WHERE PollID = " & intID & " ORDER BY ID"
			rsOptions.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set Hits = rsOptions("Hits")
			Set Name = rsOptions("Name")

			if rsOptions.EOF then
				Response.Write "<p>Sorry, but this poll has no answers, so it can't be voted on.</p>"
			else

				'Get the total number of votes
				intTotalVotes = 0
				do until rsOptions.EOF
					intTotalVotes =	intTotalVotes + Hits
					rsOptions.MoveNext
				loop
				rsOptions.MoveFirst
				if intTotalVotes = 0 then
	'------------------------End Code-----------------------------
	%>
					<p class=Heading align="<%=HeadingAlignment%>">Sorry, but nobody has voted yet.</p>
	<%
	'-----------------------Begin Code----------------------------
					if intNextID > 0 then Response.Write "<p><a href=voting_cast.asp?ID=" & intNextID & ">Vote on the next poll</a></p>"
				else
	'------------------------End Code-----------------------------
	%>
					<p class=Heading align="<%=HeadingAlignment%>">Results for: <%=Subject%>  <font size="-2">(<%=intTotalVotes%> votes)</font>
	<%
					if OpenToVote = 1 then
%>
					<br><span class="LinkText"><a href="voting_cast.asp?ID=<%=ID%>">Cast A Vote</a></span>
<%					end if
					if LoggedAdmin or (LoggedMember and Session("MemberID") = MemberID)  then
	%>
						<table align=<%=HeadingAlignment%>>
						<tr>
						<td align=right width="50%" class="LinkText"><a href="members_polls_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
						<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
						<a href="javascript:DeleteBox('If you delete this voting poll, there is no way to get it back.  Are you sure?', 'members_polls_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a>
						</td>
						</tr>
						</table>
	<%
					end if
					Response.Write "</p>"

					PrintTableHeader 100
					do until rsOptions.EOF
	'------------------------End Code-----------------------------
	%>
						<tr>
							<td width="40%" class="<% PrintTDMain %>">
								<%=Name%>
							</td>
	<%
	'-----------------------Begin Code----------------------------
						intNumVotes = Hits
						intPercent = Round( 100 * intNumVotes / intTotalVotes )

						%><td width="40%" align="left" valign="middle" class="<% PrintTDMain %>"><%
						if intPercent > 0 then
	'------------------------End Code-----------------------------
	%>
							<table width="<%=intPercent%>" border="0" cellspacing="0" cellpadding="0">
								<tr>
								<td bgcolor="<%=VotingBarColor%>" width="100%">&nbsp;</td>
								</tr>
							</table>
	<%
	'-----------------------Begin Code----------------------------
						else
	'------------------------End Code-----------------------------
	%>
							&nbsp;
	<%
	'-----------------------Begin Code----------------------------
						end if
							strVotes = "Votes"
							if intNumVotes = 1 then strVotes = "Vote"
	'------------------------End Code-----------------------------
	%>
							</td>
							<td width="20%" class="<% PrintTDMain %>">
								<%=intPercent%>% &nbsp;<%=intNumVotes%>&nbsp;<%=strVotes%>
							</td>
						</tr>
	<%
	'-----------------------Begin Code----------------------------
						rsOptions.MoveNext
					loop
					Response.Write "</table>"

					rsOptions.Close

					if LoggedAdmin() then
						Response.Write "<br>"
						PrintTableHeader 100
%>
						<tr>
							<td class="TDHeader">
								Who Voted
							</td>
							<td class="TDHeader">
								Number of Votes
							</td>
						</tr>
<%
						Query = "SELECT MemberID, COUNT( MemberID ) As MemberVotes FROM VotingResponses " & _
							"WHERE PollID = " & intID & " GROUP BY MemberID ORDER BY MemberVotes DESC, MemberID"
						rsOptions.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
						do until rsOptions.EOF
							intMemberID = rsOptions("MemberID")
							if intMemberID = 0 then
								strMember = "Non-Members or Members Not Logged In"
							else
								strMember = PrintTDLink(GetNickNameLink(intMemberID))
							end if
%>
						<tr>
							<td class="<% PrintTDMain %>"><%=strMember%></td>
						<td class="<% PrintTDMain %>"><%=rsOptions("MemberVotes")%></td>
						</tr>
<%
							rsOptions.MoveNext
						loop
						Response.Write "</table>"

						rsOptions.Close

					end if

					if intNextID > 0 then Response.Write "<p><a href=voting_cast.asp?ID=" & intNextID & ">Vote on the next poll</a></p>"

					'Print out the rating and reviews link
					if RateVoting = 1 then
						PrintRatingPulldown intID, "", "VotingPolls", "voting_results.asp", "poll"
					end if
					if ReviewVoting = 1 then
				%>
						<a href="review.asp?Source=voting_results.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=VotingPolls&Title=<%=Server.URLEncode(Subject)%>"">Add A Review</a><br>
				<%
					end if

				end if
			end if
		end if
		rsPage.MoveNext
	loop
	set rsOptions = Nothing
end if
rsPage.Close

'Print reviews if there are any
if intID <> "" and ReviewVoting = 1 then

	if ReviewsExist( "VotingPolls", intID ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=voting_results.asp?ID=<%=intID%>&TargetTable=VotingPolls&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if
		strTitle = "Comments On<br> '" & strSubject & "'"
		PrintReviewsNew "voting_results.asp", "VotingPolls", intID, strTitle
	end if
end if

Set rsPage = Nothing

'------------------------End Code-----------------------------
%>