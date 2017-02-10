<%
'-----------------------Begin Code----------------------------
'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")


'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if Request("Keywords") <> "" then
	Set rsList = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM MemberStories WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
	intSearchID = SingleSearch()
	Session("SearchID") = intSearchID
	rsList.Close
end if


if intSearchID <> "" then
	%><p align="<%=HeadingAlignment%>"><span class=Heading>Dick Moves</span><br>
	<span class=LinkText><a href="members_dickmoves_add.asp">Add a Dick Move</a></span></p>
	<%
	'Their search came up empty
	if intSearchID = 0 then
		if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
			<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
			Try again, or <a href="dickmoves.asp">click here</a> to view all dick moves.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			Try again, or <a href="dickmoves.asp">click here</a> to view all dick moves.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	else
		'They have search results, so lets list their results
		Query = "SELECT * FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
%>
		<form METHOD="POST" ACTION="dickmoves.asp">
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">&nbsp;</td>
			<td class="TDHeader">Date</td>
			<td class="TDHeader">Author</td>
			<td class="TDHeader">Dick</td>
			<td class="TDHeader">Dick Points</td>
			<td class="TDHeader">Subject</td>
			<td class="TDHeader">Rating</td>
			<td class="TDHeader">Public?</td>
		</tr>
<%		'Instantiate the recordset for the output
		Set rsList = Server.CreateObject("ADODB.Recordset")
		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				Query = "SELECT * FROM MemberStories WHERE ID = " & rsPage("TargetID")
				rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
'------------------------End Code-----------------------------
%>
				<tr>
					<td class="<% PrintTDMain %>" align="center"><% PrintNew(rsList("Date")) %><a href="dickmoves_read.asp.asp?ID=<%=rsList("ID")%>">Read</a></td>
					<td class="<% PrintTDMain %>"><%=FormatDateTime(rsList("Date"), 2)%></td>
					<td class="<% PrintTDMain %>"><%=GetNickNameLink(rsList("MemberID"))%></td>
					<td class="<% PrintTDMain %>"><%=GetNickNameLink(rsList("TargetID"))%></td>
					<td class="<% PrintTDMain %>"><%=rsList("Points")%></td>
					<td class="<% PrintTDMain %>"><%=rsList("Subject")%></td>
					<td class="<% PrintTDMain %>"><%=GetRating( rsList("TotalRating"), rsList("TimesRated") )%></td>
					<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(rsList("Private"))%></td>
				</tr>
<%
'-----------------------Begin Code----------------------------
				rsList.Close
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
		set rsPage = Nothing
		set rsList = Nothing
	end if
'They are just cycling through the stories.  No searching.
else
	'List everyone with dick points and a link to their stories
	'Use the sectionsearch to store the list of people with the most points
	if Request("TargetID") = "" then
		%><p align="<%=HeadingAlignment%>"><span class=Heading>Dick Moves</span><br>
		<span class=LinkText><a href="members_dickmoves_add.asp">Add a Dick Move</a></span></p>
		<%
		Query = "SELECT * FROM MemberStories WHERE (CustomerID = " & CustomerID & ") ORDER BY TargetID"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly

		Query = "SELECT * FROM SectionSearch"
		Set rsTempSearch = Server.CreateObject("ADODB.Recordset")
		rsTempSearch.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect


		intSearchID = 0
		intTargetID = 0
		do until rsList.EOF
			'First entry
			if intTargetID <> rsList("TargetID") then
				intTargetID = rsList("TargetID")
				rsTempSearch.AddNew
				rsTempSearch("TargetID") = rsList("TargetID")
				rsTempSearch("Score") = rsList("Points")
				rsTempSearch("CustomerID") = CustomerID 
				'The first record in the search results
				if intSearchID = 0 then
					rsTempSearch.Update
					rsTempSearch.MovePrevious
					rsTempSearch.MoveNext
					intSearchID = rsTempSearch("ID")
				end if
				rsTempSearch("SearchID") = intSearchID
			else
				rsTempSearch("Score") = rsTempSearch("Score") + rsList("Points")
			end if
			rsTempSearch.Update
			rsList.MoveNext
		loop

		Set rsTempSearch = Nothing
		Set rsList = Nothing

		'They have search results, so lets list their results
		Query = "SELECT * FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsPage.EOF then
%>
			<form METHOD="POST" ACTION="dickmoves.asp">
				Search Dick Moves For <input type="text" name="Keywords" size="25">
				<input type="submit" name="Submit" value="Go"><br>
			</form>

			<p>The biggest dick comes first, and so on.</p>
			<form METHOD="POST" ACTION="dickmoves.asp">
<%
				PrintPagesHeader
				PrintTableHeader 0
%>
				<tr>
					<td class="TDHeader">&nbsp;</td>
					<td class="TDHeader">Dick Place</td>
					<td class="TDHeader">Dick</td>
					<td class="TDHeader">Dick Points</td>
				</tr>			
<%
				for j = 1 to rsPage.PageSize
					if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
						<tr>
							<td class="<% PrintTDMain %>" align="right"><a href="dickmoves.asp?TargetID=<%=rsPage("TargetID")%>">View Dick Moves</a></td>
							<td class="<% PrintTDMain %>" align="center"><%=j%></td>
							<td class="<% PrintTDMain %>"><%=GetNickNameLink(rsPage("TargetID"))%></td>
							<td class="<% PrintTDMain %>" align="center"><%=rsPage("Score")%></td>
						</tr>
<%
'-----------------------Begin Code----------------------------
						rsPage.MoveNext
					end if
					
				next
				Response.Write("</table>")
		else
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there are no dicks at the moment.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	else
		intTargetID = CInt( Request("TargetID") )
		%><p class="Heading" align="<%=HeadingAlignment%>"><%=GetNickName( intTargetID )%>'s Dick Moves</p><%

		Query = "SELECT * FROM MemberStories WHERE (TargetID = " & intTargetID & " AND CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsPage.EOF then
%>
			<form METHOD="POST" ACTION="dickmoves.asp">
			<input type="hidden" name="TargetID" value="<%=intTargetID%>">
<%
				PrintPagesHeader
				PrintTableHeader 0
%>
				<tr>
					<td class="TDHeader">&nbsp;</td>
					<td class="TDHeader">Date</td>
					<td class="TDHeader">Author</td>
					<td class="TDHeader">Dick Points</td>
					<td class="TDHeader">Subject</td>
					<td class="TDHeader">Rating</td>
					<td class="TDHeader">Public?</td>
				</tr>			
<%				for j = 1 to rsPage.PageSize
					if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(rsPage("Date")) %><a href="dickmoves_read.asp?ID=<%=rsPage("ID")%>">Read</a></td>
							<td class="<% PrintTDMain %>"><%=FormatDateTime(rsPage("Date"), 2)%></td>
							<td class="<% PrintTDMain %>"><%=GetNickNameLink(rsPage("MemberID"))%></td>
							<td class="<% PrintTDMain %>"><%=rsPage("Points")%></td>
							<td class="<% PrintTDMain %>"><%=rsPage("Subject")%></td>
							<% if RateStories = 1 then %>
								<td class="<% PrintTDMain %>"><%=GetRating( rsPage("TotalRating"), rsPage("TimesRated") )%></td>
							<% end if %>
							<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(rsPage("Private"))%></td>
						</tr>
<%
'-----------------------Begin Code----------------------------
						rsPage.MoveNext
					end if
				next
				Response.Write("</table>")
		else
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there are no dick moves at the moment.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	end if
	set rsPage = Nothing
end if


'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(rsList("Subject") & " " & rsList("Body") )
End Function


'------------------------End Code-----------------------------
%>
