<p align="<%=HeadingAlignment%>"><span class=Heading>Pet Peeves</span><br>
<%
if IncludeAddButtons = 1 or LoggedMember() then
%>
	<span class=LinkText><a href="members_petpeeves_add.asp">Add a Pet Peeve</a></span>
<%
end if
%>	
</p>

<form METHOD="POST" ACTION="petpeeves.asp">
	View Pet Peeves In The Last <% PrintDaysOld %>
	<br>
	Or Search For <input type="text" name="Keywords" size="25">
	<input type="submit" name="Submit" value="Go"><br>
</form>
<%
'-----------------------Begin Code----------------------------
'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")


'They entered text to search for, so we are going to get matches and put them into the SectionSearch
if Request("Keywords") <> "" then
	Set rsList = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM PetPeeves WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
	intSearchID = SingleSearch()
	Session("SearchID") = intSearchID
	rsList.Close
end if

if intSearchID <> "" then
	'Their search came up empty
	if intSearchID = 0 then
		if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
			<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
			Try again, or <a href="PetPeeves.asp">click here</a> to view all pet peeves.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
			<p>Sorry, but your search came up empty.<br>
			Try again, or <a href="petpeeves.asp">click here</a> to view all pet peeves.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	else
		'They have search results, so lets list their results
		Query = "SELECT * FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
%>
		<form METHOD="POST" ACTION="petpeeves.asp">
		<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
		PrintPagesHeader
		PrintTableHeader 0
		PrintTableTitle
		'Instantiate the recordset for the output
		Set rsList = Server.CreateObject("ADODB.Recordset")
		for p = 1 to rsPage.PageSize
			if not rsPage.EOF then
				Query = "SELECT * FROM PetPeeves WHERE ID = " & rsPage("TargetID")
				rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
'------------------------End Code-----------------------------
%>
				<tr>
					<td class="<% PrintTDMain %>" align="center"><% PrintNew(rsList("Date")) %><a href="petpeeves_read.asp?ID=<%=rsList("ID")%>">Details</a></td>
					<td class="<% PrintTDMain %>"><%=FormatDateTime(rsList("Date"), 2)%></td>
					<td class="<% PrintTDMain %>"><%=GetNickNameLink(rsList("MemberID"))%></td>
					<td class="<% PrintTDMain %>"><%=rsList("Subject")%></td>
<%					if ReviewsExist( "PetPeeves", rsList("ID") ) then
%>						<td class="<% PrintTDMain %>" align=center><%=GetRating( rsList("TotalRating"), rsList("TimesRated") )%> 
						<font size="-2"><a href="petpeeves_read.asp?ID=<%=rsList("ID")%>">Rate and Read/Add Review</a></font></td>
<%					else
%>						<td class="<% PrintTDMain %>" align=center><%=GetRating( rsList("TotalRating"), rsList("TimesRated") )%> 
						<font size="-2"><a href="petpeeves_read.asp?ID=<%=rsList("ID")%>">Rate/Add Review</a></font></td>
<%					end if%>
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
'They are just cycling through the petpeeves.  No searching.
else
	'This is if they requested petpeeves written in a time period
	if Request("DaysOld") <> "" then
		CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
		Query = "SELECT * FROM PetPeeves WHERE (CustomerID = " & CustomerID & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
	else
		Query = "SELECT * FROM PetPeeves WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	end if
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
%>
		<form METHOD="POST" ACTION="petpeeves.asp">
		<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle
			for j = 1 to rsPage.PageSize
				if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(rsPage("Date")) %><a href="petpeeves_read.asp?ID=<%=rsPage("ID")%>">Details</a></td>
						<td class="<% PrintTDMain %>"><%=FormatDateTime(rsPage("Date"), 2)%></td>
						<td class="<% PrintTDMain %>"><%=GetNickNameLink(rsPage("MemberID"))%></td>
						<td class="<% PrintTDMain %>"><%=rsPage("Subject")%></td>
	<%					if ReviewsExist( "PetPeeves", rsPage("ID") ) then
	%>						<td class="<% PrintTDMain %>" align=center><%=GetRating( rsPage("TotalRating"), rsPage("TimesRated") )%> 
							<font size="-2"><a href="petpeeves_read.asp?ID=<%=rsPage("ID")%>">Rate and Read/Add Review</a></font></td>
	<%					else
	%>						<td class="<% PrintTDMain %>" align=center><%=GetRating( rsPage("TotalRating"), rsPage("TimesRated") )%> 
							<font size="-2"><a href="petpeeves_read.asp?ID=<%=rsPage("ID")%>">Rate/Add Review</a></font></td>
	<%					end if%>
						<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(rsPage("Private"))%></td>
					</tr>
<%
'-----------------------Begin Code----------------------------
					rsPage.MoveNext
				end if
			next
			Response.Write("</table>")
	else
		if Request("DaysOld") <> "" then
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there have been no pet peeves added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
		else
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there are no pet peeves at the moment.</p>
<%
'-----------------------Begin Code----------------------------
		end if
	end if
	rsPage.Close
	set rsPage = Nothing
end if


'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(rsList("Subject") & " " & rsList("Body") )
End Function


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader">&nbsp;</td>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">Author</td>
		<td class="TDHeader">Pet Peeve</td>
		<td class="TDHeader">Rating</td>
		<td class="TDHeader">Public?</td>
	</tr>
<%
End Sub
'------------------------End Code-----------------------------
%>
