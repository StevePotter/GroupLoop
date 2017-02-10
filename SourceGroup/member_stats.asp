<%
'-----------------------Begin Code----------------------------

	Query = "SELECT IncludeMemberStats, MemberStatsTitle FROM Configuration WHERE CustomerID = " & CustomerID
	Set rsConfig = Server.CreateObject("ADODB.Recordset")
	rsConfig.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	IncludeMemberStats = rsConfig("IncludeMemberStats")
	MemberStatsTitle = rsConfig("MemberStatsTitle")

	rsConfig.Close
	Set rsConfig = Nothing

if CBool(IncludeMemberStats) then



'Give them the link to change the section's properties
if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
	Response.Write "<div align=right><a href='admin_stats_configure.asp?Source=member.asp?ID=" & rsMember("ID") & "'>Configure Statistics</a></div>"
end if
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>"><%=GetNickName(rsMember("ID"))%>'s <%=MemberStatsTitle%><br>
Beginning <%=FormatDateTime(rsMember("Date"), 2)%>
</p>

<%
Function GetNumMemberHits( strTable )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetNumMemberHits"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@Table", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Count", adInteger, adParamOutput )
		.Parameters("@Table") = strTable
		.Parameters("@MemberID") = rsMember("ID")
		.Execute , , adExecuteNoRecords
		intCount = .Parameters("@Count")
	End With
	Set cmdTemp = Nothing

	GetNumMemberHits = intCount
End Function

Function GetNumMemberItems( strTable )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetNumMemberItems"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@Table", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Count", adInteger, adParamOutput )
		.Parameters("@Table") = strTable
		.Parameters("@MemberID") = rsMember("ID")
		.Execute , , adExecuteNoRecords
		intCount = .Parameters("@Count")
	End With
	Set cmdTemp = Nothing

	GetNumMemberItems = intCount
End Function


'-------------------------------------------------------------
'This function gets the number of items in a category
'-------------------------------------------------------------
Function GetNumInCategory( strTable, intCatID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetNumInCategory"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@Table", adVarWChar, adParamInput, 10 )
		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Count", adInteger, adParamOutput )
		.Parameters("@Table") = strTable
		.Parameters("@ItemID") = intCatID
		.Execute , , adExecuteNoRecords
		intCount = .Parameters("@Count")
	End With
	Set cmdTemp = Nothing

	GetNumInCategory = intCount
End Function




intTopMax = StatTopMax
intNumItems = 0

Set rsTopRated = Server.CreateObject("ADODB.Recordset")
Set rsPopular = Server.CreateObject("ADODB.Recordset")

rsTopRated.CacheSize = StatTopMax
rsPopular.CacheSize = StatTopMax

intID = rsMember("ID")




StatsAnnouncements
StatsStories
StatsCalendar
StatsLinks
StatsQuotes
StatsForum
StatsPhotos
StatsPhotoCaptions
StatsVotingPolls
StatsQuizzes


set rsPopular = Nothing
set rsTopRated = Nothing


end if


'-------------------------------------------------------------
'Stats for Announcements
'-------------------------------------------------------------
Sub StatsAnnouncements
	intNumItems = GetNumMemberItems("Announcements")

	if CBool( IncludeAnnouncements ) AND intNumItems > 0 AND (IncludeStatsPopularAnnouncements = 1 OR IncludeStatsRatedAnnouncements = 1 OR IncludeStatsSummaryAnnouncements = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=AnnouncementsTitle%></p>
	<%
		if IncludeStatsSummaryAnnouncements = 1 then
	%>
		<p>Number of Announcements - <%=intNumItems%><br>
		Their announcements have been read <%=GetNumMemberHits("Announcements")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularAnnouncements = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date FROM Announcements WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopSubject = rsPopular("Subject")
				Set PopMemberID = rsPopular("MemberID")
				Set PopDate = rsPopular("Date")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateAnnouncements ) and IncludeStatsRatedAnnouncements = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date, TotalRating, TimesRated FROM Announcements WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopSubject = rsTopRated("Subject")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopDate = rsTopRated("Date")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Announcement<%=PrintPlural(intPopMax, "", "s")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Announcement<%=PrintPlural(intRateMax, "", "s")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="announcements_read.asp?ID=<%=PopID%>"><%=PrintTDLink(PopSubject)%></a>  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
					%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="announcements_read.asp?ID=<%=TopID%>"><%=PrintTDLink(TopSubject)%></a>  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub


'-------------------------------------------------------------
'Stats for Stories
'-------------------------------------------------------------
Sub StatsStories
	intNumItems = GetNumMemberItems("Stories")

	if CBool( IncludeStories ) AND intNumItems > 0 AND (IncludeStatsPopularStories = 1 OR IncludeStatsRatedStories = 1 OR IncludeStatsSummaryStories = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=StoriesTitle%></p>
	<%
		if IncludeStatsSummaryStories = 1 then
	%>
		<p>Number of Stories - <%=intNumItems%><br>
		Their stories have been read <%=GetNumMemberHits("Stories")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularStories = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date FROM Stories WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopSubject = rsPopular("Subject")
				Set PopMemberID = rsPopular("MemberID")
				Set PopDate = rsPopular("Date")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateStories ) and IncludeStatsRatedStories = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date, TotalRating, TimesRated FROM Stories WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopSubject = rsTopRated("Subject")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopDate = rsTopRated("Date")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Stor<%=PrintPlural(intPopMax, "y", "ies")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Stor<%=PrintPlural(intRateMax, "y", "ies")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="stories_read.asp?ID=<%=PopID%>"><%=PrintTDLink(PopSubject)%></a>  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
					%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="stories_read.asp?ID=<%=TopID%>"><%=PrintTDLink(TopSubject)%></a>  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub



'-------------------------------------------------------------
'Stats for Calendar
'-------------------------------------------------------------
Sub StatsCalendar
	intNumItems = GetNumMemberItems("Calendar")

	if CBool( IncludeCalendar ) AND intNumItems > 0 AND (IncludeStatsPopularCalendar = 1 OR IncludeStatsRatedCalendar = 1 OR IncludeStatsSummaryCalendar = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=CalendarTitle%></p>
	<%
		if IncludeStatsSummaryCalendar = 1 then
	%>
		<p>Number of Calendar Events - <%=intNumItems%><br>
		Their events have been read <%=GetNumMemberHits("CalendarEvents")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularCalendar = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date, StartDate FROM Calendar WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopSubject = rsPopular("Subject")
				Set PopMemberID = rsPopular("MemberID")
				Set PopDate = rsPopular("Date")
				Set PopStartDate = rsPopular("StartDate")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateCalendar ) and IncludeStatsRatedCalendar = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date, StartDate, TotalRating, TimesRated FROM Calendar WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopSubject = rsTopRated("Subject")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopDate = rsTopRated("Date")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Event<%=PrintPlural(intPopMax, "", "s")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Event<%=PrintPlural(intRateMax, "", "s")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="calendar_event_read.asp?ID=<%=PopID%>"><%=PrintTDLink(PopSubject)%></a>  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
					%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="calendar_event_read.asp?ID=<%=TopID%>"><%=PrintTDLink(TopSubject)%></a>  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub




'-------------------------------------------------------------
'Stats for Links
'-------------------------------------------------------------
Sub StatsLinks
	intNumItems = GetNumMemberItems("Links")

	if CBool( IncludeLinks ) AND intNumItems > 0 AND (IncludeStatsPopularLinks = 1 OR IncludeStatsRatedLinks = 1 OR IncludeStatsSummaryLinks = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=LinksTitle%></p>
	<%
		if IncludeStatsSummaryLinks = 1 then
	%>
		<p>Number of Links - <%=intNumItems%><br>
		Their links have been read <%=GetNumMemberHits("Links")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularLinks = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Name, URL, MemberID, Date FROM Links WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopName = rsPopular("Name")
				Set PopURL = rsPopular("URL")
				Set PopMemberID = rsPopular("MemberID")
				Set PopDate = rsPopular("Date")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateLinks ) and IncludeStatsRatedLinks = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Name, URL, MemberID, Date, TimesRated, TotalRating FROM Links WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopName = rsTopRated("Name")
				Set TopURL = rsTopRated("URL")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopDate = rsTopRated("Date")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if
			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Link<%=PrintPlural(intPopMax, "", "s")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Link<%=PrintPlural(intRateMax, "", "s")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						strName = PopName
						if PopName = "" then strName = PopURL
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="links_read.asp?ID=<%=PopID%>"><%=PrintTDLink(strName)%></a>  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
						strName = TopName
						if TopName = "" then strName = TopURL
					%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="links_read.asp?ID=<%=TopID%>"><%=PrintTDLink(strName)%></a>  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub



'-------------------------------------------------------------
'Stats for Quotes
'-------------------------------------------------------------
Sub StatsQuotes
	intNumItems = GetNumMemberItems("Quotes")

	if CBool( IncludeQuotes ) AND intNumItems > 0 AND (IncludeStatsPopularQuotes = 1 OR IncludeStatsRatedQuotes = 1 OR IncludeStatsSummaryQuotes = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=QuotesTitle%></p>
	<%
		if IncludeStatsSummaryQuotes = 1 then
	%>
		<p>Number of Quotes - <%=intNumItems%><br>
		Their quotes have been read <%=GetNumMemberHits("Quotes")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularQuotes = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Author, Quote, MemberID, Date FROM Quotes WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopAuthor = rsPopular("Author")
				Set PopQuote = rsPopular("Quote")
				Set PopMemberID = rsPopular("MemberID")
				Set PopDate = rsPopular("Date")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateQuotes ) and IncludeStatsRatedQuotes = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Author, Quote, MemberID, Date, TotalRating, TimesRated FROM Quotes WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopAuthor = rsTopRated("Author")
				Set TopQuote = rsTopRated("Quote")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopDate = rsTopRated("Date")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Quote<%=PrintPlural(intPopMax, "", "s")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Quote<%=PrintPlural(intRateMax, "", "s")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="quotes_read.asp?ID=<%=PopID%>"><%=PrintTDLink(PrintStart(PopQuote) & "&quot; - " & PopAuthor)%></a>  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
					%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="quotes_read.asp?ID=<%=TopID%>"><%=PrintTDLink(PrintStart(TopQuote) & "&quot; - " & TopAuthor)%></a>  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub



'-------------------------------------------------------------
'Stats for Guestbook
'-------------------------------------------------------------
Sub StatsGuestbook
	intNumItems = GetNumMemberItems("Guestbook")

	if CBool( IncludeGuestbook ) AND intNumItems > 0 AND (IncludeStatsPopularGuestbook = 1 OR IncludeStatsRatedGuestbook = 1 OR IncludeStatsSummaryGuestbook = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=GuestbookTitle%></p>
	<%
		if IncludeStatsSummaryGuestbook = 1 then
	%>
		<p>Number of Entries - <%=intNumItems%><br>
		Their entries have been read <%=GetNumMemberHits("Guestbook")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularGuestbook = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Author, Email, Body, Date FROM Guestbook WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopAuthor = rsPopular("Author")
				Set PopEmail = rsPopular("Email")
				Set PopBody = rsPopular("Body")
				Set PopDate = rsPopular("Date")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateGuestbook ) and IncludeStatsRatedGuestbook = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Author, Email, Body, Date, TotalRating, TimesRated FROM Guestbook WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopAuthor = rsTopRated("Author")
				Set TopEmail = rsTopRated("Email")
				Set TopBody = rsTopRated("Body")
				Set TopDate = rsTopRated("Date")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Entr<%=PrintPlural(intPopMax, "y", "ies")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Entr<%=PrintPlural(intRateMax, "i", "ies")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						if InStr( PopEmail, "@" ) then
							strAuthor = "<a href='mailto:" & PopEmail & "'>" & PopAuthor & "</a>"
						else
							strAuthor = PopAuthor
						end if
						strBody = PopBody
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="guestbook_read.asp?ID=<%=PopID%>"><%=PrintTDLink(PrintStart(strBody))%></a>  &nbsp;&nbsp;<font size="-2">(<%=PrintTDLink(strAuthor)%></a>, <%=FormatDateTime(PopDate, 2)%>)</font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
						if InStr( TopEmail, "@" ) then
							strAuthor = "<a href='mailto:" & TopEmail & "'>" & TopAuthor & "</a>"
						else
							strAuthor = TopAuthor
						end if
						strBody = TopBody
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="guestbook_read.asp?ID=<%=TopID%>"><%=PrintTDLink(PrintStart(strBody))%></a>  &nbsp;&nbsp;<font size="-2">(<%=PrintTDLink(strAuthor)%></a>, <%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub


'-------------------------------------------------------------
'Stats for Forum
'-------------------------------------------------------------
Sub StatsForum
	intNumItems = GetNumMemberItems("ForumMessages")

	if CBool( IncludeForum ) AND intNumItems > 0 AND (IncludeStatsPopularForum = 1 OR IncludeStatsRatedForum = 1 OR IncludeStatsSummaryForum = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=ForumTitle%></p>
	<%
		if IncludeStatsSummaryForum = 1 then
	%>
		Number of Messages - <%=intNumItems%><br>
		Their messages have been read <%=GetNumMemberHits("ForumMessages")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularForum = 1 then

			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date, Author, EMail FROM ForumMessages WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopSubject = rsPopular("Subject")
				Set PopMemberID = rsPopular("MemberID")
				Set PopAuthor = rsPopular("Author")
				Set PopEMail = rsPopular("EMail")
				Set PopDate = rsPopular("Date")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateForum ) and IncludeStatsRatedForum = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date, Author, EMail, TotalRating, TimesRated FROM ForumMessages WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopSubject = rsTopRated("Subject")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopDate = rsTopRated("Date")
				Set TopAuthor = rsTopRated("Author")
				Set TopEMail = rsTopRated("EMail")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Message<%=PrintPlural(intPopMax, "", "s")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Message<%=PrintPlural(intRateMax, "", "s")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						if PopMemberID > 0 then
							strAuthor = GetNickNameLink( PopMemberID )
						elseif InStr( PopEMail, "@" ) then
							strAuthor = "<a href='mailto:" & PopEMail & "'>" & PopAuthor & "</a>"
						else
							strAuthor = PopAuthor
						end if
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="forum_read.asp?ID=<%=PopID%>"><%=PrintTDLink(PopSubject)%></a>  &nbsp;&nbsp;<font size="-2">(<%=PrintTDLink(strAuthor)%></a>, <%=FormatDateTime(PopDate, 2)%>)</font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
						if TopMemberID > 0 then
							strAuthor = GetNickNameLink( TopMemberID )
						elseif InStr( TopEMail, "@" ) then
							strAuthor = "<a href='mailto:" & TopEMail & "'>" & TopAuthor & "</a>"
						else
							strAuthor = TopAuthor
						end if

					%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="forum_read.asp?ID=<%=TopID%>"><%=PrintTDLink(TopSubject)%></a>  &nbsp;&nbsp;<font size="-2">(<%=PrintTDLink(strAuthor)%></a>, <%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub



'-------------------------------------------------------------
'Stats for Photos
'-------------------------------------------------------------
Sub StatsPhotos
	intNumItems = GetNumMemberItems("Photos")

	if CBool( IncludePhotos ) AND intNumItems > 0 AND (IncludeStatsPopularPhotos = 1 OR IncludeStatsRatedPhotos = 1 OR IncludeStatsSummaryPhotos = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=PhotosTitle%></p>
	<%
		if IncludeStatsSummaryPhotos = 1 then
	%>
		<p>
		Number of Photos - <%=intNumItems%><br>
		Their photos have been viewed <%=GetNumMemberHits("Photos")%> times.</p>

	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularPhotos = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Name, MemberID, Date, Thumbnail, ThumbnailExt FROM Photos WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopDate = rsPopular("Date")
				Set PopName = rsPopular("Name")
				Set PopMemberID = rsPopular("MemberID")
				Set PopThumbnail = rsPopular("Thumbnail")
				Set PopThumbnailExt = rsPopular("ThumbnailExt")	
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RatePhotos ) and IncludeStatsRatedPhotos = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Name, MemberID, Date, Thumbnail, ThumbnailExt, TotalRating, TimesRated FROM Photos WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopDate = rsTopRated("Date")
				Set TopName = rsTopRated("Name")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopThumbnail = rsTopRated("Thumbnail")
				Set TopThumbnailExt = rsTopRated("ThumbnailExt")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Photo<%=PrintPlural(intPopMax, "", "s")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Photo<%=PrintPlural(intRateMax, "", "s")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						%>
						<td class="<% PrintTDMainSwitch %>" align="center" valign="middle">
<%
							if PopThumbnail = 1 then
%>
								<a href="photos_view.asp?ID=<%=PopID%>"><img src="photos/<%=PopID%>t.<%=PopThumbnailExt%>" border=0 alt="<%=PopName%>"></a>
<%
							else
%>
								<a href="photos_view.asp?ID=<%=PopID%>"><%=PrintTDLink(PopName)%></a>
<%
							end if
%>
						</td>						
<%						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
%>
						<td class="<% PrintTDMainSwitch %>" align="center" valign="middle">
<%
						if TopThumbnail = 1 then
%>
							<a href="photos_view.asp?ID=<%=TopID%>"><img src="photos/<%=TopID%>t.<%=TopThumbnailExt%>" border=0 alt="<%=TopName%>"></a>
<%
						else
%>
							<a href="photos_view.asp?ID=<%=TopID%>"><%=PrintTDLink(TopName)%></a>
<%
						end if
%>
						</td>
<%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub



'-------------------------------------------------------------
'Stats for PhotoCaptions
'-------------------------------------------------------------
Sub StatsPhotoCaptions
	intNumItems = GetNumMemberItems("PhotoCaptions")

	if CBool( IncludePhotoCaptions ) AND intNumItems > 0 AND (IncludeStatsPopularPhotoCaptions = 1 OR IncludeStatsRatedPhotoCaptions = 1 OR IncludeStatsSummaryPhotoCaptions = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=PhotoCaptionsTitle%></p>
	<%
		if IncludeStatsSummaryPhotoCaptions = 1 then
	%>
			<p>Number of Captions - <%=intNumItems%><br>
			Their captions have been read <%=GetNumMemberHits("PhotoCaptions")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularPhotoCaptions = 1 then
			Query = "SELECT TOP " & intTopMax  & " PhotoID, Caption, MemberID, Date, Private FROM PhotoCaptions WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopPhotoID = rsPopular("PhotoID")
				Set PopDate = rsPopular("Date")
				Set PopCaption = rsPopular("Caption")
				Set PopPrivate = rsPopular("Private")
				Set PopMemberID = rsPopular("MemberID")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RatePhotoCaptions ) and IncludeStatsRatedPhotoCaptions = 1 then
			Query = "SELECT TOP " & intTopMax  & " PhotoID, Caption, MemberID, Date, TotalRating, TimesRated, Private FROM PhotoCaptions WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopPhotoID = rsTopRated("PhotoID")
				Set TopDate = rsTopRated("Date")
				Set TopCaption = rsTopRated("Caption")
				Set TopPrivate = rsTopRated("Private")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Caption<%=PrintPlural(intPopMax, "", "s")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Caption<%=PrintPlural(intRateMax, "", "s")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						if PopPrivate = 1 AND not LoggedMember then
							%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="photos_view.asp?ID=<%=PopPhotoID%>">Private Caption <a href="login.asp?Source=stats.asp&Submit=Read"><%=PrintTDLink("Click here")%></a> to log in and read it.  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						else
							%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="photos_view.asp?ID=<%=PopPhotoID%>"><%=PrintTDLink(PrintStart(PopCaption))%></a>  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						end if
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
							if TopPrivate = 1 AND not LoggedMember then
								%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="photos_view.asp?ID=<%=TopPhotoID%>">Private Caption <a href="login.asp?Source=stats.asp&Submit=Read"><%=PrintTDLink("Click here")%></a> to log in and read it.  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
							else
								%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="photos_view.asp?ID=<%=TopPhotoID%>"><%=PrintTDLink(PrintStart(TopCaption))%></a>  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
							end if
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub



'-------------------------------------------------------------
'Stats for VotingPolls
'-------------------------------------------------------------
Sub StatsVotingPolls
	intNumItems = GetNumMemberItems("VotingPolls")

	if CBool( IncludeVotingPolls ) AND intNumItems > 0 AND (IncludeStatsPopularVotingPolls = 1 OR IncludeStatsRatedVotingPolls = 1 OR IncludeStatsSummaryVotingPolls = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=VotingPollsTitle%></p>
	<%
		if IncludeStatsSummaryVotingPolls = 1 then
	%>
		<p>Number of Polls - <%=intNumItems%><br>
		Number of Poll Answers - <%=GetNumMemberItems("VotingOptions")%><br>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularVotingPolls = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date FROM VotingPolls WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopSubject = rsPopular("Subject")
				Set PopMemberID = rsPopular("MemberID")
				Set PopDate = rsPopular("Date")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateVotingPolls ) and IncludeStatsRatedVotingPolls = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date, TotalRating, TimesRated FROM VotingPolls WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopSubject = rsTopRated("Subject")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopDate = rsTopRated("Date")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Poll<%=PrintPlural(intPopMax, "", "s")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Poll<%=PrintPlural(intRateMax, "", "s")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="voting_results.asp?ID=<%=PopID%>"><%=PrintTDLink(PopSubject)%></a>  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
					%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="voting_results.asp?ID=<%=TopID%>"><%=PrintTDLink(TopSubject)%></a>  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub



'-------------------------------------------------------------
'Stats for Quizzes
'-------------------------------------------------------------
Sub StatsQuizzes
	intNumItems = GetNumMemberItems("Quizzes")

	if CBool( IncludeQuizzes ) AND intNumItems > 0 AND (IncludeStatsPopularQuizzes = 1 OR IncludeStatsRatedQuizzes = 1 OR IncludeStatsSummaryQuizzes = 1 ) then
		intPopMax = 0
		intRateMax = 0
		intLoopMax = 0
	%>
		<p class="Heading" align="<%=HeadingAlignment%>"><%=QuizzesTitle%></p>
	<%
		if IncludeStatsSummaryQuizzes = 1 then
	%>
		<p>Number of Quizzes - <%=intNumItems%><br>
		Number of Quiz Questions - <%=GetNumMemberItems("QuizQuestions")%><br>
		Their quizzes have been taken <%=GetNumMemberHits("Quizzes")%> times.</p>
	<%
		end if

		blPopularExists = False
		blRatedExists = False

		if IncludeStatsPopularQuizzes = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date FROM Quizzes WHERE Hits > 0 AND MemberID = " & intID & " ORDER BY Hits DESC"
			rsPopular.Open Query, Connect, adOpenStatic, adLockReadOnly
			intPopMax = intTopMax
			if rsPopular.RecordCount < intTopMax then intPopMax = rsPopular.RecordCount

			if not rsPopular.EOF then
				Set PopID = rsPopular("ID")
				Set PopSubject = rsPopular("Subject")
				Set PopMemberID = rsPopular("MemberID")
				Set PopDate = rsPopular("Date")
			end if
			blPopularExists = CBool( not rsPopular.EOF )
			if not blPopularExists then rsPopular.Close

		end if

		if CBool( RateQuizzes ) and IncludeStatsRatedQuizzes = 1 then
			Query = "SELECT TOP " & intTopMax  & " ID, Subject, MemberID, Date, TotalRating, TimesRated FROM Quizzes WHERE TimesRated > 0 AND MemberID = " & intID & " ORDER BY RatingScore DESC"
			rsTopRated.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			intRateMax = intTopMax
			if rsTopRated.RecordCount < intTopMax then intRateMax = rsTopRated.RecordCount
			if intRateMax > 0 then
				Set TopID = rsTopRated("ID")
				Set TopSubject = rsTopRated("Subject")
				Set TopMemberID = rsTopRated("MemberID")
				Set TopDate = rsTopRated("Date")
				Set TopTotalRating = rsTopRated("TotalRating")
				Set TopTimesRated = rsTopRated("TimesRated")
			end if

			blRatedExists = CBool( not rsTopRated.EOF )
			if not blRatedExists then rsTopRated.Close
		end if

		if blPopularExists or blRatedExists then
			ResetTDMain
	%>
			<% PrintTableHeader 0 %>
			<tr>
				<% if blPopularExists then %>
				<td class="TDHeader" align="center">Their <%=intPopMax%> Most Popular Quiz<%=PrintPlural(intPopMax, "", "zes")%></td>
				<% end if %>
				<% if blRatedExists then %>
					<td class="TDHeader" align="center">Their <%=intRateMax%> Highest Rated Quiz<%=PrintPlural(intRateMax, "", "zes")%></td>
				<% end if %>
			</tr>

	<%
			if intPopMax > intRateMax then
				intLoopMax = intPopMax
			else
				intLoopMax = intRateMax
			end if

			for i = 1 to intLoopMax
	%>
				<tr>
	<%
				if blPopularExists then
					if not rsPopular.EOF then
						%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="quizzes_take.asp?ID=<%=PopID%>"><%=PrintTDLink(PopSubject)%></a>  &nbsp;&nbsp;<font size="-2"><%=FormatDateTime(PopDate, 2)%></font></td><%
						rsPopular.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if
				if blRatedExists then
					ChangeTDMain
					if not rsTopRated.EOF then
					%><td class="<% PrintTDMainSwitch %>" align="left"><%=i%>. <a href="quizzes_take.asp?ID=<%=TopID%>"><%=PrintTDLink(TopSubject)%></a>  &nbsp;&nbsp;<font size="-2">(<%=FormatDateTime(TopDate, 2)%>, Rating: <%=GetRating( TopTotalRating, TopTimesRated )%>)</font></td><%
						rsTopRated.MoveNext
					else
						%><td class="<% PrintTDMainSwitch %>" align="left">&nbsp;</td><%
					end if
				end if %>
				</tr>
	<%
			next
	%>
			</table>
	<%
		end if
	%>
			<br>
	<%
		if blPopularExists then rsPopular.Close
		if blRatedExists then rsTopRated.Close
	end if
End Sub


%>
