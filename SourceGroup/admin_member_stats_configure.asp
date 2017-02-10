<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_additions_configure.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Configure <%=StatsTitle%></span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Create Files----------------------------
if Request("Submit") = "Update" then
	Query = "SELECT StatTopMax, StatsTitle FROM Configuration WHERE CustomerID = " & CustomerID
	Set rsConfig = Server.CreateObject("ADODB.Recordset")
	rsConfig.Open Query, Connect, adOpenStatic, adLockOptimistic

	if Request("StatTopMax") <> "" then rsConfig("StatTopMax") = Request("StatTopMax")
	if Request("StatsTitle") <> "" then rsConfig("StatsTitle") = Format( Request("StatsTitle") )

	rsConfig.Update
	rsConfig.Close
	Set rsConfig = Nothing
%>
	<!-- #include file="write_constants.asp" -->
<%
	if strPath = "" then strPath = GetPath("")
	strImagePath = strPath & "images/"
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	Set ConstFile = FileSystem.CreateTextFile(strPath & "stats_constantstemp.inc")

	DblQuote = Chr(34)

	'-----------------------Start Writing Files----------------------------

	'Create our opening delimiter
	ConstFile.WriteLine "<" & "%"
	ConstFile.WriteBlankLines 2

	ConstFile.WriteLine "'This file contains the constants for this site's lastest additions (cust #" & CustomerID

	ConstFile.WriteBlankLines 2

	WriteLine -1, "IncludeStatsPopularMembers"
	WriteLine 1, "IncludeStatsPopularInfoPages"
	WriteLine -1, "IncludeStatsPopularAnnouncements"
	WriteLine -1, "IncludeStatsPopularMeetings"
	WriteLine -1, "IncludeStatsPopularStories"
	WriteLine -1, "IncludeStatsPopularCalendar"
	WriteLine -1, "IncludeStatsPopularLinks"
	WriteLine -1, "IncludeStatsPopularQuotes"
	WriteLine -1, "IncludeStatsPopularForumMessages"
	WriteLine -1, "IncludeStatsPopularVotingPolls"
	WriteLine -1, "IncludeStatsPopularQuizzes"
	WriteLine -1, "IncludeStatsPopularGuestbook"
	WriteLine -1, "IncludeStatsPopularMedia"
	WriteLine -1, "IncludeStatsPopularPhotos"
	WriteLine -1, "IncludeStatsPopularPhotoCaptions"
	WriteLine 1, "IncludeStatsPopularReviews"

	WriteLine -1, "IncludeStatsRatedMembers"
	WriteLine 1, "IncludeStatsRatedInfoPages"
	WriteLine -1, "IncludeStatsRatedAnnouncements"
	WriteLine -1, "IncludeStatsRatedMeetings"
	WriteLine -1, "IncludeStatsRatedStories"
	WriteLine -1, "IncludeStatsRatedCalendar"
	WriteLine -1, "IncludeStatsRatedLinks"
	WriteLine -1, "IncludeStatsRatedQuotes"
	WriteLine -1, "IncludeStatsRatedForumMessages"
	WriteLine -1, "IncludeStatsRatedVotingPolls"
	WriteLine -1, "IncludeStatsRatedQuizzes"
	WriteLine -1, "IncludeStatsRatedGuestbook"
	WriteLine -1, "IncludeStatsRatedMedia"
	WriteLine -1, "IncludeStatsRatedPhotos"
	WriteLine -1, "IncludeStatsRatedPhotoCaptions"
	WriteLine 1, "IncludeStatsRatedReviews"

	WriteLine -1, "IncludeStatsSummaryHomePage"
	WriteLine -1, "IncludeStatsSummaryMembers"
	WriteLine 1, "IncludeStatsSummaryInfoPages"
	WriteLine -1, "IncludeStatsSummaryAnnouncements"
	WriteLine -1, "IncludeStatsSummaryMeetings"
	WriteLine -1, "IncludeStatsSummaryStories"
	WriteLine -1, "IncludeStatsSummaryCalendar"
	WriteLine -1, "IncludeStatsSummaryLinks"
	WriteLine -1, "IncludeStatsSummaryQuotes"
	WriteLine -1, "IncludeStatsSummaryForumMessages"
	WriteLine -1, "IncludeStatsSummaryVotingPolls"
	WriteLine -1, "IncludeStatsSummaryQuizzes"
	WriteLine -1, "IncludeStatsSummaryGuestbook"
	WriteLine -1, "IncludeStatsSummaryMedia"
	WriteLine -1, "IncludeStatsSummaryPhotos"
	WriteLine -1, "IncludeStatsSummaryPhotoCaptions"
	WriteLine 1, "IncludeStatsSummaryReviews"



	ConstFile.WriteLine "'End Constants"

	ConstFile.WriteLine "%" & ">"

	ConstFile.Close
	Set ConstFile = Nothing

	FileSystem.CopyFile strPath & "stats_constantstemp.inc", strPath & "stats_constants.inc"

	Set FileSystem = Nothing

	if Request("Source") = "" then
		strSource = "admin_sectionoptions_edit.asp?Submit=Changed"
	else
		strSource = Request("Source")
	end if
%>
	<p>Your changes have been made. &nbsp;<a href="<%=strSource%>">Click here</a> to go back.<br>
	<a href="admin_stats_configure.asp">Change statistics options again.</a>
	</p>

<%
else
%>
	<form METHOD="post" ACTION="admin_stats_configure.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
		<input type="hidden" name="Source" value="<%=Request("Source")%>">

<%
	PrintTableHeader 0
%>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Section Label
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="StatsTitle" value="<%=StatsTitle%>" size="30">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When showing a list of the most popular or highest rated items, how many of the top items should be listed?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="StatTopMax" value="<%=StatTopMax%>" size="3">
			</td>
		</tr>
	</table>

	<p>Please choose which statistics you would like displayed:</p>
<%
	PrintTableHeader 0
%>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Home Page
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<% PrintCheckBox IncludeStatsSummaryHomePage, "IncludeStatsSummaryHomePage" %> Give home page summary (hits and searches).<br>
		</td>
	</tr>

<%
	PrintStatSection "Members", IncludeStatsSummaryMembers, IncludeStatsPopularMembers, IncludeStatsRatedMembers, "Members", "members", RateMembers
'	PrintStatSection "Information Pages", IncludeStatsSummaryInfoPages, IncludeStatsPopularInfoPages, IncludeStatsRatedInfoPages, "InfoPages", "pages", 0
	if CBool(IncludeAnnouncements) then PrintStatSection AnnouncementsTitle, IncludeStatsSummaryAnnouncements, IncludeStatsPopularAnnouncements, IncludeStatsRatedAnnouncements, "Announcements", "announcements", RateAnnouncements
	if CBool(IncludeMeetings) then PrintStatSection MeetingsTitle, IncludeStatsSummaryMeetings, IncludeStatsPopularMeetings, IncludeStatsRatedMeetings, "Meetings", "meetings", RateMeetings
	if CBool(IncludeStories) then PrintStatSection StoriesTitle, IncludeStatsSummaryStories, IncludeStatsPopularStories, IncludeStatsRatedStories, "Stories", "stories", RateStories
	if CBool(IncludeCalendar) then PrintStatSection CalendarTitle, IncludeStatsSummaryCalendar, IncludeStatsPopularCalendar, IncludeStatsRatedCalendar, "Calendar", "events", RateCalendar
	if CBool(IncludeLinks) then PrintStatSection LinksTitle, IncludeStatsSummaryLinks, IncludeStatsPopularLinks, IncludeStatsRatedLinks, "Links", "links", RateLinks
	if CBool(IncludeQuotes) then PrintStatSection QuotesTitle, IncludeStatsSummaryQuotes, IncludeStatsPopularQuotes, IncludeStatsRatedQuotes, "Quotes", "quotes", RateQuotes
	if CBool(IncludeForum) then PrintStatSection ForumTitle, IncludeStatsSummaryForumMessages, IncludeStatsPopularForumMessages, IncludeStatsRatedForumMessages, "ForumMessages", "messages", RateForum
	if CBool(IncludeVoting) then PrintStatSection VotingTitle, IncludeStatsSummaryVotingPolls, IncludeStatsPopularVotingPolls, IncludeStatsRatedVotingPolls, "VotingPolls", "polls", RateVoting
	if CBool(IncludeQuizzes) then PrintStatSection QuizzesTitle, IncludeStatsSummaryQuizzes, IncludeStatsPopularQuizzes, IncludeStatsRatedQuizzes, "Quizzes", "quizzes", RateQuizzes
	if CBool(IncludeGuestbook) then PrintStatSection GuestbookTitle, IncludeStatsSummaryGuestbook, IncludeStatsPopularGuestbook, IncludeStatsRatedGuestbook, "Guestbook", "entries", RateGuestbook
	if CBool(IncludeMedia) then PrintStatSection MediaTitle, IncludeStatsSummaryMedia, IncludeStatsPopularMedia, IncludeStatsRatedMedia, "Media", "files", RateMedia
	if CBool(IncludePhotos) then PrintStatSection PhotosTitle, IncludeStatsSummaryPhotos, IncludeStatsPopularPhotos, IncludeStatsRatedPhotos, "Photos", "photos", RatePhotos
	if CBool(IncludePhotos) then PrintStatSection "Photo Captions", IncludeStatsSummaryPhotoCaptions, IncludeStatsPopularPhotoCaptions, IncludeStatsRatedPhotoCaptions, "PhotoCaptions", "captions", RatePhotoCaptions

'	PrintStatSection "Reviews", IncludeStatsSummaryReviews, IncludeStatsPopularReviews, IncludeStatsPopularReviews, "Reviews", "reviews", 0

%>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Update">
    		</td>
		</tr>
	</table>
	</form>

<%

end if

Sub PrintStatSection( strTitle, intIncludeSummaryValue, intIncludePopularValue, intIncludeRatingValue, strName, strNoun, intPrintRating )
%>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			<%=strTitle%>
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<% PrintCheckBox intIncludeSummaryValue, "IncludeStatsSummary" & strName %> Give section summary.<br>
			<% PrintCheckBox intIncludePopularValue, "IncludeStatsPopular" & strName %> List most popular  <%=strNoun%>.<br>
<%
			if CBool(intPrintRating) then
%>
			<% PrintCheckBox intIncludeRatingValue, "IncludeStatsRated" & strName %> List highest rated  <%=strNoun%>.<br>
<%
			end if
%>
		</td>
	</tr>
<%
End Sub

Sub WriteLine( intPreMade, strName )
	if intPreMade = -1 then
		strAns = GetCheckedResult( Request(strName))
	else
		strAns = intPreMade
	end if
	ConstFile.WriteLine "Const " & strName & " = " & strAns
End Sub

%>
