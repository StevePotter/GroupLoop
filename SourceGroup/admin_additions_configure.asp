<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_additions_configure.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Configure <%=AdditionsTitle%></span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Create Files----------------------------
if Request("Submit") = "Update" then
	Query = "SELECT AdditionsDaysOld, AdditionsTitle FROM Configuration WHERE CustomerID = " & CustomerID
	Set rsConfig = Server.CreateObject("ADODB.Recordset")
	rsConfig.Open Query, Connect, adOpenStatic, adLockOptimistic

	if Request("AdditionsDaysOld") <> "" then rsConfig("AdditionsDaysOld") = Request("AdditionsDaysOld")
	if Request("AdditionsTitle") <> "" then rsConfig("AdditionsTitle") = Format( Request("AdditionsTitle") )

	rsConfig.Update
	rsConfig.Close
	Set rsConfig = Nothing
%>
	<!-- #include file="write_constants.asp" -->
<%
	if strPath = "" then strPath = GetPath("")
	strImagePath = strPath & "images/"
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	Set ConstFile = FileSystem.CreateTextFile(strPath & "additions_constantstemp.inc")

	DblQuote = Chr(34)

	'-----------------------Start Writing Files----------------------------

	'Create our opening delimiter
	ConstFile.WriteLine "<" & "%"
	ConstFile.WriteBlankLines 2

	ConstFile.WriteLine "'This file contains the constants for this site's lastest additions (cust #" & CustomerID

	ConstFile.WriteBlankLines 2

	'Do we include the section?
	ConstFile.WriteLine "Const IncludeAdditionsMembers = " & GetCheckedResult( Request("IncludeAdditionsMembers"))
	ConstFile.WriteLine "Const IncludeAdditionsInfoPages = " & GetCheckedResult( Request("IncludeAdditionsInfoPages"))
	ConstFile.WriteLine "Const IncludeAdditionsAnnouncements = " & GetCheckedResult( Request("IncludeAdditionsAnnouncements"))
	ConstFile.WriteLine "Const IncludeAdditionsMeetings = " & GetCheckedResult( Request("IncludeAdditionsMeetings"))
	ConstFile.WriteLine "Const IncludeAdditionsStories = " & GetCheckedResult( Request("IncludeAdditionsStories"))
	ConstFile.WriteLine "Const IncludeAdditionsCalendar = " & GetCheckedResult( Request("IncludeAdditionsCalendar"))
	ConstFile.WriteLine "Const IncludeAdditionsLinks = " & GetCheckedResult( Request("IncludeAdditionsLinks"))
	ConstFile.WriteLine "Const IncludeAdditionsQuotes = " & GetCheckedResult( Request("IncludeAdditionsQuotes"))
	ConstFile.WriteLine "Const IncludeAdditionsForumMessages = " & GetCheckedResult( Request("IncludeAdditionsForumMessages"))
	ConstFile.WriteLine "Const IncludeAdditionsVotingPolls = " & GetCheckedResult( Request("IncludeAdditionsVotingPolls"))
	ConstFile.WriteLine "Const IncludeAdditionsQuizzes = " & GetCheckedResult( Request("IncludeAdditionsQuizzes"))
	ConstFile.WriteLine "Const IncludeAdditionsGuestbook = " & GetCheckedResult( Request("IncludeAdditionsGuestbook"))
	ConstFile.WriteLine "Const IncludeAdditionsMedia = " & GetCheckedResult( Request("IncludeAdditionsMedia"))
	ConstFile.WriteLine "Const IncludeAdditionsPhotos = " & GetCheckedResult( Request("IncludeAdditionsPhotos"))
	ConstFile.WriteLine "Const IncludeAdditionsPhotoCaptions = " & GetCheckedResult( Request("IncludeAdditionsPhotoCaptions"))
	ConstFile.WriteLine "Const IncludeAdditionsReviews = " & GetCheckedResult( Request("IncludeAdditionsReviews"))

	'Do we include the author for each?
	ConstFile.WriteLine "Const IncludeAdditionsInfoPagesAuthor = " & GetCheckedResult( Request("IncludeAdditionsInfoPagesAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsAnnouncementsAuthor = " & GetCheckedResult( Request("IncludeAdditionsAnnouncementsAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsMeetingsAuthor = " & GetCheckedResult( Request("IncludeAdditionsMeetingsAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsStoriesAuthor = " & GetCheckedResult( Request("IncludeAdditionsStoriesAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsCalendarAuthor = " & GetCheckedResult( Request("IncludeAdditionsCalendarAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsLinksAuthor = " & GetCheckedResult( Request("IncludeAdditionsLinksAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsQuotesAuthor = " & GetCheckedResult( Request("IncludeAdditionsQuotesAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsForumMessagesAuthor = " & GetCheckedResult( Request("IncludeAdditionsForumMessagesAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsVotingPollsAuthor = " & GetCheckedResult( Request("IncludeAdditionsVotingPollsAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsQuizzesAuthor = " & GetCheckedResult( Request("IncludeAdditionsQuizzesAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsGuestbookAuthor = " & GetCheckedResult( Request("IncludeAdditionsGuestbookAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsMediaAuthor = " & GetCheckedResult( Request("IncludeAdditionsMediaAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsPhotosAuthor = " & GetCheckedResult( Request("IncludeAdditionsPhotosAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsPhotoCaptionsAuthor = " & GetCheckedResult( Request("IncludeAdditionsPhotoCaptionsAuthor"))
	ConstFile.WriteLine "Const IncludeAdditionsReviewsAuthor = " & GetCheckedResult( Request("IncludeAdditionsReviewsAuthor"))


	ConstFile.WriteLine "'End Constants"

	ConstFile.WriteLine "%" & ">"

	ConstFile.Close
	Set ConstFile = Nothing

	FileSystem.CopyFile strPath & "additions_constantstemp.inc", strPath & "additions_constants.inc"

	Set FileSystem = Nothing
%>
	<p>Your changes have been made. &nbsp;<a href="admin_sectionoptions_edit.asp">Click here</a> to make further changes.</p>

<%
else
%>
	<form METHOD="post" ACTION="admin_additions_configure.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">

<%
	PrintTableHeader 0
%>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Section Label
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="AdditionsTitle" value="<%=AdditionsTitle%>" size="30">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What is the default number of days to leave items on the list for? 
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="AdditionsDaysOld" value="<%=AdditionsDaysOld%>" size="4">
			</td>
		</tr>
	</table>

	<p>Please select which sections have their additions displayed:</p>
<%
	PrintTableHeader 0

	PrintAdditionSection "New Members", IncludeAdditionsMembers, 0, "IncludeAdditionsMembers", False
	PrintAdditionSection "Information Pages", IncludeAdditionsInfoPages, IncludeAdditionsInfoPagesAuthor, "IncludeAdditionsInfoPages", True
	if CBool(IncludeAnnouncements) then PrintAdditionSection AnnouncementsTitle, IncludeAdditionsAnnouncements, IncludeAdditionsAnnouncementsAuthor, "IncludeAdditionsAnnouncements", True
	if CBool(IncludeMeetings) then PrintAdditionSection MeetingsTitle, IncludeAdditionsMeetings, IncludeAdditionsMeetingsAuthor, "IncludeAdditionsMeetings", True
	if CBool(IncludeStories) then PrintAdditionSection StoriesTitle, IncludeAdditionsStories, IncludeAdditionsStoriesAuthor, "IncludeAdditionsStories", True
	if CBool(IncludeCalendar) then PrintAdditionSection CalendarTitle, IncludeAdditionsCalendar, IncludeAdditionsCalendarAuthor, "IncludeAdditionsCalendar", True
	if CBool(IncludeLinks) then PrintAdditionSection LinksTitle, IncludeAdditionsLinks, IncludeAdditionsLinksAuthor, "IncludeAdditionsLinks", True
	if CBool(IncludeQuotes) then PrintAdditionSection QuotesTitle, IncludeAdditionsQuotes, IncludeAdditionsQuotesAuthor, "IncludeAdditionsQuotes", True
	if CBool(IncludeForum) then PrintAdditionSection ForumTitle, IncludeAdditionsForumMessages, IncludeAdditionsForumMessagesAuthor, "IncludeAdditionsForumMessages", True
	if CBool(IncludeVoting) then PrintAdditionSection VotingTitle, IncludeAdditionsVotingPolls, IncludeAdditionsVotingPollsAuthor, "IncludeAdditionsVotingPolls", True
	if CBool(IncludeQuizzes) then PrintAdditionSection QuizzesTitle, IncludeAdditionsQuizzes, IncludeAdditionsQuizzesAuthor, "IncludeAdditionsQuizzes", True
	if CBool(IncludeGuestbook) then PrintAdditionSection GuestbookTitle, IncludeAdditionsGuestbook, IncludeAdditionsGuestbookAuthor, "IncludeAdditionsGuestbook", True
	if CBool(IncludeMedia) then PrintAdditionSection MediaTitle, IncludeAdditionsMedia, IncludeAdditionsMediaAuthor, "IncludeAdditionsMedia", True
	if CBool(IncludePhotos) then PrintAdditionSection PhotosTitle, IncludeAdditionsPhotos, IncludeAdditionsPhotosAuthor, "IncludeAdditionsPhotos", True
	if CBool(IncludePhotos) then PrintAdditionSection "Photo Captions", IncludeAdditionsPhotoCaptions, IncludeAdditionsPhotoCaptionsAuthor, "IncludeAdditionsPhotoCaptions", True

	PrintAdditionSection "Reviews", IncludeAdditionsReviews, IncludeAdditionsReviewsAuthor, "IncludeAdditionsReviews", True
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

Sub PrintAdditionSection( strTitle, intIncludeValue, intIncludeAuthorValue, strName, blPrintAuthor )
%>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			<%=strTitle%>
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<% PrintCheckBox intIncludeValue, strName %> Include Additions<br>
<%
			if blPrintAuthor then
%>
			<% PrintCheckBox intIncludeAuthorValue, strName&"Author" %> Display Author<br>
<%
			end if
%>
		</td>
	</tr>
<%
End Sub
%>
