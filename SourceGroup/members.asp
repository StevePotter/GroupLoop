<p class="Heading" align="<%=HeadingAlignment%>"><%=MembersTitle%></p>

<%
'
'-----------------------Begin Code----------------------------	
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	MemberLogin strPassword, strNickName
end if

blLoggedAdmin = LoggedAdmin()


if LoggedMember() then
	'if we need to send them to a specific page, do it
	if Request("Source") <> "" then Redirect(Request("Source"))

	Session.Timeout = 20


	strSection = Request("Section")

	if strSection <> "" then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a></p>
<%
	end if

	if strSection = "" then
'------------------------End Code-----------------------------
%>
		Hello <%=GetNickNameSession%>. Here are your options:<br>
		<a href="http://www.GroupLoop.com/manuals/members" target="_blank">Member's Manual</a><br>
	<%
		if blLoggedAdmin then
	%>
			<a href="http://www.GroupLoop.com/manuals/admin" target="_blank">Administrator's Manual</a><br>
			<br>
<%
			MembersApplied
%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp">Change Site Properties</a><br>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_pagehits.asp">Page Hits</a><br>
	<%
			if Version = "Free" then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<b><a href="https://www.OurClubPage.com/<%=SubDirectory%>/admin_account_upgrade.asp?MemberID=<%=Session("MemberID")%>&Password=<%=Session("Password")%>">Upgrade To Gold Version</a></b><br>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_account_remove.asp">Terminate My Site</a><br>
	<%
			else
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_account_modify.asp">Manage Account</a><br>
	<%
			end if

		end if

	%>


		<br>If you want to add or change any item in a section, please choose the section below:<br>
	<%
			if CBool(AllowStore) and CBool( IncludeStore ) then
%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Store"><%=StoreTitle%></a><br>
<%
			end if
			if blLoggedAdmin then
	%>
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=News"><%=NewsTitle%></a><br>
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Members">Memberships</a><br>
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=InfoPages">Information Pages</a><br>
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Inserts">Page Inserts</a><br>
	<%
			else
%>
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Members">Your Membership</a><br>
				&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_info_view.asp">View Everyone's Info</a><br>
<%
			end if
			if CBool( IncludeAnnouncements ) and (blLoggedAdmin or CBool( AnnouncementsMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Announcements"><%=AnnouncementsTitle%></a><br>
	<%
			end if
			if CBool( IncludeMeetings ) and (blLoggedAdmin or CBool( MeetingsMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Meetings"><%=MeetingsTitle%></a><br>
	<%
			end if
			if CBool( IncludeCalendar ) and (blLoggedAdmin or CBool( CalendarMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Calendar"><%=CalendarTitle%></a><br>
	<%
			end if
			if CBool( IncludeStories ) and (blLoggedAdmin or CBool( StoriesMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Stories"><%=StoriesTitle%></a><br>
	<%
			end if
			if CBool( IncludeLinks ) and (blLoggedAdmin or CBool( LinksMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Links"><%=LinksTitle%></a><br>
	<%
			end if
			if CBool( IncludeQuotes ) and (blLoggedAdmin or CBool( QuotesMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Quotes"><%=QuotesTitle%></a><br>
	<%
			end if
			if CBool(IncludeGuestbook) and blLoggedAdmin then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Guestbook"><%=GuestbookTitle%></a><br>
	<%
			end if
			if CBool(IncludeForum) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Forum"><%=ForumTitle%></a><br>
	<%
			end if
			if CBool( IncludePhotos ) and (blLoggedAdmin or CBool( PhotosMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Photos"><%=PhotosTitle%></a><br>
	<%
			end if
			if CBool( IncludeVoting ) and (blLoggedAdmin or CBool( VotingMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Voting"><%=VotingTitle%></a><br>
	<%
			end if
			if CBool( IncludeQuizzes ) and (blLoggedAdmin or CBool( QuizzesMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Quizzes"><%=QuizzesTitle%></a><br>
	<%
			end if
			if CBool( IncludeMedia ) AND (blLoggedAdmin or CBool( MediaMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Media"><%=MediaTitle%></a><br>
	<%
			end if
			if CBool( IncludeNewsletter ) AND (blLoggedAdmin or CBool( NewsletterMembers )) then
	%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="members.asp?Section=Newsletter"><%=NewsletterTitle%></a><br>
	<%
			end if
			PrintCustom
	%>
			<br>
			<a href="members_relog.asp">Log in as a different member</a><br>
<%
	elseif strSection = "News" and blLoggedAdmin then
%>
		<strong><%=NewsTitle%></strong><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_news_add.asp">Add A News Update</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_news_modify.asp">Modify News</a><br>
<%
	elseif strSection = "Members" then
		if not blLoggedAdmin then strSection = "Your Membership Information"
%>
		<strong><%=strSection%></strong><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_info_edit.asp">Change Your Information</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_preferences_edit.asp">Change Your Preferences</a><br>
	<%
		if blLoggedAdmin then
	%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_info_view.asp">View Everyone's Info</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_members_add.asp">Add A New Member</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_members_modify.asp">Modify Members</a><br>

	<%
		end if
	elseif strSection = "InfoPages" and blLoggedAdmin then
%>
	<strong>Information Pages</strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_pages_add.asp">Add A New Info Page</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=NonSecurePath%>members_pages_modify.asp">Modify Info Pages</a><br>
<%
	elseif strSection = "Inserts" and blLoggedAdmin then
%>
	<strong>Inserts</strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=NonSecurePath%>members_inserts_modify.asp?Submit=Add">Add Inserts</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=NonSecurePath%>members_inserts_modify.asp">Modify Inserts</a><br>

<%
	elseif strSection = "Announcements" and CBool(IncludeAnnouncements) then
%>
	<strong><%=AnnouncementsTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_announcements_add.asp">Add An Announcement</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_announcements_modify.asp">Modify Announcements</a><br>

<%
	elseif strSection = "Meetings" and (blLoggedAdmin or CBool( MeetingsMembers )) then
%>
	<strong><%=MeetingsTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_meetings_add.asp">Add A Meeting</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_meetings_modify.asp">Modify Meetings</a><br>
<%

	elseif strSection = "Calendar" and CBool(IncludeCalendar) then
%>
	<strong><%=CalendarTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_calendar_add.asp">Add An Event</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_calendar_modify.asp">Modify Events</a><br>

<%
	elseif strSection = "Stories" and CBool(IncludeStories) then
%>
	<strong><%=StoriesTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_stories_add.asp">Add A Story</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_stories_modify.asp">Modify Stories</a><br>

<%
	elseif strSection = "Links" and CBool(IncludeLinks) then
%>
	<strong><%=LinksTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_links_add.asp">Add A Link</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_links_modify.asp">Modify Links</a><br>
<%
	elseif strSection = "Quotes" and CBool(IncludeQuotes) then
%>
	<strong><%=QuotesTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_quotes_add.asp">Add A Quote</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_quotes_modify.asp">Modify Quotes</a><br>
<%
	elseif strSection = "Guestbook" and blLoggedAdmin and CBool(IncludeGuestbook) then
%>
	<strong><%=GuestbookTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_guestbook_modify.asp">Modify Entries</a><br>
<%
	elseif strSection = "Forum" and CBool(IncludeForum) then
%>
	<strong><%=ForumTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="forum.asp">Add Messages</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_forum_modify.asp">Modify Messages</a><br>
<%
		if blLoggedAdmin then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_forumcategories_add.asp">Add A New Topic</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_forumcategories_modify.asp">Modify Topics</a><br>
<%
		end if
	elseif strSection = "Photos" and CBool( IncludePhotos ) and (blLoggedAdmin or CBool( PhotosMembers )) then
%>
	<strong><%=PhotosTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_photos_add.asp">Add A New Photo</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_photos_modify.asp">Modify Photos</a><br>
<%
		if blLoggedAdmin then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_photocategories_add.asp">Add A New Category</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_photocategories_modify.asp">Modify Categories</a><br>
<%
		end if
	elseif strSection = "Voting" and CBool( IncludeVoting ) and (blLoggedAdmin or CBool( VotingMembers )) then
%>
	<strong><%=VotingTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_polls_add.asp">Add A New Poll</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_polls_modify.asp">Modify Polls</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_polls_close.asp">Close A Poll</a><br>
<%
	elseif strSection = "Quizzes" and CBool( IncludeQuizzes ) and (blLoggedAdmin or CBool( QuizzesMembers )) then
%>
	<strong><%=QuizzesTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_quizzes_add.asp">Add A New Quiz</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_quizzes_modify.asp">Modify Quizzes</a><br>
<%
	elseif strSection = "Media" and CBool( IncludeMedia ) AND (blLoggedAdmin or CBool( MediaMembers )) then
%>
	<strong><%=MediaTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_media_add.asp">Add A File</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_media_modify.asp">Modify Files</a><br>
<%
		if blLoggedAdmin then
%>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_mediacategories_add.asp">Add A New Category</a><br>
			&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_mediacategories_modify.asp">Modify Categories</a><br>
<%
		end if
	elseif strSection = "Newsletter" and CBool( IncludeNewsletter ) and (blLoggedAdmin or CBool( NewsletterMembers )) then
%>
	<strong><%=NewsletterTitle%></strong><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_newsletter_subscribers.asp">Manage Subscriptions</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_newsletter_add.asp">Send A New Newsletter</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="members_newsletter_modify.asp">Modify Old Newsletters</a><br>
<%
	elseif strSection = "Store" and CBool(AllowStore) and CBool( IncludeStore ) then
%>
			<strong><%=StoreTitle%></strong><br>
<%
		if blLoggedAdmin then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=NonSecurePath%>admin_store_configure.asp">Configure Your Store</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_storeshipping_configure.asp">Configure Shipping Options</a><br>
<%
		end if
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=NonSecurePath%>members_store_modify.asp">Modify Store Contents</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=NonSecurePath%>members_storespecials_modify.asp">Modify Store Specials</a><br>
<%
		if blLoggedAdmin or OrderAccess() then
			if NewOrdersExist() then
%>
			&nbsp;&nbsp;&nbsp;&nbsp;<b><a href="<%=SecurePath%>members_storeorders_new.asp">View New Orders</a></b><br>
<%
			end if
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="<%=SecurePath%>members_storeorders_modify.asp">View/Modify Old Orders</a><br>
<%
		end if

	end if
%>


<%
else
	Redirect("login.asp?Source=members.asp")
end if

Sub MembersApplied()
			intMembersApplied = GetNumItems("MembersApplied")
			if intMembersApplied > 0 then
				if intMembersApplied = 1 then
					strPer = "person has"
				else
					strPer = "people have"
				end if
%>
			&nbsp;&nbsp;&nbsp;&nbsp;<b><font size="+1"><a href="admin_members_applied_add.asp"><%=intMembersApplied%>&nbsp;<%=strPer%>  applied for membership.  Click here to accept/decline them.</a></font></b><p></p>
<%
			end if
End Sub

'-------------------------------------------------------------
'This function sees if a category exists
'-------------------------------------------------------------
Function NewOrdersExist()
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "StoreNewOrdersExist"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blResult = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	NewOrdersExist = CBool(blResult)
End Function

Function OrderAccess()
	OrderAccess = LoggedMember and Session("OrderAccess") = "Y"
End Function
%>
