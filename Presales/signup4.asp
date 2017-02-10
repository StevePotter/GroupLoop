<!-- #include file="header.asp" -->
<!-- #include file="functions.asp" -->
<!-- #include file="dsn.asp" -->
<% AddHit "signup4.asp" %>
<!-- #include file="closedsn.asp" -->

<%
'We are creating a new child site.. secret!
if Request("ParentID") <> "" then intParentID = Request("ParentID")

if Request("Version") = "" or ( Request("Version") <> "Gold" and Request("Version") <> "Free" and Request("Version") <> "Parent"  ) then Redirect("error.asp?Message=" & Server.URLEncode("You haven't chose which version you want.  Please go through the sign-up process from the beginning."))

strType = Request("Version")

intSchemeID = Request("ID")
if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the initial look.  Please go through the sign-up process from the beginning."))
%>

<p class=Heading align=center>
Step 4. Configure Your Site Sections
</p>

Select the initial options you would like for your site.  Remember, you may change the options 
later on.  This is just a starting point.

<form METHOD="post" ACTION="signup5.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
<input type="hidden" name="ID" value="<%=intSchemeID%>">
<input type="hidden" name="Version" value="<%=strType%>">
<%
	if intParentID <> "" then
%>
		<input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
		<input type="hidden" name="ParentID" value="<%=intParentID%>">
<%
	end if
	PrintTableHeader 0 %>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Is this site completely for members only?  If so, people 
			must log in to see anything, even the home page.  This is 
			recommended only if you want to keep your site totally private (sensitive content, don't want certain people 
			to see, etc).  
			If you choose 'No', members <b>can still have private posts (stories, messages, etc.).</b>  
			We usually recommend 'No'.
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 0, "SiteMembersOnly" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Can people apply for membership?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "AllowMemberApplications" %>
		</td>
	</tr>
	<tr>
		<td class="TDHeader" valign="middle" align="center" colspan="2">
			<strong>Newsletter</strong> - Send out an e-mail letter to anyone who subscribes.
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Include the Newsletter section?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "IncludeNewsletter" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Can any member manage subscriptions and send newsletters?  If not, only administrators (you) can.
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 0, "NewsletterMembers" %>
		</td>
	</tr>


	<tr>
		<td class="TDHeader" valign="middle" align="center" colspan="2">
			<strong>Announcements</strong> - members can post their latest announcements (obviously).
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Include the Announcements section?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "IncludeAnnouncements" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Allow people to rate announcements?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "RateAnnouncements" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Allow people to write reviews of announcements?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "ReviewAnnouncements" %>
		</td>
	</tr>

	<tr>
		<td class="TDHeader" valign="middle" align="center" colspan="2">
			<strong>Meeting Minutes</strong> - if your group has meetings, use this section to archive 
			summaries of each meeting.
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Include the Meeting Minutes section?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "IncludeMeetings" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Can any member add/modify meeting minutes?  If not, only administrators (you) can.
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 0, "MeetingsMembers" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Allow people to rate meeting minutes?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "RateMeetings" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Allow people to write reviews of meeting minutes?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "ReviewMeetings" %>
		</td>
	</tr>

	<tr>
		<td class="TDHeader" valign="middle" align="center" colspan="2">
			<strong>Stories</strong> - members use this section to tell personal narratives about 
			weekend parties, sports games, family gatherings, etc.
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Include the stories section?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "IncludeStories" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Allow people to rate stories?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "RateStories" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Allow people to write reviews of stories?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "ReviewStories" %>
		</td>
	</tr>
	<tr>
		<td class="TDHeader" valign="middle" align="center" colspan="2">
			<strong>Calendar</strong> - post up events, appointments, and other happenings and view them 
			in a nice calendar format.
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Include the calendar section?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "IncludeCalendar" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Automatically show members' birthdays in the calendar?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "CalendarShowBirthdays" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Allow people to rate calendar events?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "RateCalendar" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Allow people to write reviews of calendar events?
		</td>
		<td class="<% PrintTDMainSwitch %>" align="left">
			<% PrintRadio 1, "ReviewCalendar" %>
		</td>
	</tr>

	<tr>
		<td class="TDHeader" valign="middle" align="center" colspan="2">
			<strong>Links</strong> - members can put up links to their favorite web sites.
		</td>
	</tr>
	<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Include the Links section?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludeLinks" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to rate links?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "RateLinks" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to write reviews of links?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "ReviewLinks" %>
			</td>
		</tr>


		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Quotes</strong> - put up your favorite quotes here.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Include the quotes section?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludeQuotes" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to rate quotes?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "RateQuotes" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to write reviews of quotes?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "ReviewQuotes" %>
			</td>
		</tr>


		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Quizzes</strong> - create hilarious quizzes with as many questions as you want.  
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Include the quizzes section?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludeQuizzes" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Can members add/change/delete their own Quizzes?  If not, only administrators (you) can.
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "QuizzesMembers" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to rate quizzes?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "RateQuizzes" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to write reviews of quizzes?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "ReviewQuizzes" %>
			</td>
		</tr>


		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Voting Polls</strong> - put up polls to get feedback/opinions from visitors and members.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Include the voting section?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludeVoting" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Can members add/change/delete their own Voting Polls?  If not, only administrators (you) can.
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "VotingMembers" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to rate voting polls?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "RateVoting" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to write reviews of voting polls?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "ReviewVoting" %>
			</td>
		</tr>



		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Photos</strong> - easily add photos and laugh as members write up captions.  This 
				is probably the most popular and widely used section (and is extremely easy to use).
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Include the photos section?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludePhotos" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Can members add/change/delete their own Photos?  If not, only administrators (you) can.
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "PhotosMembers" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to rate photos?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "RatePhotos" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow members to write photo captions?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludePhotoCaptions" %>
			</td>
		</tr>



		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Message Forum</strong> - The message forum section is most likely 
				going to be your busiest section.  Members and/or visitors can interact with each other 
				and discuss any topic you choose.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Include the message forum?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludeForum" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to rate messages?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "RateForum" %>
			</td>
		</tr>



		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Guestbook</strong> - find out who is visiting your site with this section.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Include the guestbook?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludeGuestbook" %>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to rate entries?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "RateGuestbook" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to write reviews of entries?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "ReviewGuestbook" %>
			</td>
		</tr>


		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Media</strong> - Upload your favorite sounds, movies, and documents (such as Word 
				or Excel files) for everyone to download.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Include the media section?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "IncludeMedia" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Can members add/change/delete their own files?  If not, only administrators (you) can.
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "MediaMembers" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to rate files?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "RateMedia" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Allow people to write reviews of files?
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<% PrintRadio 1, "ReviewMedia" %>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="I'm Done">
			</td>
		</tr>

	</table>
	</form>

<!-- #include file="footer.asp" -->