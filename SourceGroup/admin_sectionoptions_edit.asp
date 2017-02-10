<!-- #include file="admin_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_sectionoptions_edit.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change Site Properties</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Update" then
	Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic



	Query = "SELECT * FROM Configuration WHERE CustomerID = " & CustomerID
	Set rsConfig = Server.CreateObject("ADODB.Recordset")
	rsConfig.Open Query, Connect, adOpenStatic, adLockOptimistic

	if (ParentSiteExists() or ChildSiteExists()) and Request("ShortTitle") <> "" then rsConfig("ShortTitle") = Request("ShortTitle")
	if Request("Title") <> "" then rsConfig("Title") = Format(Request("Title"))
	if Request("UsernameLabel") <> "" then rsConfig("UsernameLabel") = Format(Request("UsernameLabel"))




	if Request("Description") <> "" then rsLook("Description") = Request("Description")
	if Request("Keywords") <> "" then rsLook("Keywords") = Request("Keywords")
	if Request("FooterSource") <> "" then rsLook("FooterSource") = GetTextArea( Request("FooterSource") )


	if Request("AdditionsDaysOld") <> "" then rsConfig("AdditionsDaysOld") = Request("AdditionsDaysOld")
	if Request("StatTopMax") <> "" then rsConfig("StatTopMax") = Request("StatTopMax")
	if Request("PhotosPerRow") <> "" then rsConfig("PhotosPerRow") = Request("PhotosPerRow")
	if Request("VotingBarColor") <> "" then rsConfig("VotingBarColor") = GetColor("VotingBarColor")

	if Request("AllowMemberApplications") <> "" then rsConfig("AllowMemberApplications") = Request("AllowMemberApplications")

	if Request("CalendarShowBirthdays") <> "" then rsConfig("CalendarShowBirthdays") = Request("CalendarShowBirthdays")
	if Request("NewsShowEvents") <> "" then rsConfig("NewsShowEvents") = Request("NewsShowEvents")
	if Request("CalendarBirthdayMessage") <> "" then rsConfig("CalendarBirthdayMessage") = Format( Request("CalendarBirthdayMessage") )

	if Request("QuizResult90") <> "" then rsConfig("QuizResult90") = Format( Request("QuizResult90") )
	if Request("QuizResult60") <> "" then rsConfig("QuizResult60") = Format( Request("QuizResult60") )
	if Request("QuizResult0") <> "" then rsConfig("QuizResult0") = Format( Request("QuizResult0") )



	if Request("MemberNameDisplay") <> "" then rsConfig("MemberNameDisplay") = Request("MemberNameDisplay")
	if Request("IncludeAddButtons") <> "" then rsConfig("IncludeAddButtons") = Request("IncludeAddButtons")
	if Request("IncludeEditSectionPropButtons") <> "" then rsConfig("IncludeEditSectionPropButtons") = Request("IncludeEditSectionPropButtons")
	if Request("IncludeMemberStats") <> "" then rsConfig("IncludeMemberStats") = Request("IncludeMemberStats")



	if CBool( AllowStore ) and Request("IncludeStore") <> "" then rsConfig("IncludeStore") = Request("IncludeStore")


	if Request("PageSize") <> "" then rsConfig("PageSize") = Request("PageSize")
	if Request("RatingMax") <> "" then rsConfig("RatingMax") = Request("RatingMax")
	if Request("TellNew") <> "" then rsConfig("TellNew") = Request("TellNew")
	if Request("NewDaysOld") <> "" then rsConfig("NewDaysOld") = Request("NewDaysOld")

	if Request("SiteMembersOnly") <> "" then rsConfig("SiteMembersOnly") = Request("SiteMembersOnly")
	if Request("SecureLogin") <> "" and Version = "Gold" then rsConfig("SecureLogin") = Request("SecureLogin")


	SetFields "", "Members", rsConfig
	SetFields "", "Title", rsConfig
	SetFields "Include", "", rsConfig
	SetFields "Rate", "", rsConfig
	SetFields "Review", "", rsConfig
	SetFields "SectionView", "", rsConfig
	SetFields "ListType", "", rsLook
	SetFields "DisplayDaysOld", "", rsLook
	SetFields "DisplayDateList", "", rsLook
	SetFields "DisplayAuthorList", "", rsLook
	SetFields "DisplayPrivacyList", "", rsLook
	SetFields "DisplayDateItem", "", rsLook
	SetFields "DisplayAuthorItem", "", rsLook
	SetFields "DisplaySubjectItem", "", rsLook
	SetFields "IncludePrivacy", "", rsLook
	SetInfoText
	SetFields "DisplaySearch", "", rsLook
	SetFields "DisplayDaysOld", "", rsLook


	SetField "DisplayNickNameListMembers", rsLook
	SetField "DisplayPhotoListMembers", rsLook
	SetField "DisplayFullNameListMembers", rsLook
	SetField "DisplayBirthdayListMembers", rsLook
	SetField "DisplayEMailListMembers", rsLook
	SetField "DisplayHomeAddressListMembers", rsLook
	SetField "DisplaySecondaryAddressListMembers", rsLook
	SetField "DisplayBeeperListMembers", rsLook
	SetField "DisplayCellPhoneListMembers", rsLook
	SetField "DisplayMembershipLevelListMembers", rsLook

	SetField "DisplayNickNameItemMembers", rsLook
	SetField "DisplayPhotoItemMembers", rsLook
	SetField "DisplayFullNameItemMembers", rsLook
	SetField "DisplayBirthdayItemMembers", rsLook
	SetField "DisplayEMailItemMembers", rsLook
	SetField "DisplayHomeAddressItemMembers", rsLook
	SetField "DisplaySecondaryAddressItemMembers", rsLook
	SetField "DisplayBeeperItemMembers", rsLook
	SetField "DisplayCellPhoneItemMembers", rsLook

	Sub SetField( strFieldName, rsObject )
		if Request(strFieldName) <> "" then rsObject(strFieldName) = Request(strFieldName)
	End Sub

	Sub SetFields( strFieldStart, strFieldEnd, rsObject )

	if Request(strFieldStart & "Additions" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Additions" & strFieldEnd ) = Request(strFieldStart & "Additions" & strFieldEnd )
	if Request(strFieldStart & "Announcements" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Announcements" & strFieldEnd ) = Request(strFieldStart & "Announcements" & strFieldEnd )
	if Request(strFieldStart & "Calendar" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Calendar" & strFieldEnd ) = Request(strFieldStart & "Calendar" & strFieldEnd )
	if Request(strFieldStart & "Forum" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Forum" & strFieldEnd ) = Request(strFieldStart & "Forum" & strFieldEnd )
	if Request(strFieldStart & "Guestbook" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Guestbook" & strFieldEnd ) = Request(strFieldStart & "Guestbook" & strFieldEnd )
	if Request(strFieldStart & "Links" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Links" & strFieldEnd ) = Request(strFieldStart & "Links" & strFieldEnd )
	if Request(strFieldStart & "Media" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Media" & strFieldEnd ) = Request(strFieldStart & "Media" & strFieldEnd )
	if Request(strFieldStart & "Meetings" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Meetings" & strFieldEnd ) = Request(strFieldStart & "Meetings" & strFieldEnd )
	if Request(strFieldStart & "Members" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Members" & strFieldEnd ) = Request(strFieldStart & "Members" & strFieldEnd )
	if Request(strFieldStart & "News" & strFieldEnd ) <> "" then rsObject(strFieldStart & "News" & strFieldEnd ) = Request(strFieldStart & "News" & strFieldEnd )
	if Request(strFieldStart & "Newsletter" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Newsletter" & strFieldEnd ) = Request(strFieldStart & "Newsletter" & strFieldEnd )
	if Request(strFieldStart & "Photos" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Photos" & strFieldEnd ) = Request(strFieldStart & "Photos" & strFieldEnd )
	if Request(strFieldStart & "PhotoCaptions" & strFieldEnd ) <> "" then rsObject(strFieldStart & "PhotoCaptions" & strFieldEnd ) = Request(strFieldStart & "PhotoCaptions" & strFieldEnd )
	if Request(strFieldStart & "Quizzes" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Quizzes" & strFieldEnd ) = Request(strFieldStart & "Quizzes" & strFieldEnd )
	if Request(strFieldStart & "Quotes" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Quotes" & strFieldEnd ) = Request(strFieldStart & "Quotes" & strFieldEnd )
	if Request(strFieldStart & "Stats" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Stats" & strFieldEnd ) = Request(strFieldStart & "Stats" & strFieldEnd )
	if Request(strFieldStart & "Store" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Store" & strFieldEnd ) = Request(strFieldStart & "Store" & strFieldEnd )
	if Request(strFieldStart & "Stories" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Stories" & strFieldEnd ) = Request(strFieldStart & "Stories" & strFieldEnd )
	if Request(strFieldStart & "Voting" & strFieldEnd ) <> "" then rsObject(strFieldStart & "Voting" & strFieldEnd ) = Request(strFieldStart & "Voting" & strFieldEnd )

	End Sub

	Sub SetInfoText

	if Request("GetInfoTextAdditions"  ) = "YES" then rsLook("InfoTextAdditions"  ) = GetTextArea(Request("InfoTextAdditions"  ) )
	if Request("GetInfoTextAnnouncements"  ) = "YES" then rsLook("InfoTextAnnouncements"  ) = GetTextArea(Request("InfoTextAnnouncements"  ) )
	if Request("GetInfoTextCalendar"  ) = "YES" then rsLook("InfoTextCalendar"  ) = GetTextArea(Request("InfoTextCalendar"  ) )
	if Request("GetInfoTextForum"  ) = "YES" then rsLook("InfoTextForum"  ) = GetTextArea(Request("InfoTextForum"  ) )
	if Request("GetInfoTextGuestbook"  ) = "YES" then rsLook("InfoTextGuestbook"  ) = GetTextArea(Request("InfoTextGuestbook"  ) )
	if Request("GetInfoTextLinks"  ) = "YES" then rsLook("InfoTextLinks"  ) = GetTextArea(Request("InfoTextLinks"  ) )
	if Request("GetInfoTextMedia"  ) = "YES" then rsLook("InfoTextMedia"  ) = GetTextArea(Request("InfoTextMedia"  ) )
	if Request("GetInfoTextMeetings"  ) = "YES" then rsLook("InfoTextMeetings"  ) = GetTextArea(Request("InfoTextMeetings"  ) )
	if Request("GetInfoTextMembers"  ) = "YES" then rsLook("InfoTextMembers"  ) = GetTextArea(Request("InfoTextMembers"  ) )
	if Request("GetInfoTextNews"  ) = "YES" then rsLook("InfoTextNews"  ) = GetTextArea(Request("InfoTextNews"  ) )
	if Request("GetInfoTextNewsletter"  ) = "YES" then rsLook("InfoTextNewsletter"  ) = GetTextArea(Request("InfoTextNewsletter"  ) )
	if Request("GetInfoTextPhotos"  ) = "YES" then rsLook("InfoTextPhotos"  ) = GetTextArea(Request("InfoTextPhotos"  ) )
	if Request("GetInfoTextPhotoCaptions"  ) = "YES" then rsLook("InfoTextPhotoCaptions"  ) = GetTextArea(Request("InfoTextPhotoCaptions"  ) )
	if Request("GetInfoTextQuizzes"  ) = "YES" then rsLook("InfoTextQuizzes"  ) = GetTextArea(Request("InfoTextQuizzes"  ) )
	if Request("GetInfoTextQuotes"  ) = "YES" then rsLook("InfoTextQuotes"  ) = GetTextArea(Request("InfoTextQuotes"  ) )
	if Request("GetInfoTextStats"  ) = "YES" then rsLook("InfoTextStats"  ) = GetTextArea(Request("InfoTextStats"  ) )
	if Request("GetInfoTextStore"  ) = "YES" then rsLook("InfoTextStore"  ) = GetTextArea(Request("InfoTextStore"  ) )
	if Request("GetInfoTextStories"  ) = "YES" then rsLook("InfoTextStories"  ) = GetTextArea(Request("InfoTextStories"  ) )
	if Request("GetInfoTextVoting"  ) = "YES" then rsLook("InfoTextVoting"  ) = GetTextArea(Request("InfoTextVoting"  ) )

	End Sub


	rsConfig.Update
	rsConfig.Close
	Set rsConfig = Nothing
	rsLook.Update
	set rsLook = Nothing
%>
	<!-- #include file="write_constants.asp" -->
	<!-- #include file="write_index.asp" -->
<%
	if Request("Source") = "" then
		strSource = "admin_sectionoptions_edit.asp?Submit=Changed"
	else
		strSource = Request("Source")
	end if
	Redirect("write_header_footer.asp?Source=" & strSource)
elseif Request("Submit") = "Changed" then
	'This is here so changes can be seen right away
'------------------------End Code-----------------------------
%>
		<p>Your changes have been made. &nbsp;<a href="admin_sectionoptions_edit.asp">Click here</a> to make more changes.</p>
<%
'-----------------------Begin Code----------------------------
elseif Request("Type") = "" then
%>
		<b>What would you like to do?</b><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Site">Change main site options</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Visuals">Customize site visuals</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Membership">Change membership options</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Sections">Choose which sections to use</a><br>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties">Change section properties</a><br>
<%
else
	Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockReadOnly


	if Request("Type") = "Visuals" then
'------------------------End Code-----------------------------
%>
	<b>Visual Customization</b><br>
	There are all kinds of ways to change the look of your site.  Simply select what you want to change below:<br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_look_edit.asp">Colors and Fonts</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_buttons_modify.asp">Menu Buttons</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_layout_edit.asp">Page Layout</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_images_edit.asp">Graphics</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_schemes.asp">Scheme Manager</a><br>
	&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_advanced_visuals_edit.asp">Other Visuals</a><br>


<%
	elseif Request("Type") = "Site" then
'------------------------End Code-----------------------------
%>


		<form METHOD="post" ACTION="admin_sectionoptions_edit.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
		<input type="hidden" name="Source" value="<%=Request("Source")%>">
<%
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Main Site Information</strong>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Site title
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="Title" value="<%=Title%>" size="50">
			</td>
		</tr>
		<%
			if ParentSiteExists() or ChildSiteExists() then
				Query = "SELECT ShortTitle FROM Configuration WHERE CustomerID = " & CustomerID
				Set rsConfig = Server.CreateObject("ADODB.Recordset")
				rsConfig.Open Query, Connect, adOpenStatic, adLockReadOnly
		%>
				<tr>
					<td class="<% PrintTDMain %>" valign="middle" align="right">
						Shortened Title - Used for links to this site on the menus of your other sites.
					</td>
					<td class="<% PrintTDMain %>" align="left">
						<input type="text" name="ShortTitle" value="<%=rsConfig("ShortTitle")%>" size="50">
					</td>
				</tr>
		<%
				Set rsConfig = Nothing
			end if
		%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				A short description of your site.
			</td>
			<td class="<% PrintTDMain %>" align="left">
    			<textarea name="Description" cols="55" rows="4" wrap="PHYSICAL"><%=rsLook("Description") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				A list of keywords to describe your site.
			</td>
			<td class="<% PrintTDMain %>" align="left">
    			<textarea name="Keywords" cols="55" rows="4" wrap="PHYSICAL"><%=rsLook("Keywords") %></textarea>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What text should be shown in the footer?  This is could be contact information, or anything else 
				that deserves to be on the bottom of every page.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% TextArea "FooterSource", 55, 4, True, rsLook("FooterSource") %>
			</td>
		</tr>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong>Site Security</strong>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Is this site completely for members only?  If so, anyone 
				must log in to see anything, even the home page.  This is nh
				recommended only if you want to keep your site totally private.  
				If you choose 'No', members can still have private posts (stories, messages, etc.).
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio SiteMembersOnly, "SiteMembersOnly" %>
			</td>
		</tr>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Other
			</td>
		</tr>
<%
		if not cBool(SiteMembersOnly) then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Display the "Add A..." link automatically at the top of your sections?  If you click no, the link will hidden from non-members 
				or anyone without access to add an item.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeAddButtons, "IncludeAddButtons" %>
			</td>
		</tr>
<%
		end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Display the "Change This Section's Options" link automatically at the bottoms of your sections?  These links can <b>ONLY</b> be seen by you, 
				not regular members.  Just choose no if you don't want to see the link ever.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeEditSectionPropButtons, "IncludeEditSectionPropButtons" %>
			</td>
		</tr>


		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When listing items (such as stories), how many items can be displayed per page (no more than 40 is recommended)?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" size="3" name="PageSize" value="<%=PageSize%>">
			</td>
		</tr>


		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When someone rates an item, what is the maximum rating it can get?  For allowing 1-5 ratings, enter 5.  For 1-10, enter 10.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" size="3" name="RatingMax" value="<%=RatingMax%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Should new items have 'New!' in front of them?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio TellNew, "TellNew" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				If yes, for how many days should the 'New!' be displayed?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" size="2" name="NewDaysOld" value="<%=NewDaysOld%>">
			</td>
		</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
					<input type="submit" name="Submit" value="Update">
				</td>
			</tr>

		</table>
		</form>


<%
'-----------------------Begin Code----------------------------
	elseif Request("Type") = "Membership" then
		Query = "SELECT IncludeMemberStats FROM Configuration WHERE CustomerID = " & CustomerID
		Set rsConf = Server.CreateObject("ADODB.Recordset")
		rsConf.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

		IncludeMemberStats = rsConf("IncludeMemberStats")

		rsConf.Close
		Set rsConf = Nothing
'------------------------End Code-----------------------------
%>
	<script language="JavaScript1.1"><!--
	function FixCheckboxes(what) {
		for (var i=0, j=what.elements.length; i<j; i++) {
			myType = what.elements[i].type;
			if (myType == 'checkbox') {
				if (!what.elements[i].checked){
					what.elements[i].value = '0';
					what.elements[i].checked = true;
				}
			}
		}
	}
	//--></script>

		<form METHOD="post" ACTION="admin_sectionoptions_edit.asp" name="MyForm" onSubmit="FixCheckboxes(this); if (this.submitted) return false; this.submitted = true; return true">
		<input type="hidden" name="Source" value="<%=Request("Source")%>">

<%
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong><%=MembersTitle%></strong>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Section Label for the <%=MembersTitle%> section.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="MembersTitle" value="<%=MembersTitle%>" size="30">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Each user has a special, unique user name for their site.  Initially it is called a Nickname, but can also be called "Username", "Screen Name", 
				or whatever you choose.  This  <%=MembersTitle%> section.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="UsernameLabel" value="<%=UsernameLabel%>" size="30">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Can anyone apply for membership?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio AllowMemberApplications, "AllowMemberApplications" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Let people
			</td>
			<td class="<% PrintTDMain %>" align="left">
			<%	PrintCheckBox RateMembers, "RateMembers" %> Rate site members<br>
			<%	PrintCheckBox ReviewMembers, "ReviewMembers" %> Write reviews for site members<br>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When listing members in the <a href="members_info_view.asp"><%=PrintTDLink("View Everyone's Info")%></a> section, which information about each member should be shown?
			</td>
			<td class="<% PrintTDMain %>" align="left">
			<%	PrintCheckBox rsLook("DisplayNickNameListMembers"), "DisplayNickNameListMembers" %> <%=UsernameLabel%><br>
			<%	PrintCheckBox rsLook("DisplayPhotoListMembers"), "DisplayPhotoListMembers" %> A thumbnail (small picture) of themselves, if they have added one.<br>


			<%	PrintCheckBox rsLook("DisplayFullNameListMembers"), "DisplayFullNameListMembers" %> Full Name<br>
			<%	PrintCheckBox rsLook("DisplayBirthdayListMembers"), "DisplayBirthdayListMembers" %> Birthday<br>
			<%	PrintCheckBox rsLook("DisplayEMailListMembers"), "DisplayEMailListMembers" %> E-Mail Address<br>
			<%	PrintCheckBox rsLook("DisplayHomeAddressListMembers"), "DisplayHomeAddressListMembers" %> Home Address<br>
			<%	PrintCheckBox rsLook("DisplaySecondaryAddressListMembers"), "DisplaySecondaryAddressListMembers" %> Secondary Address<br>
			<%	PrintCheckBox rsLook("DisplayBeeperListMembers"), "DisplayBeeperListMembers" %> Beeper Number<br>
			<%	PrintCheckBox rsLook("DisplayCellPhoneListMembers"), "DisplayCellPhoneListMembers" %> Cell Phone Number<br>
			<%	PrintCheckBox rsLook("DisplayMembershipLevelListMembers"), "DisplayMembershipLevelListMembers" %> Membership Level (regular member, administrator)<br>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When viewing an individual member, which information should be shown?
			</td>
			<td class="<% PrintTDMain %>" align="left">
			<%	PrintCheckBox rsLook("DisplayNickNameItemMembers"), "DisplayNickNameItemMembers" %> <%=UsernameLabel%><br>
			<%	PrintCheckBox rsLook("DisplayPhotoItemMembers"), "DisplayPhotoItemMembers" %> A thumbnail (small picture) of themselves, if they have added one.<br>
			<%	PrintCheckBox rsLook("DisplayFullNameItemMembers"), "DisplayFullNameItemMembers" %> Full Name<br>
			<%	PrintCheckBox rsLook("DisplayBirthdayItemMembers"), "DisplayBirthdayItemMembers" %> Birthday<br>
			<%	PrintCheckBox rsLook("DisplayEMailItemMembers"), "DisplayEMailItemMembers" %> E-Mail Address<br>
			<%	PrintCheckBox rsLook("DisplayHomeAddressItemMembers"), "DisplayHomeAddressItemMembers" %> Home Address<br>
			<%	PrintCheckBox rsLook("DisplaySecondaryAddressItemMembers"), "DisplaySecondaryAddressItemMembers" %> Secondary Address<br>
			<%	PrintCheckBox rsLook("DisplayBeeperItemMembers"), "DisplayBeeperItemMembers" %> Beeper Number<br>
			<%	PrintCheckBox rsLook("DisplayCellPhoneItemMembers"), "DisplayCellPhoneItemMembers" %> Cell Phone Number<br>
			</td>
		</tr>


		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When a member's name is displayed on the site (usually along with an item they added), what should be displayed?  
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Their 
				<select name="MemberNameDisplay" size="1">
				<%
					WriteOption "NickName", "NickName", MemberNameDisplay
					WriteOption "FirstName", "First Name", MemberNameDisplay
					WriteOption "LastName", "Last Name", MemberNameDisplay
					WriteOption "FullName", "First and Last Name", MemberNameDisplay
				%>
				</select>				
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				When someone clicks on a member's name, should their individual highlights (statistics) be displayed?
			</td>
			<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio IncludeMemberStats, "IncludeMemberStats" %>
			</td>
		</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
					<input type="submit" name="Submit" value="Update">
				</td>
			</tr>

		</table>
		</form>
<%
'-----------------------Begin Code----------------------------
	elseif Request("Type") = "Sections" then
'------------------------End Code-----------------------------
%>
		<form METHOD="post" ACTION="admin_sectionoptions_edit.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
		<input type="hidden" name="Source" value="<%=Request("Source")%>">

<%
		PrintTableHeader 0

		if CBool(AllowStore) then
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Use the <%=StoreTitle%> section?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio IncludeStore, "IncludeStore" %>
				</td>
			</tr>
<%
		end if
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=AnnouncementsTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeAnnouncements, "IncludeAnnouncements" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=MeetingsTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeMeetings, "IncludeMeetings" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=CalendarTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeCalendar, "IncludeCalendar" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=StoriesTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeStories, "IncludeStories" %>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=LinksTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeLinks, "IncludeLinks" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=QuotesTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeQuotes, "IncludeQuotes" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=GuestbookTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeGuestbook, "IncludeGuestbook" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=ForumTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeForum, "IncludeForum" %>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=VotingTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeVoting, "IncludeVoting" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=QuizzesTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeQuizzes, "IncludeQuizzes" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=PhotosTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludePhotos, "IncludePhotos" %>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=MediaTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeMedia, "IncludeMedia" %>
			</td>
		</tr>

		

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=NewsletterTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeNewsletter, "IncludeNewsletter" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=StatsTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeStats, "IncludeStats" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Use the <%=AdditionsTitle%> section?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio IncludeAdditions, "IncludeAdditions" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Update">
			</td>
		</tr>

	</table>
	</form>		
	
<%
	elseif Request("Section") = "" then
%>

		<b>What section's properties would you like to change?</b><br>
<%
		if CBool(IncludeStore) and CBool(AllowStore) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Store"><%=StoreTitle%></a><br>
<%
		end if
		if CBool(IncludeAnnouncements) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Announcements"><%=AnnouncementsTitle%></a><br>
<%
		end if
		if CBool(IncludeMeetings) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Meetings"><%=MeetingsTitle%></a><br>
<%
		end if
		if CBool(IncludeCalendar) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Calendar"><%=CalendarTitle%></a><br>
<%
		end if
		if CBool(IncludeStories) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Stories"><%=StoriesTitle%></a><br>
<%
		end if
		if CBool(IncludeLinks) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Links"><%=LinksTitle%></a><br>
<%
		end if
		if CBool(IncludeQuotes) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Quotes"><%=QuotesTitle%></a><br>
<%
		end if
		if CBool(IncludeGuestbook) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Guestbook"><%=GuestbookTitle%></a><br>
<%
		end if
		if CBool(IncludeForum) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Forum"><%=ForumTitle%></a><br>
<%
		end if
		if CBool(IncludeVoting) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Voting"><%=VotingTitle%></a><br>
<%
		end if
		if CBool(IncludeQuizzes) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Quizzes"><%=QuizzesTitle%></a><br>
<%
		end if
		if CBool(IncludePhotos) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Photos"><%=PhotosTitle%></a><br>
<%
		end if
		if CBool(IncludeMedia) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Media"><%=MediaTitle%></a><br>
<%
		end if
		if CBool(IncludeNewsletter) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Newsletter"><%=NewsletterTitle%></a><br>
<%
		end if
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=News"><%=NewsTitle%></a><br>
<%
		if CBool(IncludeStats) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_sectionoptions_edit.asp?Type=Properties&Section=Stats"><%=StatsTitle%></a><br>
<%
		end if
		if CBool(IncludeAdditions) then
%>
		&nbsp;&nbsp;&nbsp;&nbsp;<a href="admin_additions_configure.asp"><%=AdditionsTitle%></a><br>
<%
		end if

	else
%>
<script language="JavaScript1.1"><!--
function FixCheckboxes(what) {
    for (var i=0, j=what.elements.length; i<j; i++) {
        myType = what.elements[i].type;
        if (myType == 'checkbox') {
            if (!what.elements[i].checked){
				what.elements[i].value = '0';
				what.elements[i].checked = true;
			}
        }
    }
}
//--></script>
		<form METHOD="post" ACTION="admin_sectionoptions_edit.asp" name="MyForm">
		<input type="hidden" name="Source" value="<%=Request("Source")%>">

<%
		PrintTableHeader 100

		if Request("Section") = "Store" and CBool( AllowStore ) then
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=StoreTitle%></strong>
				</td>

			</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Section Label
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="StoreTitle" value="<%=StoreTitle%>" size="30">
				</td>
			</tr>

<%
		elseif Request("Section") = "Announcements" then

			PrintSection "Announcements", AnnouncementsTitle, IncludeAnnouncements, "announcement", "announcements", true

		elseif Request("Section") = "Meetings" then

			PrintSection "Meetings", MeetingsTitle, IncludeMeetings, "meeting", "meetings", true

		elseif Request("Section") = "Calendar" then
			PrintSection "Calendar", CalendarTitle, IncludeCalendar, "event", "calendar events", true
%>
			</table>

			<%	PrintTableHeader 100	%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Automatically show members' birthdays in the calendar?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadioNew CalendarShowBirthdays, "CalendarShowBirthdays", "show('BDays');", "hide('BDays');" %>
				</td>
			</tr>
			<span id="BDays" <%=GetDisplay(CalendarShowBirthdays)%>>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Message displayed when someone clicks a on person's birthday (same message is shown in the News)
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="CalendarBirthdayMessage" value="<%=CalendarBirthdayMessage%>" size="30">
				</td>
			</tr>
			</table>
			</span>

			<%	PrintTableHeader 100	%>
<%
		elseif Request("Section") = "Stories" then

			PrintSection "Stories", StoriesTitle, IncludeStories, "story", "stories", true

		elseif Request("Section") = "Links" then

			PrintSection "Links", LinksTitle, IncludeLinks, "link", "links", false

		elseif Request("Section") = "Quotes" then

			PrintSection "Quotes", QuotesTitle, IncludeQuotes, "quote", "quotes", false

		elseif Request("Section") = "Guestbook" then
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=GuestbookTitle%></strong>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Section Label
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="GuestbookTitle" value="<%=GuestbookTitle%>" size="30">
				</td>
			</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Let people
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox RateGuestbook, "RateGuestbook" %> Rate guestbook entries<br>
				<%	PrintCheckBox ReviewGuestbook, "ReviewGuestbook" %> Review guestbook entries<br>
				<%	PrintCheckBox rsLook("DisplaySearchGuestbook"), "DisplaySearchGuestbook" %> Search guestbook entries<br>
				<%	PrintCheckBox rsLook("DisplayDaysOldGuestbook"), "DisplayDaysOldGuestbook" %> View guestbook  entries added in the last x number of days<br>
				</td>
			</tr>

			<tr> 
    			<td class="<% PrintTDMain %>" align="right" valign="middle">If you would like a message to appear when someone visits the  
				 <%=GuestbookTitle%> section, please enter it here.</td>
    			<td class="<% PrintTDMain %>"> 
					<input type="hidden" name="GetInfoText<%=Request("Section")%>" value="YES">
					<% TextArea "InfoText"&Request("Section"), 55, 10, True, rsLook("InfoText"&Request("Section")) %>
    			</td>
			</tr>

<%
		elseif Request("Section") = "Forum" then
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=ForumTitle%></strong>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Section Label
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="ForumTitle" value="<%=ForumTitle%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Once a member posts a message, can they change/delete it later?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio ForumMembers, "ForumMembers" %>
				</td>
			</tr>
<%
			if not cBool(SiteMembersOnly) then
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Can there be messages for members only?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "IncludePrivacyForum", 1, "Yes, allow private messages<br>", rsLook("IncludePrivacyForum")
					PrintRadioOption "IncludePrivacyForum", 0, "No, all messages can be read by anyone (non-members included)<br>", rsLook("IncludePrivacyForum")
%>
				</td>
			</tr>
<%
			end if
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Let people
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox RateForum, "RateForum" %> Rate messages<br>
				<%	PrintCheckBox rsLook("DisplaySearchForum"), "DisplaySearchForum" %> Search for messages<br>
				</td>
			</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When listing messages, which information about each message should be shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox rsLook("DisplayDateListForum"), "DisplayDateListForum" %> Date Written<br>
				<%  if not cBool(SiteMembersOnly) then %>
					<%	PrintCheckBox rsLook("DisplayPrivacyListForum"), "DisplayPrivacyListForum" %> Privacy (Public or Private)
				<%  end if %>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When viewing an individual message, which information should be shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox rsLook("DisplayDateItemForum"), "DisplayDateItemForum" %> Date Written<br>
				<%	PrintCheckBox rsLook("DisplaySubjectItemForum"), "DisplaySubjectItemForum" %> Subject<br>
				</td>
			</tr>
			<tr> 
    			<td class="<% PrintTDMain %>" align="right" valign="middle">If you would like a message to appear when someone visits the  
				 <%=ForumTitle%> section, please enter it here.</td>
    			<td class="<% PrintTDMain %>"> 
					<input type="hidden" name="GetInfoText<%=Request("Section")%>" value="YES">
					<% TextArea "InfoText"&Request("Section"), 55, 10, True, rsLook("InfoText"&Request("Section")) %>
    			</td>
			</tr>
<%
		elseif Request("Section") = "Voting" then
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=VotingTitle%></strong>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Section Label
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="VotingTitle" value="<%=VotingTitle%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Color of the voting results bars
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintColors "VotingBarColor", VotingBarColor %>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Who can add voting polls?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "VotingMembers", 1, "Any Member<br>", VotingMembers
					PrintRadioOption "VotingMembers", 0, "Just Administrators<br>", VotingMembers
%>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Let people
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox RateVoting, "RateVoting" %> Rate Voting Polls<br>
				<%	PrintCheckBox ReviewVoting, "ReviewVoting" %> Review Voting Polls<br>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					How should polls be listed?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "ListTypeVoting", "Table", "In a Table<br>", rsLook("ListTypeVoting")
					PrintRadioOption "ListTypeVoting", "Bulleted", "Bulleted List<br>", rsLook("ListTypeVoting")
					PrintRadioOption "ListTypeVoting", "Numbered", "Numbered List<br>", rsLook("ListTypeVoting")
					PrintRadioOption "ListTypeVoting", "Plain", "Plain, Unordered List<br>", rsLook("ListTypeVoting")
%>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When listing polls, which information about each poll should be shown?
				</td>



				<td class="<% PrintTDMain %>" align="left">

				<%	PrintCheckBox rsLook("DisplayDateListVoting"), "DisplayDateListVoting" %> Date Written<br>
				<%	PrintCheckBox rsLook("DisplayAuthorListVoting"), "DisplayAuthorListVoting" %> Author<br>
				<%  if not cBool(SiteMembersOnly) then %>
					<%	PrintCheckBox rsLook("DisplayPrivacyListVoting"), "DisplayPrivacyListVoting" %> Privacy (Public or Private)
				<%  end if %>
				</td>
			</tr>
			<tr> 
    			<td class="<% PrintTDMain %>" align="right" valign="middle">If you would like a message to appear when someone visits the  
				 <%=VotingTitle%> section, please enter it here.</td>
    			<td class="<% PrintTDMain %>"> 
					<input type="hidden" name="GetInfoText<%=Request("Section")%>" value="YES">
					<% TextArea "InfoText"&Request("Section"), 55, 10, True, rsLook("InfoText"&Request("Section")) %>
    			</td>
			</tr>
<%
		elseif Request("Section") = "Quizzes" then
%>

			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=QuizzesTitle%></strong>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Section Label
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="QuizzesTitle" value="<%=QuizzesTitle%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Who can add quizzes?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "QuizzesMembers", 1, "Any Member<br>", QuizzesMembers
					PrintRadioOption "QuizzesMembers", 0, "Just Administrators<br>", QuizzesMembers
%>
				</td>
			</tr>

<%
			if not cBool(SiteMembersOnly) then
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Can there be quizzes for members only?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "IncludePrivacyQuizzes", 1, "Yes, allow private quizzes<br>", rsLook("IncludePrivacyQuizzes")
					PrintRadioOption "IncludePrivacyQuizzes", 0, "No, all quizzes can be read by anyone (non-members included)<br>", rsLook("IncludePrivacyQuizzes")
%>
				</td>
			</tr>
<%
			end if
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Let people
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox RateQuizzes, "RateQuizzes" %> Rate quizzes<br>
				<%	PrintCheckBox ReviewQuizzes, "ReviewQuizzes" %> Review quizzes<br>
				</td>
			</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					How should quizzes be listed?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "ListTypeQuizzes", "Table", "In a Table<br>", rsLook("ListTypeQuizzes")
					PrintRadioOption "ListTypeQuizzes", "Bulleted", "Bulleted List<br>", rsLook("ListTypeQuizzes")
					PrintRadioOption "ListTypeQuizzes", "Numbered", "Numbered List<br>", rsLook("ListTypeQuizzes")
					PrintRadioOption "ListTypeQuizzes", "Plain", "Plain, Unordered List<br>", rsLook("ListTypeQuizzes")
%>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When listing quizzes, which information about each quiz should be shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox rsLook("DisplayDateListQuizzes"), "DisplayDateListQuizzes" %> Date Written<br>
				<%	PrintCheckBox rsLook("DisplayAuthorListQuizzes"), "DisplayAuthorListQuizzes" %> Author<br>
				<%  if not cBool(SiteMembersOnly) then %>
					<%	PrintCheckBox rsLook("DisplayPrivacyListQuizzes"), "DisplayPrivacyListQuizzes" %> Privacy (Public or Private)
				<%  end if %>
				</td>
			</tr>

			<tr> 
    			<td class="<% PrintTDMain %>" align="right" valign="middle">If you would like a message to appear when someone visits the  
				 <%=QuizzesTitle%> section, please enter it here.</td>
    			<td class="<% PrintTDMain %>"> 
					<input type="hidden" name="GetInfoText<%=Request("Section")%>" value="YES">
					<% TextArea "InfoText"&Request("Section"), 55, 10, True, rsLook("InfoText"&Request("Section")) %>
    			</td>
			</tr>


			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When someone scores a 90  or better on a quiz, what message is shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="QuizResult90" value="<%=QuizResult90%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When someone scores betwee a 60 and 90 on a quiz, what message is shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="QuizResult60" value="<%=QuizResult60%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When someone scores below a 60 on a quiz, what message is shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="QuizResult0" value="<%=QuizResult0%>" size="30">
				</td>
			</tr>
<%
		elseif Request("Section") = "Photos" then
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=PhotosTitle%></strong>
				</td>
			</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Section Label
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="PhotosTitle" value="<%=PhotosTitle%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Who can add photos?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "PhotosMembers", 1, "Any Member<br>", PhotosMembers
					PrintRadioOption "PhotosMembers", 0, "Just Administrators<br>", PhotosMembers
%>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Let people
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox RatePhotos, "RatePhotos" %> Rate photos<br>
				<%	PrintCheckBox rsLook("DisplaySearchPhotos"), "DisplaySearchPhotos" %> Search photos<br>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Number of thumbnails shown per row when listing the photos
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="PhotosPerRow" value="<%=PhotosPerRow%>" size="2">
				</td>
			</tr>
			<tr> 
    			<td class="<% PrintTDMain %>" align="right" valign="middle">If you would like a message to appear when someone visits the  
				 <%=PhotosTitle%> section, please enter it here.</td>
    			<td class="<% PrintTDMain %>"> 
					<input type="hidden" name="GetInfoText<%=Request("Section")%>" value="YES">
					<% TextArea "InfoText"&Request("Section"), 55, 10, True, rsLook("InfoText"&Request("Section")) %>
    			</td>
			</tr>

			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong>Photo Captions</strong>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Allow members to write photo captions?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio IncludePhotoCaptions, "IncludePhotoCaptions" %>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Let people
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox RatePhotoCaptions, "RatePhotoCaptions" %> Rate photos captions<br>
				<%	PrintCheckBox ReviewPhotoCaptions, "ReviewPhotoCaptions" %> Review photo captions<br>
				</td>
			</tr>

<%
		elseif Request("Section") = "Media" then
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=MediaTitle%></strong>
				</td>
			</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Title
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="MediaTitle" value="<%=MediaTitle%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Can members add/change/delete their own files?  If not, only administrators can.
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio MediaMembers, "MediaMembers" %>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Can people rate files?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio RateMedia, "RateMedia" %>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Can people review files?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio ReviewMedia, "ReviewMedia" %>
				</td>
			</tr>

<%
		elseif Request("Section") = "Newsletter" then
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=NewsletterTitle%></strong>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Title
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="NewsletterTitle" value="<%=NewsletterTitle%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Can members manage subscriptions, send newsletters, and modify old newsletters?  If not, only administrators can.
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio NewsletterMembers, "NewsletterMembers" %>
				</td>
			</tr>
<%
		elseif Request("Section") = "News" then
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><%=NewsTitle%></strong>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Title
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="NewsTitle" value="<%=NewsTitle%>" size="30">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Automatically show today's calendar events in the news?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadio NewsShowEvents, "NewsShowEvents" %>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					How should the news be listed?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "ListTypeNews", "Table", "In a Table<br>", rsLook("ListTypeNews")
					PrintRadioOption "ListTypeNews", "Bulleted", "Bulleted List<br>", rsLook("ListTypeNews")
					PrintRadioOption "ListTypeNews", "Numbered", "Numbered List<br>", rsLook("ListTypeNews")
					PrintRadioOption "ListTypeNews", "Plain", "Plain, Unordered List<br>", rsLook("ListTypeNews")
%>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When listing news, which information about each news piece should be shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox rsLook("DisplayDateListNews"), "DisplayDateListNews" %> Date Written<br>
				<%	PrintCheckBox rsLook("DisplayAuthorListNews"), "DisplayAuthorListNews" %> Author<br>
				</td>
			</tr>

<%
		elseif Request("Section") = "Stats" then
%>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<strong><%=StatsTitle%></strong>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Title
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


<%
		end if
%>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
					<input type="submit" name="Submit" value="Update" onClick="FixCheckboxes(this.form);<%if UseWYSIWYGEdit() then Response.Write "ae_onSubmit();"%> if (this.form.submitted) return false; this.form.submitted = true; return true">
				</td>
			</tr>

		</table>
		</form>
<%
	end if

end if
'------------------------End Code-----------------------------
			Sub PrintRadioNew( intBool, strName, strOnClickYes, strOnClickNo )
				strChecked1 = ""
				strChecked2 = ""
				if intBool = 1 then
					strChecked1 = "checked"
				else
					strChecked2 = "checked"
				end if
			%>
					<input type="radio" name="<%=strName%>" value="1" <%=strChecked1%> onClick="<%=strOnClickYes%>">
					Yes 
					<input type="radio" name="<%=strName%>" value="0" <%=strChecked2%> onClick="<%=strOnClickNo%>">
					No 	
			<%
			End Sub

			Function GetDisplay( blDisplay )
				blDisplay = CBool(blDisplay)
				if blDisplay then
					GetDisplay = ""
				else
					GetDisplay = " style=" & chr(34) & "display: none;" & chr(34)
				end if
			End Function

'-------------------------------------------------------------
'This function writes a pulldown menu for colors or a radio table for colors if needed
'-------------------------------------------------------------
Sub PrintColors( strName, strSelectColor )
	Query = "SELECT Color FROM Colors ORDER BY ID"
	Set rsColors = Server.CreateObject("ADODB.Recordset")
	rsColors.CacheSize = 100
	rsColors.Open Query, Connect, adOpenStatic, adLockReadOnly

	Set Color = rsColors("Color")

	'If we didn't highlight the color (and it isn't empty), then it will be put in the custom box
	blFound = false
	'Well, no color, so nothing in the custom box
	if strSelectColor = "" then blFound = true

	if Request("Colors") = "Radio" then
%>
		<table cellspacing="0" cellpadding="1"><tr>
			<td class="<% PrintTDMain %>"><input type="radio" name="<%=strName%>" value="" checked>
			None</td>
<%
		p = 0
		do until rsColors.EOF
			p = p + 1
			if p mod 7 = 0 then Response.Write "</tr><tr>"
			if Color = strSelectColor then
				blFound = true
%>
				<td class="<% PrintTDMain %>"><input type="radio" name="<%=strName%>" value="<%=Color%>" checked>
				<font color="<%=Color%>"><%=Color%> </font></td>
<%
			else
%>
				<td class="<% PrintTDMain %>"><input type="radio" name="<%=strName%>" value="<%=Color%>">
				<font color="<%=Color%>"><%=Color%> </font></td>
<%
			end if
			rsColors.MoveNext
		loop
		Response.Write "</tr></table>"
	else
%>
		<select name="<%=strName%>">
<%
		if strSelectColor = "" then strSelect = " selected"
%>		<option value="" <%=strSelect%>>None</option>
<%
		do until rsColors.EOF
			if Color = strSelectColor then
				blFound = true
%>
				<option value="<%=Color%>" style="BACKGROUND: <%=Color%>;" selected><%=Color%></option>
<%
			else
%>
				<option value="<%=Color%>" style="BACKGROUND: <%=Color%>;"><%=Color%></option>
<%
			end if
			rsColors.MoveNext
		loop
%>
		</select>
<%
	end if
	'So if they already had a custom color, then put it in the box
	strValue = ""
	if blFound = false then strValue = strSelectColor
%>
	  or enter a custom color <input type=text size=7 name="<%=strName%>Cust" value="<%=strValue%>">
<%
	set rsColors = Nothing
End Sub



Sub PrintSection( strSection, strTitle, intInclude, strSingular, strPlural, blItemSubject )
%>
			<tr>
				<td class="TDHeader" valign="middle" align="center" colspan="2">
					<strong><span id="SLabel"><%=strTitle%></span></strong>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right" width="50%">
					Use the <span id="SLabel"><%=strTitle%></span> section?
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<% PrintRadioNew intInclude, "Include" & strSection, "show('SectionDetails');", "hide('SectionDetails');" %>
				</td>
			</tr>
			</table>

			<span id="SectionDetails" <%=GetDisplay(intInclude)%>>
			<%	PrintTableHeader 100	%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Section label.  This is the main title of the section.
				</td>
				<td class="<% PrintTDMain %>" align="left">
					<input type="text" name="strTitle" value="<%=strTitle%>" size="30" onChange="changeWord('SLabel', this.value);">
				</td>
			</tr>
<%
			if not cBool(SiteMembersOnly) then
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Who can view the <span id="SLabel"><%=strTitle%></span> section?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					if not cBool(SiteMembersOnly) then PrintRadioOption "SectionView" & strSection, "Anyone", "Anyone<br>", SectionViewMeetings
					PrintRadioOption "SectionView" & strSection, "Members", "Site Members<br>", SectionViewMeetings
'					PrintRadioOption "SectionView" & strSection, "Administrators", "Only Administrators<br>", SectionViewMeetings
%>
				</td>
			</tr>
<%
			end if
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Who can add <%=strPlural%>?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "MeetingsMembers", 1, "Any Member<br>", MeetingsMembers
					PrintRadioOption "MeetingsMembers", 0, "Just Administrators<br>", MeetingsMembers
%>
				</td>
			</tr>

<%
			if not cBool(SiteMembersOnly) then
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Can there be <%=strPlural%> for members only?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "IncludePrivacy" & strSection, 1, "Yes, allow private " & strPlural & "<br>", rsLook("IncludePrivacy" & strSection)
					PrintRadioOption "IncludePrivacy" & strSection, 0, "No, all " & strPlural & " can be read by anyone (non-members included)<br>", rsLook("IncludePrivacy" & strSection)
%>
				</td>
			</tr>
<%
			end if
%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					Let people
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox RateMeetings, "Rate" & strSection %> Rate <%=strPlural%><br>
				<%	PrintCheckBox ReviewMeetings, "Review" & strSection %> Review <%=strPlural%><br>
				<%	PrintCheckBox rsLook("DisplaySearch" & strSection), "DisplaySearch" & strSection %> Search <%=strPlural%><br>
				<%	PrintCheckBox rsLook("DisplayDaysOld" & strSection), "DisplayDaysOld" & strSection %> View <%=strPlural%> added in the last x number of days<br>
				</td>
			</tr>

			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					How should <%=strPlural%> be listed?
				</td>
				<td class="<% PrintTDMain %>" align="left">
<%
					PrintRadioOption "ListType" & strSection, "Table", "In a Table<br>", rsLook("ListType" & strSection)
					PrintRadioOption "ListType" & strSection, "Bulleted", "Bulleted List<br>", rsLook("ListType" & strSection)
					PrintRadioOption "ListType" & strSection, "Numbered", "Numbered List<br>", rsLook("ListType" & strSection)
					PrintRadioOption "ListType" & strSection, "Plain", "Plain, Unordered List<br>", rsLook("ListType" & strSection)
%>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When listing <%=strPlural%>, which information about each <%=strSingular%> should be shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox rsLook("DisplayDateList" & strSection), "DisplayDateList" & strSection %> Date Written<br>
				<%	PrintCheckBox rsLook("DisplayAuthorList" & strSection), "DisplayAuthorList" & strSection %> Author<br>
				<%  if not cBool(SiteMembersOnly) then %>
					<%	PrintCheckBox rsLook("DisplayPrivacyList" & strSection), "DisplayPrivacyList" & strSection %> Privacy (Public or Private)
				<%  end if %>
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					When viewing an individual <%=strSingular%>, which information should be shown?
				</td>
				<td class="<% PrintTDMain %>" align="left">
				<%	PrintCheckBox rsLook("DisplayDateItem" & strSection), "DisplayDateItem" & strSection %> Date Written<br>
				<%	PrintCheckBox rsLook("DisplayAuthorItem" & strSection), "DisplayAuthorItem" & strSection %> Author<br>
<%
				if blItemSubject then
%>
				<%	PrintCheckBox rsLook("DisplaySubjectItem" & strSection), "DisplaySubjectItem" & strSection %> Subject<br>

<%
				end if
%>
				</td>
			</tr>
			<tr> 
    			<td class="<% PrintTDMain %>" align="right" valign="middle">If you would like a message to appear when someone visits the  
				 <span id="SLabel"><%=strTitle%></span> section, please enter it here.</td>
    			<td class="<% PrintTDMain %>"> 
					<input type="hidden" name="GetInfoText<%=Request("Section")%>" value="YES">
					<% TextArea "InfoText"&Request("Section"), 55, 10, True, rsLook("InfoText"&strSection) %>
    			</td>
			</tr>
			</table>
			</span>

			<%	PrintTableHeader 100	%>

<%
End Sub
%>
