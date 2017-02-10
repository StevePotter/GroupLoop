<!-- #include file="admin_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_images_edit.asp")
Session.Timeout = 20
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change Graphics</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
if Request("Submit") = "Changed" then
'------------------------End Code-----------------------------
%>
		<p>The changes have been made. 
		If you added new images (such as a background) and can't see the changes, simply press the Reload or Refresh button on your browser.</p>
		<p>You can &nbsp;<a href="admin_images_edit.asp">make more changes</a> or 
		<a href="admin_sectionoptions_edit.asp?Type=Visuals">go back to visual customization</a>. 
		</p>

<%
'-----------------------Begin Code----------------------------
else
	Function GetDisplay( blDisplay )
		blDisplay = CBool(blDisplay)
		if blDisplay then
			GetDisplay = ""
		else
			GetDisplay = " style=" & chr(34) & "display: none;" & chr(34)
		end if
	End Function

Set FileSystem = CreateObject("Scripting.FileSystemObject")
strImagePath = GetPath("images")

strDisplay = Request("Display")
	if strDisplay = "Show" then
'------------------------End Code-----------------------------
%>
	<p><a href="admin_images_edit.asp">Collapse All Boxes</a></p>
<%
	else
%>
	<p><a href="admin_images_edit.asp?Display=Show">Expand All Boxes</a></p>
<%
	end if


	Sub PrintTitle( strTitle, strDivName )

%>
			<% PrintTableHeader 50 %>
				<tr>
					<td class="TDHeader" valign="middle" align="left" colspan="2">
						<a class="TDHeader" HREF="javascript://" onClick="switchdisplay('<%=strDivName%>div'); return false"><%=strTitle%></a>
					</td>
				</tr>
			</table>
<%
		if strDisplay = "Show" then
			'insert the "div" after the name just to avoid dupes that are already there
%>
			<span id="<%=strDivName%>div" <%=GetDisplay(1)%>>
<%
		else
%>
			<span id="<%=strDivName%>div" <%=GetDisplay(0)%>>
<%
		end if
		PrintTableHeader 0%>
<%

	End Sub
%>
	<a href="admin_advanced_images_edit.asp">Click here</a> for the advanced graphics settings.<br>
	<p>If you already have graphics and would like to keep them, just leave the button at 'Yes' and leave the file box blank. 
	If you have a graphic you would like to delete, just click the 'No' button, and leave the file box blank.  
	A rollover images is a graphic that gets switched with the original once someone puts their mouse pointer 
	over it.  That's how you can get the buttons on your site to 'light up' when someone goes to click on them.  
	Rollover images are optional, so if you don't have any, don't worry.
	</p> 
	<p>Remember that only image files may be uploaded.  Anything else will be rejected.</p>
	<form enctype="multipart/form-data" method="post" ACTION="<%=SecurePath%>admin_images_edit_process.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">

<%
	PrintTitle "Site Title", "title"
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Title Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a title image? 	<% PrintRadio ImageExistsInt("TitleImage"), "TitleImage" %><br>
				Image <input type="file" name="UpTitleImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Title Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a title rollover image? 	<% PrintRadio ImageExistsInt("TitleRolloverImage"), "TitleRolloverImage" %><br>
				Image <input type="file" name="UpTitleRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Background Image for Title
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a background image? 	<% PrintRadio ImageExistsInt("TitleMenuBackgroundImage"), "TitleMenuBackgroundImage" %><br>
				Image <input type="file" name="UpTitleMenuBackgroundImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a title header image?  (This will appear before the title)
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a title header image? 	<% PrintRadio ImageExistsInt("TitleTopImage"), "TitleTopImage" %><br>
				Image <input type="file" name="UpTitleTopImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a title header rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a title header rollover image? 	<% PrintRadio ImageExistsInt("TitleTopRolloverImage"), "TitleTopRolloverImage" %><br>
				Image File<input type="file" name="UpTitleTopRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a title footer image?  (This will appear after the title)
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a title footer image? 	<% PrintRadio ImageExistsInt("TitleBottomImage"), "TitleBottomImage" %><br>
				Image File <input type="file" name="UpTitleBottomImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a title footer rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a title footer rollover image? 	<% PrintRadio ImageExistsInt("TitleBottomRolloverImage"), "TitleBottomRolloverImage" %><br>
				Image File <input type="file" name="UpTitleBottomRolloverImage">
			</td>
		</tr>
	</table>
	</span><br>


<%
	PrintTitle "Site Background", "background"
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Background Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a background image? 	<% PrintRadio ImageExistsInt("BackgroundImage"), "BackgroundImage" %><br>
				Image <input type="file" name="UpBackgroundImage">
			</td>
		</tr>
	</table>
	</span><br>


<%
	PrintTitle "Bullet - If you are creating lists, you may use a bullet to separate items.", "bullet"
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Bullet Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a bullet image? 	<% PrintRadio ImageExistsInt("BulletImage"), "BulletImage" %><br>
				Image <input type="file" name="UpBulletImage">
			</td>
		</tr>
	</table>
	</span><br>

<%
	PrintTitle "Home Button (on menu)", "home"
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Home Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("HomeImage"), "HomeImage" %><br>
				Image <input type="file" name="UpHomeImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Home Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("HomeRolloverImage"), "HomeRolloverImage" %><br>
				Image <input type="file" name="UpHomeRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Home Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("HomeHeaderImage"), "HomeHeaderImage" %><br>
				Image <input type="file" name="UpHomeHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Home Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("HomeFooterImage"), "HomeFooterImage" %><br>
				Image <input type="file" name="UpHomeFooterImage">
			</td>
		</tr>

	</table>
	</span><br>
<%
'----------------------Begin Code----------------------------

	if ParentSiteExists() then
		GetParent intParentID, strShortTitle, strSubDirectory
		PrintTitle strShortTitle & " Button (on menu)",  "parent" & intParentID
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=strShortTitle%> Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("Parent" & intParentID & "Image"), "Parent" & intParentID & "Image" %><br>
				Image <input type="file" name="UpParent<%=intParentID%>Image">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=strShortTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("Parent" & intParentID & "RolloverImage"), "Parent" & intParentID & "RolloverImage" %><br>
				Image <input type="file" name="UpParent<%=intParentID%>RolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=strShortTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("Parent" & intParentID & "HeaderImage"), "Parent" & intParentID & "HeaderImage" %><br>
				Image <input type="file" name="UpParent<%=intParentID%>HeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=strShortTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("Parent" & intParentID & "FooterImage"), "Parent" & intParentID & "FooterImage" %><br>
				Image <input type="file" name="UpParent<%=intParentID%>FooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
	end if

	Set rsPages = Server.CreateObject("ADODB.Recordset")

	if ChildSiteExists() then
		Query = "SELECT ID FROM Customers WHERE ParentID = " & CustomerID
		rsPages.CacheSize = 20
		rsPages.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
		if not rsPages.EOF then Set ChildID = rsPages("ID")
		do until rsPages.EOF
			GetChild ChildID, strShortTitle, strSubDirectory
			PrintTitle strShortTitle & " Button (on menu)",  "child" & ChildID
%>
				<tr>
					<td class="<% PrintTDMain %>" valign="middle" align="right">
						<%=strShortTitle%> Image
					</td>
					<td class="<% PrintTDMain %>" align="left">
						Use an image? 	<% PrintRadio ImageExistsInt("Child" & ChildID & "Image"), "Child" & ChildID & "Image" %><br>
						Image <input type="file" name="UpChild<%=ChildID%>Image">
					</td>
				</tr>
				<tr>
					<td class="<% PrintTDMain %>" valign="middle" align="right">
						<%=strShortTitle%> Button Rollover Image
					</td>
					<td class="<% PrintTDMain %>" align="left">
						Use a rollover image? 	<% PrintRadio ImageExistsInt("Child" & ChildID & "RolloverImage"), "Child" & ChildID & "RolloverImage" %><br>
						Image <input type="file" name="UpChild<%=ChildID%>RolloverImage">
					</td>
				</tr>
				<tr>
					<td class="<% PrintTDMain %>" valign="middle" align="right">
						<%=strShortTitle%> Button Header Image
					</td>
					<td class="<% PrintTDMain %>" align="left">
						Use a Header image? 	<% PrintRadio ImageExistsInt("Child" & ChildID & "HeaderImage"), "Child" & ChildID & "HeaderImage" %><br>
						Image <input type="file" name="UpChild<%=ChildID%>HeaderImage">
					</td>
				</tr>
				<tr>
					<td class="<% PrintTDMain %>" valign="middle" align="right">
						<%=strShortTitle%> Button Footer Image
					</td>
					<td class="<% PrintTDMain %>" align="left">
						Use a Footer image? 	<% PrintRadio ImageExistsInt("Child" & ChildID & "FooterImage"), "Child" & ChildID & "FooterImage" %><br>
						Image <input type="file" name="UpChild<%=ChildID%>FooterImage">
					</td>
				</tr>
			</table>
			</span><br>
<%
			rsPages.MoveNext
		loop
		rsPages.Close
	end if


	Query = "SELECT ID, Title FROM InfoPages WHERE Title <> 'Home Page' AND CustomerID = " & CustomerID & " AND ShowButton = 1"
	rsPages.CacheSize = 20
	rsPages.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	if not rsPages.EOF then
		Set ID = rsPages("ID")
		Set PageTitle = rsPages("Title")
		do until rsPages.EOF

			PrintTitle PageTitle & " Button (on menu)", "info" & ID

%>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					<%=PageTitle%> Image
				</td>
				<td class="<% PrintTDMain %>" align="left">
					Use an image? 	<% PrintRadio ImageExistsInt("InfoPage" & ID & "Image"), "InfoPage" & ID & "Image" %><br>
					Image <input type="file" name="UpInfoPage<%=ID%>Image">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					<%=PageTitle%> Button Rollover Image
				</td>
				<td class="<% PrintTDMain %>" align="left">
					Use a rollover image? 	<% PrintRadio ImageExistsInt("InfoPage" & ID & "RolloverImage"), "InfoPage" & ID & "RolloverImage" %><br>
					Image <input type="file" name="UpInfoPage<%=ID%>RolloverImage">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					<%=PageTitle%> Button Header Image
				</td>
				<td class="<% PrintTDMain %>" align="left">
					Use a Header image? 	<% PrintRadio ImageExistsInt("InfoPage" & ID & "HeaderImage"), "InfoPage" & ID & "HeaderImage" %><br>
					Image <input type="file" name="UpInfoPage<%=ID%>HeaderImage">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="middle" align="right">
					<%=PageTitle%> Button Footer Image
				</td>
				<td class="<% PrintTDMain %>" align="left">
					Use a Footer image? 	<% PrintRadio ImageExistsInt("InfoPage" & ID & "FooterImage"), "InfoPage" & ID & "FooterImage" %><br>
					Image <input type="file" name="UpInfoPage<%=ID%>FooterImage">
				</td>
			</tr>
			</table>
			</span><br>
<%
			rsPages.MoveNext
		loop
	end if
	rsPages.Close

	Set rsPages = Nothing


		if CBool( IncludeAnnouncements ) then
			PrintTitle AnnouncementsTitle & " Button (on menu)", "AnnouncementsTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=AnnouncementsTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("AnnouncementsImage"), "AnnouncementsImage" %><br>
				Image <input type="file" name="UpAnnouncementsImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=AnnouncementsTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("AnnouncementsRolloverImage"), "AnnouncementsRolloverImage" %><br>
				Image <input type="file" name="UpAnnouncementsRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=AnnouncementsTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("AnnouncementsHeaderImage"), "AnnouncementsHeaderImage" %><br>
				Image <input type="file" name="UpAnnouncementsHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=AnnouncementsTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("AnnouncementsFooterImage"), "AnnouncementsFooterImage" %><br>
				Image <input type="file" name="UpAnnouncementsFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeMeetings ) then
			PrintTitle MeetingsTitle & " Button (on menu)", "MeetingsTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MeetingsTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("MeetingsImage"), "MeetingsImage" %><br>
				Image <input type="file" name="UpMeetingsImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MeetingsTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("MeetingsRolloverImage"), "MeetingsRolloverImage" %><br>
				Image <input type="file" name="UpMeetingsRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MeetingsTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("MeetingsHeaderImage"), "MeetingsHeaderImage" %><br>
				Image <input type="file" name="UpMeetingsHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MeetingsTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("MeetingsFooterImage"), "MeetingsFooterImage" %><br>
				Image <input type="file" name="UpMeetingsFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeStories ) then
			PrintTitle StoriesTitle & " Button (on menu)", "StoriesTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StoriesTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("StoriesImage"), "StoriesImage" %><br>
				Image <input type="file" name="UpStoriesImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StoriesTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("StoriesRolloverImage"), "StoriesRolloverImage" %><br>
				Image <input type="file" name="UpStoriesRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StoriesTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("StoriesHeaderImage"), "StoriesHeaderImage" %><br>
				Image <input type="file" name="UpStoriesHeaderImage">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StoriesTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("StoriesFooterImage"), "StoriesFooterImage" %><br>
				Image <input type="file" name="UpStoriesFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeCalendar ) then
			PrintTitle CalendarTitle & " Button (on menu)", "CalendarTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=CalendarTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("CalendarImage"), "CalendarImage" %><br>
				Image <input type="file" name="UpCalendarImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=CalendarTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("CalendarRolloverImage"), "CalendarRolloverImage" %><br>
				Image <input type="file" name="UpCalendarRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=CalendarTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("CalendarHeaderImage"), "CalendarHeaderImage" %><br>
				Image <input type="file" name="UpCalendarHeaderImage">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=CalendarTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("CalendarFooterImage"), "CalendarFooterImage" %><br>
				Image <input type="file" name="UpCalendarFooterImage">
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<%=CalendarTitle%> Last and Next Month Buttons
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Last Month Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio LastMonthImage, "LastMonthImage" %><br>
				Image <input type="file" name="UpLastMonthImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Last Month Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio LastMonthRolloverImage, "LastMonthRolloverImage" %><br>
				Image <input type="file" name="UpLastMonthRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Next Month Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio NextMonthImage, "NextMonthImage" %><br>
				Image <input type="file" name="UpNextMonthImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Next Month Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio NextMonthRolloverImage, "NextMonthRolloverImage" %><br>
				Image <input type="file" name="UpNextMonthRolloverImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeLinks ) then
			PrintTitle LinksTitle & " Button (on menu)", "LinksTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=LinksTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("LinksImage"), "LinksImage" %><br>
				Image <input type="file" name="UpLinksImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=LinksTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("LinksRolloverImage"), "LinksRolloverImage" %><br>
				Image <input type="file" name="UpLinksRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=LinksTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("LinksHeaderImage"), "LinksHeaderImage" %><br>
				Image <input type="file" name="UpLinksHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=LinksTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("LinksFooterImage"), "LinksFooterImage" %><br>
				Image <input type="file" name="UpLinksFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeQuotes ) then
			PrintTitle QuotesTitle & " Button (on menu)", "QuotesTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=QuotesTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("QuotesImage"), "QuotesImage" %><br>
				Image <input type="file" name="UpQuotesImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=QuotesTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("QuotesRolloverImage"), "QuotesRolloverImage" %><br>
				Image <input type="file" name="UpQuotesRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=QuotesTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("QuotesHeaderImage"), "QuotesHeaderImage" %><br>
				Image <input type="file" name="UpQuotesHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=QuotesTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("QuotesFooterImage"), "QuotesFooterImage" %><br>
				Image <input type="file" name="UpQuotesFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeGuestbook ) then
			PrintTitle GuestbookTitle & " Button (on menu)", "GuestbookTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=GuestbookTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("GuestbookImage"), "GuestbookImage" %><br>
				Image <input type="file" name="UpGuestbookImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=GuestbookTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("GuestbookRolloverImage"), "GuestbookRolloverImage" %><br>
				Image <input type="file" name="UpGuestbookRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=GuestbookTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("GuestbookHeaderImage"), "GuestbookHeaderImage" %><br>
				Image <input type="file" name="UpGuestbookHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=GuestbookTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("GuestbookFooterImage"), "GuestbookFooterImage" %><br>
				Image <input type="file" name="UpGuestbookFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeForum ) then
			PrintTitle ForumTitle & " Button (on menu)", "ForumTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=ForumTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("ForumImage"), "ForumImage" %><br>
				Image <input type="file" name="UpForumImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=ForumTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("ForumRolloverImage"), "ForumRolloverImage" %><br>
				Image <input type="file" name="UpForumRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=ForumTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("ForumHeaderImage"), "ForumHeaderImage" %><br>
				Image <input type="file" name="UpForumHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=ForumTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("ForumFooterImage"), "ForumFooterImage" %><br>
				Image <input type="file" name="UpForumFooterImage">
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				<%=ForumTitle%> + and - Buttons
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" align="center" colspan=2>
				The + and - buttons are used to show all the replies to a message.  When a message has replies, there is an automatic '+' next to it, and once you click on the '+', a '-' shows.  If you want to replace the + and -, you can do it here.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				+ Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ForumPlusImage, "ForumPlusImage" %><br>
				Image <input type="file" name="UpForumPlusImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				- Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ForumMinusImage, "ForumMinusImage" %><br>
				Image <input type="file" name="UpForumMinusImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludePhotos ) then
			PrintTitle PhotosTitle & " Button (on menu)", "PhotosTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=PhotosTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("PhotosImage"), "PhotosImage" %><br>
				Image <input type="file" name="UpPhotosImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=PhotosTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("PhotosRolloverImage"), "PhotosRolloverImage" %><br>
				Image <input type="file" name="UpPhotosRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=PhotosTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("PhotosHeaderImage"), "PhotosHeaderImage" %><br>
				Image <input type="file" name="UpPhotosHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=PhotosTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("PhotosFooterImage"), "PhotosFooterImage" %><br>
				Image <input type="file" name="UpPhotosFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeVoting ) then
			PrintTitle VotingTitle & " Button (on menu)", "VotingTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=VotingTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("VotingImage"), "VotingImage" %><br>
				Image <input type="file" name="UpVotingImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=VotingTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("VotingRolloverImage"), "VotingRolloverImage" %><br>
				Image <input type="file" name="UpVotingRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=VotingTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("VotingHeaderImage"), "VotingHeaderImage" %><br>
				Image <input type="file" name="UpVotingHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=VotingTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("VotingFooterImage"), "VotingFooterImage" %><br>
				Image <input type="file" name="UpVotingFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeQuizzes ) then
			PrintTitle QuizzesTitle & " Button (on menu)", "QuizzesTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=QuizzesTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("QuizzesImage"), "QuizzesImage" %><br>
				Image <input type="file" name="UpQuizzesImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=QuizzesTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("QuizzesRolloverImage"), "QuizzesRolloverImage" %><br>
				Image <input type="file" name="UpQuizzesRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=QuizzesTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("QuizzesHeaderImage"), "QuizzesHeaderImage" %><br>
				Image <input type="file" name="UpQuizzesHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=QuizzesTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("QuizzesFooterImage"), "QuizzesFooterImage" %><br>
				Image <input type="file" name="UpQuizzesFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeMedia ) then
			PrintTitle MediaTitle & " Button (on menu)", "MediaTitle"
'------------------------End Code-----------------------------
%>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MediaTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("MediaImage"), "MediaImage" %><br>
				Image <input type="file" name="UpMediaImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MediaTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("MediaRolloverImage"), "MediaRolloverImage" %><br>
				Image <input type="file" name="UpMediaRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MediaTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("MediaHeaderImage"), "MediaHeaderImage" %><br>
				Image <input type="file" name="UpMediaHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MediaTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("MediaFooterImage"), "MediaFooterImage" %><br>
				Image <input type="file" name="UpMediaFooterImage">
			</td>
		</tr>
		</table>
		</span><br>

<%
'----------------------Begin Code----------------------------
		end if
		if CBool( IncludeNewsletter ) then
			PrintTitle NewsletterTitle & " Button (on menu)", "NewsletterTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=NewsletterTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("NewsletterImage"), "NewsletterImage" %><br>
				Image <input type="file" name="UpNewsletterImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=NewsletterTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("NewsletterRolloverImage"), "NewsletterRolloverImage" %><br>
				Image <input type="file" name="UpNewsletterRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=NewsletterTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("NewsletterHeaderImage"), "NewsletterHeaderImage" %><br>
				Image <input type="file" name="UpNewsletterHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=NewsletterTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("NewsletterFooterImage"), "NewsletterFooterImage" %><br>
				Image <input type="file" name="UpNewsletterFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if

		if CBool( AllowStore ) AND CBool( IncludeStore ) then
			PrintTitle StoreTitle & " Button (on menu)", "StoreTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StoreTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("StoreImage"), "StoreImage" %><br>
				Image <input type="file" name="UpStoreImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StoreTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("StoreRolloverImage"), "StoreRolloverImage" %><br>
				Image <input type="file" name="UpStoreRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StoreTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("StoreHeaderImage"), "StoreHeaderImage" %><br>
				Image <input type="file" name="UpStoreHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StoreTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("StoreFooterImage"), "StoreFooterImage" %><br>
				Image <input type="file" name="UpStoreFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if

		if CBool( IncludeStats ) then
			PrintTitle StatsTitle & " Button (on menu)", "StatsTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StatsTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("StatsImage"), "StatsImage" %><br>
				Image <input type="file" name="UpStatsImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StatsTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("StatsRolloverImage"), "StatsRolloverImage" %><br>
				Image <input type="file" name="UpStatsRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StatsTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("StatsHeaderImage"), "StatsHeaderImage" %><br>
				Image <input type="file" name="UpStatsHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=StatsTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("StatsFooterImage"), "StatsFooterImage" %><br>
				Image <input type="file" name="UpStatsFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
				PrintTitle MembersTitle & " Button (on menu)", "MembersTitle"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MembersTitle%> Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("MembersImage"), "MembersImage" %><br>
				Image <input type="file" name="UpMembersImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MembersTitle%> Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("MembersRolloverImage"), "MembersRolloverImage" %><br>
				Image <input type="file" name="UpMembersRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MembersTitle%> Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("MembersHeaderImage"), "MembersHeaderImage" %><br>
				Image <input type="file" name="UpMembersHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				<%=MembersTitle%> Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("MembersFooterImage"), "MembersFooterImage" %><br>
				Image <input type="file" name="UpMembersFooterImage">
			</td>
		</tr>
		</table>
		</span><br>



<%
		PrintTitle "Search Button (on menu)", "Search"
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Search Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("SearchImage"), "SearchImage" %><br>
				Image <input type="file" name="UpSearchImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Search Button Rollover Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a rollover image? 	<% PrintRadio ImageExistsInt("SearchRolloverImage"), "SearchRolloverImage" %><br>
				Image <input type="file" name="UpSearchRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Search Button Header Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Header image? 	<% PrintRadio ImageExistsInt("SearchHeaderImage"), "SearchHeaderImage" %><br>
				Image <input type="file" name="UpSearchHeaderImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Search Button Footer Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a Footer image? 	<% PrintRadio ImageExistsInt("SearchFooterImage"), "SearchFooterImage" %><br>
				Image <input type="file" name="UpSearchFooterImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
		if SectorHasButtons("Top") then
			PrintTitle "Top Menu Images", "TopMenu"
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="Left" colspan=2>
				The menu on the top of the screen can have a header image before all the buttons and a footer image after all the buttons.  
				It can have it's own background image, BUT that image is above until "Title Background Image".  Because the title and top menu 
				share the same space, they must have the same background.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="Right">
				Background image for menu
			</td>
			<td class="<% PrintTDMain %>" align="Left">
				Use a background image? 	<% PrintRadio ImageExistsInt("TopMenuBackgroundImage"), "TopMenuBackgroundImage" %><br>
				Image <input type="file" name="UpTopMenuBackgroundImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Button separator image for menu
			</td>
			<td class="<% PrintTDMain %>" align="Top">
				Use a background image? 	<% PrintRadio ImageExistsInt("TopMenuSeparatorImage"), "TopMenuSeparatorImage" %><br>
				Image <input type="file" name="UpTopMenuSeparatorImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="Right">
				Do you want to use a menu header image?  (This will appear before the buttons)
			</td>
			<td class="<% PrintTDMain %>" align="Left">
				Use a menu header image? 	<% PrintRadio ImageExistsInt("TopMenuTopImage"), "TopMenuTopImage" %><br>
				Image <input type="file" name="UpTopMenuTopImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="Right">
				Do you want to use a menu header rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="Left">
				Use a menu header rollover image? 	<% PrintRadio ImageExistsInt("TopMenuTopRolloverImage"), "TopMenuTopRolloverImage" %><br>
				Image File<input type="file" name="UpTopMenuTopRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="Right">
				Do you want to use a menu footer image?  (This will appear after the buttons)
			</td>
			<td class="<% PrintTDMain %>" align="Left">
				Use a menu footer image? 	<% PrintRadio ImageExistsInt("TopMenuBottomImage"), "TopMenuBottomImage" %><br>
				Image File <input type="file" name="UpTopMenuBottomImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="Right">
				Do you want to use a menu footer rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="Left">
				Use a menu footer rollover image? 	<% PrintRadio ImageExistsInt("TopMenuBottomRolloverImage"), "TopMenuBottomRolloverImage" %><br>
				Image File <input type="file" name="UpTopMenuBottomRolloverImage">
			</td>
		</tr>

		</table>
		</span><br>
<%
		end if

		if SectorHasButtons("Left") then
			PrintTitle "Left-Side Menu Images", "LeftMenu"
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="left" colspan=2>
				The menu on the left side of the screen can have a header image before all the buttons and a footer image after all the buttons.  
				It can also have it's own background image.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Background image for menu
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a background image? 	<% PrintRadio ImageExistsInt("LeftMenuBackgroundImage"), "LeftMenuBackgroundImage" %><br>
				Image <input type="file" name="UpLeftMenuBackgroundImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Button separator image for menu
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a background image? 	<% PrintRadio ImageExistsInt("LeftMenuSeparatorImage"), "LeftMenuSeparatorImage" %><br>
				Image <input type="file" name="UpLeftMenuSeparatorImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu header image?  (This will appear before the buttons)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu header image? 	<% PrintRadio ImageExistsInt("LeftMenuTopImage"), "LeftMenuTopImage" %><br>
				Image <input type="file" name="UpLeftMenuTopImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu header rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu header rollover image? 	<% PrintRadio ImageExistsInt("LeftMenuTopRolloverImage"), "LeftMenuTopRolloverImage" %><br>
				Image File<input type="file" name="UpLeftMenuTopRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu footer image?  (This will appear after the buttons)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu footer image? 	<% PrintRadio ImageExistsInt("LeftMenuBottomImage"), "LeftMenuBottomImage" %><br>
				Image File <input type="file" name="UpLeftMenuBottomImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu footer rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu footer rollover image? 	<% PrintRadio ImageExistsInt("LeftMenuBottomRolloverImage"), "LeftMenuBottomRolloverImage" %><br>
				Image File <input type="file" name="UpLeftMenuBottomRolloverImage">
			</td>
		</tr>



		</table>
		</span><br>
<%
		end if

		if SectorHasButtons("Right") then
			PrintTitle "Right-Side Menu Images", "RightMenu"
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="Right" colspan=2>
				The menu on the right side of the screen can have a header image before all the buttons and a footer image after all the buttons.  
				It can also have it's own background image.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Background image for menu
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a background image? 	<% PrintRadio ImageExistsInt("RightMenuBackgroundImage"), "RightMenuBackgroundImage" %><br>
				Image <input type="file" name="UpRightMenuBackgroundImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Button separator image for menu
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a background image? 	<% PrintRadio ImageExistsInt("RightMenuSeparatorImage"), "RightMenuSeparatorImage" %><br>
				Image <input type="file" name="UpRightMenuSeparatorImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu header image?  (This will appear before the buttons)
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a menu header image? 	<% PrintRadio ImageExistsInt("RightMenuTopImage"), "RightMenuTopImage" %><br>
				Image <input type="file" name="UpRightMenuTopImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu header rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a menu header rollover image? 	<% PrintRadio ImageExistsInt("RightMenuTopRolloverImage"), "RightMenuTopRolloverImage" %><br>
				Image File<input type="file" name="UpRightMenuTopRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu footer image?  (This will appear after the buttons)
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a menu footer image? 	<% PrintRadio ImageExistsInt("RightMenuBottomImage"), "RightMenuBottomImage" %><br>
				Image File <input type="file" name="UpRightMenuBottomImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu footer rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="Right">
				Use a menu footer rollover image? 	<% PrintRadio ImageExistsInt("RightMenuBottomRolloverImage"), "RightMenuBottomRolloverImage" %><br>
				Image File <input type="file" name="UpRightMenuBottomRolloverImage">
			</td>
		</tr>


		</table>
		</span><br>
<%
		end if
		PrintTitle "Main Body", "BodyBack"

%>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Main Body Background Image.  This will ONLY cover the space by the regular page body, and 
				does not cover menus, titles, or things outside the main body space.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a background image? 	<% PrintRadio ImageExistsInt("BodyMenuBackgroundImage"), "BodyMenuBackgroundImage" %><br>
				Image <input type="file" name="UpBodyMenuBackgroundImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		if CBool( TellNew ) then
			PrintTitle "'New' Button", "newbutton"
'------------------------End Code-----------------------------
%>
		<tr>
			<td class="<% PrintTDMain %>" align="center" colspan=2>
				If you chose so, items added in the last <%=NewDaysOld%> days have a 'New!' displayed in front of them.  You can replace that 'New!' with an image below.
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				'New' Button Image
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use an image? 	<% PrintRadio ImageExistsInt("NewImage"), "NewImage" %><br>
				Image <input type="file" name="UpNewImage">
			</td>
		</tr>
		</table>
		</span><br>
<%
'----------------------Begin Code----------------------------
		end if
'------------------------End Code-----------------------------
%>
		<% 	PrintTableHeader 50%>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
			<input type="submit" name="Submit" value="Update (click once)"  onClick="alert('If you are not uploading any files, just click okay and dont worry about this message.  You may be uploading many images, so please wait as long as it takes.  After pressing OK, your files will upload.  Please do not constantly press the Update button.')"></td>
			</td>
		</tr>
	</table>
	</form>
<%
'----------------------Begin Code----------------------------
	Set FileSystem = Nothing
end if

Function ImageExistsInt( strImage )
	strExt = ""
	if ImageExists( strImage, strExt) then
		ImageExistsInt = 1
	else
		ImageExistsInt = 0
	end if

End Function


Sub GetParent( intParentID, strShortTitle, strSubDirectory )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetParentSiteInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@ParentID", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@ShortTitle", adVarWChar, adParamOutput, 100 )
		.Parameters.Append .CreateParameter ("@SubDirectory", adVarWChar, adParamOutput, 100 )
		.Parameters("@CustomerID") = CustomerID
		.Execute , , adExecuteNoRecords
		intParentID = .Parameters("@ParentID")
		strShortTitle = .Parameters("@ShortTitle")
		strSubDirectory = .Parameters("@SubDirectory")
	End With
	Set cmdTemp = Nothing
End Sub


Sub GetChild( intChildID, strShortTitle, strSubDirectory )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetChildSiteInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@ShortTitle", adVarWChar, adParamOutput, 100 )
		.Parameters.Append .CreateParameter ("@SubDirectory", adVarWChar, adParamOutput, 100 )
		.Parameters("@CustomerID") = intChildID
		.Execute , , adExecuteNoRecords
		strShortTitle = .Parameters("@ShortTitle")
		strSubDirectory = .Parameters("@SubDirectory")
	End With
	Set cmdTemp = Nothing

	GetChildID = intChildID
End Sub

'------------------------End Code-----------------------------
%>