<%
'-----------------------Begin Code----------------------------
strTable = "Announcements"
strNoun = "Announcement"
strPluralNoun = "Announcements"
strListSource = "announcements.asp"
strListSourceName = "Announcements"
strAddSource = "members_announcements_add.asp"
strModSource = "members_announcements_modify.asp"
strViewSource = "announcements_read.asp?ID="

strOrderBy = "Date DESC"
strFields = "ID, Date, MemberID, Subject, TotalRating, TimesRated, Private"

blLoggedAdmin = LoggedAdmin()
blLoggedMember = LoggedMember()

intItemsPerRow = 1

'strViewAction = "Read"


'Make sure they can enter this section
CheckSection IncludeAnnouncements, SectionViewAnnouncements, strListSource

'This toggles the display buttons
blShowModify = DisplayModifyCol()

PrintTitle strTitle, AnnouncementsMembers, strAddSource, strNoun


Public ListType, DisplayDate, DisplayAuthor, DisplayPrivacy, blBulletImg, ItemNumber

strImagePath = GetPath("images")
blBulletImg = ImageExists("BulletImage", strBulletExt)
ItemNumber = 0	'This will be set by the PrintPagesHeader sub

Query = "SELECT IncludePrivacy" & strTable & ", DisplaySearch" & strTable & ", DisplayDaysOld" & strTable & ", InfoText" & strTable & ", ListType" & strTable & ", DisplayDateList" & strTable & ", DisplayAuthorList" & strTable & ", DisplayPrivacyList" & strTable & "  FROM Look WHERE CustomerID = " & CustomerID
Set rsList = Server.CreateObject("ADODB.Recordset")
rsList.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplaySearch = CBool(rsList("DisplaySearch" & strTable ))
	DisplayDaysOld = CBool(rsList("DisplayDaysOld" & strTable ))
	InfoText = rsList("InfoText" & strTable )
	ListType = rsList("ListType" & strTable )
	DisplayDate = CBool(rsList("DisplayDateList" & strTable ))
	DisplayAuthor = CBool(rsList("DisplayAuthorList" & strTable ))
	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = (CBool(rsList("DisplayPrivacyList" & strTable )) and CBool(rsList("IncludePrivacy" & strTable ))) and not cBool(SiteMembersOnly)
rsList.Close



PrintSearch DisplaySearch, DisplayDaysOld, strListSource, strPluralNoun


Query = "SELECT ID, Date, MemberID, Subject, Body FROM Announcements WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"



GoList




'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(Subject & Body & ItemDate & GetNickName(MemberID) )
End Function

'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintListHeader
	if ListType = "Table" then
		PrintTableHeader 0
		Response.Write "<tr>"
		if DisplayDate then Response.Write PrintHeaderCol( "Date", "" )
		if DisplayAuthor then Response.Write PrintHeaderCol( "Author", "" )
		Response.Write PrintHeaderCol( "Subject", "" )
		if RateAnnouncements = 1 then Response.Write PrintHeaderCol( "Rating", "" )
		if ReviewAnnouncements = 1 then Response.Write PrintHeaderCol( "Review", "" )
		if blShowModify then Response.Write PrintHeaderCol( "", "" )
		Response.Write "</tr>"
	elseif ListType = "Bulleted" and not blBulletImg then
		Response.Write "<ul>"
	else
		Response.Write "<p>"
	end if
End Sub


%>




		
<%
'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
	if ListType = "Table" then
%>
	<tr>
		<% if DisplayDate then %>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<% end if %>	
		<% if DisplayAuthor then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
		<% end if %>	
		<td class="<% PrintTDMain %>"><a href="announcements_read.asp?ID=<%=ID%>"><%=PrintTDLink( Subject )%></a></td>
<%		if RateAnnouncements = 1 and ReviewAnnouncements = 0 then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="announcements_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate ")%></a></font></td>
<%		elseif RateAnnouncements = 0 and ReviewAnnouncements = 1 then
			if ReviewsExist( "Announcements", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="announcements_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Read/Add Review ")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="announcements_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Add Review ")%></a></font></td>
<%			end if
		elseif RateAnnouncements = 1 and ReviewAnnouncements = 1 then
			if ReviewsExist( "Announcements", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="announcements_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate and Read/Add Review ")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="announcements_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate/Add Review ")%></a></font></td>
<%			end if
		end if%>
		<% if DisplayPrivacy  then %>
		<td class="<% PrintTDMain %>"><%=PrintPublic(IsPrivate)%></td>
		<% end if %>
		<% if  blShowModify and (blLoggedAdmin or (blLoggedMember and Session("MemberID") = MemberID)) then %>
		<td class="<% PrintTDMain %>">
			<a href="members_announcements_modify.asp?Submit=Edit&ID=<%=ID%>"><%=PrintTDLink("Edit")%></a>&nbsp;
			<a href="javascript:DeleteBox('If you delete this announcement (<%=Subject%>), there is no way to get it back.  Are you sure?', 'members_announcements_modify.asp?Submit=Delete&ID=<%=ID%>')"><%=PrintTDLink("Delete")%></a>&nbsp;
			<%if ReviewsExist( "Announcements", ID ) AND blLoggedAdmin then%>
				<a href="javascript:Redirect('admin_reviews_modify.asp?Source=announcements.asp&TargetTable=Announcements&TargetID=<%=ID%>')"><%=PrintTDLink("Modify Reviews")%></a>
			<%end if%>
		</td>
		<% elseif blShowModify then %>
		<td class="<% PrintTDMain %>">&nbsp;</td>
		<% end if %>			
	</tr>
<%
		ChangeTDMain
	else
		strHeader = ""
		strFooter = "<br>"
		if ListType = "Bulleted" then
			if blBulletImg then
				strHeader = "<img src='images/BulletImage." & strBulletExt & "'>"
			else
				strHeader = "<li>"
				strFooter = "</li>"
			end if
		elseif ListType = "Numbered" then
				ItemNumber = ItemNumber + 1
				strHeader = ItemNumber & ".&nbsp;"
		end if
%>
		<%=strHeader%>
		<a href="announcements_read.asp?ID=<%=ID%>"><%=Subject%></a>&nbsp;&nbsp;&nbsp;&nbsp;
		<% if DisplayDate then %>
		<% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayAuthor then %>
		By: <%=GetNickNameLink(MemberID)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayPrivacy and IsPrivate = 1 then Response.Write "Private&nbsp;&nbsp;"
		if RateAnnouncements = 1 and ReviewAnnouncements = 0 then
%>			<%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="announcements_read.asp?ID=<%=ID%>">Rate</a></font>&nbsp;&nbsp;
<%		elseif RateAnnouncements = 0 and ReviewAnnouncements = 1 then
			if ReviewsExist( "Announcements", ID ) then
%>				<font size="-2"><a href="announcements_read.asp?ID=<%=ID%>">Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<font size="-2"><a href="announcements_read.asp?ID=<%=ID%>">Add Review</a></font>&nbsp;&nbsp;
<%			end if
		elseif RateAnnouncements = 1 and ReviewAnnouncements = 1 then
			if ReviewsExist( "Announcements", ID ) then
%>				<%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="announcements_read.asp?ID=<%=ID%>">Rate and Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="announcements_read.asp?ID=<%=ID%>">Rate/Add Review</a></font>&nbsp;&nbsp;
<%			end if
		end if


		if  blShowModify and (blLoggedAdmin or (blLoggedMember and Session("MemberID") = MemberID)) then %>
		<td class="<% PrintTDMain %>">
			<a href="members_announcements_modify.asp?Submit=Edit&ID=<%=ID%>"><%=PrintTDLink("Edit")%></a>&nbsp;
			<a href="javascript:DeleteBox('If you delete this announcement (<%=Subject%>), there is no way to get it back.  Are you sure?', 'members_announcements_modify.asp?Submit=Delete&ID=<%=ID%>')"><%=PrintTDLink("Delete")%></a>&nbsp;
			<%if ReviewsExist( "Announcements", ID ) AND blLoggedAdmin then%>
				<a href="javascript:Redirect('admin_reviews_modify.asp?Source=announcements.asp&TargetTable=Announcements&TargetID=<%=ID%>')"><%=PrintTDLink("Modify Reviews")%></a>
			<%end if%>
		</td>
<%
		end if



		Response.Write strFooter
	end if
End Sub

'------------------------End Code-----------------------------
%>

<div id="divTooltip"></div>
