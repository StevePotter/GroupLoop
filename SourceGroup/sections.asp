<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<%
'
'-----------------------Begin Code----------------------------
'List their sections
if Request("ID") = "" then
	Query = "SELECT * FROM Sections WHERE CustomerID = " & CustomerID
	Set rsSection = Server.CreateObject("ADODB.Recordset")
	rsSection.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	do until rsSection.EOF
%>
		<a href="sections.asp?ID=<%=rsSection("ID")%>"><%=rsSection("Title")%></a><br>
<%
		rsSection.MoveNext
	loop
	rsSection.Close
	Set rsSection = Nothing


else

	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intSectionID = CInt(Request("ID"))

	Query = "SELECT * FROM Sections WHERE CustomerID = " & CustomerID & " AND ID = " & intSectionID
	Set rsSection = Server.CreateObject("ADODB.Recordset")
	rsSection.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if rsSection.EOF then Redirect("error.asp")

	Query = "SELECT * FROM CustomSectionItems WHERE SectionID = " & intSectionID
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	blLoggedAdmin = LoggedAdmin()
	blLoggedMember = LoggedMember()

	SectionTitle = rsSection("Title")
	Noun = rsSection("NounSingular")
	PluralNoun = rsSection("NounPlural")

	if rsSection("SectionViewSecurity") = "Members" and not blLoggedMember then Redirect("login.asp?Source=sections.asp&ID=" & intSectionID & "&Message=" & Server.URLEncode("Only members can view " & SectionTitle & " section.  If you are a member, please log in with your information below.  Otherwise, sorry, but you may not view this section."))
	if rsSection("SectionViewSecurity") = "Administrators" and not blLoggedAdmin then
		if blLoggedMember then
			Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but the " & SectionTitle & " section can only be viewed by an administrator."))
		else
			Redirect("login.asp?Source=sections.asp&ID=" & intSectionID & "Message=" & Server.URLEncode("Only <b>site administrators</b> can view this section.  If you are an administrator, please log in with your information below.  If you are a regular member or a non-member, sorry, but you may not view this section."))
		end if
	end if


	'The title and the add an item link
	AllowModify = CBool(blLoggedAdmin or (blLoggedMember and rsSection("ModifySecurity") = "Members"))

	ShowAddLink = CBool( rsSection("DisplayAddItems") = 1 and AllowModify )
	if rsSection("DisplayTitleList") = 1 or ShowAddLink then
		Response.Write "<p align='" & HeadingAlignment & "'>"
		if rsSection("DisplayTitleList") = 1 then Response.Write "<span class=Heading>" & SectionTitle & "</span><br>"
		if ShowAddLink then Response.Write "<span class=LinkText><a href=members_sectionitems_add.asp?ID=" & intSectionID & ">Add " & PrintAn(Noun) & " " & PrintFirstCap(Noun) & "</a></span>"
		Response.Write "</p>"
	end if



	'This toggles the display buttons
	blShowModify = False

	if AllowModify then
		blShowModify = True
		if Request("Modify") = "Yes" then
			Session("ModifySectionItems") = "Yes"
		elseif Request("Modify") = "No" then
			Session("ModifySectionItems") = "No"
		elseif Session("ModifySectionItems") = "" then	'Their first time, set the modify to yes
			Session("ModifySectionItems") = "Yes"
		end if
	end if

	strImagePath = GetPath("posts")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	ItemNumber = 0	'This will be set by the PrintPagesHeader sub

	DisplaySearch = CBool(rsSection("DisplaySearch"))
	DisplayDaysOld = CBool(rsSection("DisplayDaysOld"))
	InfoText = rsSection("InfoText")
	ListType = rsSection("ListType")
	DisplayDate = CBool(rsSection("DisplayDateList"))
	DisplayAuthor = CBool(rsSection("DisplayAuthorList"))
	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = (CBool(rsSection("DisplayPrivacyList")) and CBool(rsSection("IncludePrivacy"))) and not cBool(SiteMembersOnly)

	RateItems = CBool(rsSection("RateItems"))
	ReviewItems = CBool(rsSection("ReviewItems"))



	if DisplaySearch or DisplayDaysOld then
	%>
		<form METHOD="POST" ACTION="sections.asp?ID=<%=intSectionID%>">
	<%	if DisplayDaysOld then	%>
		View <%=PluralNoun%> In The Last <% PrintDaysOld %>
		<br>
	<%		if DisplaySearch then Response.Write "Or "
		end if
		if DisplaySearch then	%>
		Search For <input type="text" name="Keywords" size="25">
		<input type="submit" name="Submit" value="Go"><br>
	<%	end if	%>	
		</form>
	<%
	end if

	'They did a search
	if intSearchID <> "" then
%>


<%
	'They are just cycling through the items.  No searching.
	else
		if not rsPage.EOF then
	%>
			<form METHOD="POST" ACTION="sections.asp?ID=<%=intSectionID%>">
			<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
	<%
			Set ID = rsPage("ID")
			Set ItemDate = rsPage("Date")
			Set MemberID = rsPage("MemberID")
			Set TotalRating = rsPage("TotalRating")
			Set TimesRated = rsPage("TimesRated")
			Set IsPrivate = rsPage("Privacy")

			PrintPagesHeader
			PrintListHeader
			for j = 1 to rsPage.PageSize
				if not rsPage.EOF then
					PrintTableData
					rsPage.MoveNext
				end if
			next

			PrintListFooter
		else
			if Request("DaysOld") <> "" then
	'------------------------End Code-----------------------------
	%>
				<p>Sorry, but there have been no <%=PluralNoun%> added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
	<%
	'-----------------------Begin Code----------------------------
			else
	'------------------------End Code-----------------------------
	%>
				<p>Sorry, but there are no <%=PluralNoun%> at the moment.</p>
	<%
	'-----------------------Begin Code----------------------------
			end if
		end if
		rsPage.Close
		set rsPage = Nothing
	end if

	rsSection.Close
	Set rsSection = Nothing

	Set FileSystem = Nothing
end if


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintListHeader
	if ListType = "Table" then
		PrintTableHeader 0
%>		
	<tr>
		<% if DisplayDate then %>
		<td class="TDHeader">Date</td>
		<% end if %>	
		<% if DisplayAuthor then %>
		<td class="TDHeader">Author</td>
		<% end if %>	
<%
		for i = 1 to 10
			if rsSection("FieldName" & i) <> "" and rsSection("DisplayListField" & i) = 1 then
%>
			<td class="TDHeader" align=center><%=rsSection("FieldName" & i)%> </td>
<%
			end if
		next
%>
		<% if RateItems then %>
			<td class="TDHeader" align=center>Rating</td>
		<% elseif not RateItems  and ReviewItems then %>
			<td class="TDHeader" align=center>Review</td>
		<% end if %>	
		<% if DisplayPrivacy then %>
		<td class="TDHeader">Public?</td>
		<% end if %>	
<%
		if blShowModify and Session("ModifySectionItems") = "Yes" then
%>
		<td class="TDHeader">&nbsp;</td>
<%
		end if
%>
	</tr>
<%
	elseif ListType = "Bulleted" and not blBulletImg then
		Response.Write "<ul>"
	else
		Response.Write "<p>"
	end if
End Sub

'-------------------------------------------------------------
'This prints the closing for the list
'-------------------------------------------------------------
Sub PrintListFooter
	if ListType = "Table" then
		Response.Write("</table>")

	elseif ListType = "Bulleted" and not blBulletImg then
		Response.Write "</ul>"
	else
		Response.Write "</p>"
	end if

	'They can modify items, and aren't already showing the edit/delete buttons
	if blShowModify and Session("ModifySectionItems") = "No" then
		if blLoggedAdmin then
			strYour = ""
		else
			strYour = "Your"
		end if
		Response.Write "<br><br><p align=left><a href='sections.asp?ID=" & intSectionID & "&Modify=Yes'>Show Edit/Delete Buttons For " & strYour & " " & PluralNoun & "</a></p>"
	elseif blShowModify and Session("ModifySectionItems") = "Yes" then
		Response.Write "<br><br><p align=left><a href='sections.asp?ID=" & intSectionID & "&Modify=No'>Hide Edit/Delete Buttons</a></p>"
	end if

'	'Give them the link to change the section's properties
'	if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
'		Response.Write "<div align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Announcements&Source=announcements.asp'>Change Section Options</a></div>"
'	end if

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
<%
		for lpVar = 1 to 10
			if rsSection("FieldName" & lpVar) <> "" and rsSection("DisplayListField" & lpVar) = 1 then
%>
			<td class="<% PrintTDMain %>" align=center>
<%
				LinkToItem = CBool(rsSection("LinkToItemField"&lpVar))

				if LinkToItem then
					strHeader = "<a href='sectionitems_view.asp?ID=" & rsPage("ID")& "'>"
					strFooter = "</a>"
				else
					strHeader = ""
					strFooter = ""
				end if

				if rsSection("FieldType" & lpVar) = "Link" then
					PrintLink rsPage("Field" & lpVari)
				elseif rsSection("FieldType" & lpVar) = "Photo" and ImageExists("CustomSectionItems"&rsPage("ID")&"-"&lpVar, strExt) then

					if not LinkToItem then strHeader = "<a href='section_photo_view.asp?ID=" & rsPage("ID") & "&FieldNum=" & lpVar & "'>"
					'We have a thumbnail
					if ImageExists("CustomSectionItems"&rsPage("ID")&"-"&lpVar&"t", strThumbExt) then
%>
						<%=strHeader%><img src="posts/CustomSectionItems<%=rsPage("ID")%>-<%=lpVar%>t.<%=strThumbExt%>" border="0"></a>
<%
					else
%>
						<%=strHeader%>View Photo</a>
<%
					end if
				elseif rsSection("FieldType" & lpVar) = "TextBox" then
					Response.Write strHeader & rsPage("FieldLongText" & lpVar) & strFooter
				else
					Response.Write strHeader & rsPage("Field" & lpVar) & strFooter
				end if
				Response.Write "</td>"
			end if
		next
%>

<%		if RateItems and not ReviewItems then
%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate ")%></a></font></td>
<%		elseif not RateItems and ReviewItems then
			if ReviewsExist( "CustomSectionItems", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>"><%=PrintTDLink( "Read/Add Review ")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>"><%=PrintTDLink( "Add Review ")%></a></font></td>
<%			end if
		elseif RateItems and ReviewItems then
			if ReviewsExist( "CustomSectionItems", ID ) then
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate and Read/Add Review ")%></a></font></td>
<%			else
%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate/Add Review ")%></a></font></td>
<%			end if
		end if%>
		<% if DisplayPrivacy  then %>
		<td class="<% PrintTDMain %>"><%=IsPrivate%></td>
		<% end if %>
		<% if  blShowModify and Session("ModifySectionItems") = "Yes" and (blLoggedAdmin or (blLoggedMember and Session("MemberID") = MemberID)) then %>
		<td class="<% PrintTDMain %>">
			<a href="members_sectionitems_modify.asp?Submit=Edit&ID=<%=ID%>"><%=PrintTDLink("Edit")%></a>&nbsp;
			<a href="javascript:DeleteBox('If you delete this <%=Noun%> (<%=Subject%>), there is no way to get it back.  Are you sure?', 'members_sectionitems_modify.asp?Submit=Delete&ID=<%=ID%>')"><%=PrintTDLink("Delete")%></a>
			<%if ReviewsExist( "CustomSectionItems", ID ) AND blLoggedAdmin then%>
				<a href="javascript:Redirect('admin_reviews_modify.asp?Source=sections.asp?ID=<%=intSectionID%>&TargetTable=CustomSectionItems&TargetID=<%=ID%>')">Modify Reviews</a>
			<%end if%>
		</td>
		<% elseif blShowModify and Session("ModifySectionItems") = "Yes" then %>
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
		<a href="sectionitems_view.asp?ID=<%=ID%>"><%=Subject%></a>&nbsp;&nbsp;&nbsp;&nbsp;
		<% if DisplayDate then %>
		<% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%>&nbsp;&nbsp;
		<% end if %>	
		<% if DisplayAuthor then %>
		By: <%=GetNickNameLink(MemberID)%>&nbsp;&nbsp;
		<% end if %>	
		<% 'if DisplayPrivacy and IsPrivate = 1 then Response.Write "Private&nbsp;&nbsp;"
		if RateItems and not ReviewItems then
%>			<%=GetRating( TotalRating, TimesRated )%> 
			<font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>">Rate</a></font>&nbsp;&nbsp;
<%		elseif not RateItems and ReviewItems then
			if ReviewsExist( "CustomSectionItems", ID ) then
%>				<font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>">Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>">Add Review</a></font>&nbsp;&nbsp;
<%			end if
		elseif RateItems and ReviewItems then
			if ReviewsExist( "CustomSectionItems", ID ) then
%>				<%=GetRating( TotalRating, TimesRated )%> 
					<font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>">Rate and Read/Add Review</a></font>&nbsp;&nbsp;
<%			else
%>				<%=GetRating( TotalRating, TimesRated )%> 
				<font size="-2"><a href="sectionitems_view.asp?ID=<%=ID%>">Rate/Add Review</a></font>&nbsp;&nbsp;
<%			end if
		end if
		Response.Write strFooter
	end if
End Sub
'------------------------End Code-----------------------------

Function PrintAn(strFollowWord)
	if IsNull( strFollowWord) then
		PrintAn = "a"
	elseif strFollowWord = "" then
		PrintAn = "a"
	else
		testchar = Lcase( Left( strFollowWord,1 ) )

		if Instr( testchar, "aeiou" ) then
			PrintAn = "an"
		else
			PrintAn = "a"
		end if
	end if

End Function

Sub PrintLink( strPassedLink )
	strLabel = strPassedLink
	strLink = strPassedLink
	if InStr( strLink, "@" ) then 'e-mail address
		if not InStr( strLink, "mailto:" ) then strLink = "mailto:" & strLink
	else
		if not InStr( strLink, "http://" ) then strLink = "http://" & strLink
	end if

	Response.Write "<a href=" & Chr(34) & strLink & Chr(34) & ">" & strLabel & "</a>"
End Sub

Function PrintFirstCap( strWord )
	PrintFirstCap = UCase( Left(strWord, 1) ) & Right(strWord, Len(strWord)-1)

End Function
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->