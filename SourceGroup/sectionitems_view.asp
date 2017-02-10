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
	intItemID = CInt(Request("ID"))

	Query = "SELECT * FROM CustomSectionItems WHERE ID = " & intItemID
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	if rsPage.EOF then Redirect("error.asp")

	intSectionID = rsPage("SectionID")

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	Query = "SELECT * FROM Sections WHERE CustomerID = " & CustomerID & " AND ID = " & intSectionID
	Set rsSection = Server.CreateObject("ADODB.Recordset")
	rsSection.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if rsSection.EOF then Redirect("error.asp")


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

	if rsPage("Privacy") = "Members" and not blLoggedMember then Redirect("login.asp?Source=sectionitems_view.asp&ID=" & intItemID & "&Message=" & Server.URLEncode("Only members can view this " & Noun & ".  If you are a member, please log in with your information below.  Otherwise, sorry, but you may not view this " & Noun & "."))
	if rsPage("Privacy") = "Administrators" and not blLoggedAdmin then
		if blLoggedMember then
			Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this " & Noun & " can only be viewed by an administrator."))
		else
			Redirect("login.asp?Source=sectionitems_view.asp&ID=" & intItemID & "Message=" & Server.URLEncode("Only <b>site administrators</b> can view this " & Noun & ".  If you are an administrator, please log in with your information below.  If you are a regular member or a non-member, sorry, but you may not view this " & Noun & "."))
		end if
	end if


	'The title and the add an item link
	AllowModify = CBool(blLoggedAdmin or (blLoggedMember and rsPage("MemberID") = Session("MemberID")))

	if rsSection("DisplayTitleItem") = 1 then
		Response.Write "<p align='" & HeadingAlignment & "'>"
		Response.Write "<span class=Heading>" & SectionTitle & "</span><br>"
		Response.Write "</p>"
	end if

	strImagePath = GetPath("posts")
	blBulletImg = ImageExists("BulletImage", strBulletExt)
	ItemNumber = 0	'This will be set by the PrintPagesHeader sub

	DisplayDate = CBool(rsSection("DisplayDateItem"))
	DisplayAuthor = CBool(rsSection("DisplayAuthorItem"))

	RateItems = CBool(rsSection("RateItems"))
	ReviewItems = CBool(rsSection("ReviewItems"))


	if Request("Rating") <> "" and RateItems then
	%>
		<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
		<a href="javascript:history.go(-1)">Back To <%=PrintFirstCap(Noun)%></a></p>
	<%
		AddRating intItemID, "CustomSectionItems"
	else
		IncrementHits intItemID, "CustomSectionItems"
	%>
		<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
	<%
		if AllowModify then
	%>
			<table align=<%=HeadingAlignment%>>
			<tr>
			<td align=right width="50%" class="LinkText"><a href="members_sectionitems_modify.asp?Submit=Edit&ID=<%=intItemID%>">Edit</a>&nbsp;&nbsp;</td>
			<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
			<a href="javascript:DeleteBox('If you delete this <%=Noun%>, there is no way to get it back.  Are you sure?', 'members_sectionitems_modify.asp?Submit=Delete&ID=<%=intItemID%>')">Delete</a>
			</td>
			</tr>
			</table>
	<%
		end if
		Response.Write "</p>"
	end if

	if DisplayDate or DisplayAuthor then
		PrintTableHeader 100
		if DisplayDate or DisplayAuthor then %>

		<tr>
			<td colspan="2" class="<% PrintTDMain %>">
			<table width=100% cellspacing=0 cellpadding=0>
			<tr>
	
			<% if DisplayAuthor then %>
			<td class="<% PrintTDMain %>" align="left">Author: <%=PrintTDLink(GetNickNameLink(rsPage("MemberID")))%></td>
			<% end if %>	
			<% if DisplayDate then
				strAlign = "left"
				if DisplayAuthor then strAlign = "right"
			%>
			<td class="<% PrintTDMainSwitch %>" align="<%=strAlign%>">Date Written: <%=FormatDateTime(rsPage("Date"), 2)%></td>
			<% end if %>		
			</tr>	
			</table>
			</td>
		</tr>
<%
		end if
		if DisplaySubject then
%>	
		<tr>
			<td class="<% PrintTDMainSwitch %>" align="left" colspan="2">Subject: <%=rsPage("Subject")%></td>
		</tr>
<%
		end if
		Response.Write "</table>"
	end if

	ListNum = 1


	for lpVar = 1 to 10
		if rsSection("FieldName" & lpVar) <> "" and rsSection("DisplayItemField" & lpVar) = 1 then
			strHeader = ""
			strFooter = ""

			strAlign = rsSection("ItemAlignmentField" & lpVar)
			strFormat = rsSection("ItemFormatField" & lpVar)

			if strFormat = "plain" then
				strHeader = "<div align=" & strAlign & ">"
				strFooter = "</div><br>"
			elseif strFormat = "bullet" then
				strHeader = "<div align=" & strAlign & "><ul><li>"
				strFooter = "</ul></li></div>"
			elseif strFormat = "table" then
				strHeader = GetTableHeader(0) & "<tr><td class=" & GetTDMain & " align=" & strAlign & " >"
				strFooter = "</td></tr></table>"
				ChangeTDMain
			elseif strFormat = "paragraph" then
				strHeader = "<p align=" & strAlign & ">"
				strFooter = "</p>"
			elseif strFormat = "numbered" then
				strHeader = "<div align=" & strAlign & ">" & ListNum & ".&nbsp;"
				strFooter = "</div><br>"
				ListNum = ListNum + 1
			end if

			if rsSection("FieldType" & lpVar) = "Link" then
				PrintLink rsPage("Field" & lpVari)
			elseif rsSection("FieldType" & lpVar) = "Photo" and ImageExists("CustomSectionItems"&rsPage("ID")&"-"&lpVar, strExt) then

				if not LinkToItem then strHeader = strHeader & "<a href='section_photo_view.asp?ID=" & rsPage("ID") & "&FieldNum=" & lpVar & "'>"
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
		end if
	next


'-----------------------Begin Code----------------------------
	if RateItems then
		PrintRatingPulldown intItemID, "", "CustomSectionItems", "sectionitems_view.asp?ID=" & intItemID, Noun
	end if
	if ReviewItems then
%>
		<a href="review.asp?Source=sectionitems_view.asp?ID=<%=intItemID%>&TargetID=<%=intItemID%>&Table=CustomSectionItems">Add A Review</a><br>
<%
		if ReviewsExist( "CustomSectionItems", intItemID ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=announcements_read.asp?ID=<%=intItemID%>&TargetTable=Announcements&TargetID=<%=intItemID%>">Modify Reviews</a><br>
<%
			end if
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "sectionitems_view.asp?ID=" & intItemID, "CustomSectionItems", intItemID
			Set rsPage = Nothing
		end if
	end if
end if


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