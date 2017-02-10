<!-- #include file="forum_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeForum ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>"><%=ForumTitle%></p>

<%
'-----------------------Begin Code----------------------------
intCategoryID = Request("ID")
if intCategoryID <> "" then intCategoryID = CInt(intCategoryID)


'Get the page number
if Request("PageNum") <> "" then
	intPage = CInt(Request("PageNum"))
else
	intPage = 1
end if

'Get the searchID from the last page.  May be blank.
intSearchID = Request("SearchID")
if intSearchID <> "" then intSearchID = CInt(intSearchID)

intRateForum = RateForum

Public ListTypeForum, DisplayDate, DisplayAuthor, DisplayPrivacy, blBulletImg, ItemNumber
	strImagePath = GetPath("images")
	blBulletImg = ImageExists("BulletImage", strBulletExt)

Query = "SELECT DisplaySearchForum, InfoTextForum, DisplayDateListForum, DisplayAuthorListForum, DisplayPrivacyListForum, IncludePrivacyForum  FROM Look WHERE CustomerID = " & CustomerID
Set rsList = Server.CreateObject("ADODB.Recordset")
rsList.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplaySearch = CBool(rsList("DisplaySearchForum"))
	InfoText = rsList("InfoTextForum")
	DisplayDate = CBool(rsList("DisplayDateListForum"))
	DisplayAuthor = CBool(rsList("DisplayAuthorListForum"))
	DisplayPrivacy = (CBool(rsList("DisplayPrivacyListForum")) and CBool(rsList("IncludePrivacyForum"))) and not cBool(SiteMembersOnly)
rsList.Close


'Start them off in the first category if none is specified
if intCategoryID = "" then

	'They entered text to search for, so we are going to get matches and put them into the SectionSearch
	if Request("Keywords") <> "" AND intSearchID = "" then
		Query = "SELECT ID, Date, Email, Author, Subject, Body, MemberID FROM ForumMessages WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set Email = rsList("Email")
			Set Author = rsList("Author")
			Set Subject = rsList("Subject")
			Set Body = rsList("Body")
			Set MemberID = rsList("MemberID")
		intSearchID = SingleSearch()
		rsList.Close
		set rsList = Nothing
	end if

	if intSearchID = "" then
		if DisplaySearch then
%>
		<form METHOD="POST" ACTION="forum.asp">
			Search For <input type="text" name="Keywords" size="15">
			<input type="submit" name="Submit" value="Go"><br>
		</form>
<%
		end if
		if InfoText <> " " and InfoText <> "" then Response.Write "<p>" & InfoText & "</p>"

		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then Response.Write "<p><a href='admin_forumcategories_modify.asp'>Modify Categories</a></p>"

		Set rsPage = Server.CreateObject("ADODB.Recordset")
		PrintCategoryMenu "forum.asp"
		Set rsPage = Nothing

		'Give them the link to change the section's properties
		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Forum&Source=forum.asp'>Change Section Options</a></p>"
		end if

	else
		'Their search came up empty
		if intSearchID = 0 then
			if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
				<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
				Try again, or <a href="forum.asp">click here</a> to go back to the topic list.</p>
<%
'-----------------------Begin Code----------------------------
			else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="forum.asp">click here</a> to go back to the topic list.</p>
<%
'-----------------------Begin Code----------------------------
			end if
		else
			'They have search results, so lets list their results
			Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			rsPage.CacheSize = PageSize
			Set TargetID = rsPage("TargetID")
%>
			<p><a href="forum.asp">Click here</a> to go back to the topic list.</p>

			<form METHOD="POST" ACTION="forum.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
			PrintPagesHeader
			PrintTableHeader 0
%>
			<tr>
				<% if DisplayDate then %>
				<td class="TDHeader">Date</td>
				<% end if %>
				<td class="TDHeader">Author</td>
				<td class="TDHeader">Subject</td>
				<td class="TDHeader">Topic</td>
				<% if intRateForum = 1 then %>
					<td class="TDHeader">Rating</td>
				<% end if %>
				<% if DisplayPrivacy then %>

				<td class="TDHeader">Public?</td>
				<% end if %>
		</tr>
<%
			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Date, Email, Author, Subject, CategoryID, MemberID, Private, TimesRated, TotalRating FROM ForumMessages WHERE CustomerID = " & CustomerID
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set Email = rsList("Email")
			Set Author = rsList("Author")
			Set Subject = rsList("Subject")
			Set CategoryID = rsList("CategoryID")
			Set MemberID = rsList("MemberID")
			Set TotalRating = rsList("TotalRating")
			Set TimesRated = rsList("TimesRated")
			Set IsPrivate = rsList("Private")

			for p = 1 to rsPage.PageSize
				if not rsPage.EOF then
					rsList.Filter = "ID = " & TargetID

					if MemberID > 0 then
						strAuthor = GetNickNameLink( MemberID )
					elseif InStr( Email, "@" ) then
						strAuthor = "<a href='mailto:" & Email & "'>" & PrintTDLink(Author) & "</a>"
					else
						strAuthor = Author
					end if
%>
					<tr>
						<% if DisplayDate then %>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
						<% end if %>
						<td class="<% PrintTDMain %>"><%=strAuthor%></td>
						<td class="<% PrintTDMain %>"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( Subject )%></a></td>
						<td class="<% PrintTDMain %>"><%=GetForumCategory(CategoryID)%></td>
				<%		if intRateForum = 1 and intReviewForum = 0 then
				%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
							<font size="-2"><a href="forum_read.asp?ID=<%=ID%>">Rate</a></font></td>
				<%		elseif intRateForum = 0 and intReviewForum = 1 then
							if ReviewsExist( "Forum", ID ) then
				%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Read/Add Review" )%></a></font></td>
				<%			else
				%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Add Review" )%></a></font></td>
				<%			end if
						elseif intRateForum = 1 and intReviewForum = 1 then
							if ReviewsExist( "Forum", ID ) then
				%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
									<font size="-2"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate and Read/Add Review" )%></a></font></td>
				<%			else
				%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
								<font size="-2"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate/Add Review" )%></a></font></td>
				<%			end if
						end if%>
					<% if DisplayPrivacy then %>
						<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
					<% end if %>
					</tr>

<%

					rsPage.MoveNext
				end if
			next
			Response.Write("</table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing
		end if
	end if
else
	if not ValidCategory(intCategoryID) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but that is not a valid category."))

	GetCategory intCategoryID, strName, blPrivate, blMembersOnly

	if blPrivate AND not LoggedMember then Redirect( "login.asp?Source=forum.asp&ID=" & intCategoryID & "&Submit=Go" )

	'Keep track of shit
	IncrementHits intCategoryID, "ForumCategories"

	if DisplaySearch then
'------------------------End Code-----------------------------
%>

	<form METHOD="POST" ACTION="forum.asp">
		<input type="hidden" name="ID" value="<%=intCategoryID%>">
		Search <%=strName%> For <input type="text" name="Keywords" size="15">
		<input type="submit" name="Submit" value="Go"><br>
	</form>
<%
	end if
%>

	<table width="100%">
		<tr>
			<td align="left">
				<span class="Heading">Topic: <%=strName%> 
<%
'-----------------------Begin Code----------------------------			
				if blMembersOnly then
					%></span><font size="-2">(only members may post messages)</font><%
				else
					%></span><%
				end if
'------------------------End Code-----------------------------
%>
			</td>
<%
			if NeedCategoryMenu("ForumCategories") then
%>
			<td align="right">
				<form action="forum.asp" method="post">
					<font size="-1">Change Topic To:</font><br>
					<% PrintCategoryPullDown intCategoryID %>
					<input type="Submit" value="Switch">
				</form>
			</td>
<%
			end if
%>
		</tr>
	</table>

	<span class="LinkText"><a HREF="forum_post.asp?CategoryID=<%=intCategoryID%>">Post New</a></span><br>


<%
'-----------------------Begin Code----------------------------


	'They entered text to search for, so we are going to get matches and put them into the SectionSearch
	if Request("Keywords") <> "" then
		Query = "SELECT ID, Date, Email, Author, Subject, Body, MemberID FROM ForumMessages WHERE (CategoryID = " & intCategoryID & " AND CustomerID = " & CustomerID & ") ORDER BY Date DESC"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set Email = rsList("Email")
			Set Author = rsList("Author")
			Set Subject = rsList("Subject")
			Set Body = rsList("Body")
			Set MemberID = rsList("MemberID")
		intSearchID = SingleSearch()
		rsList.Close
		set rsList = Nothing
	end if

	if intSearchID <> "" then
		'Their search came up empty
		if intSearchID = 0 then
			if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
				<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
				Try again, or <a href="forum.asp?ID=<%=intCategoryID%>">click here</a> to view all messages in this topic.</p>
<%
'-----------------------Begin Code----------------------------
			else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="forum.asp?ID=<%=intCategoryID%>">click here</a> to view all messages in this topic.</p>
<%
'-----------------------Begin Code----------------------------
			end if
		else
			'They have search results, so lets list their results
			Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			rsPage.CacheSize = PageSize
			Set TargetID = rsPage("TargetID")
%>
			<form METHOD="POST" ACTION="forum.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
			PrintPagesHeader
			PrintTableHeader 0
%>
			<tr>
				<% if DisplayDate then %>
				<td class="TDHeader">Date</td>
				<% end if %>
				<td class="TDHeader">Author</td>
				<td class="TDHeader">Subject</td>
				<% if intRateForum = 1 then %>
					<td class="TDHeader">Rating</td>
				<% end if %>
				<% if DisplayPrivacy then %>
				<td class="TDHeader">Public?</td>
				<% end if%>
			</tr>
<%
			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Date, Email, Author, Subject, MemberID, Private, TimesRated, TotalRating FROM ForumMessages WHERE CustomerID = " & CustomerID
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set Email = rsList("Email")
			Set Author = rsList("Author")
			Set Subject = rsList("Subject")
			Set MemberID = rsList("MemberID")
			Set TotalRating = rsList("TotalRating")
			Set TimesRated = rsList("TimesRated")
			Set IsPrivate = rsList("Private")

			for p = 1 to rsPage.PageSize
				if not rsPage.EOF then
					rsList.Filter = "ID = " & TargetID

					if MemberID > 0 then
						strAuthor = GetNickNameLink( MemberID )
					elseif InStr( Email, "@" ) then
						strAuthor = "<a href='mailto:" & Email & "'>" & PrintTDLink(Author) & "</a>"
					else
						strAuthor = Author
					end if
%>
					<tr>
						<% if DisplayDate then %>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
						<% end if %>
						<td class="<% PrintTDMain %>"><%=strAuthor%></td>
						<td class="<% PrintTDMain %>"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
				<%		if intRateForum = 1 and intReviewForum = 0 then
				%>			<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
							<font size="-2"><a href="forum_read.asp?ID=<%=ID%>">Rate</a></font></td>
				<%		elseif intRateForum = 0 and intReviewForum = 1 then
							if ReviewsExist( "Forum", ID ) then
				%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Read/Add Review" )%></a></font></td>
				<%			else
				%>				<td class="<% PrintTDMain %>" align=center><font size="-2"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Add Review" )%></a></font></td>
				<%			end if
						elseif intRateForum = 1 and intReviewForum = 1 then
							if ReviewsExist( "Forum", ID ) then
				%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
									<font size="-2"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate and Read/Add Review" )%></a></font></td>
				<%			else
				%>				<td class="<% PrintTDMain %>" align=center><%=GetRating( TotalRating, TimesRated )%> 
								<font size="-2"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink( "Rate/Add Review" )%></a></font></td>
				<%			end if
						end if%>
						<% if DisplayPrivacy then %>
						<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
						<% end if %>
					</tr>
<%
					rsPage.MoveNext
				end if
			next
			Response.Write("</table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing
		end if

	'They are just cycling through the messages.  No searching.
	else
		intTempMessageID = Request("MessageID")

		if Request("Action") = "Expand" then
			if intTempMessageID = "" then Redirect("error.asp")
			Query = "SELECT SessionID, MessageID FROM ForumThreadExpanded"
			Set rsExpand = Server.CreateObject("ADODB.Recordset")
			rsExpand.Open Query, Connect, adOpenStatic, adLockOptimistic
			if Session("UserID") = "" then
				if rsExpand.EOF then
					intSessionID = 1
				else
					rsExpand.MoveLast
					intSessionID = ( rsExpand("SessionID") + 1 )
				end if
				Session("UserID") = intSessionID
			end if
			rsExpand.AddNew
			rsExpand("SessionID") = Session("UserID")
			rsExpand("MessageID") = intTempMessageID
			rsExpand.Update
			rsExpand.Close
			set rsExpand = Nothing
		end if

		if Request("Action") = "Collapse" then
			if intTempMessageID = "" then Redirect("error.asp")
			Query = "DELETE ForumThreadExpanded WHERE (SessionID = " & Session("UserID") & " AND MessageID = " & intTempMessageID & ")"
			Connect.Execute Query, , adCmdText + adExecuteNoRecords
		end if

		'lets get the base messages
		Query = "SELECT ID, Date, Email, Author, Subject, Private, MemberID, TimesRated, TotalRating FROM ForumMessages WHERE CustomerID = " & CustomerID & " AND CategoryID = " & intCategoryID & " AND BaseID = 0 ORDER BY Date DESC"
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly

			Set ID = rsPage("ID")
			Set ItemDate = rsPage("Date")
			Set Email = rsPage("Email")
			Set Author = rsPage("Author")
			Set Subject = rsPage("Subject")
			Set MemberID = rsPage("MemberID")
			Set TotalRating = rsPage("TotalRating")
			Set TimesRated = rsPage("TimesRated")
			Set IsPrivate = rsPage("Private")


		'Don't navigate if it's empty
		if not rsPage.EOF then
%>
			<form METHOD="POST" ACTION="forum.asp">
			<input type="hidden" name="ID" value="<%=intCategoryID%>">
<%
			PrintPagesHeader

			Set rsReplies = Server.CreateObject("ADODB.Recordset")
			rsReplies.CacheSize = 20

			for p = 1 to rsPage.PageSize

				if not rsPage.EOF then
					if not HasReplies( ID ) then

						if DisplayPrivacy and IsPrivate = 1 and not blPrivate then
							strPrivate = "Private, "
						else
							strPrivate = ""
						end if
						'Check the email address
						if MemberID > 0 then
							strAuthor = GetNickNameLink( MemberID )
						elseif InStr( Email, "@" ) then
							strAuthor = "<a href='mailto:" & Email & "'>" & Author & "</a>"
						else
							strAuthor = Author
						end if
						if TimesRated > 0 and intRateForum = 1 then
							strRating = ", Rating: " & GetRating( TotalRating, TimesRated )
						else
							strRating = ""
						end if

						strDate = ""
						if DisplayDate then strDate = ", " & FormatDateTime(ItemDate, 2)
						%>

						&nbsp;&nbsp;&nbsp;<% PrintNew(ItemDate) %> <A HREF="forum_read.asp?ID=<%=ID%>"><%=Subject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%> <%=strDate%> <%=strRating%> )</font><br>
						<%
					else
						if Session("UserID") = "" then
							blExpanded = false
						else
							blExpanded = IsExpanded( Session("UserID"), ID )
						end if

						if DisplayPrivacy and IsPrivate = 1 and not blPrivate then
							strPrivate = "Private, "
						else
							strPrivate = ""
						end if
						'Check the email address
						if MemberID > 0 then
							strAuthor = GetNickNameLink( MemberID )
						elseif InStr( Email, "@" ) then
							strAuthor = "<a href='mailto:" & Email & "'>" & Author & "</a>"
						else
							strAuthor = Author
						end if
						if TimesRated > 0 and intRateForum = 1 then
							strRating = ", Rating: " & GetRating( TotalRating, TimesRated )
						else
							strRating = ""
						end if

						strDate = ""
						if DisplayDate then strDate = ", " & FormatDateTime(ItemDate, 2)

						if blExpanded = False then
				%>
							<a href="forum.asp?Action=Expand&MessageID=<%=ID%>&PageNum=<%=intPage%>&ID=<%=intCategoryID%>"><% PrintPlus %></a>&nbsp;<% PrintNew(ItemDate) %> <A HREF="forum_read.asp?ID=<%=ID%>"><%=Subject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%> <%=strDate%> <%=strRating%> )</font><br>
				<%
						else
				%>
							<a href="forum.asp?Action=Collapse&MessageID=<%=ID%>&PageNum=<%=intPage%>&ID=<%=intCategoryID%>"><% PrintMinus %></a>&nbsp;<% PrintNew(ItemDate) %> <A HREF="forum_read.asp?ID=<%=ID%>"><%=Subject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%> <%=strDate%> <%=strRating%> )</font><br>
				<%
							'Let's check if there are replies.  If so, just print out the link with no + or -
							Query = "SELECT ID, Date, Email, Author, Subject, Private, MemberID, TimesRated, TotalRating FROM ForumMessages WHERE BaseID = " & ID & " ORDER BY Date"
							rsReplies.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

							Set RepID = rsReplies("ID")
							Set RepItemDate = rsReplies("Date")
							Set RepEmail = rsReplies("Email")
							Set RepAuthor = rsReplies("Author")
							Set RepSubject = rsReplies("Subject")
							Set RepIsPrivate = rsReplies("Private")
							Set RepMemberID = rsReplies("MemberID")
							Set RepTimesRated = rsReplies("TimesRated")
							Set RepTotalRating = rsReplies("TotalRating")

							'Print the replies
							do until rsReplies.EOF
								'Check the email address
								if RepMemberID > 0 then
									strAuthor = GetNickNameLink( RepMemberID )
								elseif InStr( RepEmail, "@" ) then
									strAuthor = "<a href='mailto:" & RepEmail & "'>" & RepAuthor & "</a>"
								else
									strAuthor = RepAuthor
								end if
								if DisplayPrivacy and RepIsPrivate = 1 and not blPrivate then
									strPrivate = "Private, "
								else
									strPrivate = ""
								end if

								if TimesRated > 0 and intRateForum = 1 then
									strRating = ", Rating: " & GetRating( RepTotalRating, RepTimesRated )
								else
									strRating = ""
								end if

								strDate = ""
								if DisplayDate then strDate = ", " & FormatDateTime(RepItemDate, 2)

				%>
									&nbsp;&nbsp;&nbsp;<a href="forum.asp?Action=Collapse&MessageID=<%=ID%>&PageNum=<%=intPage%>&ID=<%=intCategoryID%>"><% PrintMinus %></a>&nbsp;<% PrintNew(RepItemDate) %> <a href="forum_read.asp?ID=<%=RepID%>"><%=RepSubject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%> <%=strDate%> <%=strRating%> )</font><br>
				<%
								rsReplies.MoveNext
							loop
							rsReplies.Close
						end if
					end if
					rsPage.MoveNext
				end if
			next
			rsPage.Close

			Set rsReplies = Nothing


		'Give them the link to change the section's properties
		if LoggedAdmin() and IncludeEditSectionPropButtons = 1 then
			Response.Write "<br><br><p align=right><a href='admin_sectionoptions_edit.asp?Type=Properties&Section=Forum&Source=forum.asp'>Change Section Options</a></p>"
		end if

		else
			'If there are no available messages
%>
			<p>Sorry, but there are no messages in this topic.</p>
<%
		end if

		set rsPage = Nothing
	end if
end if


'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	if MemberID > 0 then
		GetDesc = UCASE(Subject & Body & ItemDate & GetNickName(MemberID) )
	else
		GetDesc = UCASE(Subject & Body & ItemDate & Author & Email )
	end if
End Function


'------------------------End Code-----------------------------
%>
