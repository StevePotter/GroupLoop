<!-- #include file="forum_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
blLoggedAdmin = LoggedAdmin
blLoggedMember = LoggedMember
if not CBool( IncludeForum ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not (blLoggedAdmin or CBool( ForumMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
if not LoggedMember then Redirect("members.asp?Source=members_forum_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Messages</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------

if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if (blLoggedAdmin and Request("Date") = "") or Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")

	Query = "SELECT * FROM ForumMessages WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockPessimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if Request("Private") = "1" then 
		rsUpdate("Private") = 1
	else
		rsUpdate("Private") = 0
	end if
	if blLoggedAdmin then rsUpdate("Date") = Request("Date")
	rsUpdate("Subject") = Format( Request("Subject") )
	rsUpdate("Body") = GetTextArea( Request("Body") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update

	intCategoryID = rsUpdate("CategoryID")

	'All the administrator updates
	if blLoggedAdmin then
		intCategoryID = CInt(Request("CategoryID"))

		'Change the category to all the messages in the thread
		if Request("CatAllChange") = "1" and rsUpdate("CategoryID") <> intCategoryID then
			'This is the head
			if rsUpdate("BaseID") = 0 then
				intBaseID = intID
			else
				'This is a reply, so update the head
				intBaseID = rsUpdate("BaseID")
			end if

			Query = "UPDATE ForumMessages SET CategoryID = '" & intCategoryID & "' WHERE ID = " & intBaseID
			Connect.Execute Query, , adCmdText + adExecuteNoRecords

			Query = "UPDATE ForumMessages SET CategoryID = '" & intCategoryID & "' WHERE BaseID = " & intBaseID
			Connect.Execute Query, , adCmdText + adExecuteNoRecords

		'Only change this ones category
		elseif Request("CatAllChange") = "" and rsUpdate("CategoryID") <> intCategoryID then
			'Update the category
			rsUpdate("CategoryID") = intCategoryID

			'If this is a base message then make the first reply the new head in the old category
			Query = "SELECT ID, IP, ModifiedID, BaseID FROM ForumMessages WHERE BaseID = " & intID & " ORDER BY ID"
			Set rsRepliesUpdate = Server.CreateObject("ADODB.Recordset")
			rsRepliesUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
			if not rsRepliesUpdate.EOF then
				intNewBaseID = rsRepliesUpdate("ID")
				rsRepliesUpdate("BaseID") = 0
				rsRepliesUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
				rsRepliesUpdate("ModifiedID") = Session("MemberID")
				rsRepliesUpdate.Update
				rsRepliesUpdate.Close
				set rsRepliesUpdate = Nothing
			end if
			'Now make every other one in the old thread point to the new head
			Query = "UPDATE ForumMessages SET BaseID = '" & intNewBaseID & "' WHERE BaseID = " & intID
			Connect.Execute Query, , adCmdText + adExecuteNoRecords
		end if

		'Make it not a reply
		if Request("Reply") = "" then rsUpdate("BaseID") = 0

		'Update the shit if it wasn't a member's post
		if rsUpdate("MemberID") = 0 then
			rsUpdate("Author") = Format(Request("Author"))
			rsUpdate("EMail") = Request("EMail")
		end if
	end if

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The message has been edited. &nbsp;<a href="members_forum_modify.asp">Click here</a> to modify another.<br>
	<a href="members_forum_modify.asp?ID=<%=intCategoryID%>">Click here</a> to modify another message in <%=GetForumCategory(intCategoryID)%>.<br>
	<a href="forum_read.asp?ID=<%=intID%>">Click here</a> to read the modified message.
	</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Delete" or strSubmit = "Delete Replies Also" or strSubmit = "DeleteReplies" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	'Delete replies also
	if blLoggedAdmin and (strSubmit = "Delete Replies Also" or strSubmit = "DeleteReplies") then
		Query = "SELECT BaseID FROM ForumMessages WHERE ID = " & intID
		Set rsUpdate = Server.CreateObject("ADODB.Recordset")
		rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
		if rsUpdate.EOF then
			set rsUpdate = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
		end if

		'Get the base message
		if rsUpdate("BaseID") > 0 then
			intBaseID = rsUpdate("BaseID")
		else
			intBaseID = intID
		end if

		rsUpdate.Close
		set rsUpdate = Nothing

		Query = "DELETE ForumMessages WHERE ID = " & intBaseID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		Query = "DELETE ForumMessages WHERE BaseID = " & intBaseID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
%>
		<p>The message and its replies have been deleted. &nbsp;<a href="members_forum_modify.asp">Click here</a> to modify another.</p>
<%
	'Just delete the message
	else
		Query = "SELECT BaseID FROM ForumMessages WHERE ID = " & intID
		Set rsUpdate = Server.CreateObject("ADODB.Recordset")
		rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
		if rsUpdate.EOF then
			set rsUpdate = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
		end if

		'It's a reply, so just delete it
		if rsUpdate("BaseID") > 0 then
			rsUpdate.Delete
			rsUpdate.Update
		'Base, so make sure replies stay
		else
			intBaseID = intID
			rsUpdate.Delete
			rsUpdate.Update
			rsUpdate.Close
			'If this is a base message then make the first reply the new head
			Query = "SELECT ID, BaseID, IP, ModifiedID FROM ForumMessages WHERE BaseID = " & intBaseID & " ORDER BY ID"
			rsUpdate.CacheSize = 20
			rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
			if not rsUpdate.EOF then
				intNewBaseID = rsUpdate("ID")
				rsUpdate("BaseID") = 0
				rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
				rsUpdate("ModifiedID") = Session("MemberID")
				rsUpdate.Update
			end if
			rsUpdate.Close
			set rsUpdate = Nothing

			'Make the rest of the replies point to the new head
			Query = "UPDATE ForumMessages SET BaseID = '" & intNewBaseID & "' WHERE BaseID = " & intBaseID
			Connect.Execute Query, , adCmdText + adExecuteNoRecords
		end if
%>
		<p>The message has been deleted. &nbsp;<a href="members_forum_modify.asp">Click here</a> to modify another.</p>
<%
	end if
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT Private, Subject, Date, Body, CategoryID, BaseID, Author, Email, MemberID FROM ForumMessages WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if rsEdit("Private") = 1 then 
		strChecked = "checked"
	else
		strChecked = ""
	end if
'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	<a href="inserts_view.asp?Table=InfoPages" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_forum_modify.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=intID%>">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<%PrintTableHeader 0%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>
<%	if blLoggedAdmin then %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      		<td class="<% PrintTDMain %>"> 
				<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
			</td>
		</tr>
<%		if rsEdit("MemberID") = 0 then	%>
			<tr> 
      			<td class="<% PrintTDMain %>" align="right">* Author</td>
      			<td class="<% PrintTDMain %>"> 
       				<input type="text" name="Author" size="55" value="<%=rsEdit("Author") %>">
     			</td>
			</tr>
			<tr> 
      			<td class="<% PrintTDMain %>" align="right">* EMail</td>
      			<td class="<% PrintTDMain %>"> 
       				<input type="text" name="EMail" size="55" value="<%=rsEdit("EMail") %>">
     			</td>
			</tr>
<%		end if
		if rsEdit("BaseID") > 0 then	%>
			<tr> 
				<td class="<% PrintTDMain %>" align="right">Keep this message as a reply to something?</td>
				<td class="<% PrintTDMain %>"> 
					<input type="checkbox" name="Reply" value="1" checked>
     			</td>
   			</tr>
<%		end if %>	
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Category</td>
			<td class="<% PrintTDMain %>">
<%
		Query = "SELECT ID, Name, Private, MembersOnly FROM ForumCategories WHERE (CustomerID = " & CustomerID & ")"
		Set rsTempCats = Server.CreateObject("ADODB.Recordset")
		rsTempCats.CacheSize = 20
		rsTempCats.Open Query, Connect, adOpenStatic, adLockReadOnly

		'Make the size 3 if there are many members
		if rsTempCats.RecordCount <= 30 then
			%><select name="CategoryID" size="1"><%
		else
			%><select name="CategoryID" size="3"><%
		end if

		do until rsTempCats.EOF
			strSelect = ""
			if rsTempCats("ID") = rsEdit("CategoryID") then strSelect = "SELECTED"
			Response.Write "<option value = '" & rsTempCats("ID") & "' " & strSelect & ">" & rsTempCats("Name") & "</option>" & vbCrlf
			rsTempCats.MoveNext
		loop
		rsTempCats.Close
		set rsTempCats = Nothing
		Response.Write("</select>")

		'Now make sure they want to apply changes to all messages or not
		'this is a base message if there are replies
		if rsEdit("BaseID") = 0 then
			if HasReplies( intID ) then
%>
				<br>Apply any category change to all replies?  <input type="checkbox" name="CatAllChange" value="1" checked>
<%
			end if
		'this is a reply
		else
%>
			<br>Apply any category change to the head message and all other replies?  <input type="checkbox" name="CatAllChange" value="1" checked>
<%
		end if

	end if
%>
     	</td>
    </tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">*Subject</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Subject" size="55" value="<%=FormatEdit( rsEdit("Subject") )%>">
     	</td>
    </tr>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">*Message (inserts allowed)</td>
    	<td class="<% PrintTDMain %>"> 
			<% TextArea "Body", 55, 20, True, ExtractInserts( rsEdit("Body") ) %>
    	</td>
    </tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
    </tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing
else
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

	'Start them off in the first category if none is specified
	if intCategoryID = "" then

		'They entered text to search for, so we are going to get matches and put them into the SectionSearch
		if Request("Keywords") <> "" AND intSearchID = "" then
			Query = "SELECT ID, Date, Email, Author, Subject, Body, MemberID FROM ForumMessages WHERE " & strMatch & " ORDER BY Date DESC"
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

		if intSearchID = "" then
	%>
			<form METHOD="POST" ACTION="members_forum_modify.asp">
				Search For <input type="text" name="Keywords" size="15">
				<input type="submit" name="Submit" value="Go"><br>
			</form>
	<%
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintCategoryMenu "members_forum_modify.asp"
			Set rsPage = Nothing
		else
			'Their search came up empty
			if intSearchID = 0 then
				if Session("MemberID") <> "" then
	'-----------------------End Code----------------------------
	%>
					<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
					Try again, or <a href="members_forum_modify.asp">click here</a> to go back to the topic list.</p>
	<%
	'-----------------------Begin Code----------------------------
				else
	'-----------------------End Code----------------------------
	%>
					<p>Sorry, but your search came up empty.<br>
					Try again, or <a href="members_forum_modify.asp">click here</a> to go back to the topic list.</p>
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
				<form METHOD="POST" ACTION="members_forum_modify.asp">
				<input type="hidden" name="SearchID" value="<%=intSearchID%>">
	<%
				PrintPagesHeader
				PrintTableHeader 0
	%>
				<tr>
					<td class="TDHeader">Date</td>
					<% if blLoggedAdmin then %>
					<td class="TDHeader">Author</td>
					<% end if %>
					<td class="TDHeader">Subject</td>
					<td class="TDHeader">Topic</td>
					<td class="TDHeader">Public?</td>
					<td class="TDHeader">&nbsp;</td>
				</tr>
	<%
				'Instantiate the recordset for the output
				Set rsList = Server.CreateObject("ADODB.Recordset")
				Query = "SELECT ID, Date, Email, Author, Subject, CategoryID, MemberID, Private, TimesRated, TotalRating FROM ForumMessages WHERE " & strMatch
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
							strAuthor = "<a href='mailto:" & Email & "'>" & Author & "</a>"
						else
							strAuthor = PrintTDLink(Author)
						end if
	%>
						<form METHOD="POST" ACTION="members_forum_modify.asp">
						<input type="hidden" name="ID" value="<%=ID%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
							<% if blLoggedAdmin then %>
							<td class="<% PrintTDMain %>"><%=strAuthor%></td>
							<% end if %>
							<td class="<% PrintTDMain %>"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
							<td class="<% PrintTDMain %>"><%=GetForumCategory(CategoryID)%></td>
							<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
							<td class="<% PrintTDMainSwitch %>">
								<input type="submit" name="Submit" value="Edit">
								<input type="button" value="Delete" onClick="DeleteBox('If you delete this message, there is no way to get it back.  Are you sure?', 'members_forum_modify.asp?Submit=Delete&ID=<%=ID%>')">			
							</td>
						</tr>
						</form>
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

		if blPrivate AND not LoggedMember then Redirect( "login.asp?Source=members_forum_modify.asp&ID=" & intCategoryID & "&Submit=Go" )

		'Keep track of shit
		IncrementHits intCategoryID, "ForumCategories"

	'------------------------End Code-----------------------------
	%>

		<form METHOD="POST" ACTION="members_forum_modify.asp">
			<input type="hidden" name="ID" value="<%=intCategoryID%>">
			Search <%=strName%> For <input type="text" name="Keywords" size="15">
			<input type="submit" name="Submit" value="Go"><br>
		</form>

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
				<td align="right">
					<form action="members_forum_modify.asp" method="post">
						<font size="-1">Change Topic To:</font><br>
						<% PrintCategoryPullDown intCategoryID %>
						<input type="Submit" value="Switch">
					</form>
				</td>
			</tr>
		</table>

		<span class="LinkText"><a HREF="forum_post.asp?CategoryID=<%=intCategoryID%>">Post New</a></span><br>


	<%
	'-----------------------Begin Code----------------------------


		'They entered text to search for, so we are going to get matches and put them into the SectionSearch
		if Request("Keywords") <> "" then
			Query = "SELECT ID, Date, Email, Author, Subject, Body, MemberID FROM ForumMessages WHERE (CategoryID = " & intCategoryID & " AND " & strMatch & ") ORDER BY Date DESC"
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
					Try again, or <a href="members_forum_modify.asp?ID=<%=intCategoryID%>">click here</a> to view all messages in this topic.</p>
	<%
	'-----------------------Begin Code----------------------------
				else
	'-----------------------End Code----------------------------
	%>
					<p>Sorry, but your search came up empty.<br>
					Try again, or <a href="members_forum_modify.asp?ID=<%=intCategoryID%>">click here</a> to view all messages in this topic.</p>
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
				<form METHOD="POST" ACTION="members_forum_modify.asp">
				<input type="hidden" name="SearchID" value="<%=intSearchID%>">
	<%
				PrintPagesHeader
				PrintTableHeader 0
	%>
				<tr>
					<td class="TDHeader">Date</td>
					<% if blLoggedAdmin then %>
					<td class="TDHeader">Author</td>
					<% end if %>
					<td class="TDHeader">Subject</td>
					<td class="TDHeader">Public?</td>
					<td class="TDHeader">&nbsp;</td>
				</tr>
	<%
				'Instantiate the recordset for the output
				Set rsList = Server.CreateObject("ADODB.Recordset")
				Query = "SELECT ID, Date, Email, Author, Subject, MemberID, Private, TimesRated, TotalRating FROM ForumMessages WHERE " & strMatch
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
							strAuthor = "<a href='mailto:" & Email & "'>" & Author & "</a>"
						else
							strAuthor = PrintTDLink(Author)
						end if
	%>
						<form METHOD="POST" ACTION="members_forum_modify.asp">
						<input type="hidden" name="ID" value="<%=ID%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
					<%		if blLoggedAdmin then %>
							<td class="<% PrintTDMain %>"><%=strAuthor%></td>
					<%		end if %>
							<td class="<% PrintTDMain %>"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
							<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
							<td class="<% PrintTDMainSwitch %>">
								<input type="submit" name="Submit" value="Edit">
								<input type="button" value="Delete" onClick="DeleteBox('If you delete this message, there is no way to get it back.  Are you sure?', 'members_forum_modify.asp?Submit=Delete&ID=<%=ID%>')">			
							</td>

						</tr>
						</form>
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
			if blLoggedAdmin then
				'lets get the base messages
				Query = "SELECT ID, Date, Email, Author, Subject, Private, MemberID, TimesRated, TotalRating FROM ForumMessages WHERE CustomerID = " & CustomerID & " AND CategoryID = " & intCategoryID & " AND BaseID = 0 ORDER BY Date DESC"
			else
				Query = "SELECT ID, Date, Email, Author, Subject, Private, MemberID, TimesRated, TotalRating FROM ForumMessages WHERE " & strMatch & " ORDER BY Date DESC"
			end if

			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.CacheSize = PageSize
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly


			'Don't navigate if it's empty
			if not rsPage.EOF then
				Set ID = rsPage("ID")
				Set ItemDate = rsPage("Date")
				Set Email = rsPage("Email")
				Set Author = rsPage("Author")
				Set Subject = rsPage("Subject")
				Set MemberID = rsPage("MemberID")
				Set TotalRating = rsPage("TotalRating")
				Set TimesRated = rsPage("TimesRated")
				Set IsPrivate = rsPage("Private")
	%>
				<form METHOD="POST" ACTION="members_forum_modify.asp">
				<input type="hidden" name="ID" value="<%=intCategoryID%>">
	<%
				PrintPagesHeader

				if blLoggedAdmin then

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


					Set rsReplies = Server.CreateObject("ADODB.Recordset")
					rsReplies.CacheSize = 20

					for p = 1 to rsPage.PageSize

						if not rsPage.EOF then
							if not HasReplies( ID ) then

								if IsPrivate = 1 and not blPrivate then
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
								%>
								&nbsp;&nbsp;&nbsp;<% PrintNew(ItemDate) %> 
								<a href="members_forum_modify.asp?Submit=Edit&ID=<%=ID%>">Edit</a>&nbsp; 
								<a href="javascript:DeleteBox('If you delete this message, there is no way to get it back.  Are you sure?', 'members_forum_modify.asp?Submit=Delete&ID=<%=ID%>')">Delete</a>&nbsp;&nbsp; 
								<A HREF="forum_read.asp?ID=<%=ID%>"><%=Subject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%>, <%=FormatDateTime(ItemDate, 2)%> <%=strRating%> )</font><br>
								<%
							else
								if Session("UserID") = "" then
									blExpanded = false
								else
									blExpanded = IsExpanded( Session("UserID"), ID )
								end if

								if IsPrivate = 1 and not blPrivate then
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

								if blExpanded = False then
						%>
									<a href="members_forum_modify.asp?Action=Expand&PageNum=<%=intPage%>&MessageID=<%=ID%>&ID=<%=intCategoryID%>"><% PrintPlus %></a>&nbsp;<% PrintNew(ItemDate) %> 
									<a href="members_forum_modify.asp?Submit=Edit&ID=<%=ID%>">Edit</a>&nbsp; 
									<a href="javascript:DeleteBox('If you delete this message, there is no way to get it back.  Are you sure?', 'members_forum_modify.asp?Submit=Delete&ID=<%=ID%>')">Delete</a>&nbsp; 
									<a href="javascript:DeleteBox('If you delete this message and its replies, there is no way to get them back.  Are you sure?', 'members_forum_modify.asp?Submit=DeleteReplies&ID=<%=ID%>')">Delete Replies Also</a>&nbsp;&nbsp; 
									<A HREF="forum_read.asp?ID=<%=ID%>"><%=Subject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%>, <%=FormatDateTime(ItemDate, 2)%> <%=strRating%> )</font><br>
						<%
								else
						%>
									<a href="members_forum_modify.asp?Action=Collapse&PageNum=<%=intPage%>&MessageID=<%=ID%>&ID=<%=intCategoryID%>"><% PrintMinus %></a>&nbsp;<% PrintNew(ItemDate) %> 
									<a href="members_forum_modify.asp?Submit=Edit&ID=<%=ID%>">Edit</a>&nbsp; 
									<a href="javascript:DeleteBox('If you delete this message, there is no way to get it back.  Are you sure?', 'members_forum_modify.asp?Submit=Delete&ID=<%=ID%>')">Delete</a>&nbsp; 
									<a href="javascript:DeleteBox('If you delete this message and its replies, there is no way to get them back.  Are you sure?', 'members_forum_modify.asp?Submit=DeleteReplies&ID=<%=ID%>')">Delete Replies Also</a>&nbsp;&nbsp; 
									<A HREF="forum_read.asp?ID=<%=ID%>"><%=Subject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%>, <%=FormatDateTime(ItemDate, 2)%> <%=strRating%> )</font><br>
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
										if RepIsPrivate = 1 and not blPrivate then
											strPrivate = "Private, "
										else
											strPrivate = ""
										end if

										if TimesRated > 0 and intRateForum = 1 then
											strRating = ", Rating: " & GetRating( RepTotalRating, RepTimesRated )
										else
											strRating = ""
										end if
						%>
											&nbsp;&nbsp;&nbsp;<a href="members_forum_modify.asp?Action=Collapse&MessageID=<%=ID%>&PageNum=<%=intPage%>&ID=<%=intCategoryID%>"><% PrintMinus %></a>&nbsp;
											<a href="members_forum_modify.asp?Submit=Edit&ID=<%=RepID%>">Edit</a>&nbsp; 
											<a href="javascript:DeleteBox('If you delete this message, there is no way to get it back.  Are you sure?', 'members_forum_modify.asp?Submit=Delete&ID=<%=RepID%>')">Delete</a>&nbsp;&nbsp; 			
											<% PrintNew(RepItemDate) %> <a href="forum_read.asp?ID=<%=RepID%>"><%=RepSubject%></a> <font size="-2"> ( <%=strPrivate%> <%=strAuthor%>, <%=FormatDateTime(RepItemDate, 2)%> <%=strRating%> )</font><br>
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

				else
					PrintTableHeader 0	%>
					<tr>
						<td class="TDHeader">Date</td>
						<td class="TDHeader">Subject</td>
						<td class="TDHeader">Public?</td>
						<td class="TDHeader">&nbsp;</td>
					</tr>
<%					for p = 1 to rsPage.PageSize
						if not rsPage.EOF then
							if MemberID > 0 then
								strAuthor = GetNickNameLink( MemberID )
							elseif InStr( Email, "@" ) then
								strAuthor = "<a href='mailto:" & Email & "'>" & Author & "</a>"
							else
								strAuthor = Author
							end if
%>
							<form METHOD="POST" ACTION="members_forum_modify.asp">
							<input type="hidden" name="ID" value="<%=ID%>">
							<tr>
								<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
								<td class="<% PrintTDMain %>"><a href="forum_read.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
								<td class="<% PrintTDMainSwitch %>"><%=PrintPublic(IsPrivate)%></td>
								<td class="<% PrintTDMainSwitch %>">
									<input type="submit" name="Submit" value="Edit">
									<input type="button" value="Delete" onClick="DeleteBox('If you delete this message, there is no way to get it back.  Are you sure?', 'members_forum_modify.asp?Submit=Delete&ID=<%=ID%>')">			
								</td>

							</tr>
							</form>
<%							rsPage.MoveNext
						end if
					next
					Response.Write("</table>")
					rsPage.Close

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
%>