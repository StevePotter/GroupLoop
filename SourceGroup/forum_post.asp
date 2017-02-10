<!-- #include file="forum_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeForum ) then Redirect("error.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Post A Message</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>
<%
'-----------------------Begin Code----------------------------
blLoggedMember = LoggedMember()


if Request("Submit") = "Post" then
	if not blLoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")

	if blLoggedMember then
		if Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")
	else
		if Request("Author") = "" or Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")
	end if

	intBaseID = CInt(Request("BaseID"))
	intCategoryID = CInt(Request("CategoryID"))

	'Make sure the base ID is valid
	if intBaseID > 0 then
		Query = "SELECT Private FROM ForumMessages WHERE ID = " & intBaseID & " AND CustomerID = " & CustomerID
		Set rsTemp = Server.CreateObject("ADODB.Recordset")
		rsTemp.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

		if rsTemp.EOF then
			Set rsTemp = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("You are trying to reply to an invalid message.  Go back to the forum and try replying to the message again.") )
		end if

		'Make sure a non-member didn't try to reply to a private message
		if rsTemp("Private") = 1 AND not blLoggedMember then
			set rsTemp = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("You are being very sneaky and trying to reply to a message you can't even read. Give it up already.") )
		end if
		rsTemp.Close
		set rsTemp = Nothing
	else
		intCategoryID = CInt(Request("CategoryID2"))
	end if

	if not ValidCategory(intCategoryID) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but that is not a valid topic."))

	GetCategory intCategoryID, strName, blPrivate, blMembersOnly
	if (blMembersOnly or blPrivate) AND not blLoggedMember then Redirect("error.asp?Message=" & Server.URLEncode("You are a non-member trying to post to a members only topic.  Don't do that."))

	Query = "SELECT ID, Private, Subject, Body, CustomerID, CategoryID, BaseID, Author, EMail, BaseID, IP, MemberID, ModifiedID FROM ForumMessages"
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Update the fields
	rsNew.AddNew
		if Request("Private") = "1" or blPrivate then
			rsNew("Private") = 1
		else
			rsNew("Private") = 0
		end if
		if blLoggedMember then
			rsNew("Subject") = Format( Request("Subject") )
			rsNew("Body") = GetTextArea( Request("Body") )
			rsNew("MemberID") = Session("MemberID")
			rsNew("ModifiedID") = Session("MemberID")
		else
			rsNew("Subject") = FormatNonMember( Request("Subject") )
			rsNew("Body") = GetTextArea( Request("Body") )
			rsNew("Author") = FormatNonMember(Request("Author"))
			rsNew("EMail") = Request("EMail")
		end if
		rsNew("CustomerID") = CustomerID
		rsNew("CategoryID") = intCategoryID
		rsNew("BaseID") = intBaseID
		rsNew("IP") = Request.ServerVariables("REMOTE_HOST")
	rsNew.Update
	rsNew.MovePrevious
	rsNew.MoveNext
	intID = rsNew("ID")
	Set rsNew = Nothing

'------------------------End Code-----------------------------
%>
	<p>Your message has been added. &nbsp;&nbsp;<a href="forum_read.asp?ID=<%=intID%>">Click here</a> to read it.
	<br><a href="forum.asp?ID=<%=intCategoryID%>">Click here</a> to go back to all messages in <%=strName%>.</p>
<%
'-----------------------Begin Code----------------------------
else
Set rsNew = Server.CreateObject("ADODB.Recordset")


	Public DisplayPrivacy

	Query = "SELECT IncludePrivacyForum, DisplaySearchForum, DisplayDaysOldForum, InfoTextForum, ListTypeForum, DisplayDateListForum, DisplayAuthorListForum, DisplayPrivacyListForum  FROM Look WHERE CustomerID = " & CustomerID
	rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = CBool(rsNew("IncludePrivacyForum")) and not cBool(SiteMembersOnly)

	Set rsNew = Nothing

	intCategoryID = Request("CategoryID")
	intBaseID = Request("ReplyID")
	'We want to promote the members logging in before writing messages, so put a login form
	'here but only once, and not if they are logged in

	'If they are nonmembers and clicked so, make sure we know
	if Request("Type") = "NonMember" then Session("NonMember") = "Y"

	'Log in members who typed in their info
	if Request("Password") <> "" or Request("NickName") <> "" then MemberLogin Request("Password"), Request("NickName")
	blLoggedMember = LoggedMember()

	if Session("NonMember") <> "Y" AND not blLoggedMember then
		strLink = "forum_post.asp?ReplyID=" & Request("ReplyID") & "&CategoryID=" & Request("CategoryID")
		if Request("Password") = "" and Request("NickName") = "" then
%>
			<p>If you are a member, please enter your information and log in below. &nbsp;<br><b>If you aren't a member, <a href="<%=strLink%>&Type=NonMember">click here</a> to add your message.</b></p>
<%
		else
%>
			<p>Sorry, but that name and password don't work.  Please try again, or if you aren't a member, <a href="<%=strLink%>&Type=NonMember">click here</a> to add your message.</b></p>
<%
		end if
		PrintLogin strLink, "Log In"
	else
		'They are posting a new message
		if intBaseID = "" then
			intCategoryID = CInt(intCategoryID)
			intBaseID = 0

			if not ValidCategory(intCategoryID) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but that is not a valid topic."))
			GetCategory intCategoryID, strName, blPrivate, blMembersOnly
			if (blMembersOnly or blPrivate) AND not blLoggedMember then Redirect( "login.asp?Source=forum_post.asp?CategoryID=" & intCategoryID & "&Submit=Log+In" )

			strSubject = ""
'------------------------End Code-----------------------------
%>
			<p class="BodyText">New Message In: <b><%=strName%></b>
<%
'-----------------------Begin Code----------------------------
		'They are replying to a message
		else
			intBaseID = CInt(intBaseID)
			Query = "SELECT ID, BaseID, Private, CategoryID, Subject FROM ForumMessages WHERE ID = " & intBaseID & " AND CustomerID = " & CustomerID
			Set rsBase = Server.CreateObject("ADODB.Recordset")
			rsBase.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

			if rsBase.EOF then
				Set rsBase = Nothing
				Redirect("error.asp?Message=" & Server.URLEncode("You are trying to reply to an invalid message.") )
			end if
			if rsBase("Private") = 1 AND not blLoggedMember then Redirect("error.asp?Message=" & Server.URLEncode(" You are a non-member trying to reply to a private message."))
			intCategoryID = rsBase("CategoryID")

			'Don't reply to a reply.  Reply to the base message.  this will keep threads even, and not crazy
			if rsBase("BaseID") <> 0 then intBaseID = rsBase("BaseID")

			strBaseSubject = rsBase("Subject")
			strBaseID = rsBase("ID")

			if InStr( strBaseSubject, "RE:" ) then
				strSubject = strBaseSubject
			else
				strSubject = "RE: " & strBaseSubject
			end if

			rsBase.Close
			set rsBase = Nothing

			if not ValidCategory(intCategoryID) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but that is not a valid topic."))
			GetCategory intCategoryID, strName, blPrivate, blMembersOnly
			if (blMembersOnly or blPrivate) AND not blLoggedMember then Redirect( "login.asp?Source=forum_post.asp?BaseID=" & intBaseID & "&Submit=Log+In" )
'------------------------End Code-----------------------------
%>
			<p class="BodyText">Reply To: <b><a href="forum_read.asp?ID=<%=strBaseID%>"><%=strBaseSubject%></a></b>
<%
'-----------------------Begin Code----------------------------
		end if

		'If the person posting the message is a member, fill in their name and e-mail for them
		if blLoggedMember then
			if blPrivate then
				Response.Write "<p>Because this is a private topic, only members may read this message.</p>"
			else
				Response.Write "<p>If you only want members to be able to read it, you should check the private box.</p>"
			end if
		else
			'The links to the login page must be different for new messages and replies (categoryID and replyID
			if Request("CategoryID") = "" then
'------------------------End Code-----------------------------
%>
				<br>If you are a member and want to make this message private, <a HREF="login.asp?Source=forum_post.asp?ReplyID=<%=intBaseID%>&Submit=Log In">click here</a>.
<%
'-----------------------Begin Code----------------------------
			else
			'Include the categoryID, not the replyID
'------------------------End Code-----------------------------
%>
				<br>If you are a member and want to make this message private, <a HREF="login.asp?Source=forum_post.asp?CategoryID=<%=intCategoryID%>&Submit=Log In">click here</a>.
<%
'-----------------------Begin Code----------------------------
			end if
		end if
'------------------------End Code-----------------------------
%>
		</p>
		<a href="inserts_view.asp" target="_blank">Click here</a> for page inserts.<br>
		<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips (non-members cannot use HTML, but can use the formatting keywords shown).<br>

		* indicates required information<br>
		<form method="post" action="forum_post.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
<%		if blLoggedMember then	%>
		<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
		<input type="hidden" name="Password" value="<%=Session("Password")%>">
<%		end if	%>
		<input type="hidden" name="BaseID" value="<%=intBaseID%>">
		<input type="hidden" name="CategoryID" value="<%=intCategoryID%>">
		<% PrintTableHeader 0 %>
		<% if blLoggedMember and not blPrivate and DisplayPrivacy then %>
			<tr> 
				<td class="<% PrintTDMain %>" align="right">Private?</td>
				<td class="<% PrintTDMain %>"> 
					<input type="checkbox" name="Private" value="1">
				</td>
   			</tr>
		<% end if 
			'If this isn't a reply, they can change the category it's in
			if intBaseID = 0 then
%>
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
						%><select name="CategoryID2" size="1"><%
					else
						%><select name="CategoryID2" size="3"><%
					end if

					do until rsTempCats.EOF
						strSelect = ""
						if rsTempCats("ID") = intCategoryID then strSelect = "SELECTED"
						'Print category unless it's a non-member and a member posting category
						if not (not blLoggedMember AND (rsTempCats("Private") = 1 or rsTempCats("MembersOnly") = 1)) then
							Response.Write "<option value = '" & rsTempCats("ID") & "' " & strSelect & ">" & rsTempCats("Name") & "</option>" & vbCrlf
						end if
						rsTempCats.MoveNext
					loop
					rsTempCats.Close
					set rsTempCats = Nothing
					Response.Write("</select>")
%>
					</td>
				</tr>
<%
			end if
			'members names are automatically entered...
			if not blLoggedMember then
%>
				<tr>
					<td class="<% PrintTDMain %>" align="right">
						* Your Name
					</td>
					<td class="<% PrintTDMain %>">
						<input type="text" size="25" name="Author">
					</td>
				</tr>
				<tr>
					<td class="<% PrintTDMain %>" align="right">
						Your E-Mail
					</td>
					<td class="<% PrintTDMain %>">
						<input type="text" size="25" name="EMail">
					</td>
				</tr>
<%			end if %>
			<tr>
				<td class="<% PrintTDMain %>" align="right">
					* Subject
				</td>
				<td  class="<% PrintTDMain %>">
					<input type="text" size = "40" name="Subject" value="<%=strSubject%>">
				</td>
			</tr>
			<tr>
				<td class="<% PrintTDMain %>" valign="top" align="right">
					* Message (inserts allowed)
				</td>
				<td class="<% PrintTDMain %>">
					<% TextArea "Body", 55, 30, True, "" %>
				</td>
			</tr>
			<tr>
    			<td colspan="2" align="center" class="<% PrintTDMain %>">
					<input type="submit" name="Submit" value="Post">
    			</td>
			</tr>
		</table>
		</form>
<%
'-----------------------Begin Code----------------------------
	end if
end if
'------------------------End Code-----------------------------
%>