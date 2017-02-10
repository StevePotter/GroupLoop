<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_members_applied_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Membership Applications</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------


strSubmit = Request("Submit")

'Add the story
if strSubmit = "Accept" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Set cmdMember = Server.CreateObject("ADODB.Command")
	With cmdMember
		'CREATE THE NEW MEMBER
		.ActiveConnection = Connect
		.CommandText = "AddAppliedMember"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@AppliedID") = intID
		.Parameters("@CustomerID") = CustomerID
		.Execute , , adExecuteNoRecords
		intMemberID = .Parameters("@MemberID")

	End With

	Set cmdMember = Nothing

	Query = "SELECT * FROM Members WHERE ID = " & intMemberID
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	Query = "SELECT * FROM MembersApplied WHERE ID = " & intID
	Set rsApplied = Server.CreateObject("ADODB.Recordset")
	rsApplied.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	rsNew("FirstName") = rsApplied("FirstName")
	rsNew("LastName") = rsApplied("LastName")
	rsNew("NickName") = rsApplied("NickName")
	rsNew("Password") = rsApplied("Password")
	strEMail = rsApplied("EMail1")
	rsNew("EMail1") = strEMail
	rsNew("EMail2") = rsApplied("EMail2")

	rsNew("Beeper") = rsApplied("Beeper")
	rsNew("CellPhone") = rsApplied("CellPhone")
	rsNew("Birthdate") = AssembleDate("Birthdate")
	rsNew("HomeStreet") = rsApplied("HomeStreet")
	rsNew("HomeCity") = rsApplied("HomeCity")
	rsNew("HomeState") = rsApplied("HomeState")
	rsNew("HomeZip") = rsApplied("HomeZip")
	rsNew("HomeCountry") = rsApplied("HomeCountry")
	rsNew("HomePhone") = rsApplied("HomePhone")
	rsNew("SecondaryDescription") = Format( rsApplied("SecondaryDescription") )
	rsNew("SecondaryStreet") = rsApplied("SecondaryStreet")
	rsNew("SecondaryCity") = rsApplied("SecondaryCity")
	rsNew("SecondaryState") = rsApplied("SecondaryState")
	rsNew("SecondaryZip") = rsApplied("SecondaryZip")
	rsNew("SecondaryCountry") = rsApplied("SecondaryCountry")
	rsNew("SecondaryPhone") = rsApplied("SecondaryPhone")
	rsNew("SecondaryPExt") = rsApplied("SecondaryPExt")

	rsNew("Custom1") = rsApplied("Custom1")
	rsNew("Custom2") = rsApplied("Custom2")

	rsNew("PrivateName") = rsApplied("PrivateName")
	rsNew("PrivateBirthdate") = rsApplied("PrivateBirthdate")
	rsNew("PrivateEMail") = rsApplied("PrivateEMail")
	rsNew("PrivateHome") = rsApplied("PrivateHome")
	rsNew("PrivateSecondary") = rsApplied("PrivateSecondary")
	rsNew("PrivateBeeper") = rsApplied("PrivateBeeper")
	rsNew("PrivateCellPhone") = rsApplied("PrivateCellPhone")


	rsNew.Update
	rsNew.Close
	Set rsNew = Nothing

	rsApplied.Delete
	rsApplied.Update
	rsApplied.Close
	Set rsApplied = Nothing

	if strEMail <> "" then SendEMail(intMemberID)
%>
	<p>The applicant has been accepted.&nbsp;
<%

	if GetNumItems("MembersApplied") > 0 then
%>
		 <a href="admin_members_applied_add.asp">Click here to accept/decline another applicant.</a></p>
<%
	else
%>
		 <a href="members.asp">Click here to return to <%=MembersTitle%>.</a></p>
<%
	end if
elseif strSubmit = "Decline" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM MembersApplied WHERE ID = " & intID & " AND CustomerID = " & CustomerID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The applicant you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

'------------------------End Code-----------------------------
%>
	<p>The applicant has been declined.&nbsp;
<%
'-----------------------Begin Code----------------------------
	if GetNumItems("MembersApplied") > 0 then
%>
		 <a href="admin_members_applied_add.asp">Click here to accept/decline another applicant.</a></p>
<%
	else
%>
		 <a href="members.asp">Click here to return to <%=MembersTitle%>.</a></p>
<%
	end if
else
	Query = "SELECT * FROM MembersApplied WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"

	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then

		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set FirstName = rsPage("FirstName")
		Set LastName = rsPage("LastName")
		Set EMail = rsPage("EMail1")
		Set NickName = rsPage("NickName")

'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="admin_members_applied_add.asp">
<%
'-----------------------Begin Code----------------------------

		PrintPagesHeader
		PrintTableHeader 0
		PrintTableTitle
		for j = 1 to rsPage.PageSize
			if not rsPage.EOF then
				PrintTableData
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
	else
'------------------------End Code-----------------------------
%>
			<p>Sorry, but there are no available applicants right now.  Either nobody has applied, or someone else accepted/declined the existing applicants.</p>
<%
'-----------------------Begin Code----------------------------
	end if
	rsPage.Close
	set rsPage = Nothing
end if




'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader">Date Applied</td>
		<td class="TDHeader">Name</td>
		<td class="TDHeader"><%=UsernameLabel%></td>
		<td class="TDHeader">E-Mail</td>
		<td class="TDHeader">Address</td>
		<td class="TDHeader">&nbsp;</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
End Sub

'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData

%>
	<form METHOD="POST" ACTION="admin_members_applied_add.asp">
	<input type="hidden" name="ID" value="<%=ID%>">
	<tr>
		<td class="<% PrintTDMain %>"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<td class="<% PrintTDMain %>"><%=FirstName%>&nbsp;<%=LastName%></td>
		<td class="<% PrintTDMain %>"><%=NickName%></td>
		<td class="<% PrintTDMain %>">
<%		if rsPage("EMail1") <> "" then %>
		
		<a href="mailto:<%=EMail%>"><%=EMail%></a>
<%
		end if
		if not rsPage("EMail2") = "" then
'------------------------End Code-----------------------------
%>
			<br><a href="mailto:<%=rsPage("EMail2")%>"><%=rsPage("EMail2")%></a>
<%'-----------------------Begin Code----------------------------
		end if
%>		
		</td>
<%
		if rsPage("HomeStreet") <> "" and rsPage("HomeCity") <> "" and rsPage("HomeState") <> "" then
'------------------------End Code-----------------------------
%>
			<td class="<% PrintTDMain %>"><%=rsPage("HomeStreet")%><br>
			    <%=rsPage("HomeCity")%>,&nbsp;<%=rsPage("HomeState")%>&nbsp;<%=rsPage("HomeZip")%>
			    <br><%=rsPage("HomePhone")%>
			</td>
<%'-----------------------Begin Code----------------------------
		else
'------------------------End Code-----------------------------
%>
			<td class="<% PrintTDMain %>">&nbsp;</td>
<%'-----------------------Begin Code----------------------------
		end if
%>
		<td class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Accept">
		</td>
		<td class="<% PrintTDMainSwitch %>">
			<input type="submit" name="Submit" value="Decline">
		</td>
		</tr>
	</form>

<%
End Sub



Sub SendEMail(intMemberID)
	'The subject line
	strSubject = "You have been added as a member to '" & Title

	Query = "SELECT ID, FirstName, LastName, EMail1, NickName, Password FROM Members WHERE EMail1 <> '' AND ID = " & intMemberID
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	strFirstName = rsNew("FirstName")
	strLastName = rsNew("LastName")
	strEMail = rsNew("EMail1")
	strPassword = rsNew("Password")
	strNickName = rsNew("NickName")

	Set rsNew = Nothing

	'The body
	strBody = "<p align=center><a href='http://www.GroupLoop.com'><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></a></p>" & _
		"<p><i>Dear " & strFirstName & " " & strLastName & ",</i><br>" & _
		" &nbsp;&nbsp;You have been added as a member to '" & Title & "' a group Web Site part of the <a href='http://www.GroupLoop.com'>GroupLoop.com</a> network."


		strBody = strBody & ".</p><p>" &  _
			" &nbsp;&nbsp;If this is the first time you have heard about this Web Site, you need filled in.  You are not just part of a regular Web Site, " &_
			"you are part of a <i>Virtual Community</i>.  This Web Site may have such sections as: announcements, stories, quizzes, photos, voting, message forums, and a calendar.  " &_
			"As a member you have the ability to participate in any of these sections.  You may easily add to any of these sections, and you may view items only for members' eyes.  "  &_
			"</p><p> &nbsp;&nbsp;This site was made possible by <a href='http://www.GroupLoop.com'>GroupLoop.com</a>, and was made to be as user-friendly as possible.  So, if you have hardly any Internet experience, or have been using "  &_
			"it for years, this site will be a pleasure to visit.  Explore the site, and you will discover how exciting and powerful being a member of a unique, e-community can be." & _
			"</p><p>With that said, here is the information you need to get started.  Please <i>write your password down</i>, because you will need it to log into the Site.<br>" & _
			"Your " & UsernameLabel & ": <b>" & strNickName  & "</b><br>" & _
			"Your Password: <b>" & Format(strPassword)  & "</b><br>" & _
			"Site Address: <b><a href='" & NonSecurePath & "'>" & NonSecurePath & "</a></b></p>"

		'Give the closing and 
		strBody = strBody &	VbCrLf & "<p>Have fun!  To access the Members Only section, simply click on the <b>Members Only</b> button at any time.  We recommend first viewing the Member Manual (inside the Members section), and then changing your member information." & VbCrLf

		strBody = strBody & VbCrLf & _
				"</p><p>Please do not respond to this e-mail.  Your questions can be answered by the GroupLoop staff, or other members of your site.<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf


		strBody = strBody & "<p>Please read GroupLoop.com's Terms Of Service <a href='http://www.GroupLoop.com/homegroup/tos.asp'>here</a>.  Your signing into your site verifies that you have read and accept the Terms Of Service, so please read it carefully.</p>" & VbCrLf

		strHeader = "<html><title>" & strSubject & "</title><body>"
		strFooter = "</body></html>"

		strRecipName = strFirstName & " " & strLastName

		'Set the rest of the mailing info and send it
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.ContentType = "text/html"
		Mailer.IgnoreMalformedAddress = true
		Mailer.RemoteHost  = "mail4.burlee.com"
		Mailer.FromName    = MailerFromName
		Mailer.FromAddress = "support@grouploop.com"
		Mailer.AddRecipient strRecipName, strEMail
		Mailer.Subject    = strSubject
		Mailer.BodyText   = strHeader & strBody & strFooter

		if not Mailer.SendMail then 
%>				<p>There has been an error, and the email has not been sent.  Please e-mail the member, 
				informing them of their membership.  All they need is the site address, their <%=UsernameLabel%> (<%=strNickName%>), and 
				their password (<%=strPassword%>).
				</p>
				If the problem
				persists, e-mail <a href="mailto:support@keist.com">support@grouploop.com</a>.  Please include the error below.<br>
				Error was '<%=Mailer.Response%>'
<%		end if
		Set Mailer = Nothing
End Sub

'------------------------End Code-----------------------------
%>