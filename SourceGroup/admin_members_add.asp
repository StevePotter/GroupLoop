<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_members_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add A New Member</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------

'Add the story
if Request("Submit") = "Add" then
	strNickName = Format(Request("NickName"))
	strFirstName = Format(Request("FirstName"))
	strLastName = Format(Request("LastName"))
	strEMail = Request("EMail1")
	if strFirstName = "" or strLastName = "" or strNickName = "" or Request("PW1") = "" then Redirect("incomplete.asp")

	if Request("SendMail") = "YES" then
		SendMail = true
	else
		SendMail = false
	end if


	Set cmdMember = Server.CreateObject("ADODB.Command")
	With cmdMember
		'CREATE THE NEW MEMBER
		.ActiveConnection = Connect
		.CommandText = "AddMember"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh

		if Request("Admin") = "YES" then 
			.Parameters("@Admin") = 1
		else
			.Parameters("@Admin") = 0
		end if

		.Parameters("@FirstName") = strFirstName
		.Parameters("@LastName") = strLastName
		.Parameters("@NickName") = strNickName
		.Parameters("@EMail1") = strEMail
		.Parameters("@Password") = Request("PW1")
	End With

	Set rsInfo = Server.CreateObject("ADODB.Recordset")

	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.ContentType = "text/html"
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = MailerFromName
	Mailer.FromAddress = "support@grouploop.com"

	blMultiSiteMember = MultiSiteMember()

	intCommonID = 0

	'If they can add to more than one site....
	if blMultiSiteMember then
		SitesToAdd = Request("SiteCustID")
		'Get the list of sites

		Set rsSites = Server.CreateObject("ADODB.Recordset")

		GetMemberSitesRecordset rsSites

		if not rsSites.EOF then
			Set SiteCustID = rsSites("CustomerID")
			Set SiteTitle = rsSites("Title")
		end if

		do until rsSites.EOF
			'If they chose this site to be added to, or all the sites
			if SitesToAdd = "All" or InStr( SitesToAdd, SiteCustID ) then
				AddMember SiteCustID
			end if

			rsSites.MoveNext
		loop
		rsSites.Close
		Set rsSites = Nothing
		Response.Write "<a href=admin_members_add.asp>Click here</a> to add another.<br>"
	else
		AddMember CustomerID
	end if

	Set rsInfo = Nothing
	Set cmdMember = Nothing
	Set Mailer = Nothing
else
	'give the link if they are already a member
	if ParentSiteExists() or ChildSiteExists() then
%>
		<b><a href="admin_members_existing_add.asp">Click here if the member is already a member of another site.</a></b><br>
<%
	end if
%>
	* indicates required information<br>

	<form method="post" action="admin_members_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<% PrintTableHeader 0 %>
<%
		if MultiSiteMember() then
%>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">What sites should this member be added to?  To select more than one, hold down the Control ('Ctrl') key.</td>
			<td class="<% PrintTDMain %>"> 
				<% PrintMemberSites %>
			</td>
   		</tr>
<%
		end if
%>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* First Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="FirstName" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Last Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="LastName" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* <%=UsernameLabel%></td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="NickName" size="55">
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Password</td>
    		<td class="<% PrintTDMain %>"> 
    			<input type="text" name="PW1" size="55">
    		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">E-Mail Address (remember @aol.com for AOL members)</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="EMail1" size="55">
     		</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">E-Mail <%=UsernameLabel%>, Password, and Instructions To New Member?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="SendMail" value="YES" checked>
			</td>
   		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Enter any personal message to be sent with the e-mail here</td>
    		<td class="<% PrintTDMain %>"> 
    			<textarea name="Message" cols="50" rows="3" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Give Administrator Access?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="Admin" value="YES">
			</td>
   		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Add">
    		</td>
		</tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if


Sub AddMember( CustID )
	strSiteTitle = GetSiteTitle( CustID )

	if NickNameTaken(strNickName, CustID) then
		if blMultiSiteMember then
			Response.Write "<b>" & strNickName & " could not be added to " & strSiteTitle & " because the " & UsernameLabel & " was already taken.</b><br>"
		else
			Redirect("message.asp?Message=" & Server.URLEncode(strNickName & " could not be added because the " & UsernameLabel & " was already taken.<br>"))
		end if
		Exit Sub
	end if

	'Add the person to the
	With cmdMember
		.Parameters("@CustomerID") = CustID
		.Parameters("@CommonID") = intCommonID
		.Execute , , adExecuteNoRecords
		if intCommonID = 0 then intCommonID = .Parameters("@MemberID")
	End With


	'Now email the person their new info
	if SendMail and strEMail <> "" then
		'Get the customer info
		Query = "SELECT FirstName, LastName, DomainName, UseDomain, Subdirectory, Organization FROM Customers WHERE ID = " & CustID
		rsInfo.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

		strCreator = rsInfo("FirstName") & " " & rsInfo("LastName")

		intUseDomain = rsInfo("UseDomain")
		strSubdirectory = rsInfo("Subdirectory")
		strOrganization = rsInfo("Organization")

		'if they used a domain name, it may not be transferred yet, so we are going to include the alternate URL
		if intUseDomain = 1 then
			strURL = rsInfo("DomainName")
			strAltURL = "http://www.GroupLoop.com/" & strSubdirectory
		else
			strURL = "http://www.GroupLoop.com/" & strSubdirectory
		end if

		rsInfo.Close
		Query = "SELECT FirstName, LastName, EMail1 FROM Members WHERE ID = " & Session("MemberID")
		rsInfo.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

		strAdminName = rsInfo("FirstName") & " " & rsInfo("LastName")
		strAdminEMail = rsInfo("EMail1")

		rsInfo.Close

		'The subject line
		strSubject = "You have been added as a member to '" & strSiteTitle & "' by " & strAdminName

		'The body
		strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
		"<p><i>Dear " & strFirstName & " " & strLastName & ",</i><br>" & _
		" &nbsp;&nbsp;You have been added as a member to '" & strSiteTitle & "'.  This Web Site, part of the <a href='http://www.GroupLoop.com'>GroupLoop.com</a> network, was started by " & strCreator

		'Add the organization name if there is one
		if strOrg <> "" then strBody = strBody & ", part of " & strOrg

		'Continue
		strBody = strBody & ".</p><p>" &  _
			" &nbsp;&nbsp;If this is the first time you have heard about this Web Site, you need filled in.  You are not just part of a regular Web Site, " &_
			"you are part of a <i>Virtual Community</i>.  This Web Site may have such sections as: announcements, stories, quizzes, photos, voting, message forums, and a calendar.  " &_
			"As a member you have the ability to participate in any of these sections.  You may easily add to any of these sections, and you may view items only for members' eyes.  "  &_
			"</p><p> &nbsp;&nbsp;This site was made possible by <a href='http://www.GroupLoop.com'>GroupLoop.com</a>, and was made to be as user-friendly as possible.  So, if you have hardly any Internet experience, or have been using "  &_
			"it for years, this site will be a pleasure to visit.  Explore the site, and you will discover how exciting and powerful being a member of a unique, e-community can be." & _
			"</p><p>With that said, here is the information you need to get started.  Please <i>write your password down</i>, because you will need it to log into the Site.<br>" & _
			"Your Nickname: <b>" & strNickName  & "</b><br>" & _
			"Your Password: <b>" & Format(Request("PW1"))  & "</b><br>" & _
			"Site Address: <b><a href='" & strURL & "'>" & strURL & "</a></b></p>"

		'If they need the alternate URL
		if intUseDomain = 1 then
			strBody = strBody & "<p><b>Note:</b> If the above link does not work, that's because the domain name hasn't been set up yet.  Sometimes it takes a little while to set up the name.  But " & _
			"don't worry, you can still get on the page.  Use this link instead: <b><a href='" & strAltURL & "'>" & strAltURL & "</a></b><br>" & _
			"Once the domain name is set up, you can use that instead.</p>"
		end if

		if Request("Message") <> "" then
			strBody = strBody & "<p>" & strAdminName & " has left you the following personal message:&nbsp;&nbsp;" & Format(Request("Message")) & "</p>"
		end if

		'Give the closing and 
		strBody = strBody &	VbCrLf & "<p>Have fun!  To access the Members Only section, simply click on the <b>Members Only</b> button at any time.  We recommend first viewing the Member Manual (inside the Members section), and then changing your member information." & VbCrLf
		if intAdmin = 1 then
			strBody = strBody & VbCrLf & "</p><p>CONGRATULATIONS!  You have been given administrator access on the Web Site.  As an administrator, you have the ability to " & _
				"change everything on your site.  You can change the way the site looks, add new members, determine which sections are allowed, what members can " & _
				"and cannot do, and many other things.  All your options are located in the 'Members' section.  Please keep in mind that as an administrator, you have power over everything.  So please know what you are doing by reading the administrator's " & _
				"manual, which can be accessed from the 'Members' section."
		end if
		strBody = strBody & VbCrLf & _
				"</p><p>Please do not respond to this e-mail.  Your questions may be answered by the administrator who added you or the GroupLoop.com staff.<br>" & VbCrLf

		if strAdminEMail <> "" then
			strBody = strBody &	"You may e-mail the administrator who added you (" &  strAdminName  & ") at: <a href=mailto:" & strAdminEMail & ">" & strAdminEMail & "</a><br>" & VbCrLf
		end if

		strBody = strBody & "You may e-mail the GroupLoop.com staff at: <a href=mailto:support@grouploop.com>support@grouploop.com</a><br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf


		strBody = strBody & "<p>Please read GroupLoop.com's Terms Of Service <a href='http://www.GroupLoop.com/homegroup/tos.asp'>here</a>.  Your signing into your site verifies that you have read and accept the Terms Of Service, so please read it carefully.</p>" & VbCrLf

		strHeader = "<html><title>" & strSubject & "</title><body>"
		strFooter = "</body></html>"

		strRecipName = strFirstName & " " & strLastName

		'Set the rest of the mailing info and send it
		Mailer.ClearRecipients
		Mailer.AddRecipient strRecipName, strEMail
		Mailer.Subject    = strSubject
		Mailer.BodyText   = strHeader & strBody & strFooter

		if not Mailer.SendMail then 
%>				<p>There has been an error, and the email has not been sent to the new member.  Please make sure you had a valid e-mail address entered.  Try again, and if the problem
				persists, e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a>.  Please include the error below.<br>
				Error was '<%=Mailer.Response%>'
<%		end if
	end if
	if blMultiSiteMember then
%>
		<%=strRecipName%> has been added to <%=strSiteTitle%>.<br>
<%
	else
%>
		<p><%=strRecipName%> has been added.  <a href="admin_members_add.asp">Click here</a> to add another.</p>
<%
	end if
End Sub

'------------------------End Code-----------------------------
%>