<p align="<%=HeadingAlignment%>"><span class=Heading>Apply for Membership</span></p>
<%
'-----------------------Begin Code----------------------------
if AllowMemberApplications <> 1 then Redirect "error.asp"

'Add the story
if Request("Submit") = "Apply" then
	strNickName = Format(Request("NickName"))
	strFirstName = Format(Request("FirstName"))
	strLastName = Format(Request("LastName"))
	strEMail = Request("EMail1")
	strPW = Request("PW1")
	if strFirstName = "" or strLastName = "" or strNickName = "" or Request("PW1") = "" then Redirect("incomplete.asp")

	if NickNameTaken(strNickName, CustomerID) then
%>
	Sorry, but the <%=UsernameLabel%> <%=strNickName%> is already taken.
<%
	else
		'Get the customer info
		Query = "SELECT * FROM MembersApplied"
		Set rsNew = Server.CreateObject("ADODB.Recordset")
		rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic

		rsNew.AddNew
		rsNew("CustomerID") = CustomerID

		rsNew("FirstName") = strFirstName
		rsNew("LastName") = strLastName
		rsNew("NickName") = strNickName
		rsNew("EMail1") = strEMail
		rsNew("Password") = strPW

		rsNew.Update
		rsNew.Close

		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.ContentType = "text/html"
		Mailer.RemoteHost  = "mail4.burlee.com"
		Mailer.FromName    = MailerFromName
		Mailer.FromAddress = "support@grouploop.com"

		strSubject = Request("FirstName") & " " & Request("LastName") & " has applied for membership to " & Title & "."
		Mailer.Subject    = FormatEdit(strSubject)


		Query = "SELECT ID, FirstName, LastName, EMail1 FROM Members WHERE EMail1 <> '' AND CustomerID = " & CustomerID & " AND Admin = 1"
		rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

		do until rsNew.EOF

			'The body
			strBody = "<p align=center><a href='http://www.GroupLoop.com'><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></a></p>" & _
			"<p><i>Dear " & rsNew("FirstName") & ",</i><br>" & _
			"<a href=mailto:" & strEMail & ">" & strFirstName & " " & strLastName & "</a> has applied as a member to your site.  <br>" & _
			"<a href=" & Chr(34) & "http://www.GroupLoop.com/" & SubDirectory & "/members.asp" & Chr(34) & ">Click here to log in and accept/decline their membership.</a></p>"

			strHeader = "<html><title>" & strSubject & "</title><body>"
			strFooter = "</body></html>"

			strRecipName = rsNew("FirstName") & " " & rsNew("LastName")

			'Set the rest of the mailing info and send it
			Mailer.ClearRecipients
			Mailer.AddRecipient strRecipName, rsNew("EMail1")
			Mailer.ClearBodyText
			Mailer.BodyText   = strHeader & strBody & strFooter

				if not Mailer.SendMail then 
	%>			<p>There has been an error, and the administrator has not been notified.  Try again, and if the problem
					persists, e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a>.  Please include the error below.<br>
					Error was '<%=Mailer.Response%>'
	<%		end if
			rsNew.MoveNext
		loop

		Set rsNew = Nothing
		Set Mailer = Nothing
	'------------------------End Code-----------------------------
	%>
			<p>Your application has been sent.  If you are accepted, you will receive an e-mail.</p>
	<%
'-----------------------Begin Code----------------------------
	end if
else
%>
	<p>To apply for membership, simply fill out this form.  If you get accepted, you will receive an e-mail welcoming 
	you to the site.</p>

	* indicates required information<br>

	<form method="post" action="members_apply.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<% PrintTableHeader 0 %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Your First Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="FirstName" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Your Last Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="LastName" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Your Site <%=UsernameLabel%> (this is required to log in, so make it something you can remember)</td>
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
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Apply">
    		</td>
		</tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if
'------------------------End Code-----------------------------
%>