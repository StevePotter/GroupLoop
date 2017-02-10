<!-- #include file="header.asp" -->
<!-- #include file="..\dsn.asp" -->
<!-- #include file="..\functions.asp" -->

<p align="center"><span class=Heading>Request Custom Designing</span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("CustomerID") = "" then Redirect("message.asp?Message=" & Server.URLEncode("You are missing your Customer ID.  Please go back to the Modify Account menu and use the links there."))
intCustomerID = CInt(Request("CustomerID"))

strSubmit = Request("Submit")

if strSubmit = "Send My Request" then
	if Request("CCName") = "" or Request("CCNumber") = "" or Request("EMail") = "" then Redirect("incomplete.asp")

	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")

	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'Check the scheme to make sure the CC info is correct
		.ActiveConnection = Connect
		.CommandText = "ValidCustomerInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@EMail") = strEMail
		.Parameters("@CCName") = strCCName
		.Parameters("@CCNumber") = strCCNumber
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Valid"))
		'Wrong info
		if not blValid then
			Set Command = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode("The information you entered did not exactly match that of the account.  Remember that the credit card being billed for the account is the only one that will work.  Please try again."))
		end if

		'Get the subdirectory
		.CommandText = "GetCustomerInfo"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords
		strSubDir = .Parameters("@SubDirectory")
		strFirstName = .Parameters("@FirstName")
		strLastName = .Parameters("@LastName")
		strVersion = .Parameters("@Version")
		if strVersion = "Free" then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("You must upgrade to the Gold Version before you can do this."))
		end if
		if strSubDir = "" then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("Your Subdirectory was not in our records.  Please e-mail <a href=mailto:support@grouploop.com>support@grouploop.com</a> immediately.  Include your Credit Card information and your CustomerID (" & intCustomerID & ")"))
		end if
	End With
	Set Command = Nothing

	strBody = strFirstName & " " & strLastName & ", Customer# " & intCustomerID & " (http://www.GroupLoop.com/" & strSubDir & ")" & VbCrLf & _
		"Work to be done:" & VbCrLf & Request("Work") & VbCrLf & _
		"How to contact:" & VbCrLf & Request("Contact") & VbCrLf

	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.IgnoreMalformedAddress = true
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = "GroupLoop.com"
	Mailer.FromAddress = "support@grouploop.com"
	Mailer.AddRecipient "Custom Work", "customwork@grouploop.com"
	Mailer.Subject    = "Custom Work For CustomerID: " & intCustomerID & ", " & strFirstName & " " & strLastName
	Mailer.BodyText   = strBody

	if not Mailer.SendMail then 
%>		<p>There has been an error, and the request has not been sent.<br>
			Error was '<%=Mailer.Response%>'<br>
			Please try again, and if the problem persists, please e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a>.
		</p>
<%	end if
	Set Mailer = Nothing
'------------------------End Code-----------------------------
%>
	<p>Your request has been sent.  To return to your site, <a href="http://www.GroupLoop.com/<%=strSubDir%>">click here</a>.  Thanks!</p>
<%
'-----------------------Begin Code---------------------------
elseif strSubmit = "Verify" then
	if Request("CCName") = "" or Request("CCNumber") = "" or Request("EMail") = "" then Redirect("incomplete.asp")
	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")
	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'Check to make sure the CC info is correct
		.ActiveConnection = Connect
		.CommandText = "ValidCustomerInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@EMail") = strEMail
		.Parameters("@CCName") = strCCName
		.Parameters("@CCNumber") = strCCNumber
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Valid"))
		'Wrong info
		if not blValid then
			Set Command = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode("The information you entered did not exactly match that of the account.  Remember that the credit card being billed for the account is the only one that will work.  Please try again."))
		end if
	End With
	Set Command = Nothing
%>
	<form METHOD="POST" ACTION="account_custom_add.asp">
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
	<input type="hidden" name="EMail" value="<%=strEMail%>">
	<input type="hidden" name="CCName" value="<%=strCCName%>">
	<input type="hidden" name="CCNumber" value="<%=strCCNumber%>">
	<% PrintTableHeader 0 %>
		<tr>
    		<td class="TDHeader" colspan=2 align="center"> 
    			Custom Work.
    		</td>
		</tr>
		<tr>
    		<td class="<% PrintTDMain %>" colspan=2 align="left"> 
    			We have programmers and graphics designers ready to make your site really stand out.  We can create the graphics look 
				you've always wanted.  Also, if you need a custom sections (Pet Peeves for examples), it can be integrated into 
				your site easier than you think.  All we need is an idea of what you want and how to get in touch with you.  
				You will hear from us shortly with a quote.  Please remember that you will be charged once the work is done.  This 
				request does not obligate you to anything.
			</td>
		</tr>
		<tr> 
   			<td class="<% PrintTDMain %>" align="right">Please describe, in detail, what you would like done.</td>
			<td class="<% PrintTDMain %>"> 
				<textarea name="Work" cols="55" rows="20" wrap="PHYSICAL"></textarea>
			</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">How can we contact you?</td>
			<td class="<% PrintTDMain %>"> 
				<textarea name="Contact" cols="55" rows="4" wrap="PHYSICAL"></textarea>
			</td>
		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Send My Request">
			</td>
		</tr>
  	</table>
	</form>
<%
'Get their info
else
%>
	<p>Before we can edit your account, we must validate your account information.  Please enter 
	your information <b>exactly</b> like you did when you signed up.  Otherwise, it won't work.</p>
	<form METHOD="POST" ACTION="account_custom_add.asp">
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
	<% PrintTableHeader 0 %>
		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Verify Account Information
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Account E-Mail Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="EMail" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Name on Credit Card</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCName" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Credit Card Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCNumber" size="20">
     		</td>
		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Verify">
	   		</td>
		</tr>
  	</table>
	</form>
<%
end if
%>

<!-- #include file="..\closedsn.asp" -->

<!-- #include file="footer.asp" -->