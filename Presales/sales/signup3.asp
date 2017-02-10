<!-- #include file="header.asp" -->
<!-- #include file="..\homegroup\dsn.asp" -->
<!-- #include file="functions.asp" -->

<%
'-----------------------Begin Code----------------------------

strFirstName = Format(Request("FirstName"))
strLastName = Format(Request("LastName"))
strNickName = Format(Request("NickName"))
strPassword = Request("PW1")
strEMail = Request("EMail")
Birthdate = AssembleDate("BirthDate")

strStreet1 = Format(Request("Street1"))
strStreet2 = Format(Request("Street2"))
strCity = Format(Request("City"))
strState = Request("State")
strCountry = Request("Country")
strZip = Request("Zip")
strPhone = Request("Phone")



if strFirstName = "" or strLastName = "" or strNickName = "" or strPassword = "" or strEMail = "" or _
strStreet1 = "" or strCity = "" or strState = "" or strCountry = "" or strZip = "" or strPhone = "" then Redirect("message.asp?Message=" & Server.URLEncode("You are missing information."))


if EmployeeNickNameTaken( strNickName ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but someone else has that nickname.  Please choose another.<br>"))



Set Command = Server.CreateObject("ADODB.Command")


intReferralID = 0
'If there is a salesman ID sent, make sure it is valid
CheckSalesmanID

With Command
	'Check to make sure the CC info is correct
	.ActiveConnection = Connect
	.CommandText = "AddEmployee"
	.CommandType = adCmdStoredProc
	.Parameters.Refresh
	.Parameters("@FirstName") = strFirstName
	.Parameters("@LastName") = strLastName
	.Parameters("@NickName") = strNickName
	.Parameters("@Password") = strPassword
	.Parameters("@Birthdate") = Birthdate
	.Parameters("@HomeStreet1") = strStreet1
	.Parameters("@HomeStreet2") = strStreet2
	.Parameters("@HomeCity") = strCity
	.Parameters("@HomeState") = strState
	.Parameters("@HomeZip") = strZip
	.Parameters("@HomePhone") = strPhone
	.Parameters("@HomeCountry") = strCountry
	.Parameters("@EMail1") = strEMail
	.Parameters("@ReferralID") = intReferralID

	.Execute , , adExecuteNoRecords
	intEmployeeID = .Parameters("@EmployeeID")
End With
Set Command = Nothing

Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.ContentType = "text/html"
Mailer.IgnoreMalformedAddress = true
Mailer.RemoteHost  = "mail4.burlee.com"
Mailer.FromName    = "GroupLoop.com"
Mailer.FromAddress = "support@grouploop.com"
Mailer.Subject    = "Your GroupLoop Salesperson Information"

strBody = "<html><body><p align=center><img src='http://www.GroupLoop.com/sales/title.gif' border=0></p>" & _
"<p>Congratulations " & strFirstName & "!  You have become part of our continually rewarding sales program. Our unique system doesn't just pay you commission once on a sale.  We pay you 20% each month for your sales!</p>" & _
"<p>Everything you need to get started is available at our sales Web Site.  You will need to remember the following information, so please write it down or print it out:<br>" & _
"Your NickName: <b>" & strNickName & "</b><br>" & _
"Your Password: <b>" & strPassword & "</b></p>" & _
"Your Salesman ID Number: <b>" & intEmployeeID & "</b></p>" & _

"<p>Just go to <i><a href='http://www.GroupLoop.com/sales'>www.GroupLoop.com/Sales</a></i> to log in and begin your exciting journey!</p>" & _

"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you so much for being a part of GroupLoop,<br>" & VbCrLf & _
"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p></body></html>" & VbCrLf

Mailer.BodyText   = strBody
Mailer.AddRecipient strFirstName & " "  & strLastName, strEMail
Mailer.BodyText   = strNewBody

Mailer.Sendmail

Mailer.ClearRecipients
Mailer.AddRecipient "Accounts", "accounts@grouploop.com"
Mailer.Subject    = "New Salesman - #" & intEmployeeID & " - " & strFirstName & " "  & strLastName

Mailer.Sendmail

Set Mailer = Nothing


Sub CheckSalesmanID()
	With Command
		'Check the salesman
		intReferralID = 0
		if Request("Referral") <> "" then
			intReferralID = CInt(Request("Referral"))
			.ActiveConnection = Connect
			.CommandText = "EmployeeIDExists"
			.CommandType = adCmdStoredProc
			.Parameters.Refresh
			.Parameters("@EmployeeID") = intReferralID
			.Execute , , adExecuteNoRecords
			blExists = CBool(.Parameters("@Exists"))
			'They don't exist! Get the FUCK out
			if not blExists then
				Set FileSystem = Nothing
				Set Command = Nothing
				Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you have an invalid referral ID number.  Please check it and re-enter."))
			end if
		end if
	End With
End Sub
'------------------------End Code-----------------------------
%>
<p>You are all ready to go!  An e-mail has been sent to <%=strEMail%>, with all the information you need to remember.</p>

<p><b>If you want to get started right away, just <a href="login.asp?NickName=<%=strNickName%>&Password=<%=strPassword%>">click here to log in!</a></b></p>

<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->

