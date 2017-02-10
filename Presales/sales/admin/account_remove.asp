<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\dsn.asp" -->

<p align="center"><span class=Heading>Remove Your Account</span></p>
<%
'-----------------------Begin Code----------------------------
if Request("CustomerID") = "" or Request("MemberID") = "" or Request("Password") = "" then Redirect("message.asp?Message=" & Server.URLEncode("You are missing information."))
intCustomerID = CInt(Request("CustomerID"))
intMemberID = CInt(Request("MemberID"))
strPassword = Request("Password")

Set FileSystem = CreateObject("Scripting.FileSystemObject")
Set Command = Server.CreateObject("ADODB.Command")

With Command
	'Check to make sure the CC info is correct
	.ActiveConnection = Connect
	.CommandText = "GetOwnerMemberID"
	.CommandType = adCmdStoredProc
	.Parameters.Refresh
	.Parameters("@CustomerID") = intCustomerID
	.Execute , , adExecuteNoRecords
	intOwnerID = .Parameters("@MemberID")
	checkPassword = .Parameters("@Password")

	'Check on their password
	if checkPassword <> strPassword then Redirect("error.asp?Message=" & Server.URLEncode("Invalid password passed."))

	'Double check their info, probably don't need this.  oh well, just double check
	.CommandText = "ValidMember"
	.Parameters.Refresh
	.Parameters("@CustomerID") = intCustomerID
	.Parameters("@MemberID") = intOwnerID
	.Parameters("@NickName") = ""
	.Parameters("@Password") = strPassword
	.Execute , , adExecuteNoRecords
	blValid = CBool(.Parameters("@Valid"))
	'Wrong info
	if not blValid then	Redirect("message.asp?Message=" & Server.URLEncode("Non-existing member.  If you already deleted your site, please close your browser."))

	'Get the subdirectory
	.CommandText = "GetCustomerInfo"
	.Parameters.Refresh
	.Parameters("@CustomerID") = intCustomerID
	.Execute , , adExecuteNoRecords
	strSubDir = .Parameters("@SubDirectory")
	strFirstName = .Parameters("@FirstName")
	strLastName = .Parameters("@LastName")
	strTitle = .Parameters("@Title")
	if strSubDir = "" then
		Set Command = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("Your Subdirectory was not in our records.  Please e-mail <a href=mailto:support@grouploop.com>support@grouploop.com</a> immediately.  Include your Credit Card information and your CustomerID (" & intCustomerID & ")"))
	end if

	'Check the folder.  better be there
	strFolder = Server.MapPath("../" & strSubDir)
	if not FileSystem.FolderExists(strFolder) then
		Set Command = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("Your Subdirectory could not be found on our server.  Please e-mail <a href=mailto:support@grouploop.com>support@grouploop.com</a> immediately.  Include your Credit Card information and your CustomerID (" & intCustomerID & ")"))
	end if

	if LCase(strFolder) = "e:\webs\websites\ourclubpage.com" then Redirect "error.asp"

	SendEMails

	'Delete the customer in the db
	.CommandText = "DeleteCustomer"
	.Parameters.Refresh
	.Parameters("@CustomerID") = intCustomerID
	.Execute , , adExecuteNoRecords

	'Delete the folder
	FileSystem.DeleteFolder strFolder, True
End With
Set Command = Nothing

Set FileSystem = Nothing


Sub SendEMails()
	'Open up all the members
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	rsNew.CacheSize = 50
	Query = "SELECT ID, FirstName, LastName, EMail1 FROM Members WHERE EMail1 <> '' AND CustomerID = " & intCustomerID
	rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	intNumEMails = 0

	if not rsNew.EOF then
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.ContentType = "text/html"
		Mailer.IgnoreMalformedAddress = true
		Mailer.RemoteHost  = "mail4.burlee.com"
		Mailer.FromName    = "GroupLoop.com"
		Mailer.FromAddress = "support@grouploop.com"
		Mailer.Subject    = "Your GroupLoop Site - " & strTitle & " - Has Been Removed"

		strBody = "<html><body><p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
		"&nbsp;&nbsp;&nbsp;Your GroupLoop.com site, " & strTitle & " has been removed.<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;We'd like to thank you for being with us, and we are sorry to lose you.  This company was founded on a " & _
		"dream and you helped to make that dream a reality.  Please check back with us " & _
		"once in a while to look for new additions and improvements.  And please tell your friends!<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you so much for being a part of GroupLoop,<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p></body></html>" & VbCrLf

		Mailer.BodyText   = strBody

		do until rsNew.EOF
			strEMail = rsNew("EMail1")

			'Send the confirmation email
			strRecipName = rsNew("FirstName") & " " & rsNew("LastName")

			'We have 50 recipients, send out the email
			if intNumEMails = 50 then
				intNumEMails = 0
				Mailer.SendMail
				Mailer.ClearRecipients
			end if


			intNumEMails = intNumEMails + 1

			Mailer.AddRecipient strRecipName, strEMail

			rsNew.MoveNext

		loop

		'We have other than 50 recipients, send out the email
		if intNumEMails <> 50 then Mailer.SendMail


	end if
	rsNew.Close

	Set rsNew = Nothing


	'Send an email alerting me
	strNewBody = "CustomerID: " & intCustomerID & " has removed their account"

	Mailer.ClearRecipients
	Mailer.ClearBodyText
	Mailer.AddRecipient "Accounts", "accounts@grouploop.com"
	Mailer.Subject    = "GroupLoop Site # " & intCustomerID & " Has Been Removed"
	Mailer.BodyText   = strNewBody
	Mailer.Sendmail

	Set Mailer = Nothing
End Sub
'------------------------End Code-----------------------------
%>
<p>Your account has been removed.  Thank you so much for choosing us, and check back with us from time to time, and please recommend us to others.</p>
<%
'-----------------------Begin Code---------------------------
%>

<!-- #include file="..\sourcegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->