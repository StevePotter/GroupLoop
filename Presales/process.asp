<!-- #include file="header.asp" -->
<!-- #include file="dsn.asp" -->
<!-- #include file="functions.asp" -->
<!-- #include file="..\sourcegroup\functions.asp" -->

<!-- #include file="..\templategroup\button_info.inc" -->

<%
'-----------------------Begin Code----------------------------
'This script signs a motherfucker up.  I payed special attention to when they click twice, 
'even though I made that pretty fucking impossible.  So, the e-mail comes first because two emails 
'won't hurt anything.  There is an error checker on the AddCustomer command, so it will return a 
'bool if the person has already been added.  Although the domain name/subdir checker should get them 
'before that.  Then last, the folder is created, which I also made sure isn't duplicated.  So this 
'bitch is pretty fucking failsafe.  Metal condom fo ma cock.
'1. Check everything and get our variables set up.
'2. Send e-mail
'3. Create customer, look, config, etc. tables
'4. Create folder, copy template files, set up DSN.asp
'5. Set Scheme for the initial look

'Record that they were here
AddHit "process.asp"


'Did they try twice?  If so, this will send them the e-mail they need
SendMailOnly



if Request("Version") = "" or ( Request("Version") <> "Gold" and Request("Version") <> "Free" and Request("Version") <> "Parent"  ) then Redirect("error.asp?Message=" & Server.URLEncode("You haven't chose which version you want.  Please go through the sign-up process from the beginning."))

Version = Request("Version")

'Make sure all the data has been passed
CheckIncomplete


intCredits = 0
intSalesmanID = 0
intCustomerID = 0
strNewFolder = ""

	'Put everything into variables.  Pretty fucking obvious.
	strFirstName = Format(Request("FirstName"))
	strLastName = Format(Request("LastName"))
	strNickName = Format(Request("NickName"))
	strPassword = Request("PW1")
	strEMail = Request("EMail")
	strStreet1 = Format(Request("Street1"))
	strStreet2 = Format(Request("Street2"))
	strCity = Format(Request("City"))
	strState = Request("State")
	strCountry = Request("Country")
	strZip = Request("Zip")
	strPhone = Request("Phone")
	strTitle = Format(Request("Title"))
	strOrganization = Format(Request("Organization"))
	intUseDomain = CInt(Request("UseDomain"))
	blUseDomain = CBool(intUseDomain)
	intSchemeID = CInt(Request("SchemeID"))
	if blUseDomain then
		strDomainName = Request("DomainName")
		if not InStr( strDomainName, "http://" ) then strDomainName = "http://" & strDomainName
		strDomainAction = Request("DomainAction")
		strSubDirectory = ""
	else
		strDomainName = ""
		strDomainAction = ""
		strSubDirectory = Request("SubDirectory")
	end if

	if Version = "Gold" or Version = "Parent" then

		strCCFirstName = Request("CCFirstName")
		strCCLastName = Request("CCLastName")
		strCCCompany = Request("CCCompany")
		strCCType = Request("CCType")
		strCCNumber = Request("CCNumber")
		intCCExpMonth = CInt(Request("CCExpMonth"))
		intCCExpYear = CInt(Request("CCExpYear"))
		CCExpDate = CDate(intCCExpMonth & "/01/" & intCCExpYear)

		if Request("CCStreet1") = "" then
			strCCStreet1 = strStreet1
		else
			strCCStreet1 = Request("CCStreet1")
		end if
		if Request("CCStreet2") = "" then
			strCCStreet2 = strStreet2
		else
			strCCStreet2 = Request("CCStreet2")
		end if
		if Request("CCCity") = "" then
			strCCCity = strCity
		else
			strCCCity = Request("CCCity")
		end if
		if Request("CCState") = "" then
			strCCState = strState
		else
			strCCState = Request("CCState")
		end if
		if Request("CCZip") = "" then
			strCCZip = strZip
		else
			strCCZip = Request("CCZip")
		end if
		if Request("CCCountry") = "" then
			strCCCountry = strCountry
		else
			strCCCountry = Request("CCCountry")
		end if

'		VerifyCard
	else
		strCCFirstName = ""
		strCCLastName = ""
		strCCCompany = ""
		strCCType = ""
		strCCNumber = ""
	end if

'If we are creating a child site, this gets all the data needed
CheckChildSite


'Open up the filesystem
Set FileSystem = CreateObject("Scripting.FileSystemObject")

'Create our command
Set Command = Server.CreateObject("ADODB.Command")
With Command
	'Check the scheme to make sure they aren't fucking around on us
	.ActiveConnection = Connect
	.CommandType = adCmdStoredProc
End With


'Set up the mailer object
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.ContentType = "text/html"
Mailer.IgnoreMalformedAddress = true
Mailer.RemoteHost  = "mail4.burlee.com"
Mailer.FromName    = "GroupLoop.com"
Mailer.FromAddress = "support@grouploop.com"

'Make sure they have a valid signup scheme
ValidScheme

'If there is a salesman ID sent, make sure it is valid
CheckSalesmanID

'Make sure they don't already exist. If they do, give them the link for the email to be sent
CustomerExists

'Make sure the directory isn't taken
SiteExists

'If they entered a promo code, make sure it is correct
GetPromo

'If they had a referral, check that shiznit
GetReferral

'Add the customer to the database
intCustomerID = AddCustomer()

'Some functions use customerID
CustomerID = intCustomerID

'Add in their additonal members and send their emails
AddAdditionalMembers

'Create the site's files/folders
strNewFolder = CreateFiles()



'Set the scheme in place
GetLook intSchemeID
GetGraphics intSchemeID, strNewFolder

'Write the header and footer

'Set the data in the config table (didn't use stored proc just because I'm lazy)
SetConfigTable

strSource = "No"
strPath = strNewFolder
%>
<!-- #include file="..\sourcegroup\write_constants.asp" -->

<!-- #include file="..\sourcegroup\write_header_footer.asp" -->
<%

'Send the confirmation e-mail
if Version = "Gold" or Version = "Parent" then
	SendGoldEMail intCustomerID
else
	SendFreeEMail intCustomerID
end if
intStep = intStep + 1


if intSalesmanID > 0 then
	SendSalesmanEMail intCustomerID
end if


Set Mailer = Nothing
Set Command = Nothing
Set FileSystem = Nothing

%>
<p>
Your site is set up and ready to go!  An e-mail receipt with your registration, member info, etc. 
has been sent to <%=strEMail%>. 
</p>
<%
if blUseDomain then
%>
	<p>
	Your site address will be: <a href="<%=strDomainName%>"><%=strDomainName%></a>.  However, the domain name may 
	take some time to get set up.  So, until the domain is get up, your site can be found at: 
	<a href="http://www.GroupLoop.com/<%=intCustomerID%>/write_header_footer.asp">http://www.GroupLoop.com/<%=intCustomerID%></a>.
	</p>
<%
else
%>
	<p>
	Your site address is: <a href="http://www.GroupLoop.com/<%=strSubDirectory%>/write_header_footer.asp">http://www.GroupLoop.com/<%=strSubDirectory%></a>.<br>
	</p>
<%
end if
if Version <> "Parent" then
%>
	<p>
	To get started quickly and easily, please click on the 'Members Only' link on your site.  
	There you will be prompted for your NickName and Password.  Enter them to log in, and you will be given all your 
	options.  We recommend reading the Member and Administrator Manuals (links are right on the Members Only page) before 
	you dive into the site.  They will tutor you on using your new site so you can become a aquainted as quickly as possible. 
	If you run into a problem that you can't solve, you may e-mail the GroupLoop.com staff at <a href="mailto:support@grouploop.com">support@grouploop.com</a>
	</p>
	<p>
	Thank you and enjoy!
	</p>
<%
else
'<form METHOD="post" ACTION="signup10.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
'	<input type="hidden" name="CustomerID" value="<%=intCustomerID">
'	<input type="hidden" name="EMail" value="<%=strEMail">
'	<input type="submit" name="Submit" value="Set Up My Child Sites">
'</form>
%>
	We are currently changing the setup for the multi-site version.  We can still set your child sites up.  Just 
	e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a> and give us a list of your child sites.  
	For now, you can enjoy your main site.  Thank you for your patience!

<%
end if

Sub SendGoldEMail( intCustomerID )
	if not IsObject(Mailer) then Set Mailer = Server.CreateObject("SMTPsvg.Mailer")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		'Get the customer's info
		.ActiveConnection = Connect
		.CommandText = "GetCustomerInfoForEMail"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords

		strFirstName = .Parameters("@FirstName")
		strLastName = .Parameters("@LastName")
		strEMail = .Parameters("@EMail")
		intMemberID = .Parameters("@MemberID")
		strCCNumber = .Parameters("@CCNumber")
		dateCreated = .Parameters("@SignUpDate")
		blUseDomain = CBool(.Parameters("@UseDomain"))
		strDomainName = .Parameters("@DomainName")
		strDomainAction = .Parameters("@DomainAction")
		strSubDirectory = .Parameters("@SubDirectory")

		if strSubDirectory = "" then strSubDirectory = intCustomerID


		'Get the customer's info
		.CommandText = "GetNickNamePassword"
		.Parameters.Refresh
		.Parameters("@MemberID") = intMemberID
		.Execute , , adExecuteNoRecords
		strNickName = .Parameters("@NickName")
		strPassword = .Parameters("@Password")
	End With
	Set cmdTemp = Nothing


	strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
	"<p><i>Dear " & strFirstName & " " & strLastName & ",</i><br>" & _
	" &nbsp;&nbsp;Your GroupLoop.com site has been set up!  Here is a summary of your new account, which will be billed " & _
	"monthly, beginning next month.  If during the first month you decide to cancel your account, please do " & _
	"so from the Member's Only section of your site, under Modify Account.  Your account:<br>" & VbCrLf & _
	"Customer ID: <b>" & intCustomerID & "</b><br>" & VbCrLf & _
	"Site Created On: <b>" & FormatDateTime(dateCreated, 1) & "</b><br>" & VbCrLf & _
	"Credit Card Number: <b>" & Left( strCCNumber, 4 ) & "...." & Right( strCCNumber, 4 ) & "</b><br>" & VbCrLf & _
	"SubDirectory: <b>" & strSubDirectory & "</b></p>" & VbCrLf

	if blUseDomain then
		strBody = strBody & "<p>Domain Name: <b>" & strDomainName & "</b><br>" & VbCrLf & _
			"Domain Action: <b>" & strDomainAction & "</b><br>"
		if LCase(strDomainAction) = "new" then
			strBody = strBody & "The domain will be registered soon, and you should receive a confirmation e-mail and a charge from Register.com" & VbCrLf
		else
			strBody = strBody & "Whoever registered the domain name should receive directions on how to transfer the domain." & VbCrLf
		end if
		strBody = strBody & "<br>Please Note: You will be charged a $2/month for the domain name.</p>" & VbCrLf
	end if

	if blUseDomain then
		strBody = strBody & VbCrLf & "<p>Your site address will be: <b><a href='" & strDomainName & "'>" & strDomainName & "</a></b>.  However, the domain name may " & _
		"take some time to get set up.  So, until the domain is get up, your site can be found at: <b><a href='http://www.GroupLoop.com/" & intCustomerID & "'>" & "http://www.GroupLoop.com/" & intCustomerID & "</a></b>.</p>" & VbCrLf & VbCrLf
	else
		strBody = strBody & "<p>Your site address is: <b><a href='http://www.GroupLoop.com/" & strSubDirectory & "'>http://www.GroupLoop.com/" & strSubDirectory & "</a></b>.</p>" & VbCrLf
	end if

	if Child then
		strBody = strBody & "<p><b>Your nickname and password are the same!</b><br>"

	else
		strBody = strBody & "<p>Your member information:<br>" & VbCrLf & _
			"Your NickName: <b>" & strNickName & "</b><br>" & _
			"Your Password: <b>" & strPassword & "</b><br>" & VbCrLf
	end if


	strBody = strBody & "<p>To get started quickly and easily, please click on the 'Members Only' link on your site.  " & _
		"There you will be prompted for your NickName and Password.  Enter them to log in, and you will be given all your " & _
		"options.  We recommend reading the Member and Administrator Manuals (links are right on the Members Only page) before " & _
		"you dive into the site.  They will tutor you on using your new site so you can become a aquainted as quickly as possible. " & _
		"If you run into a problem that you can't solve, you may e-mail the GroupLoop.com staff at: <a href=mailto:support@grouploop.com>support@grouploop.com</a><br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf

	strRecipName = strFirstName & " " & strLastName

	strHeader = "<html><title>Your GroupLoop Site Has Been Set Up!</title><body>"
	strFooter = "</body></html>"

	'Set the rest of the mailing info and send it
	Mailer.ClearRecipients
	Mailer.ClearBodyText
	Mailer.AddRecipient strRecipName, strEMail
	Mailer.Subject    = "Your GroupLoop Site Has Been Set Up!"
	Mailer.BodyText   = strHeader & strBody & strFooter

	if not Mailer.SendMail then 
%>		<p>There has been an error, and the email has not been sent.  Please e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a> and include the error below.<br>
			Error was '<%=Mailer.Response%>'
		</p>
<%	end if

	'Send the email to accounts
	Mailer.ClearRecipients
	Mailer.ClearBodyText
	Mailer.AddRecipient "Accounts", "accounts@grouploop.com"

	strDup = ""
	if Request("Action") = "MailOnly" then strDup = "RESENT MAIL - "
	if blUseDomain then
		Mailer.Subject = strDup & "New Gold Account - CustomerID:" & intCustomerID & " - DOMAIN:" & strDomainName & " - Name:" & strRecipName
	else
		Mailer.Subject = strDup & "New Gold Account - CustomerID:" & intCustomerID & " - SubDir:" & strSubDirectory & " - Name:" & strRecipName
	end if

	Mailer.BodyText   = strBody
	Mailer.Sendmail

End Sub


Sub SendFreeEMail( intCustomerID )
	if not IsObject(Mailer) then Set Mailer = Server.CreateObject("SMTPsvg.Mailer")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		'Get the customer's info
		.ActiveConnection = Connect
		.CommandText = "GetCustomerInfoForEMail"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords

		strFirstName = .Parameters("@FirstName")
		strLastName = .Parameters("@LastName")
		strEMail = .Parameters("@EMail")
		intMemberID = .Parameters("@MemberID")
		strCCNumber = .Parameters("@CCNumber")
		dateCreated = .Parameters("@SignUpDate")
		blUseDomain = CBool(.Parameters("@UseDomain"))
		strDomainName = .Parameters("@DomainName")
		strDomainAction = .Parameters("@DomainAction")
		strSubDirectory = .Parameters("@SubDirectory")

		if strSubDirectory = "" then strSubDirectory = intCustomerID


		'Get the customer's info
		.CommandText = "GetNickNamePassword"
		.Parameters.Refresh
		.Parameters("@MemberID") = intMemberID
		.Execute , , adExecuteNoRecords
		strNickName = .Parameters("@NickName")
		strPassword = .Parameters("@Password")
	End With
	Set cmdTemp = Nothing

	strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
	"<p><i>Dear " & strFirstName & " " & strLastName & ",</i><br>" & _
	"Your GroupLoop.com site has been set up!  Here is a summary of your new site:<br>" & VbCrLf & _
	"Customer ID: <b>" & intCustomerID & "</b><br>" & _
	"Site Created On: <b>" & FormatDateTime(dateCreated, 1) & "</b><br>" & _
	"SubDirectory: <b>" & strSubDirectory & "</b><br>" & _
	"Site address: <b><a href='http://www.GroupLoop.com/" & strSubDirectory & "'>http://www.GroupLoop.com/" & strSubDirectory & "</a></b></p>" & _
	"<p>Your member information:<br>" & _
	"Your NickName: <b>" & strNickName & "</b><br>" & _
	"Your Password: <b>" & strPassword & "</b></p>" & _
	 "<p>To get started quickly and easily, please click on the 'Members Only' link on your site.  " & _
	"There you will be prompted for your NickName and Password.  Enter them to log in, and you will be given all your " & _
	"options.  We recommend reading the Member and Administrator Manuals (links are right on the Members Only page) before " & _
	"you dive into the site.  They will tutor you on using your new site so you can become a aquainted as quickly as possible. " & _
	"If you run into a problem that you can't solve, you may e-mail the GroupLoop.com staff at: <a href=mailto:support@grouploop.com>support@grouploop.com</a><br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf


	strRecipName = strFirstName & " " & strLastName

	strHeader = "<html><title>Your GroupLoop Site Has Been Set Up!</title><body>"
	strFooter = "</body></html>"

	'Set the rest of the mailing info and send it
	Mailer.ClearRecipients
	Mailer.ClearBodyText
	Mailer.AddRecipient strRecipName, strEMail
	Mailer.Subject    = "Your GroupLoop Site Has Been Set Up!"
	Mailer.BodyText   = strHeader & strBody & strFooter

	if not Mailer.SendMail then 
%>		<p>There has been an error, and the email has not been sent.  Please e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a> and include the error below.<br>
			Error was '<%=Mailer.Response%>'
		</p>
<%	end if

	'Send the email to accounts
	Mailer.ClearRecipients
	Mailer.ClearBodyText
	Mailer.AddRecipient "Accounts", "accounts@grouploop.com"

	strDup = ""
	if Request("Action") = "MailOnly" then strDup = "RESENT MAIL - "
	if blUseDomain then
		Mailer.Subject = strDup & "New Free Account - CustomerID:" & intCustomerID & " - DOMAIN:" & strDomainName & " - Name:" & strRecipName
	else
		Mailer.Subject = strDup & "New Free Account - CustomerID:" & intCustomerID & " - SubDir:" & strSubDirectory & " - Name:" & strRecipName
	end if

	Mailer.BodyText   = strBody
	Mailer.Sendmail

End Sub



Sub SendSalesmanEMail( intCustomerID )
	if not IsObject(Mailer) then Set Mailer = Server.CreateObject("SMTPsvg.Mailer")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		'Get the customer's info
		.ActiveConnection = Connect
		.CommandText = "GetCustomerInfoForEMail"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords

		strFirstName = .Parameters("@FirstName")
		strLastName = .Parameters("@LastName")
		strEMail = .Parameters("@EMail")
		intMemberID = .Parameters("@MemberID")
		strCCNumber = .Parameters("@CCNumber")
		dateCreated = .Parameters("@SignUpDate")
		blUseDomain = CBool(.Parameters("@UseDomain"))
		strDomainName = .Parameters("@DomainName")
		strDomainAction = .Parameters("@DomainAction")
		strSubDirectory = .Parameters("@SubDirectory")

		if strSubDirectory = "" then strSubDirectory = intCustomerID


		'Get the customer's info
		.CommandText = "GetNickNamePassword"
		.Parameters.Refresh
		.Parameters("@MemberID") = intMemberID
		.Execute , , adExecuteNoRecords
		strNickName = .Parameters("@NickName")
		strPassword = .Parameters("@Password")

		'Get the customer's info
		.CommandText = "GetEmployeeInfo"
		.Parameters.Refresh
		.Parameters("@EmployeeID") = intSalesmanID
		.Execute , , adExecuteNoRecords
		strSalesmanName = .Parameters("@FirstName")
		strSalesmanFullName = strSalesmanName & " " & .Parameters("@LastName")
		strSalesmanEMail = .Parameters("@EMail")
		intNumCustomers = .Parameters("@NumCustomers")
	End With
	Set cmdTemp = Nothing

	'Display the number of customers
	strNumCusts = ""
	if intNumCustomers = 0 then
		strNumCusts = "first"
	elseif intNumCustomers = 1 then
		strNumCusts = "second"
	elseif intNumCustomers = 2 then
		strNumCusts = "third"
	else
		strNumCusts = intNumCustomers & "th"
	end if


	strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
	"<p><i>Dear " & strSalesmanName & ",</i><br>" & _
	"Congratulations!  Your " & strNumCusts & " customer has signed up!<br>" & VbCrLf & _
	"The person who signed up is: <b>" & strFirstName & " " & strLastName & "</b><br>" & _
	"The site's name is: <b>" & Format(Request("Title")) & "</b><br>" & _
	"Site address: <b><a href='http://www.GroupLoop.com/" & strSubDirectory & "'>http://www.GroupLoop.com/" & strSubDirectory & "</a></b></p>" & _
	"<p>Great job.  Your hard work will continuously pay off.  Thank you, and we urge you to continue to get more customers!</p>" & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Good Luck,<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf


	strHeader = "<html><title>Your " & strNumCusts & " customer has signed up</title><body>"
	strFooter = "</body></html>"

	'Set the rest of the mailing info and send it
	Mailer.ClearRecipients
	Mailer.ClearBodyText
	Mailer.AddRecipient strSalesmanFullName, strSalesmanEMail
	Mailer.Subject    = "Your " & strNumCusts & " customer has signed up"
	Mailer.BodyText   = strHeader & strBody & strFooter

	Mailer.SendMail

	'Send the email to accounts
	Mailer.ClearRecipients
	Mailer.ClearBodyText
	Mailer.AddRecipient "Accounts", "accounts@grouploop.com"

	Mailer.Subject = "Salesman #" & intSalesmanID & ", " & strSalesmanFullName & " has gotten their " & strNumCusts & " customer"

	Mailer.BodyText   = strBody
	Mailer.Sendmail
End Sub



'This sends an email to the person that referred the new customer, telling them they have a credit
Sub SendReferrerEMail( intCustomerID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		'Get the customer's info
		.ActiveConnection = Connect
		.CommandText = "GetCustomerInfoForEMail"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords

		strFirstName = .Parameters("@FirstName")
		strLastName = .Parameters("@LastName")
		strEMail = .Parameters("@EMail")
	End With
	Set cmdTemp = Nothing

	strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
		"Congratulations!<br><br>" & VbCrLf & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;" & Request("FirstName") & "&nbsp;" & Request("LastName") & " has signed up with GroupLoop.com, and " & _
		"put you as a reference.  This means you get a free month (only includes the $20 base fee, not any additional " & _
		"fees)!  To avoid possible scams, your credit will be applied in one month.  This is to prevent someone from signing up, giving you " & _
		"a free month, and then terminating the site during the first month.  Your credit will be applied at the first possible billing cycle.  " & _
		"Remember that your free credits can pile up, so get as many people to sign up as you can, " & _
		"and save yourself some money!  We thank you so much for telling others about the " & _
		"great service we offer.  This company was founded on a dream and you are making that dream possible.<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you again,<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf

	strRecipName = strFirstName & " " & strLastName

	'Set the rest of the mailing info and send it
	Mailer.ClearRecipients
	Mailer.ClearBodyText
	Mailer.AddRecipient strRecipName, strEMail
	Mailer.Subject    = "You have receive a free month's credit to your GroupLoop.com site!!!"
	Mailer.BodyText   = strBody
	Mailer.SendMail
End Sub



Sub GetLook( intID )
		'Open a new scheme
		Set rsScheme = Server.CreateObject("ADODB.Recordset")

		Query = "SELECT * FROM MenuButtons WHERE CustomerID = " & CustomerID
		rsScheme.Open Query, Connect, adOpenStatic, adLockOptimistic

		Query = "SELECT * FROM MenuButtons WHERE Custom = 0 AND SchemeID = " & intID
		Set rsLook = Server.CreateObject("ADODB.Recordset")
		rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic
		do until rsLook.EOF
			rsScheme.AddNew
			rsScheme("CustomerID") = CustomerID
			rsScheme("Position") = rsLook("Position")
			rsScheme("Name") = rsLook("Name")
			rsScheme("Show") = rsLook("Show")
			rsScheme("Align") = rsLook("Align")
			rsScheme.Update
			rsLook.MoveNext
		loop
		rsScheme.Close
		rsLook.Close


		Query = "SELECT * FROM Look WHERE SchemeID = " & intID
		rsScheme.Open Query, Connect, adOpenStatic, adLockReadOnly


		'Open up the look recordset
		Query = "SELECT * FROM Look WHERE CustomerID = " & CustomerID
		rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic



		for i = 0 to rsLook.Fields.Count - 1
			strField = rsLook(i).Name

			'Don't include these here
			blExclude = cBool( Instr(strField, "ID") or Left(strField, 8) = "InfoText" or Left(strField, 7) = "Include" or Left(strField, 8) = "ListType" or Left(strField, 7) = "Display" )

			if not blExclude then rsLook(strField) = rsScheme(strField)
		next

		rsLook.Update
		rsLook.Close
		set rsLook = Nothing

		rsScheme.Close
		set rsScheme = Nothing
End Sub


Sub GetGraphics( intSchemeID, strNewPath )

		'Get the folder paths
		strSchemeFolder = Server.MapPath("schemes") & "\" & intSchemeID & "\"
		strImageFolder = strNewPath & "images"
		'Not make sure the scheme folder exists.  If not, the graphics scheme has been lost, and we have an error
		Set FileSystem = CreateObject("Scripting.FileSystemObject")
		if FileSystem.FolderExists( strSchemeFolder ) then

		FileSystem.CopyFile strSchemeFolder&"*.*", strImageFolder
		end if
End Sub


'Custom sub for the buttons
Sub PrintCustomMenu
	Exit Sub
End Sub

'Custom sub for the buttons
Sub PrintCustomFooter
	Exit Sub
End Sub

'Custom sub for the buttons
Function GetCustomPreload
	GetCustomPreload = ""
End Function

Sub CheckIncomplete()
	'Error checking, just to make sure
	if Version = "Free" then
		if Request("SchemeID") = "" or Request("FirstName") = "" or Request("LastName") = "" or _
		Request("NickName") = "" or Request("EMail") = "" or 	Request("Title") = "" or Request("PW1") = "" or _
		Request("SubDirectory") = "" then Redirect("incomplete.asp")
	else
		if Request("SchemeID") = "" or Request("FirstName") = "" or Request("LastName") = "" or _
		Request("NickName") = "" or Request("EMail") = "" or Request("Street1") = "" or _
		Request("City") = "" or _
		( ( Request("Country") = "USA" or Request("Country") = "CAN" ) AND Request("State") = "" ) or Request("Zip") = "" or _
		Request("Phone") = "" or Request("Title") = "" or _
		(Request("CCFirstName") = "" and Request("CCLastName") = "" and Request("CCCompany") = "" ) or _
		Request("CCNumber") = "" or Request("PW1") = "" or _
		(Request("DomainAction") = "0" AND Request("SubDirectory") = "") or _
		(Request("DomainAction") = "1" AND Request("DomainName") = "") or _
		Request("Zip") = "" then Redirect("incomplete.asp")
	end if
End Sub


Sub CreateVariables()
'asd
End Sub


Public Child

Sub CheckChildSite()

	'We are creating a new child site.. secret!
	if Request("ParentID") <> "" then
		intParentID = CInt(Request("ParentID"))

		intMemberID = Request("MemberID")
		if intMemberID <> "" then intMemberID = CInt(intMemberID)

		Child = true
		Set Command = Server.CreateObject("ADODB.Command")
		With Command
			'Get the subdirectory
			.ActiveConnection = Connect
			.CommandText = "GetCustomerInfo"
			.CommandType = adCmdStoredProc
			.Parameters.Refresh
			.Parameters("@CustomerID") = intParentID
			.Execute , , adExecuteNoRecords
			ParentDir = .Parameters("@SubDirectory")
			'If there are more than one levels, just get the first subdirectory, because we don't want to go more than one level deep
			'because the include files iwth the template won't work
			if ParentDir <> "" and Instr( ParentDir, "/" ) then
				intPos = Instr( ParentDir, "/" )
				ParentDir = Left( ParentDir, (intPos - 1) )
			end if

			strSubDirectory = ParentDir & "/" & strSubDirectory

		End With
		Set Command = Nothing
	else
		intParentID = 0
		Child = false
	end if
End Sub



Sub SendMailOnly()

	'This will resend the confirmation e-mail
	if Request("Action") = "MailOnly" then
		blMailOnly = true
		intCustomerID = CInt(Request("CustomerID"))
		strEMail = Request("EMail")
		Set Command = Server.CreateObject("ADODB.Command")
		With Command
			'Check the scheme to make sure they aren't fuckign around on us
			.ActiveConnection = Connect
			.CommandType = adCmdStoredProc
			.CommandText = "GetCustomerEMail"
			.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
			.Parameters.Append .CreateParameter ("@EMail", adVarWChar, adParamOutput, 100 )
			.Parameters.Append .CreateParameter ("@Version", adVarWChar, adParamOutput, 100 )
			.Parameters("@CustomerID") = intCustomerID
			.Execute , , adExecuteNoRecords
			if UCase(strEMail) <> UCase(.Parameters("@EMail")) then
				Set Command = Nothing
				Redirect("error.asp?Source=noback&Message=" & Server.URLEncode("Sorry, we can't send the e-mail because of an invalid email passed.  If you didn't do this on purpose trying to be cute, please email <a href=mailto:support@grouploop.com>support@grouploop.com</a>."))
			end if
			strVersion = .Parameters("@Version")
		End With
		Set Command = Nothing

		if strVersion = "Free" then
			SendFreeEMail intCustomerID
		else
			SendGoldEMail intCustomerID
		end if

		Redirect "message.asp?Source=noback&Message=" & Server.URLEncode("The e-mail has been sent to " & strEMail & ".  The e-mail contains everything you need to get started.")
	end if
End Sub


Sub ValidScheme()
	With Command
		.CommandText = "ValidSignupScheme"
		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )
		.Parameters("@ItemID") = intSchemeID
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Exists"))
		if not blValid then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but you have an invalid scheme."))
		end if
	End With
End Sub



Sub CheckSalesmanID()
	With Command
		'Check the salesman
		intSalesmanID = 0
		if Request("SalesmanID") <> "" then
			intSalesmanID = CInt(Request("SalesmanID"))
			.CommandText = "EmployeeIDExists"
			.Parameters.Refresh
			.Parameters("@EmployeeID") = intSalesmanID
			.Execute , , adExecuteNoRecords
			blExists = CBool(.Parameters("@Exists"))
			'They don't exist! Get the FUCK out
			if not blExists then
				Set FileSystem = Nothing
				Set Command = Nothing
				Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you have an invalid salesman ID number.  Please check it and re-enter."))
			end if
		end if
	End With
End Sub



Sub CustomerExists()
	With Command
		'Check to make sure they haven't clicked Add twice
		.CommandText = "CustomerExists"
		.Parameters.Refresh
		.Parameters("@FirstName") = strFirstName
		.Parameters("@LastName") = strLastName
		.Parameters("@Street1") = strStreet1
		.Parameters("@City") = strCity
		.Parameters("@State") = strState
		.Parameters("@Country") = strCountry
		.Parameters("@Zip") = strZip
		.Parameters("@EMail") = strEMail
		.Parameters("@CCFirstName") = strCCFirstName
		.Parameters("@CCLastName") = strCCLastName
		.Parameters("@CCCompany") = strCCCompany
		.Parameters("@CCNumber") = strCCNumber
		.Parameters("@DomainName") = strDomainName
		.Parameters("@UseDomain") = intUseDomain
		.Parameters("@SubDirectory") = strSubDirectory

		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))
		intCustomerID = .Parameters("@CustomerID")

		'They already exist! Get the FUCK out
		if blExists then
			Set FileSystem = Nothing
			Set Command = Nothing
			Redirect("message.asp?Source=noback&Message=" & Server.URLEncode("We detected that you have already set up your account.  This usually happens if you clicked Sign Up more than once.  You should still have your confirmation e-mail waiting.  If you don't have it, please wait a few minutes.  If you still don't have it, <a href='process.asp?EMail=" & strEMail & "&Action=MailOnly&CustomerID=" & intCustomerID & "'>Click here</a> to have us send another.  If you experience any further problems, please e-mail <a href=mailto:support@grouploop.com>support@grouploop.com</a> and tell us.  Please include your name.  Thanks!"))
		end if
	End With
End Sub


Sub SiteExists()
	With Command
		'Check the subdirectory and 
		.CommandText = "SiteExists"
		.Parameters.Refresh
		.Parameters("@DomainName") = strDomainName
		.Parameters("@UseDomain") = intUseDomain
		.Parameters("@SubDirectory") = strSubDirectory
		.Execute , , adExecuteNoRecords
		blExists = CBool(.Parameters("@Exists"))

		if not blUseDomain then
			strPath = Server.MapPath("..\" & strSubDirectory)
			if FileSystem.FolderExists( strPath ) then blExists = true
		end if

		'Site already exists
		if blExists then
			Set Command = Nothing
			if blUseDomain then
				Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but that domain name has already been taken.  Please choose a different one."))
			else
				Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but that subdirectory has already been taken.  Please choose a different one."))
			end if
		end if
	End With
End Sub




Sub GetPromo()
	With Command
		'They had a promo, enter it
		if strPromoCode <> "" then
			.CommandText = "GetPromo"
			.Parameters.Refresh
			.Parameters("@Code") = strPromoCode
			.Execute , , adExecuteNoRecords
			blExists = CBool(.Parameters("@Exists"))
			if not blExists then
					Set Mailer = Nothing
					Set Command = Nothing
					Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but we could not find the promo code you entered.  Please try again."))
			end if
			intCredits = .Parameters("@Credits")
			'Response.Write intStep & ".  Reading promo code<br>"
			intStep = intStep + 1

		end if
	End With
End Sub


Sub GetReferral()
	With Command
		'Check the referral shit here
		strReferFirstName = Request("ReferFirstName")
		strReferLastName = Request("ReferLastName")
		strReferURL = Request("ReferURL")
		strPromoCode = Request("PromoCode")
		'They had a reference, let's check em
		if strReferFirstName <> "" and strReferLastName <> "" and strReferURL <> "" then
			if not InStr( strReferURL, "http://" ) then strReferURL = "http://" & strReferURL
			.CommandText = "CustomerReference"
			.Parameters.Refresh
			.Parameters("@FirstName") = strReferFirstName
			.Parameters("@LastName") = strReferLastName
			.Parameters("@URL") = strReferURL
			.Execute , , adExecuteNoRecords
			intSuccess  = .Parameters("@Success")
			intCustomerID  = .Parameters("@CustomerID")

			if intSuccess = 0 then
				Set Mailer = Nothing
				Set Command = Nothing
				Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but we could not find a site with the address you entered.  Please go back and correct the referrer's address."))
			elseif intSuccess = 1 then
				Set Mailer = Nothing
				Set Command = Nothing
				Redirect("message.asp?Message=" & Server.URLEncode("Sorry, we found a referrer with the site address you entered, but could not match the person's name.  Please go back and correct the referrer's name."))
			elseif intSuccess = 2 then
				'send an email to the referrer, telling them they got a free month
				SendReferrerEMail intCustomerID
				'Response.Write intStep & ".  Recording reference<br>"
				intStep = intStep + 1
			else
				Set Mailer = Nothing
				Set Command = Nothing
				Redirect("message.asp?Message=" & Server.URLEncode("Invalid success result when checking the reference."))
			end if
		end if
	End With
End Sub


Function AddCustomer()
	With Command
		'Add the customer
		if Version = "Gold" or Version = "Parent" then

			.CommandText = "AddGoldCustomer"
			.Parameters.Refresh

			.Parameters("@Version") = Version
			.Parameters("@SalesmanID") = intSalesmanID
			.Parameters("@Credits") = intCredits
			.Parameters("@Organization") = strOrganization
			.Parameters("@FirstName") = strFirstName
			.Parameters("@LastName") = strLastName
			.Parameters("@Street1") = strStreet1
			.Parameters("@Street2") = strStreet2
			.Parameters("@City") = strCity
			.Parameters("@State") = strState
			.Parameters("@Country") = strCountry
			.Parameters("@Zip") = strZip
			.Parameters("@Phone") = strPhone
			.Parameters("@EMail") = strEMail
			.Parameters("@CCType") = strCCType
			.Parameters("@CCFirstName") = strCCFirstName
			.Parameters("@CCLastName") = strCCLastName
			.Parameters("@CCCompany") = strCCCompany
			.Parameters("@CCExpdate") = CCExpDate
			.Parameters("@CCNumber") = strCCNumber


			.Parameters("@BillingStreet1") = strCCStreet1
			.Parameters("@BillingStreet2") = strCCStreet2
			.Parameters("@BillingCity") = strCCCity
			.Parameters("@BillingState") = strCCState
			.Parameters("@BillingZip") = strCCZip
			.Parameters("@BillingCountry") = strCCCountry

			.Parameters("@DomainName") = strDomainName
			.Parameters("@UseDomain") = intUseDomain
			.Parameters("@DomainAction") = strDomainAction
			.Parameters("@SubDirectory") = strSubDirectory
			.Parameters("@Title") = strTitle
			.Parameters("@NickName") = strNickName
			.Parameters("@Password") = strPassword

			.Parameters("@ParentID") = intParentID


			if Child then
				.Parameters("@MemberLinkID") = intMemberID
			else
				.Parameters("@MemberLinkID") = 0
			end if


			.Execute , , adExecuteNoRecords
			intCustomerID = .Parameters("@CustomerID")

		else
			.CommandText = "AddFreeCustomer"
			.Parameters.Refresh
			.Parameters("@SalesmanID") = intSalesmanID
			.Parameters("@Credits") = intCredits
			.Parameters("@Organization") = strOrganization
			.Parameters("@FirstName") = strFirstName
			.Parameters("@LastName") = strLastName
			.Parameters("@Street1") = strStreet1
			.Parameters("@Street2") = strStreet2
			.Parameters("@City") = strCity
			.Parameters("@State") = strState
			.Parameters("@Country") = strCountry
			.Parameters("@Zip") = strZip
			.Parameters("@Phone") = strPhone
			.Parameters("@EMail") = strEMail
			.Parameters("@SubDirectory") = strSubDirectory
			.Parameters("@Title") = strTitle
			.Parameters("@NickName") = strNickName
			.Parameters("@Password") = strPassword
			.Execute , , adExecuteNoRecords
			intCustomerID = .Parameters("@CustomerID")

		end if
	End With

	AddCustomer = intCustomerID
End Function




Sub AddAdditionalMembers()
	With Command
		'We are going to add their initial members
		for i = 1 to 4
			strTempNick = Format(Request("NickName"&i))
			'If they have all the info and they didn't enter their own nickname
			if not ( Request("FirstName"&i) = "" and Request("LastName"&i) = "" and strTempNick = "" and _
				Request("Password"&i) = "" and Request("EMail"&i) = "" ) and strTempNick <> strNickName and _
				not NickNameTaken(strTempNick, CustomerID) then

				strTempFirstName = Format(Request("FirstName"&i))
				strTempLastName = Format(Request("LastName"&i))

				.ActiveConnection = Connect
				.CommandText = "AddMember"
				.CommandType = adCmdStoredProc
				.Parameters.Refresh
				.Parameters("@Admin") = 0
				.Parameters("@CustomerID") = intCustomerID
				.Parameters("@FirstName") = strTempFirstName
				.Parameters("@LastName") = strTempLastName
				.Parameters("@NickName") = strTempNick
				.Parameters("@EMail1") = Request("EMail"&i)
				.Parameters("@Password") = Request("Password"&i)

				.Execute , , adExecuteNoRecords

				'if they used a domain name, it may not be transferred yet, so we are going to include the alternate URL
				if intUseDomain = 1 then
					strURL = strDomainName
					strAltURL = "http://www.GroupLoop.com/" & strSubdirectory
				else
					strURL = "http://www.GroupLoop.com/" & strSubdirectory
				end if

				strAdminName = strFirstName & " " & strLastName

				'The subject line
				strSubject = "You have been added as a member to '" & strTitle & "'"


				'The body
				strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
				"<p><i>Dear " & strTempFirstName & " " & strTempLastName & ",</i><br>" & _
				" &nbsp;&nbsp;You have been added as a member to '" & strTitle & "'.  This Web Site, part of the <a href='http://www.GroupLoop.com'>GroupLoop.com</a> network, was started by " & strFirstName & "&nbsp;" & strLastName

				'Add the organization name if there is one
				if strOrganization <> "" then strBody = strBody & ", part of " & strOrganization

				'Continue
				strBody = strBody & ".</p><p>" &  _
					" &nbsp;&nbsp;If this is the first time you have heard about this Web Site, you need filled in.  You are not just part of a regular Web Site, " &_
					"you are part of a <i>Virtual Community</i>.  This Web Site may have such sections as: announcements, stories, quizzes, photos, voting, message forums, and a calendar.  " &_
					"As a member you have the ability to participate in any of these sections.  You may easily add to any of these sections, and you may view items only for members' eyes.  "  &_
					"</p><p> &nbsp;&nbsp;This site was made possible by <a href='http://www.GroupLoop.com'>GroupLoop.com</a>, and was made to be as user-friendly as possible.  So, if you have hardly any Internet experience, or have been using "  &_
					"it for years, this site will be a pleasure to visit.  Explore the site, and you will discover how exciting and powerful being a member of a unique, e-community can be." & _
					"</p><p>With that said, here is the information you need to get started.  Please <i>write your password down</i>, because you will need it to log into the Site.<br>" & _
					"Your Nickname: <b>" & strTempNick  & "</b><br>" & _
					"Your Password: <b>" & Request("Password"&i)  & "</b><br>" & _
					"Site Address: <b><a href='" & strURL & "'>" & strURL & "</a></b></p>"


				'If they need the alternate URL
				if intUseDomain = 1 then
					strBody = strBody & "<p><b>Note:</b> If the above link does not work, that's because the domain name hasn't been set up yet.  Sometimes it takes a little while to set up the name.  But " & _
					"don't worry, you can still get on the page.  Use this link instead: <b><a href='" & strAltURL & "'>" & strAltURL & "</a></b><br>" & _
					"Once the domain name is set up, you can use that instead.</p>"
				end if


				'Give the closing and 
				strBody = strBody &	VbCrLf & "<p>Have fun!  To access the Members Only section, simply click on the <b>Members Only</b> button at any time.  We recommend first viewing the Member Manual (inside the Members section), and then changing your member information." & _
				"</p><p>Please do not respond to this e-mail. " & _
				"You may e-mail the administrator who added you (" &  strAdminName  & ") at: <a href=mailto:" & strEMail & ">" & strEMail & "</a><br>" & _
				"You may e-mail the GroupLoop.com staff at: <a href=mailto:support@grouploop.com>support@grouploop.com</a><br>" & VbCrLf & _
					"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
					"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
					"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & _
					 "<p>Please read GroupLoop.com's Terms Of Service at <a href='http://www.GroupLoop.com/homegroup/tos.asp'>http://www.GroupLoop.com/homegroup/tos.asp</a>.  Your signing into your site verifies that you have read and accept the Terms Of Service, so please read it carefully.</p>" & VbCrLf

				strHeader = "<html><title>" & strSubject & "</title><body>"
				strFooter = "</body></html>"


				strRecipName = Request("FirstName"&i) & " " & Request("LastName"&i)

				'Set the rest of the mailing info and send it
				Mailer.ClearRecipients
				Mailer.ClearBodyText
				Mailer.AddRecipient strRecipName, Request("EMail"&i)
				Mailer.Subject    = strSubject
				Mailer.BodyText   = strHeader & strBody & strFooter

				if not Mailer.SendMail then 
		%>
						<p>Error: the email to <%=Request("FirstName"&i)%>&nbsp;<%=Request("LastName"&i)%> was not sent.  The error was <%=Mailer.Response%>.</p>
		<%		end if
			end if
		next

	End With
End Sub



Function CreateFiles()
	'Now create the directory for the site
	if blUseDomain then strSubDirectory = intCustomerID	'Use the CustomerID if they have a domain name

	strNewFolder = Server.MapPath("..\" & strSubDirectory) & "\"
	if Child then
		strTempFolder = Server.MapPath("..\templategroup2") & "\"
	else
		strTempFolder = Server.MapPath("..\templategroup") & "\"
	end if


	FileSystem.CreateFolder strNewFolder
	FileSystem.CreateFolder strNewFolder & "images\"
	FileSystem.CreateFolder strNewFolder & "photos\"
	FileSystem.CreateFolder strNewFolder & "schemes\"
	FileSystem.CreateFolder strNewFolder & "storeitems\"
	FileSystem.CreateFolder strNewFolder & "storegroups\"
	FileSystem.CreateFolder strNewFolder & "media\"
	FileSystem.CreateFolder strNewFolder & "inserts\"

	'Copy the template files
	FileSystem.CopyFile strTempFolder&"*.*", strNewFolder

	CreateFiles = strNewFolder

End Function


Sub SetConfigTable()
	'Open up the config table for the write header and footer script
	Query = "SELECT * FROM Configuration WHERE CustomerID = " & intCustomerID
	Set rsSite = Server.CreateObject("ADODB.Recordset")
	rsSite.Open Query, Connect, adOpenStatic, adLockOptimistic

	rsSite("AllowMemberApplications") = CInt(Request("AllowMemberApplications"))
	rsSite("SiteMembersOnly") = CInt(Request("SiteMembersOnly"))
	rsSite("IncludeNewsletter") = CInt(Request("IncludeNewsletter"))
	rsSite("NewsletterMembers") = CInt(Request("NewsletterMembers"))
	rsSite("IncludeAnnouncements") = CInt(Request("IncludeAnnouncements"))
	rsSite("RateAnnouncements") = CInt(Request("RateAnnouncements"))
	rsSite("ReviewAnnouncements") = CInt(Request("ReviewAnnouncements"))
	rsSite("IncludeMeetings") = CInt(Request("IncludeMeetings"))
	rsSite("MeetingsMembers") = CInt(Request("MeetingsMembers"))
	rsSite("RateMeetings") = CInt(Request("RateMeetings"))
	rsSite("ReviewMeetings") = CInt(Request("ReviewMeetings"))
	rsSite("IncludeStories") = CInt(Request("IncludeStories"))
	rsSite("RateStories") = CInt(Request("RateStories"))
	rsSite("ReviewStories") = CInt(Request("ReviewStories"))
	rsSite("IncludeCalendar") = CInt(Request("IncludeCalendar"))
	rsSite("CalendarShowBirthdays") = CInt(Request("CalendarShowBirthdays"))
	rsSite("RateCalendar") = CInt(Request("RateCalendar"))
	rsSite("ReviewCalendar") = CInt(Request("ReviewCalendar"))
	rsSite("IncludeLinks") = CInt(Request("IncludeLinks"))
	rsSite("RateLinks") = CInt(Request("RateLinks"))
	rsSite("ReviewLinks") = CInt(Request("ReviewLinks"))
	rsSite("IncludeQuotes") = CInt(Request("IncludeQuotes"))
	rsSite("RateQuotes") = CInt(Request("RateQuotes"))
	rsSite("ReviewQuotes") = CInt(Request("ReviewQuotes"))
	rsSite("IncludeQuizzes") = CInt(Request("IncludeQuizzes"))
	rsSite("QuizzesMembers") = CInt(Request("QuizzesMembers"))
	rsSite("RateQuizzes") = CInt(Request("RateQuizzes"))
	rsSite("ReviewQuizzes") = CInt(Request("ReviewQuizzes"))
	rsSite("IncludeVoting") = CInt(Request("IncludeVoting"))
	rsSite("VotingMembers") = CInt(Request("VotingMembers"))
	rsSite("RateVoting") = CInt(Request("RateVoting"))
	rsSite("ReviewVoting") = CInt(Request("ReviewVoting"))
	rsSite("IncludePhotos") = CInt(Request("IncludePhotos"))
	rsSite("PhotosMembers") = CInt(Request("PhotosMembers"))
	rsSite("RatePhotos") = CInt(Request("RatePhotos"))
	rsSite("IncludePhotoCaptions") = CInt(Request("IncludePhotoCaptions"))
	rsSite("IncludeForum") = CInt(Request("IncludeForum"))
	rsSite("RateForum") = CInt(Request("RateForum"))
	rsSite("IncludeGuestbook") = CInt(Request("IncludeGuestbook"))
	rsSite("RateGuestbook") = CInt(Request("RateGuestbook"))
	rsSite("ReviewGuestbook") = CInt(Request("ReviewGuestbook"))
	rsSite("IncludeMedia") = CInt(Request("IncludeMedia"))
	rsSite("MediaMembers") = CInt(Request("MediaMembers"))
	rsSite("RateMedia") = CInt(Request("RateMedia"))
	rsSite("ReviewMedia") = CInt(Request("ReviewMedia"))

	rsSite.Update
	rsSite.Close
	Set rsSite = Nothing
End Sub


Sub VerifyCard
	'Verify the Credit Card HERE

	' Create xAuthorize object.
	Dim objAuthorize
	Set objAuthorize = Server.CreateObject("xAuthorize.Process")
	' Initialize xAuthorize for a new transaction.
	objAuthorize.Initialize

	' Set object properties.
	objAuthorize.Processor = "AUTHORIZE_NET"

	objAuthorize.FirstName = strCCFirstName
	objAuthorize.LastName = strCCLastName
	objAuthorize.Company = strCCCompany
	objAuthorize.Address = strCCStreet1 & " " & strCCStreet2
	objAuthorize.City = strCCCity
	objAuthorize.State = strState
	objAuthorize.Zip = strCCZip
	objAuthorize.Country = strCCCountry

	objAuthorize.CustomerID = 1000
	objAuthorize.InvoiceNumber = 1000

	objAuthorize.Login = "OurPage"
	objAuthorize.Password = "hgf554jh"

	objAuthorize.CardNumber = strCCNumber
	objAuthorize.CardType = strCCType
	objAuthorize.ExpDate = intCCExpMonth & "/" & intCCExpYear

	objAuthorize.Amount = 20
	objAuthorize.TransType = "AUTH_ONLY"
	objAuthorize.Description = "GroupLoop.com monthly charge."

	objAuthorize.EmailMerchant = false
	objAuthorize.EmailCustomer = false

	' Initiate transation processing
	objAuthorize.Process

	strTransID = objAuthorize.TransID

	If objAuthorize.ErrorCode = 0 Then
		 ' Communication was successful.
		 ' Examine Results
		 If objAuthorize.ResponseCode <> 1 then
			strError = "Sorry, but there is the following problem with the card you entered:<br><font size='+1'>" & objAuthorize.ResponseReasonText & "</font><br><br>Make sure you have the correct address, card number, and expiration date."
		end if
	Else
		Select Case objAuthorize.ErrorCode
			Case -1
				strError = "Sorry, a connection could not be established with the authorization network.  Please try again."
			Case -2
				strError = "Sorry, a connection could not be established with the authorization network.  Please try again."
			Case Else
				strError = "Sorry, an unknown error occured with the authorization network.  Please try again, and notify support@keist.com if this keeps happening."
		End Select
	End If

	Set objAuthorize = Nothing

	if strError <> "" then Redirect("message.asp?Message=" & Server.URLEncode(strError))
End Sub


'------------------------End Code-----------------------------
%>

<!-- #include file="closedsn.asp" -->

<!-- #include file="footer.asp" -->