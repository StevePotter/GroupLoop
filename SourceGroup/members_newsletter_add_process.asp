<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeNewsletter ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))

Session.Timeout = 20
Server.ScriptTimeout = 5400
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Send a New Newsletter</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	strPath = GetPath("posts")
	upl.Path = strPath

	if not LoggedMember and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
	if not LoggedMember then Redirect("members.asp?Source=members_newsletter_add.asp")
	if not (LoggedAdmin or CBool( MeetingsMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))


	'Add the newsletter to the database
	if upl.Form("Subject") = "" or ( upl.Form("Body") = "" and upl.Form("File").IsEmpty ) then Redirect("incomplete.asp")

	strSubject = Format( upl.Form("Subject") )
	strBody = GetTextArea( upl.Form("Body") )

	'Save the file
	if upl.Form("File").IsEmpty then
		blFile = False		'This is just for the e-mail sendout
	else
		blFile = True
		'Get the extension and get the name to save as

		'Get rid of the directories and stuff, and get the extension
		strFileName = FormatFileName(Mid(upl.Form("File").UserFilename, InstrRev(upl.Form("File").UserFilename, "\") + 1))
		strExt = GetExtension(strFileName)

		'Make sure it isn't executable
		if lcase(strExt) = ".exe" or lcase(strExt) = ".asp" or lcase(strExt) = ".com" or lcase(strExt) = ".bat" then
			strError = strError & "You are trying to update an invalid type of file."
			Redirect "message.asp?Message=" & Server.URLEncode(strError)
		end if

		'We can't have duplicate file names in the folder, so keep adding numbers to the end
		intNum = 1
		do until not FileSystem.FileExists( strPath & strFileName )
			strFileName = GetJustFile( strFileName ) & intNum & "." & GetExtension( strFileName )
			intNum = intNum + 1
		loop

		'Save this badboy file
		upl.Form("File").SaveAs strFileName

	end if


	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddNewsletter"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh

		.Parameters("@CustomerID") = CustomerID
		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Subject") = strSubject
		.Parameters("@FileName") = strFileName

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ItemID")
	End With
	Set cmdTemp = Nothing



	Set rsNews = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT ID, Body FROM Newsletters WHERE ID = " & intID
	rsNews.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Update the fields

	rsNews("Body") = strBody

	rsNews.Update
	rsNews.Close

	'Open up all the members
	Query = "SELECT ID, FirstName, LastName, EMail1 FROM Members WHERE CustomerID = " & CustomerID & " AND SubscribeSiteNewsletter = 1 AND EMail1 <> ''"
	rsNews.CacheSize = 100
	rsNews.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if not rsNews.EOF then
		Set FirstName = rsNews("FirstName")
		Set LastName = rsNews("LastName")
		Set EMail = rsNews("EMail1")
	end if

	'Set up the mailer object
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.ContentType = "text/html"
	Mailer.IgnoreMalformedAddress = true
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = MailerFromName
	Mailer.FromAddress = "support@grouploop.com"
	Mailer.Subject    = FormatEdit(strSubject)
	'Members will be sent to the unsubscribe thing for members

	strHeader = "<html><title>" & strSubject & "</title><body>"
	strFooter = "</body></html>"
	'Put in links with the full URL for the e-mail programs
	strBody = Replace(strBody, "inserts", "http://www.GroupLoop.com/" & SubDirectory & "/inserts")

	if blFile then
		strBody = "<p><a href='" & NonSecurePath & "posts/" & strFileName & "'>Click here to view the meeting review.</a></p>" & strBody
	end if


	Mailer.BodyText   = strHeader & strBody & VbCrLf & VbCrLf & _
		"<p align=center>You received this newsletter because you signed up or someone else signed you up. &nbsp;" & _
		"To unsubscribe, simply <a href='" & NonSecurePath & "newsletter.asp?Submit=GoUnSubscribe'>click here</a>.</p>" & strFooter

	'Send to all the members, one at a time
	do until rsNews.EOF
		Mailer.ClearRecipients
		Mailer.AddRecipient FirstName & " " & LastName, EMail
		Mailer.SendMail

		rsNews.MoveNext
	loop

	rsNews.Close

	'Open up all the subscribers
	Query = "SELECT Name, EMail FROM NewsletterSubscribers WHERE CustomerID = " & CustomerID
	rsNews.CacheSize = 500
	rsNews.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if not rsNews.EOF then
		Set Name = rsNews("Name")
		Set EMail = rsNews("EMail")
	end if

	'Send to all the subscribers
	do until rsNews.EOF
		Mailer.ClearRecipients
		if Name <> "" then
			Mailer.AddRecipient Name, EMail
		else
			Mailer.AddRecipient EMail, EMail
		end if

		Mailer.ClearBodyText
		'They get a direct link to unsubscribe
		Mailer.BodyText   = strHeader & strBody & VbCrLf & VbCrLf & _
			"<p>To unsubscribe, <a href='" & NonSecurePath & "newsletter.asp?Submit=UnSubscribe&EMail=" & EMail & "'>click here</a>.</p>" & strFooter

		Mailer.SendMail

		rsNews.MoveNext
	loop
	rsNews.Close

	Set rsNews = Nothing
	Set Mailer = Nothing
	Set upl = Nothing
	Set FileSystem = Nothing
'------------------------End Code-----------------------------
%>
	<p>The newsletter has been sent out. &nbsp;<a href="members_newsletter_add.asp">Click here</a> to add another.<br>
	<a href="newsletter_read.asp?ID=<%=intID%>">Click here</a> to read it.
	</p>
