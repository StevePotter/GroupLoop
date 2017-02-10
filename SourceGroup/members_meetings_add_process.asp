<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeMeetings ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
Session.Timeout = 20
Server.ScriptTimeout = 5400
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Add A Meeting</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	strPath = GetPath("posts")
	upl.Path = strPath

	if not LoggedMember and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
	if not LoggedMember then Redirect("members.asp?Source=members_meetings_add.asp")
	if not (LoggedAdmin or CBool( MeetingsMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))

	'Add the meeting to the database
	if upl.Form("Subject") = "" or ( upl.Form("Body") = "" and upl.Form("File").IsEmpty ) then Redirect("incomplete.asp")
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT ID, Private, MemberID, Subject, Body, CustomerID, IP, ModifiedID, CommitteeID, FileName, FileLinkDirect FROM Meetings"
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Update the fields
	rsNew.AddNew
		if upl.Form("Private") = "1" then 
			rsNew("Private") = 1
		else
			rsNew("Private") = 0
		end if

		rsNew("FileLinkDirect") = GetCheckedResult(upl.Form("FileLinkDirect"))

		rsNew("MemberID") = Session("MemberID")
		rsNew("ModifiedID") = Session("MemberID")
		rsNew("Subject") = Format( upl.Form("Subject") )
		rsNew("Body") = GetTextArea( upl.Form("Body") )
		rsNew("CustomerID") = CustomerID
		rsNew("IP") = Request.ServerVariables("REMOTE_HOST")
		if IncludeCommittees = 1 then rsNew("CommitteeID") = CInt(upl.Form("CommitteeID"))


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
		end if

		'Now make sure we still don't have a problem
		if strError <> "" then
			Set upl = Nothing
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

		rsNew("FileName") = strFileName
	end if

	rsNew.Update
	rsNew.MoveNext
	rsNew.MovePrevious
	intID = rsNew("ID")
	rsNew.Close
	Set rsNew = Nothing





	if upl.Form("EMail") = "1" then
		'Open up all the members
		Query = "SELECT ID, FirstName, LastName, EMail1 FROM Members WHERE CustomerID = " & CustomerID & " AND EMail1 <> '' AND SubscribeMeetings = 1"
		Set rsMembers = Server.CreateObject("ADODB.Recordset")
		rsMembers.CacheSize = 50
		rsMembers.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

		Set FirstName = rsMembers("FirstName")
		Set LastName = rsMembers("LastName")
		Set EMail = rsMembers("EMail1")

		'Get the author's info
		rsMembers.Filter = "ID = " & Session("MemberID")

		strAuthor = FirstName & " " & LastName
		strAuthorEMail = EMail

		rsMembers.Filter = ""

		strSubject = Title & " - meeting review by " & strAuthor & " - " & upl.Form("Subject")

		strBody = "This meeting review was automatically sent to you by " & strAuthor & "'s request.  Please do not respond to this e-mail.  " & _
			"You may reach " & strAuthor & " at " & strAuthorEMail & "<br><br>" 

		if upl.Form("CommitteeID") <> "" then
			strBody = strBody & "This meeting was held by the " & GetCommittee( upl.Form("CommitteeID") ) & " committee.<br><br>"
		end if

		if blFile then
			strBody = strBody & "<p><a href='" & NonSecurePath & "posts/" & strFileName & "'>Click here to view the meeting review.</a></p>"
		end if

		strBody = strBody &	GetTextArea(upl.Form("Body"))

		'Set the rest of the mailing info and send it
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.ContentType = "text/html"
		Mailer.IgnoreMalformedAddress = true
		Mailer.RemoteHost  = "mail4.burlee.com"
		Mailer.FromName    = MailerFromName
		Mailer.FromAddress = "support@grouploop.com"
		Mailer.Subject    = FormatEdit(strSubject)
		Mailer.BodyText   = strBody

		do until rsMembers.EOF
			Mailer.ClearRecipients
			Mailer.AddRecipient FirstName & " " & LastName, EMail
			Mailer.SendMail

			rsMembers.MoveNext
		loop

		rsMembers.Close
		Set rsMembers = Nothing

		Set Mailer = Nothing
	end if


	Set FileSystem = Nothing
	Set upl = Nothing
'------------------------End Code-----------------------------
%>
	<p>Your meeting has been added. &nbsp;<a href="members_meetings_add.asp">Click here</a> to add another.<br>
	<a href="meetings_read.asp?ID=<%=intID%>">Click here</a> to view it.
	</p>
