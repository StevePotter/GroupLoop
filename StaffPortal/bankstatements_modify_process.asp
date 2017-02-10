<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Bank Statements</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
Set upl = Server.CreateObject("SoftArtisans.FileUp")
strPath = GetPath ("posts")
upl.Path = strPath

if not LoggedStaff() then Redirect("login.asp?Source=bankstatements_modify.asp&ID=" & upl.form("ID") & "&Submit=" & upl.form("Submit"))
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

	intStatementID = CInt(upl.form("ID"))

	Query = "SELECT * FROM BankStatements WHERE ID = " & intStatementID
	Set rsAccount = Server.CreateObject("ADODB.Recordset")
	rsAccount.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	'upload this new file
	if not upl.Form("File").IsEmpty then
		if rsAccount("FileName") <> "" then if FileSystem.FileExists( GetPath("posts") & rsAccount("FileName") ) then FileSystem.DeleteFile( GetPath("posts") & rsAccount("FileName") )


		orgFileName = Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1)

		'Get rid of the directories and stuff, and get the extension
		strFileName = orgFileName
		strFileName = FormatFileName(strFileName)
		strExt = GetExtension(strFileName)
		strFileName = "bankstatements" & intStatementID & "." & strExt

		rsAccount("FileName") = strFileName
		rsAccount("OriginalFileName") = orgFileName


		upl.Form("File").SaveAs strFileName


	elseif upl.Form("UseFile") = "0" then
		if FileSystem.FileExists( GetPath("posts") & rsAccount("FileName") ) then FileSystem.DeleteFile( GetPath("posts") & rsAccount("FileName") )

		rsAccount("FileName") = ""
		rsAccount("OriginalFileName") = ""

	end if

	Set FileSytem = Nothing

	rsAccount("DateStarted") = AssembleDate("DateStarted")
	rsAccount("DateEnded") = AssembleDate("DateEnded")
	rsAccount("StartingBalance") = upl.Form("StartingBalance")
	rsAccount("EndingBalance") = upl.Form("EndingBalance")
	rsAccount("AccountID") = upl.Form("AccountID")
	rsAccount("Note") = upl.Form("Note")
	rsAccount("EmployeeID") = Session("EmployeeID")


	rsAccount.Update
	rsAccount.Close
	set rsAccount = Nothing
'------------------------End Code-----------------------------
%>
	<p>The statement has been edited. &nbsp;<a href="bankstatements_modify.asp">Click here</a> to modify another.</p>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->