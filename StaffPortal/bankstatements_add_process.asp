<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\media_functions.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Statement</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
Server.ScriptTimeout = 5400
if not LoggedStaff() then Redirect("login.asp?Source=bankstatements_add.asp")
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------
	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	upl.Path = GetPath ("posts")

	'Create the new photo
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddBankStatement"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Execute , , adExecuteNoRecords
		intStatementID = .Parameters("@ItemID")
	End With

	blProceed = true
	strError = ""

	strFileName = ""
	if not upl.Form("File").IsEmpty then
		Set FileSystem = CreateObject("Scripting.FileSystemObject")

		orgFileName = Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1)

		'Get rid of the directories and stuff, and get the extension
		strFileName = orgFileName
		strFileName = FormatFileName(strFileName)
		strExt = GetExtension(strFileName)
		strFileName = "bankstatements" & intStatementID & "." & strExt

		upl.Form("File").SaveAs strFileName

		Set FileSytem = Nothing
	end if

	Query = "SELECT * FROM BankStatements WHERE ID = " & intStatementID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic

	rsLook("DateStarted") = AssembleDate("DateStarted")
	rsLook("DateEnded") = AssembleDate("DateEnded")
	rsLook("StartingBalance") = upl.Form("StartingBalance")
	rsLook("EndingBalance") = upl.Form("EndingBalance")
	rsLook("AccountID") = upl.Form("AccountID")
	rsLook("Note") = upl.Form("Note")
	rsLook("FileName") = strFileName
	rsLook("OriginalFileName") = orgFileName
	rsLook("EmployeeID") = Session("EmployeeID")

	rsLook.Update
	Set rsLook = Nothing

	Set upl = Nothing

'------------------------End Code-----------------------------
%>
	<p>The statement has been added. <br>
	<a href="bankstatements_add.asp">Add another.</a><br>
	<a href="bankstatements_modify.asp">Modify statements.</a>
	</p>

<%
'-----------------------Begin Code----------------------------

'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->