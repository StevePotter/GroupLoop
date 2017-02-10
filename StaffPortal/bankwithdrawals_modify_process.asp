<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Bank Withdrawals</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
Set upl = Server.CreateObject("SoftArtisans.FileUp")
strPath = GetPath ("posts")
upl.Path = strPath

if not LoggedStaff() then Redirect("login.asp?Source=bankwithdrawals_modify.asp&ID=" & upl.form("ID") & "&Submit=" & upl.form("Submit"))
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

	intWithdrawalID = CInt(upl.form("ID"))

	Query = "SELECT * FROM BankWithdrawals WHERE ID = " & intWithdrawalID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	'upload this new file
	if not upl.Form("File").IsEmpty then
		if rsUpdate("FileName") <> "" then if FileSystem.FileExists( GetPath("posts") & rsUpdate("FileName") ) then FileSystem.DeleteFile( GetPath("posts") & rsUpdate("FileName") )


		orgFileName = Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1)

		'Get rid of the directories and stuff, and get the extension
		strFileName = orgFileName
		strFileName = FormatFileName(strFileName)
		strExt = GetExtension(strFileName)
		strFileName = "bankwithdrawals" & intWithdrawalID & "." & strExt

		rsUpdate("FileName") = strFileName
		rsUpdate("OriginalFileName") = orgFileName


		upl.Form("File").SaveAs strFileName


	elseif upl.Form("UseFile") = "0" then
		if FileSystem.FileExists( GetPath("posts") & rsUpdate("FileName") ) then FileSystem.DeleteFile( GetPath("posts") & rsUpdate("FileName") )

		rsUpdate("FileName") = ""
		rsUpdate("OriginalFileName") = ""

	end if

	Set FileSytem = Nothing

	if upl.Form("BankAccountID") <> "" then rsUpdate("BankAccountID") = upl.Form("BankAccountID")
	if upl.Form("BankStatementID") <> "" then rsUpdate("BankStatementID") = upl.Form("BankStatementID")
	rsUpdate("Date") = AssembleDate("Date")

	rsUpdate("Total") = upl.Form("Total")
	rsUpdate("PaidTo") = upl.Form("PaidTo")
	rsUpdate("PaymentType") = upl.Form("PaymentType")
	if upl.Form("CheckNum") <> "" then rsUpdate("CheckNum") = upl.Form("CheckNum")
	if upl.Form("InvoiceReceivedID") <> "" then rsUpdate("InvoiceReceivedID") = upl.Form("InvoiceReceivedID")
	rsUpdate("Description") = Format(upl.Form("Description"))
	rsUpdate("StaffNote") = Format(upl.Form("StaffNote"))
	rsUpdate("EmployeeID") = Session("EmployeeID")


	rsUpdate.Update
	rsUpdate.Close
	set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The withdrawal has been edited. &nbsp;<a href="bankwithdrawals_modify.asp">Click here</a> to modify another.</p>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->