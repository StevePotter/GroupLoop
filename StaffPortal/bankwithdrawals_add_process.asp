<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\media_functions.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Withdrawal</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
Server.ScriptTimeout = 5400
if not LoggedStaff() then Redirect("login.asp?Source=bankwithdrawals_add.asp")
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------
	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	upl.Path = GetPath ("posts")

	'Create the new photo
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddBankWithdrawal"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Execute , , adExecuteNoRecords
		intWithdrawalID = .Parameters("@ID")
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
		strFileName = "bankwithdrawals" & intWithdrawalID & "." & strExt

		upl.Form("File").SaveAs strFileName

		Set FileSytem = Nothing
	end if

	Query = "SELECT * FROM BankWithdrawals WHERE ID = " & intWithdrawalID
	Set rsLook = Server.CreateObject("ADODB.Recordset")
	rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic

	if upl.Form("BankAccountID") <> "" then rsLook("BankAccountID") = upl.Form("BankAccountID")
	if upl.Form("BankStatementID") <> "" then rsLook("BankStatementID") = upl.Form("BankStatementID")
	rsLook("Date") = AssembleDate("Date")

	rsLook("Total") = upl.Form("Total")
	rsLook("PaidTo") = upl.Form("PaidTo")
	rsLook("PaymentType") = upl.Form("PaymentType")
	if upl.Form("CheckNum") <> "" then rsLook("CheckNum") = upl.Form("CheckNum")
	if upl.Form("InvoiceReceivedID") <> "" then rsLook("InvoiceReceivedID") = upl.Form("InvoiceReceivedID")
	rsLook("Description") = Format(upl.Form("Description"))
	rsLook("StaffNote") = Format(upl.Form("StaffNote"))
	rsLook("FileName") = strFileName
	rsLook("OriginalFileName") = orgFileName
	rsLook("EmployeeID") = Session("EmployeeID")

	rsLook.Update
	Set rsLook = Nothing

	Set upl = Nothing

'------------------------End Code-----------------------------
%>
	<p>The withdrawal has been added. <br>
	<a href="bankwithdrawals_add.asp">Add another.</a><br>
	<a href="bankwithdrawals_modify.asp">Modify withdrawals.</a>
	</p>

<%
'-----------------------Begin Code----------------------------

'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->