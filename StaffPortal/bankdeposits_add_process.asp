<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\media_functions.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Deposit</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
Server.ScriptTimeout = 5400
if not LoggedStaff() then Redirect("login.asp?Source=bankdeposits_add.asp")
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------
	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	upl.Path = GetPath ("posts")

	'Create the new photo
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddBankDeposit"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Execute , , adExecuteNoRecords
		intDepositID = .Parameters("@ID")
	End With

	blProceed = true
	strError = ""

	strFileName = ""
	orgFileName = ""
	if not upl.Form("File").IsEmpty then
		Set FileSystem = CreateObject("Scripting.FileSystemObject")

		orgFileName = Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1)

		'Get rid of the directories and stuff, and get the extension
		strFileName = orgFileName
		strFileName = FormatFileName(strFileName)
		strExt = GetExtension(strFileName)
		strFileName = "bankdeposits" & intDepositID & "." & strExt

		upl.Form("File").SaveAs strFileName

		Set FileSytem = Nothing
	end if


	Query = "SELECT * FROM BankDeposits WHERE ID = " & intDepositID
	Set rsDeposit = Server.CreateObject("ADODB.Recordset")
	rsDeposit.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	rsDeposit("Total") = upl.Form("Total")
	if upl.Form("BankAccountID") <> "" then rsDeposit("BankAccountID") = upl.Form("BankAccountID")
	if upl.Form("BankStatementID") <> "" then rsDeposit("BankStatementID") = upl.Form("BankStatementID")
	rsDeposit("DateDeposited") = AssembleDate("DateDeposited")
	if upl.Form("CustomerID") <> "" then rsDeposit("CustomerID") = upl.Form("CustomerID")
	if upl.Form("InvoiceID") <> "" then rsDeposit("InvoiceID") = upl.Form("InvoiceID")
	if upl.Form("BillingType") <> "" then rsDeposit("BillingType") = upl.Form("BillingType")

	if upl.Form("CheckNum") <> "" then rsDeposit("CheckNum") = upl.Form("CheckNum")

	rsDeposit("Description") = Format(upl.Form("Description"))
	rsDeposit("CustomerNote") = Format(upl.Form("CustomerNote"))
	rsDeposit("StaffNote") = Format(upl.Form("StaffNote"))

	rsDeposit("FileName") = strFileName
	rsDeposit("OriginalFileName") = orgFileName
	rsDeposit("EmployeeID") = Session("EmployeeID")

	rsDeposit.Update
	rsDeposit.Close
	Set rsDeposit = Nothing


	Set upl = Nothing

'------------------------End Code-----------------------------
%>
	<p>The deposit has been added. <br>
	<a href="bankdeposits_add.asp">Add another.</a><br>
	<a href="bankdeposits_modify.asp">Modify deposits.</a>
	</p>

<%
'-----------------------Begin Code----------------------------

'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->