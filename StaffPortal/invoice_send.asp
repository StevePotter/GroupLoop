<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Send Invoice</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=invoice_send.asp&ID=" & Request("ID"))

if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the Invoice's ID."))
intInvoiceID = CInt(Request("ID"))

'Get the invoice
Set rsInvoice = Server.CreateObject("ADODB.Recordset")
Query = "SELECT * FROM CustomerInvoices WHERE ID = " & intInvoiceID
rsInvoice.CacheSize = 100
rsInvoice.Open Query, Connect, adOpenStatic, adLockOptimistic
if rsInvoice.EOF then
	set rsInvoice = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
elseif rsInvoice("InvoiceSent") = 1 then 
	set rsInvoice = Nothing
	Redirect("message.asp?Message=" & Server.URLEncode("The invoice has already been sent."))
end if

intCustomerID = rsInvoice("CustomerID")
BillingType = rsInvoice("BillingType")

'Get the command set up
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = Connect
cmd.CommandText = "GetSiteInfoRecordSet"
cmd.CommandType = adCmdStoredProc

'Get the customer
Set rsCust = Server.CreateObject("ADODB.Recordset")
rsCust.CacheSize = 100
rsCust.Open cmd, , adOpenStatic, adLockReadOnly, adCmdTableDirect
rsCust.Filter = "ID = " & intCustomerID
if rsCust.EOF then
	set rsCust = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The customer the invoice belongs to does not exist."))
end if

'Create the mailer object
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.ContentType = "text/html"
Mailer.RemoteHost  = "mail4.burlee.com"
Mailer.FromName    = "GroupLoop.com"
Mailer.FromAddress = "accounts@grouploop.com"


'Get the amount due
Query = "SELECT Date, Total, DateDeposited FROM BankDeposits WHERE InvoiceID = " & intInvoiceID
Set rsDeposits = Server.CreateObject("ADODB.Recordset")
rsDeposits.CacheSize = 100
rsDeposits.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

blDeposits = not rsDeposits.EOF   'If there are records, we have deposits, if there aren't, we don't

dblDue = rsInvoice("Total")

if blDeposits then
	do until rsDeposits.EOF
		dblDue = dblDue - rsDeposits("Total")
		rsDeposits.MoveNext
	loop

	rsDeposits.MoveFirst
end if




'Charge their credit card
if BillingType = "CreditCard" or BillingType = "Credit Card" then

	'Add the new deposit
	intDepositID = AddDeposit()

	strDescription = "Balance Due"
	ChargeCard dblDue, intDepositID, strDescription, blSuccess, true, strError, rsCust("ID"), strTransID

	if not blSuccess then
		'Didn't work, delete the deposit record
		DeleteDeposit intDepositID
		Response.Write "<b>CHARGE FOR " & FormatCurrency( dblDue ) & " COULD NOT BE MADE - TransID:" & strTransID & " - DepositID:" & intDepositID & "</b>  Error explained below.<br>" & strError & "<br>"
	else
		Query = "SELECT * FROM BankDeposits WHERE ID = " & intDepositID
		Set rsDeposit = Server.CreateObject("ADODB.Recordset")
		rsDeposit.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

		rsDeposit("CustomerID") = intCustomerID
		rsDeposit("InvoiceID") = intInvoiceID
		rsDeposit("Total") = dblDue
		dblCharge = dblDue
		rsDeposit("BillingType") = "CreditCard"
		rsDeposit("Description") = strDescription
		rsDeposit("TransID") = strTransID
		rsDeposit("StaffNote") = Format(Request("StaffNote"))
		strCustomerNote = Format(Request("CustomerNote"))
		rsDeposit("CustomerNote") = strCustomerNote
		rsDeposit("DateDeposited") = GetCurrentDate = FormatDateTime(Date, 2) 

		rsDeposit.Update
		rsDeposit.Close
		Set rsDeposit = Nothing


		Response.Write "Card successfully charged " & FormatCurrency( dblDue ) & " - TransID:" & strTransID & " - DepositID:" & intDepositID & "<br>"

		'They don't owe any money anymore
		dblDue = 0
	end if


end if


'Object that accesses web pages
Set Getter = Server.CreateObject("Getter.Get")
Getter.RequestTimeOut = 60	'Seconds
On Error Resume Next
PrintOut = Getter.OpenURL (cStr("http://www.GroupLoop.com/staff/invoice_print.asp?InvoiceID=" & intInvoiceID& "&NickName=" & Session("NickName") & "&Password=" & Session("Password")))
if Err.Number <> 0 then
	Response.Write "Getter Component Error: " &  err.Number & " " & Err.Description
end if
Set Getter = Nothing



'The body
strBody = "<p><i>Dear " & rsCust("FirstName") & ",</i><br>"

if CDbl(dblDue) > 0 then
	'The subject line
	strSubject = "Your GroupLoop.com Invoice (#" & intInvoiceID & ")"

	strBody = strBody & "Below is your latest invoice.  We appreciate your timely payment.  If you have any questions, please e-mail us at <a href='mailto:accounts@GroupLoop.com'>accounts@GroupLoop.com</a>.  " & _
	"Thank you for choosing GroupLoop!</p>"
else
	strSubject = "Receipt for GroupLoop.com Invoice (#" & intInvoiceID & ")"

	if BillingType = "CreditCard" or BillingType = "Credit Card" and blSuccess then
		strBody = strBody & "Below is the receipt for your latest invoice.  We charged " & FormatCurrency(dblCharge) & " to your credit card today.<br>" & _
		"The card charged was: " & Left(rsCust("CCNumber"), 4) & "..." & Right(rsCust("CCNumber"), 3) & "<br>" & _
		"If you have any questions, please e-mail us at <a href='mailto:accounts@GroupLoop.com'>accounts@GroupLoop.com</a>.  " & _
		"Thank you for choosing GroupLoop!</p>"

	else
		strBody = strBody & "Below is the receipt for your latest invoice.  Since you have already paid the invoice, <b>this is not a bill</b>.  This is simply for your records.  If you have any questions, please e-mail us at <a href='mailto:accounts@GroupLoop.com'>accounts@GroupLoop.com</a>.  " & _
		"Thank you for choosing GroupLoop!</p>"

	end if

end if

strBody = strBody & PrintOut


Mailer.Subject    = strSubject
Mailer.ClearBodyText
Mailer.BodyText   = strBody
Mailer.ClearRecipients
Mailer.IgnoreMalformedAddress = true

strRecip = rsCust("FirstName")&" "&rsCust("LastName")
strEmail1 = rsCust("EMail")
Mailer.AddRecipient strRecip, strEmail1
strEmail2 = rsCust("EMail1")
if strEmail2 <> "" and strEmail1 <> strEmail2 then Mailer.AddRecipient strRecip, strEmail2
strEmail3 = rsCust("EMail2")
if strEmail3 <> "" and strEmail1 <> strEmail3 and strEmail2 <> strEmail3 then Mailer.AddRecipient strRecip, strEmail3

if not Mailer.SendMail then 
	Response.Write "<b>There has been an error, and the email has not been sent to " & rsCust("FirstName") & " " & rsCust("LastName") & ", <a href=mailto:" & rsCust("EMail") & ">" & rsCust("EMail") & "</a>, CustID: " & rsCust("ID") & ", Total:" & FormatCurrency(dblChargeTotal) & ", InvoiceID:" & intInvoiceID & ".</b>  <br>Error was '" & Mailer.Response & "'<br>"
end if
Mailer.ClearRecipients
Mailer.AddRecipient "GroupLoop Accounts", "accounts@grouploop.com"
Mailer.SendMail

'Mark it sent
rsInvoice("InvoiceSent") = 1
rsInvoice.Update
rsInvoice.Close

Set rsInvoice = Nothing
Set cmd = Nothing
Set rsCust = Nothing
Set Mailer = Nothing
%>
<p>The invoice has been sent.<br>
<a href="invoice_print.asp?InvoiceID=<%=intInvoiceID%>">Print the invoice.</a><br>
<a href="invoices_modify.asp?Submit=Edit&ID=<%=intInvoiceID%>">Edit the invoice.</a><br>
<a href="customer_view.asp?ID=<%=intCustomerID%>">View this customer's details.</a><br>
</p>

<%



Function AddDeposit()
	With cmd
		'Get the customer's info
		.CommandText = "AddBankDeposit"
		.Parameters.Refresh
		.Execute , , adExecuteNoRecords

		intDepositID = .Parameters("@ID")
	End With

	AddDeposit = intDepositID
End Function


Sub DeleteDeposit(intDepositID)
	With cmd
		'Get the customer's info
		.CommandText = "DeleteBankDeposit"
		.Parameters.Refresh
		.Parameters("@ID") = intDepositID
		.Execute , , adExecuteNoRecords

	End With

End Sub


Sub ChargeCard( dblDue, intDepositID, strDescription, blSuccess, blSendErrorEMail, strError, CustID, strTransID )
	strError = ""
	blSuccess = true

	strFirstName = rsCust("CCFirstName")
	if IsNull(strFirstName) then strFirstName = ""
	strLastName = rsCust("CCLastName")
	if IsNull(strLastName) then strLastName = ""
	strCompany = rsCust("CCCompany")
	if IsNull(strCompany) then strCompany = ""

	strCCNumber = rsCust("CCNumber")
	strCCType = rsCust("CCType")
	CCExpDate = rsCust("CCExpdate")

	strEMail = rsCust("EMail")
	if IsNull(strEMail) then strEMail = ""

	strBillStreet1 = rsCust("BillingStreet1")

	strBillStreet2 = rsCust("BillingStreet2")
	if IsNull(strBillStreet2) then strBillStreet2 = ""

	strBillCity = rsCust("BillingCity")
	if IsNull(strBillCity) then strBillCity = ""

	strBillState = rsCust("BillingState")
	if IsNull(strBillState) then strBillState = ""

	strBillZip = rsCust("BillingZip")
	if IsNull(strBillZip) then strBillZip = ""

	strBillCountry = rsCust("BillingCountry")
	if IsNull(strBillCountry) then strBillCountry = ""

	strPhone = rsCust("BillingPhone")
	if IsNull(strPhone) then strPhone = ""


	Set objAuthorize = Server.CreateObject("xAuthorize.Process")
	' Initialize xAuthorize for a new transaction.
	objAuthorize.Initialize

	' Set object properties.
	objAuthorize.Processor = "AUTHORIZE_NET"

	objAuthorize.FirstName = strFirstName
	objAuthorize.LastName = strLastName
	objAuthorize.Company = strCompany
	objAuthorize.Address = strBillStreet1 & " " & strBillStreet2
	objAuthorize.City = strBillCity
	objAuthorize.State = strBillState
	objAuthorize.Zip = strBillZip
	objAuthorize.Country = "USA"
	objAuthorize.Phone = strPhone
	objAuthorize.EMail = strEMail

	objAuthorize.CustomerID = CustID
	objAuthorize.InvoiceNumber = intDepositID

	objAuthorize.Login = "OurPage"
	objAuthorize.Password = "hgf554jh"

	objAuthorize.CardNumber = strCCNumber
	objAuthorize.CardType = strCCType
	objAuthorize.ExpDate = Month(CCExpDate) & "/" & Year(CCExpDate)

	'Charge the amount from the order!!!  Add the description too
	objAuthorize.Amount = dblDue
	strTransType = "AUTH_CAPTURE"
	objAuthorize.TransType = strTransType
	objAuthorize.Description = strDescription

	objAuthorize.EmailMerchant = false
	objAuthorize.EmailCustomer = false

	' Initiate transation processing
	objAuthorize.Process

	strTransID = ""

	strTransID = objAuthorize.TransID

	If objAuthorize.ErrorCode = 0 Then
		 ' Communication was successful.
		 ' Examine Results
		 If objAuthorize.ResponseCode <> 1 then
			strError = PrintIndent & "Problem while trying to transact with account credit card (" & Left(strCCNumber, 4) & "..." & Right(strCCNumber, 4) & "):<br>" & PrintIndent & objAuthorize.ResponseReasonText
		end if
	Else
		Select Case objAuthorize.ErrorCode
			Case -1
				Response.Write "A connection could not be established with the authorization network.  Our support personnel have been notified, and will personally handle the matter.<br>"
			Case -2
				Response.Write "A connection could not be established with the authorization network.  Our support personnel have been notified, and will personally handle the matter.<br>"
			Case Else
				Response.Write "An unknown error occured with the authorization network.  Our support personnel have been notified, and will personally handle the matter.<br>"
		End Select
	End If


	if strError <> "" then
		blSuccess = false

		if blSendErrorEMail = "" then blSendErrorEMail = true

		if blSendErrorEMail = true then

			'The subject line
			strSubject = "Credit Card charge problem! CustomerID " & CustID & ", TransID " & strTransID
			'The body
			strBody = "Contact information:<br>" & _
				strFirstName & " " & strLastName & "<br>" & _
				strPhone & "<br>" & _
				"<a href=mailto:" & strEMail & ">" & strEMail & "</a><br><br>" & _
				"Transaction information:<br>" & _
				"Card Number: " & strCCNumber & "<br>" & _
				"Card Type: " & strCCType & "<br>" & _
				"Card Exp: " & Month(CCExpDate) & "/" & Year(CCExpDate) & "<br>" & _
				"TransType: " & strTransType & "<br>" & _
				"Total: " & dblDue & "<br>" & _
				"TransID: " & strTransID & "<br>" & _
				"Description: " & strDescription & "<br>" & _
				"Error given: <b>" & strError & "<br>"

			'Set the rest of the mailing info and send it
			Mailer.Subject    = strSubject
			Mailer.ClearBodyText
			Mailer.BodyText   = strBody
			Mailer.ClearRecipients
			Mailer.AddRecipient "Credit Card Support", "support@grouploop.com"
			Mailer.SendMail


		end if
	end if

	Set objAuthorize = Nothing


End Sub

%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->













