<!-- #include file="dsn.asp" -->
<!-- #include file="..\sourcecommon\functions.asp" -->
<%
strSource = "<html><head></head><body>"

'Keep track of a few things
intCardCharges = 0
intInvoices = 0
intErrors = 0
intExpirationWarnings = 0

'The command object used throughout the script
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = Connect
cmd.CommandType = adCmdStoredProc

'Object that accesses web pages
Set Getter = Server.CreateObject("Getter.Get")
Getter.RequestTimeOut = 60	'Seconds



'Special instructions passed to script
'Custom date
if Request("Date") = "" then
	ProcessDate = Date
else
	ProcessDate = CDate(Request("Date"))
end if

'Send expiration warnings or not
if Request("Warnings") = "NO" then
	blExpWarnings = false
else
	blExpWarnings = true
end if

'Process customer charges
if Request("Charges") = "NO" then
	blProcessCharges = false
else
	blProcessCharges = true
end if

'Write the index file or not
if Request("WriteIndex") = "NO" then
	blWriteIndex = false
else
	blWriteIndex = true
end if

'Write the index file or not
if Request("SendEMail") = "NO" then
	blSendEMail = false
else
	blSendEMail = true
end if

'Write the Constants file or not
if Request("WriteConstants") = "YES" then
	blWriteConstants = true
else
	blWriteConstants = false
end if

'Write the HeaderFooter file or not
if Request("WriteHeaderFooter") = "YES" then
	blWriteHeaderFooter = true
else
	blWriteHeaderFooter = false
end if

Public Border, Cellspacing, Cellpadding

Set rsPage = Server.CreateObject("ADODB.Recordset")

if RunAlready(ProcessDate) and Request("Override") <> "YES" then
	Set rsPage = Nothing
	Response.Write "Already been run today - " & ProcessDate
else
	'Add the record
	intMaintenanceID = AddMaintenance()

	'Get rid of the shit we don't need...
	DeleteInvalid

	Set rsNew = Server.CreateObject("ADODB.Recordset")

	' Create xAuthorize object.
	Set objAuthorize = Server.CreateObject("xAuthorize.Process")

	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.ContentType = "text/html"
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = "GroupLoop.com"
	Mailer.FromAddress = "support@grouploop.com"

	Set rsSite = Server.CreateObject("ADODB.Recordset")
	Set rsBDays = Server.CreateObject("ADODB.Recordset")
	Set rsCalendar = Server.CreateObject("ADODB.Recordset")
	Set rsBilling = Server.CreateObject("ADODB.Recordset")
	Set rsConfig = Server.CreateObject("ADODB.Recordset")
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")


	'get the customer recordset
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetSiteInfoRecordSet"
	cmdTemp.CommandType = adCmdStoredProc

	Set rsCust = Server.CreateObject("ADODB.Recordset")
	rsCust.CacheSize = 100
	rsCust.Open cmdTemp, , adOpenStatic, adLockOptimistic, adCmdTableDirect

	Set cmdTemp = Nothing




	intCustomerID = rsCust("ID")
	'Put the system date into the shortened format (drop the time)
	DayToday = Day(ProcessDate)

	strSource = strSource & "<p><b><u>Beginning nightly maintenence (ID # " & intMaintenanceID & ") at " & FormatDateTime( ProcessDate, 1 ) & " " & Time & ".</u></b></p>"

	do until rsCust.EOF
		intCustomerID = rsCust("ID")
		CustomerID = intCustomerID

		'Print out header
		strSource = strSource & "Customer ID: " & CustomerID & PrintIndent & "Version: " & rsCust("Version") & PrintIndent & "Name: " & rsCust("FirstName") & " "  & rsCust("LastName") & PrintIndent & "Address: <a href=http://www.GroupLoop.com/" & rsCust("SubDirectory") & ">" & "http://www.GroupLoop.com/" & rsCust("SubDirectory") & "</a><br>"

		'If they are a free site, give them a warning that their site is going to expire
		if blExpWarnings then SendExpirationWarnings

		'Make any charges to their account
		if blProcessCharges then MonthlyCharges


		'Write the current index file (this is here for birthdays and calendar events)
		if blWriteIndex then WriteIndex

		'Write the current constants file
		if blWriteConstants then WriteConstants

		'Write the current HeaderFooter file
		if blWriteHeaderFooter then WriteHeaderFooter


		rsCust.MoveNext

	loop
	rsCust.Close

	strSource = strSource & "Done!</body></html>"

	'Send output to screen
	Response.Write strSource

	'Put this into record
	UpdateMaintenance intMaintenanceID

	set rsCust = Nothing
	Set rsPage = Nothing
	Set rsBDays = Nothing
	Set rsCalendar = Nothing
	Set rsBilling = Nothing
	Set rsConfig = Nothing

	Set objAuthorize = Nothing
	Set cmd = Nothing
	Set FileSystem = Nothing
	Set Mailer = Nothing
	Set Getter = Nothing

end if

'------------------------End Code-----------------------------

'-------------------------------------------------------------
'This function deletes a customer once their site has expired
'-------------------------------------------------------------
Function DeleteCustomer( intCustomerID )

		strReturn = ""

		Set Command = Server.CreateObject("ADODB.Command")
		With Command
			'Get the subdirectory
			.ActiveConnection = Connect
			.CommandText = "GetCustomerInfo"
			.CommandType = adCmdStoredProc
			.Parameters.Refresh
			.Parameters("@CustomerID") = intCustomerID
			.Execute , , adExecuteNoRecords
			strSubDir = .Parameters("@SubDirectory")

			.CommandText = "DeleteCustomer"
			.Parameters.Refresh
			.Parameters("@CustomerID") = intCustomerID
			.Execute , , adExecuteNoRecords
			strReturn = strReturn &  "1. Removing from database<br>"
		End With
		Set Command = Nothing

		if strSubDir = "" then Redirect("error.asp?Message=" & Server.URLEncode("The customer has been deleted from the database, but the directory didn't exist."))
		strFolder = Server.MapPath("../" & strSubDir)

		if strFolder = "E:\Webs\Websites\ourclubpage.com" then
			Redirect "error.asp"
		end if

		Set FSys = CreateObject("Scripting.FileSystemObject")
		if FSys.FolderExists(strFolder) then
			strReturn = strReturn &  "2. Removing folder: " & strFolder & "<br>"
			FSys.DeleteFolder strFolder
		else
			strReturn = strReturn &  "2. FOLDER NOT FOUND - " & strFolder & "<br>"
		end if
		Set FSys = Nothing

		DeleteCustomer = strReturn
End Function


'-------------------------------------------------------------
'This sub starts up the automatic billing system
'-------------------------------------------------------------
Sub MonthlyCharges()
	'Get the customer's signup day into the shortened format
	custDate = FormatDateTime(rsCust("SignupDate"), 2)
	custDay = Day(custDate)

	'don't charge people over the 28th (avoids people that
	if DayToday <= 28 then
		'If the days match up or it's 28th and the customer signed up after the 28th
		if rsCust("FreeSite") = 0 and rsCust("Version") <> "Free" and (DayToday = custDay or (DayToday = 28 and custDay > 28)) then
			strSource = strSource & PrintIndent & "Customer up for charge.<br>"

			'Returns 0 if not charged so far this month, and the invoiceID if there has been a charge this month
			intChargedAlready = ChargedAlready()

			if intChargedAlready > 0 then
				intErrors = intErrors + 1
				strSource = strSource & PrintIndent & "Customer has already been charged this month (Invoice # " & intChargedAlready & ")<br>"
			else

				'If they have an existing invoice, we keep adding to it
				intInvoiceID = NeedNewInvoice()
				if intInvoiceID = 0 then intInvoiceID = AddInvoice("Monthly GroupLoop Charge", intMaintenanceID)


				'We gotta charge diss bitch...
				if rsCust("Version") = "Parent" then
					CreateInvoiceCharges intInvoiceID, intMaintenanceID


				'Gold site...
				elseif rsCust("Version") = "Gold" then
					CreateGoldSiteAdditional intInvoiceID, intMaintenanceID
					CreateInvoiceCharges intInvoiceID, intMaintenanceID
				else
					CreateInvoiceCharges intInvoiceID, intMaintenanceID
				end if

				dblChargeTotal = GetInvoiceTotal( intInvoiceID )


				if dblChargeTotal > 0 then
					if TimeToCharge(intInvoiceID) then ChargeCustomer dblChargeTotal, intInvoiceID
				else
					DeleteInvoice intInvoiceID
				end if

			end if

		end if
	end if


End Sub


'-------------------------------------------------------------
'This sub charges a certain amount to a customer for a given invoice
'-------------------------------------------------------------
Sub ChargeCustomer( dblChargeTotal, intInvoiceID )
	strSubject = ""

	intInvoices = intInvoices + 1

	if rsCust("BillingType") = "CreditCard" then

		ChargeCard dblChargeTotal, intInvoiceID, "Monthly GroupLoop Fee", blSuccess, true, strError, rsCust("ID"), strTransID

		intCardCharges = intCardCharges + 1

		if not blSuccess then
			UpdateInvoice intInvoiceID, dblChargeTotal, rsCust("BillingType"), 0, NULL, Date, 0
			strSource = strSource & PrintIndent & "<b>CHARGE FOR " & FormatCurrency( dblChargeTotal ) & " COULD NOT BE MADE - TransID:" & strTransID & " - InvoiceID:" & intInvoiceID & "</b>  Error explained below.<br>" & strError & "<br>"

			intErrors = intErrors + 1

			'The subject line
			strSubject = "Problem with the credit card used for your GroupLoop site."

			CCExpDate = rsCust("CCExpDate")

			'The body
			strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
			"<p><i>Dear " & rsCust("FirstName") & ",</i><br>" & _
			" &nbsp;&nbsp;We attempted to charge " & FormatCurrency( dblChargeTotal ) & " to your credit card for your GroupLoop site.<br>" & _
			"The card used was: " & Left(rsCust("CCNumber"), 4) & "..." & Right(rsCust("CCNumber"), 3) & "<br>" & _
			"Card Type: " & rsCust("CCType") & "<br>" & _
			"Card Expiration Date: " & Month(CCExpDate) & "/" & Year(CCExpDate) & "<br>" & _
			"Street: " & rsCust("BillingStreet1") & "<br>" & _
			"City: " & rsCust("BillingCity") & "<br>" & _
			"State: " & rsCust("BillingState") & "<br>" & _
			"Country: " & rsCust("BillingCountry") & "</p>" & _

			"<p>If the card information has changed, or you would like to use another credit card, please " & _
			"<a href='https://www.OurClubPage.com/admin/account_edit.asp?CustomerID=" & rsCust("ID") & "'>CLICK HERE</a>. <br>" & _
			"Please correct this right away.  We will try several more times, and if we still cannot collect payment, we will be " & _
			"forced to end your account.</p>" & _

			"<p> &nbsp;&nbsp;We would like to thank you for signing up with us.  It is our mission to constantly make your site better, easier, and faster.  " & _
			"I hope you will stick with us, and enjoy your site as long as possible!</p>" & _
			"<p><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf
		else
			strSource = strSource & PrintIndent & "Card successfully charged " & FormatCurrency( dblChargeTotal ) & " - TransID:" & strTransID & " - InvoiceID:" & intInvoiceID & "<br>"

			UpdateInvoice intInvoiceID, dblChargeTotal, rsCust("BillingType"), 1, Date, Date, 1
			AddDeposit rsCust("ID"), intInvoiceID, dblChargeTotal, strTransID

			'The subject line
			strSubject = "Your monthly GroupLoop charge receipt."

			'The body
			strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
			"<p><i>Dear " & rsCust("FirstName") & ",</i><br>" & _
			" &nbsp;&nbsp;Your credit card has been charged " & FormatCurrency( dblChargeTotal ) & " for your GroupLoop site.<br>" & _
			"The card charged was: " & Left(rsCust("CCNumber"), 4) & "..." & Right(rsCust("CCNumber"), 3) & "</p>" & _
			"<p> &nbsp;&nbsp;We would like to thank you for signing up with us.  It is our mission to constantly make your site better, easier, and faster.  " & _
			"I hope you continue to enjoy your site and use it as much as you can!</p>" & _
			"<p><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf
		end if
	else
		UpdateInvoice intInvoiceID, dblChargeTotal, rsCust("BillingType"), 0, NULL, Date, 1

		strSource = strSource & PrintIndent & "Sending invoice for " & FormatCurrency( dblChargeTotal ) & " to be paid by check - InvoiceID:" & intInvoiceID & "<br>"


		'The subject line
		strSubject = "Your monthly GroupLoop charge invoice (# " & intInvoiceID & ")."

		CutoffDate = DateAdd("d", (7), ProcessDate)

		'The body
		strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
		"<p><i>Dear " & rsCust("FirstName") & ",</i><br>" & _
		" &nbsp;&nbsp;Our records show that you have a payment due by " & FormatDateTime( CutoffDate, 2  ) & " for your GroupLoop site.<br>" & _
		"The total to send is: <b>" & FormatCurrency(dblChargeTotal) & "</b></p>" & _
		"<p><i>Please make the check out to:</i><br>" & _
		"GroupLoop.com<br>" & _
		"P.O. Box 5271<br>" & _
		"Somerville, NJ 08876-3430</p>" & _

		"<p> &nbsp;&nbsp;We appreciate your prompt payment, and would like to thank you for signing up with us.  It is our mission to constantly make your site better, easier, and faster.  " & _
		"I hope you continue to enjoy your site and use it as much as you can!</p>" & _
			"<p><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
			"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf

	end if


	if strSubject <> "" then
		'Set the rest of the mailing info and send it
		Mailer.Subject    = strSubject
		Mailer.ClearBodyText
		Mailer.BodyText   = strBody
		Mailer.ClearRecipients
		Mailer.AddRecipient rsCust("FirstName")&" "&rsCust("LastName"), rsCust("EMail")
		if not Mailer.SendMail then 
			intErrors = intErrors + 1
			strSource = strSource & PrintIndent & "<b>There has been an error, and the email has not been sent to " & rsCust("FirstName") & " " & rsCust("LastName") & ", <a href=mailto:" & rsCust("EMail") & ">" & rsCust("EMail") & "</a>, CustID: " & rsCust("ID") & ", Total:" & FormatCurrency(dblChargeTotal) & ", InvoiceID:" & intInvoiceID & ".</b>  <br>Error was '" & PrintIndent & Mailer.Response & "'<br>"
		end if
		Mailer.ClearRecipients
		Mailer.AddRecipient "Monthly Billing", "support@grouploop.com"
	end if

End Sub


'-------------------------------------------------------------
'This sub adds an invoice
'-------------------------------------------------------------
Function AddInvoice(strDescription, intMaintenanceID)
	With cmd
		'Get the customer's info
		.CommandText = "AddCustomerInvoice"
		.Parameters.Refresh
		.Parameters("@CustomerID") = rsCust("ID")
		.Parameters("@Description") = strDescription
		.Parameters("@MaintenanceID") = intMaintenanceID
		.Parameters("@Sent") = 0
		.Execute , , adExecuteNoRecords

		intID = .Parameters("@InvoiceID")
	End With

	strSource = strSource & PrintIndent & "New invoice created, #" & intID & "<br>"

	AddInvoice = intID


End Function


'-------------------------------------------------------------
'This sub opens a web site
'-------------------------------------------------------------
Sub GetURL (strURL)
	On Error Resume Next
	Getter.OpenURL (cStr(strURL))
	if Err.Number <> 0 then
		strSource = strSource & PrintIndent &  "Getter Component Error: " &  err.Number & " " & Err.Description
	end if
End Sub



'-------------------------------------------------------------
'This sub adds an invoice
'-------------------------------------------------------------
Function NeedNewInvoice()
	With cmd
		'Get the customer's info
		.CommandText = "CustomerNeedsInvoice"
		.Parameters.Refresh
		.Parameters("@CustomerID") = rsCust("ID")
		.Execute , , adExecuteNoRecords

		intNeedID = .Parameters("@InvoiceID")
	End With

	if intNeedID > 0 then strSource = strSource & PrintIndent & "Current invoice already exists, #" & intNeedID & "<br>"

	NeedNewInvoice = intNeedID

End Function


'-------------------------------------------------------------
'This sub updates the information for an invoice
'-------------------------------------------------------------
Sub UpdateInvoice(intInvoiceID, dblTotal, strBillingType, intPaid, dateReceived, dateSent, intSent)
	Query = "SELECT * FROM CustomerInvoices WHERE ID = " & intInvoiceID
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	rsUpdate("Total") = dblTotal
	rsUpdate("BillingType") = strBillingType
	rsUpdate("Paid") = intPaid
	rsUpdate("DateReceived") = dateReceived
	rsUpdate("DateSent") = dateSent
	rsUpdate("InvoiceSent") = intSent


	rsUpdate.Update
	rsUpdate.Close

End Sub

'-------------------------------------------------------------
'This sub creates a charge to a customer's credit card
'-------------------------------------------------------------
Sub ChargeCard( dblChargeTotal, intInvoiceID, strDescription, blSuccess, blSendErrorEMail, strError, CustID, strTransID )
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
	objAuthorize.InvoiceNumber = intInvoiceID

	objAuthorize.Login = "OurPage"
	objAuthorize.Password = "hgf554jh"

	objAuthorize.CardNumber = strCCNumber
	objAuthorize.CardType = strCCType
	objAuthorize.ExpDate = Month(CCExpDate) & "/" & Year(CCExpDate)

	'Charge the amount from the order!!!  Add the description too
	objAuthorize.Amount = dblChargeTotal
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
				strError = PrintIndent & "A connection could not be established with the authorization network.  Our support personnel have been notified, and will personally handle the matter."
			Case -2
				strError = PrintIndent & "A connection could not be established with the authorization network.  Our support personnel have been notified, and will personally handle the matter."
			Case Else
				strError = PrintIndent & "An unknown error occured with the authorization network.  Our support personnel have been notified, and will personally handle the matter."
		End Select
	End If


	if strError <> "" then
		blSuccess = false

		if blSendErrorEMail = "" then blSendErrorEMail = true

		if blSendErrorEMail = true then

			'The subject line
			strSubject = "Credit Card charge problem! CustomerID " & CustomerID & ", TransID " & strTransID
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
				"Total: " & dblChargeTotal & "<br>" & _
				"TransID: " & strTransID & "<br>" & _
				"Description: " & strDescription & "<br>" & _
				"Error given: <b>" & strError & "<br>"

			'Set the rest of the mailing info and send it
			Mailer.Subject    = strSubject
			Mailer.ClearBodyText
			Mailer.BodyText   = strBody
			Mailer.ClearRecipients
			Mailer.AddRecipient "Credit Card Support", "accounts@grouploop.com"
			Mailer.SendMail


		end if
	end if

End Sub

'-------------------------------------------------------------
'This function gets the total monthly cost of a site
'-------------------------------------------------------------
Function GetInvoiceTotal( intInvoiceID )
	With cmd
		.CommandText = "GetInvoiceTotal"

		.Parameters.Refresh

		.Parameters("@InvoiceID") = intInvoiceID

		.Execute , , adExecuteNoRecords

		currTotal = .Parameters("@Total")
	End With

	GetInvoiceTotal = CDbl(currTotal)
End Function



'-------------------------------------------------------------
'This function adds a bank deposit to the databse for an invoice charged with a card
'-------------------------------------------------------------
Sub AddDeposit( intCustomerID, intInvoiceID, dblChargeTotal, strTransID )
		'Add the deposit
		With cmd
			'Get the customer's info
			.CommandText = "AddBankDeposit"
			.Parameters.Refresh
			.Execute , , adExecuteNoRecords

			intDepositID = .Parameters("@ID")
		End With

		Query = "SELECT * FROM BankDeposits WHERE ID = " & intDepositID
		rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

		rsUpdate("CustomerID") = intCustomerID
		rsUpdate("InvoiceID") = intInvoiceID
		rsUpdate("Total") = dblChargeTotal
		rsUpdate("BillingType") = "CreditCard"
		rsUpdate("Description") = "GroupLoop Site Charge"
		rsUpdate("TransID") = strTransID

		rsUpdate("DateDeposited") = FormatDateTime(Date, 2) 

		rsUpdate.Update
		rsUpdate.Close
End Sub


'-------------------------------------------------------------
'This deletes an invoice
'-------------------------------------------------------------
Sub DeleteInvoice(intInvoiceID)
	With cmd
		.CommandText = "DeleteInvoice"
		.Parameters.Refresh

		.Parameters("@InvoiceID") = intInvoiceID
		.Execute , , adExecuteNoRecords

	End With

	strSource = strSource & PrintIndent & "Deleted invoice #" & intInvoiceID & "<br>"
End Sub


'-------------------------------------------------------------
'This puts all their pre-determined monthly charges into invoice charges for the current invoice
'-------------------------------------------------------------
Sub CreateInvoiceCharges(intInvoiceID, intMaintenanceID)
	Query = "SELECT ID, Total, Description FROM CustomerMonthlyCharges WHERE CustomerID = " & rsCust("ID") & " ORDER BY ID DESC"
	rsBilling.CacheSize = 10
	rsBilling.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	do until rsBilling.EOF
		AddInvoiceCharge intInvoiceID, rsBilling("Total"), rsBilling("Description"), intMaintenanceID, rsBilling("ID")
		rsBilling.MoveNext
	loop

	rsBilling.Close

End Sub


'-------------------------------------------------------------
'This function creates an invoice charge
'-------------------------------------------------------------
Sub AddInvoiceCharge(intInvoiceID, currTotal, strDescription, intMaintenanceID, intMonthlyChargeID)
	With cmd
		if .CommandText <> "AddCustomerInvoiceCharge" then
			.CommandText = "AddCustomerInvoiceCharge"
			.Parameters.Refresh
		end if

		.Parameters("@InvoiceID") = intInvoiceID
		.Parameters("@Total") = currTotal
		.Parameters("@Description") = strDescription
		.Parameters("@MaintenanceID") = intMaintenanceID
		.Parameters("@MonthlyChargeID") = intMonthlyChargeID

		.Execute , , adExecuteNoRecords

		currTotal = .Parameters("@Total")
	End With

End Sub


'-------------------------------------------------------------
'This function creates an invoice charge if they have additional charges for their site
'-------------------------------------------------------------
Sub CreateGoldSiteAdditional(intInvoiceID, intMaintenanceID)
	if rsCust("ChargeAdditionalFees") = 1 then
		With cmd
			.CommandText = "CreateGoldSiteAdditional"

			.Parameters.Refresh

			.Parameters("@CustomerID") = rsCust("ID")
			.Parameters("@InvoiceID") = intInvoiceID
			.Parameters("@MaintenanceID") = intMaintenanceID

			.Execute , , adExecuteNoRecords

			currTotal = .Parameters("@Total")
		End With

	end if
End Sub


'-------------------------------------------------------------
'This function checks if the customer has already been charged this month
'-------------------------------------------------------------
Function ChargedAlready()
	With cmd
		.CommandText = "CustomerChargedAlready"

		.Parameters.Refresh

		.Parameters("@CustomerID") = rsCust("ID")

		.Execute , , adExecuteNoRecords

		intInvoiceID = .Parameters("@InvoiceID")
	End With

	ChargedAlready = intInvoiceID
End Function


'-------------------------------------------------------------
'This function sees if it's the time in the billing cycle to charge the customer.
'-------------------------------------------------------------
Function TimeToCharge(intInvoiceID)
	With cmd
		.CommandText = "InvoiceTimeToCharge"

		.Parameters.Refresh

		.Parameters("@BillingCycleMonths") = rsCust("BillingCycleMonths")
		.Parameters("@InvoiceID") = intInvoiceID

		.Execute , , adExecuteNoRecords


		intMonthDifference = .Parameters("@MonthDifference")

		blTime = CBool(.Parameters("@Charge"))
	End With

	if not blTime then strSource = strSource & PrintIndent & "The billing cycle is " & rsCust("BillingCycleMonths") & " months.  Invoice will be sent in " & rsCust("BillingCycleMonths") - intMonthDifference & "month(s).  Do not send.<br>"

	TimeToCharge = blTime
End Function



'-------------------------------------------------------------
'This function checks if the maintainence has been run today (uses NightlyMaintenance table)
'Switch this to a stored proc later...
'-------------------------------------------------------------
Function RunAlready(ProcessDate)
	StartDate = FormatDateTime(ProcessDate, 2) & " 12:00:01 AM"
	EndDate = FormatDateTime(ProcessDate, 2) & " 11:59:59 PM"

	'Add the maintenance record
	Query = "SELECT ID FROM NightlyMaintenance WHERE ProcessDate = '" & ProcessDate & "' OR ProcessDate >= '" & StartDate & "' AND ProcessDate <= '" & EndDate & "'"
	rsPage.Open Query, Connect, adOpeForwardOnly, adLockOptimistic, adCmdTableDirect

	if rsPage.EOF then
		blResult = false
	else
		blResult = true
	end if

	rsPage.Close

	RunAlready = blResult
End Function


'-------------------------------------------------------------
'Create a maintenance record
'-------------------------------------------------------------
Function AddMaintenance()
	With cmd
		'Get the customer's info
		.CommandText = "AddNightlyMaintenance"
		.Parameters.Refresh
		.Execute , , adExecuteNoRecords

		intMaintenanceID = .Parameters("@ItemID")
	End With
	AddMaintenance = intMaintenanceID
End Function


'-------------------------------------------------------------
'This sub takes the output from today's maintenance and adds it to a record, then e-mails it
'to me for record
'-------------------------------------------------------------
Sub UpdateMaintenance( intMaintenanceID )
	'Add the maintenance record
	Query = "SELECT ID, Output, CardCharges, Invoices, Errors, ExpirationWarnings, ProcessDate FROM NightlyMaintenance WHERE ID = " & intMaintenanceID
	rsPage.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect
	'Update the fields
		rsPage("Output") = strSource
		rsPage("CardCharges") = intCardCharges
		rsPage("Invoices") = intInvoices
		rsPage("Errors") = intErrors
		rsPage("ExpirationWarnings") = intExpirationWarnings
		rsPage("ProcessDate") = ProcessDate

	rsPage.Update
	rsPage.Close


	if blSendEMail then

		'Send an e-mail to accounts with a copy of todays maintenance
		strSubject = "Nightly Maintenance: " & FormatDateTime( ProcessDate, 2 ) & " " & Time & "  CardCharges: " & intCardCharges  & _
		"  Invoices: " & intInvoices  & "  Errors: " & intErrors  & "  ExpirationWarnings: " & intExpirationWarnings

		'Set the rest of the mailing info and send it
		Mailer.ClearRecipients
		Mailer.AddRecipient "GroupLoop Accounts", "accounts@GroupLoop.com"
		Mailer.Subject    = strSubject
		Mailer.ClearBodyText
		Mailer.BodyText   = strSource
		Mailer.SendMail
	end if
End Sub



'-------------------------------------------------------------
'This sub deletes records that don't belong.  They usually get there when
'something does not complete... but they are useless.. housecleaning
'-------------------------------------------------------------
Sub DeleteInvalid()
	Query = "DELETE ForumMessages WHERE CategoryID = 0"
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	Query = "DELETE Photos WHERE CategoryID = 0"
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	Query = "DELETE Media WHERE CategoryID = 0"
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	CutoffDate = DateAdd("d", (-2), ProcessDate)

	Query = "DELETE SiteSearch WHERE ( Date < '" & CutoffDate & "')"
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
	Query = "DELETE SectionSearch WHERE ( Date < '" & CutoffDate & "')"
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
End Sub



'-------------------------------------------------------------
'This sub prints the index file for a customer.  Just includes the others...
'-------------------------------------------------------------
Sub WriteIndex()
	strSubDir = rsCust("SubDirectory")

	if rsCust("SubDirectory") = "" then
		intErrors = intErrors + 1
		strSource = strSource & PrintIndent &  "<b><u>Index could not be written because sub-directory is blank!  Check this personally!</u></b><br>"
		Exit Sub
	end if

	strPath = Server.MapPath("..\" & strSubDir ) & "\write_index.asp"

	if not FileSystem.FileExists( strPath ) then
		'intErrors = intErrors + 1
		'strSource = strSource & PrintIndent &  "<b><u>Index could not be written because " & strPath & " does not exist!</u></b><br>"
		Exit Sub
	end if

	strURL = "http://www.GroupLoop.com/" & strSubDir & "/write_index.asp"

	GetURL strURL
End Sub



'-------------------------------------------------------------
'This sub writes the constants file for a customer.
'-------------------------------------------------------------
Sub WriteConstants()
	strSubDir = rsCust("SubDirectory")

	if rsCust("SubDirectory") = "" then
		intErrors = intErrors + 1
		strSource = strSource & PrintIndent &  "<b><u>Constants file could not be written because sub-directory is blank!  Check this personally!</u></b><br>"
		Exit Sub
	end if

	strPath = Server.MapPath("..\" & strSubDir ) & "\write_constants.asp"

	if not FileSystem.FileExists( strPath ) then
		'intErrors = intErrors + 1
		'strSource = strSource & PrintIndent &  "<b><u>Constants file could not be written because " & strPath & " does not exist!</u></b><br>"
		Exit Sub
	end if

	strURL = "http://www.GroupLoop.com/" & strSubDir & "/write_constants.asp"

	GetURL strURL
End Sub


'-------------------------------------------------------------
'This sub writes the HeaderFooter file for a customer.
'-------------------------------------------------------------
Sub WriteHeaderFooter()
	strSubDir = rsCust("SubDirectory")

	if rsCust("SubDirectory") = "" then
		intErrors = intErrors + 1
		strSource = strSource & PrintIndent &  "<b><u>Header file could not be written because sub-directory is blank!  Check this personally!</u></b><br>"
		Exit Sub
	end if

	strPath = Server.MapPath("..\" & strSubDir ) & "\write_header_footer.asp"

	if not FileSystem.FileExists( strPath ) then
		'intErrors = intErrors + 1
		'strSource = strSource & PrintIndent &  "<b><u>Header file could not be written because " & strPath & " does not exist!</u></b><br>"
		Exit Sub
	end if

	strURL = "http://www.GroupLoop.com/" & strSubDir & "/write_header_footer.asp"

	GetURL strURL
End Sub


'-------------------------------------------------------------
'This function prints an indent to make things look nice... awwww... I'm drunk, and I'm not being some college tool.  Partook of Dad's bacardi
'-------------------------------------------------------------
Function PrintIndent
	PrintIndent = "&nbsp;&nbsp;&nbsp;&nbsp;"
End Function



'-------------------------------------------------------------
'This sub checks the signup date for a free site.  If it is at a certain point (30 days old, 7 days old, etc)
'it sends an e-mail.  If the site is 60 days old, delete it
'-------------------------------------------------------------
Sub SendExpirationWarnings()
	Version = rsCust("Version")

	'Only send warnings to free sites
	if Version <> "Free" then exit sub

	'Put this into a local variable...
	SignUpDate = rsCust("SignupDate")

	'If they are in the day below, send them a warning email...
	if DateDiff( "d", SignUpDate, ProcessDate ) = 16 then
		CutoffDate = DateAdd("d", (-14), ProcessDate)
		SendExpriationEMail "in 2 weeks", CutoffDate

	elseif DateDiff( "d", SignUpDate, ProcessDate ) = 23 then
		CutoffDate = DateAdd("d", (-7), ProcessDate)
		SendExpriationEMail "in 1 week", CutoffDate

	elseif DateDiff( "d", SignUpDate, ProcessDate ) = 25 then
		CutoffDate = DateAdd("d", (-5), ProcessDate)
		SendExpriationEMail "in 5 days", CutoffDate

	elseif DateDiff( "d", SignUpDate, ProcessDate ) = 28 then
		CutoffDate = DateAdd("d", (-2), ProcessDate)
		SendExpriationEMail "in 2 days", CutoffDate
		
	elseif DateDiff( "d", SignUpDate, ProcessDate ) = 29 then	'It's been 29 days... one left
		CutoffDate = DateAdd("d", (-1), ProcessDate)
		SendExpriationEMail "TOMORROW", CutoffDate

	elseif DateDiff( "d", SignUpDate, ProcessDate ) >= 30 then	'It's been 30 days
		'DELETESITE
		strSource = strSource & PrintIndent &  "<FONT SIZE='+2'>DELETE THIS SITE TODAY!</font><br>"

		strSource = strSource & DeleteCustomer( rsCust("ID") )

	end if
End Sub

'-------------------------------------------------------------
'This sub actually sends the expiration warning e-mail.  Pretty simple
'-------------------------------------------------------------
Sub SendExpriationEMail( strTime, CutoffDate )
	intExpirationWarnings = intExpirationWarnings + 1

	strSource = strSource & PrintIndent &  "Sending expiration warning: site will expire " & strTime & ".<br>"

	'The subject line
	strSubject = "Your GroupLoop site will expire " & strTime

	'The body
	strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
	"<p><i>Dear " & rsCust("FirstName") & ",</i><br>" & _
	" &nbsp;&nbsp;Your free group site has a two month time limit.  Your trial is almost up!  Your site will be automatically erased <b>" & _
	strTime & "</b>.  All the content, members, photos, and files you added will be permanently erased.</p>" & _
	"<p> &nbsp;&nbsp;You must upgrade to the Gold version to prevent this from happening.  The Gold version is only $20/month, and we are sure you " & _
	"will agree it's worth it.  You can upgrade two ways:<br>" & _
	"1. <a href='https://www.OurClubPage.com/admin/account_site_upgrade.asp?CustomerID=" & rsCust("ID") & "'>CLICK HERE</a> <br>" & _
	"2. Goto your site's member section at <a href='http://www.GroupLoop.com/" & rsCust("SubDirectory") & "/members.asp'>http://www.GroupLoop.com/" & rsCust("SubDirectory") & "</a> and " & _
	"click on Upgrade To Gold Version<br><br>" & _
	" &nbsp;&nbsp;As a Gold member, you will have the ability to have as many members, photos, and files you want!  We hope you upgrade and enjoy your site as long as possible!</p>" & _
		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf

	strHeader = "<html><title>" & strSubject & "</title><body>"
	strFooter = "</body></html>"

	strRecipName = rsCust("FirstName") & " " & rsCust("LastName")

	'Set the rest of the mailing info and send it
	Mailer.ClearRecipients

	if rsCust("EMail") = "" then
		Response.Write "<b>Customer does not have an e-mail address.</b><br>"
	else
		Mailer.AddRecipient strRecipName, rsCust("EMail")
	end if

	if not IsNull(rsCust("EMail1")) then
		if rsCust("EMail1") <> "" and rsCust("EMail") <> rsCust("EMail1") then Mailer.AddRecipient strRecipName, rsCust("EMail1")
	end if

	Mailer.Subject    = strSubject
	Mailer.ClearBodyText
	Mailer.BodyText   = strHeader & strBody & strFooter
	Mailer.SendMail
End Sub




'STEVE READ THIS SOMEDAY - Although this summer should have been the greatest one ever, you didn't let it.  Instead, you let
'the your responsiblilities hang over your head the whole time (good, shows dedication).  GroupLoop WILL make you a goddamn shitload of fucking money (hopefully with these churches).
'You WILL suprise everyone and laugh in the fucking face of people that doubt this.  You WILL make your father proud of you.  You WILL NOT be 
'some fucking guy who works for a company, comes home, and bitches about his job and the corporate world.  You WILL show
'people that corporations are powerful, but no substitute for hard, innovative work.  You WILL take the credit and reap the rewards of your own work.
'You WILL treat your employees, friends, and family right.  You WILL make your life's work doing the work you want.  You WILL NOT take some
'bullshit job because of money.  You WILL NOT marry some girl because you don't want to be alone.  You WILL raise your kid like Dad raised me.  You WILL
'let him try guitar, drums, and build a badass go-kart.  You WILL never hate yourself.  You WILL try not to hate people, only understand them.  You WILL
'live Robert Frost's poem and always take the right path, the one less traveled.  You WILL remember everything your father said, because he is the 
'smartest man you've ever known.  YOU WILL NOT drink any more of that wine from the jug.

'If anyone other than me ever reads that, please don't think I'm some cocky asshole.  I never am condescending; however, I am a bit depressed, 
'and hope the company I put my whole life into works out like I think it can.
%>


<!-- #include file ="../sourcegroup/closedsn.asp" -->

