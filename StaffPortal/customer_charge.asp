<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p class=Heading align=center><font size=+3>Charge Customer</font></p>
<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Submit" then
	if Request("CustomerID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intCustomerID = CInt(Request("CustomerID"))
	intInvoiceID = CInt(Request("InvoiceID"))

	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = Connect
	cmd.CommandType = adCmdStoredProc


	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.ContentType = "text/html"
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = "GroupLoop.com"
	Mailer.FromAddress = "support@grouploop.com"



	Set rsCust = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM Customers WHERE ID = " & intCustomerID
	rsCust.CacheSize = 100
	rsCust.Open Query, Connect, adOpenStatic, adLockOptimistic

	dblChargeTotal = CDbl(Request("Total"))
	strDescription = Format(Request("Description"))

	'Add the new deposit
	intDepositID = AddDeposit()

	ChargeCard dblChargeTotal, intDepositID, strDescription, blSuccess, true, strError, rsCust("ID"), strTransID

	if not blSuccess then
		'Didn't work, delete the deposit record
		DeleteDeposit intDepositID
		Response.Write "<b>CHARGE FOR " & FormatCurrency( dblChargeTotal ) & " COULD NOT BE MADE - TransID:" & strTransID & " - DepositID:" & intDepositID & "</b>  Error explained below.<br>" & strError & "<br>"


	else


		Query = "SELECT * FROM BankDeposits WHERE ID = " & intDepositID
		Set rsDeposit = Server.CreateObject("ADODB.Recordset")
		rsDeposit.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

		rsDeposit("CustomerID") = intCustomerID
		rsDeposit("InvoiceID") = intInvoiceID
		rsDeposit("Total") = dblChargeTotal
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


		Response.Write "Card successfully charged " & FormatCurrency( dblChargeTotal ) & " - TransID:" & strTransID & " - DepositID:" & intDepositID & "<br>"

		'The subject line
		strSubject = "Charge to your GroupLoop account"

		'The body
		strBody = "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
		"<p><i>Dear " & rsCust("FirstName") & ",</i><br>" & _
		"Your credit card has been charged " & FormatCurrency( dblChargeTotal ) & " for your GroupLoop site.<br>" & _
		"The card charged was: " & Left(rsCust("CCNumber"), 4) & "..." & Right(rsCust("CCNumber"), 3) & "<br>" & _
		"The description of the charge was: " & strDescription & "</p>"

		if strCustomerNote <> "" then
			strBody = strBody & "<p>" & strCustomerNote & "</p>"
		end if

		strBody = strBody & "<p>Thank you for being with us.  It is our mission to constantly make your site better, easier, and faster.  " & _
		"I hope you continue to enjoy your site and use it as much as you can!</p>" & _
		"<p><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you and enjoy,<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p>" & VbCrLf
	end if


	if blSuccess then
		Response.write "Sending mail to customer...<br>"
		'Set the rest of the mailing info and send it
		Mailer.Subject    = strSubject
		Mailer.ClearBodyText
		Mailer.BodyText   = strBody
		Mailer.ClearRecipients
		strRecip = rsCust("FirstName")&" "&rsCust("LastName")
		strEmail = rsCust("EMail")
		Mailer.AddRecipient strRecip, strEmail
		if not Mailer.SendMail then 
			Response.Write "<b>There has been an error, and the email has not been sent to " & rsCust("FirstName") & " " & rsCust("LastName") & ", <a href=mailto:" & rsCust("EMail") & ">" & rsCust("EMail") & "</a>, CustID: " & rsCust("ID") & ", Total:" & FormatCurrency(dblChargeTotal) & ", InvoiceID:" & intInvoiceID & ".</b>  <br>Error was '" & Mailer.Response & "'<br>"
		end if
		Mailer.ClearRecipients
		Mailer.AddRecipient "Charge Support", "support@grouploop.com"
		Mailer.SendMail
	end if

	Set cmd = Nothing

	Set Mailer = Nothing

	if blSuccess then
%>
	<p>
	The customer has been charged. <br>
	<a href="customer_view.asp?ID=<%=intCustomerID%>">Return to the customer's information.</a>
	</p>
<%
	else
%>
	<p>
	The customer was not charged. <br>
	<a href="customer_charge.asp?CustomerID=<%=intCustomerID%>&InvoiceID=<%=intInvoiceID%>">Try again.</a><br>
	
	<a href="customer_view.asp?ID=<%=intCustomerID%>">Return to the customer's information.</a>
	</p>
<%

	end if
else
	if Request("CustomerID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intCustomerID = CInt(Request("CustomerID"))
%>

	<script language="JavaScript">
	<!--

		//Throw out all the stuff we don't want ($)
		function ConvertDollar(currCheck) {
			if (!currCheck) return '';
			for (var i=0, currOutput='', valid="0123456789."; i<currCheck.length; i++)
				if (valid.indexOf(currCheck.charAt(i)) != -1)
					currOutput += currCheck.charAt(i);
			return currOutput;
		}


		function submit_page(form) {
			//Error message variable
			var strError = "";
			form.Total.value = ConvertDollar(form.Total.value)

			if (form.Total.value == "" || form.Total.value == "0.00" || form.Total.value == "0" )
				strError += "          You forgot the total. \n";

			if (form.Description.value == "" )
				strError += "          You forgot the description. \n";

			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can add this: \n" + strError;
				alert (strError);
				return false;
			}   
		}


	//-->
	</SCRIPT>
	<form method="post" action="customer_charge.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
	<table width="100%">
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Amount to Charge</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Total" size="5" value="$">
     		</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Invoice</td>
			<td class="<% PrintTDMain %>"> 
				<% PrintInvoicePullDown intCustomerID %>
			</td>
		</tr>

		<tr> 
     		<td class="<% PrintTDMain %>" align="right">* Description</td>
     		<td class="<% PrintTDMain %>"> 
    				<textarea name="Description" cols="55" rows="2" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr> 
     		<td class="<% PrintTDMain %>" align="right">Note to Customer</td>
     		<td class="<% PrintTDMain %>"> 
    				<textarea name="CustomerNote" cols="55" rows="2" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr> 
     		<td class="<% PrintTDMain %>" align="right">Note to Staff Only</td>
     		<td class="<% PrintTDMain %>"> 
    				<textarea name="StaffNote" cols="55" rows="2" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMainSwitch %>" align="center" colspan="2"><input type="submit" name="Submit" value="Submit">
		</tr>
	</table>
	</form>
<%

end if



Sub ChargeCard( dblChargeTotal, intDepositID, strDescription, blSuccess, blSendErrorEMail, strError, CustID, strTransID )
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
				"Total: " & dblChargeTotal & "<br>" & _
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





Sub PrintInvoicePullDown( intCustomerID )
	Set rsPulldown = Server.CreateObject("ADODB.Recordset")
	rsPulldown.CacheSize = 150

	intHighLightID = 0
	if Request("InvoiceID") <> "" then intHighLightID = CInt(Request("InvoiceID"))

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetInvoiceRecordset"
	cmdTemp.CommandType = adCmdStoredProc

	cmdTemp.Parameters.Refresh
	cmdTemp.Parameters("@InvoiceID") = intHighLightID
	cmdTemp.Parameters("@CustomerID") = intCustomerID


	rsPulldown.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	Set cmdTemp = Nothing

	if not rsPulldown.EOF then
	Set ID = rsPulldown("ID")
	Set Total = rsPulldown("Total")
	Set Description = rsPulldown("Description")
	end if


	%><select name="InvoiceID" size="1"><%

	Response.Write "<option value='0'>None</option>" & vbCrlf


	do until rsPulldown.EOF
		'Highlight the current category
		if intHighLightID = ID then
			Response.Write "<option value = '" & ID & "' SELECTED>ID #" & ID & " - " & FormatCurrency(Total) & " - " & Description & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & ID & "'>ID #" & ID & " - " & FormatCurrency(Total) & " - " & Description & "</option>" & vbCrlf
		end if

		rsPulldown.MoveNext
	loop
	rsPulldown.Close

	set rsPulldown = Nothing
	Response.Write("</select>")

End Sub
'------------------------End Code-----------------------------
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->