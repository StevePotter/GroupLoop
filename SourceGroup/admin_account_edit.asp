<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin() and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin() then Redirect("members.asp?Source=admin_account_edit.asp")

Set Command = Server.CreateObject("ADODB.Command")

With Command
	'Check to make sure the CC info is correct
	.ActiveConnection = Connect
	.CommandText = "GetOwnerMemberID"
	.CommandType = adCmdStoredProc
	.Parameters.Refresh
	.Parameters("@CustomerID") = CustomerID
	.Execute , , adExecuteNoRecords
	intOwnerID = .Parameters("@MemberID")
	strName = .Parameters("@FirstName") & "&nbsp;" & .Parameters("@LastName")
End With
Set Command = Nothing

if Session("MemberID") <> intOwnerID then Redirect("message.asp?Source=members.asp&Message=" & Server.URLEncode("Sorry, but only the site owner, " & strName & " can terminate the site."))

if Version <> "Gold" then Redirect("error.asp")

strSubmit = Request("Submit")
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change Account Information</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
if strSubmit = "Update My Account" then
	if Request("FirstName") = "" or Request("LastName") = "" or Request("NewEMail") = "" _
		or Request("Address1") = "" or Request("City") = "" or Request("State") = "" _
		or Request("Zip") = "" or Request("Phone") = "" then Redirect("incomplete.asp")

	strFirstName = Request("FirstName")
	strLastName = Request("LastName")
	strOrganization = Request("Organization")
	strNewEMail = Request("NewEMail")
	strStreet1 = Request("Address1")
	strStreet2 = Request("Address2")
	strCity = Request("City")
	strState = Request("State")
	strPhone = Request("Phone")
	strZip = Request("Zip")
	strNewCCNumber = Request("NewCCNumber")

	strNewCCFirstName = Request("NewCCFirstName")
	strNewCCLastName = Request("NewCCLastName")

	strNewCCType = Request("NewCCType")
	dateNewCCExp = CDate(Request("CCExpMonth") & "/1/" & Request("CCExpYear"))
	if not IsDate( dateNewCCExp ) then Redirect("error.asp?Message=" & Server.URLEncode("The expiration date is invalid."))


	if strNewCCNumber <> "" then

		if Request("CCStreet1") = "" then
			strCCStreet1 = strStreet1
		else
			strCCStreet1 = Request("CCStreet1")
		end if
		if Request("CCStreet2") = "" then
			strCCStreet2 = strStreet2
		else
			strCCStreet2 = Request("CCStreet2")
		end if
		if Request("CCCity") = "" then
			strCCCity = strCity
		else
			strCCCity = Request("CCCity")
		end if
		if Request("CCState") = "" then
			strCCState = strState
		else
			strCCState = Request("CCState")
		end if
		if Request("CCZip") = "" then
			strCCZip = strZip
		else
			strCCZip = Request("CCZip")
		end if
		if Request("CCCountry") = "" then
			strCCCountry = strCountry
		else
			strCCCountry = Request("CCCountry")
		end if

		strCCCompany = Request("CCCompany")

		' Create xAuthorize object.
		Dim objAuthorize
		Set objAuthorize = Server.CreateObject("xAuthorize.Process")
		' Initialize xAuthorize for a new transaction.
		objAuthorize.Initialize

		' Set object properties.
		objAuthorize.Processor = "AUTHORIZE_NET"

		objAuthorize.FirstName = strNewCCFirstName
		objAuthorize.LastName = strNewCCLastName
		objAuthorize.Company = strCCCompany


		objAuthorize.Address = strCCStreet1 & " " & strCCStreet2
		objAuthorize.City = strCCCity
		objAuthorize.State = strCCState
		objAuthorize.Zip = strCCZip
		objAuthorize.Country = strCCCountry

		objAuthorize.CustomerID = CustomerID
		objAuthorize.InvoiceNumber = CustomerID

		objAuthorize.Login = "OurPage"
		objAuthorize.Password = "hgf554jh"

		objAuthorize.CardNumber = strNewCCNumber
		objAuthorize.CardType = strNewCCType
		objAuthorize.ExpDate = Month(dateNewCCExp) & "/" & Year(dateNewCCExp)

		objAuthorize.Amount = 20
		objAuthorize.TransType = "AUTH_ONLY"
		objAuthorize.Description = "GroupLoop.com monthly charge."

		objAuthorize.EmailMerchant = false
		objAuthorize.EmailCustomer = false

		' Initiate transation processing
		objAuthorize.Process

		strTransID = objAuthorize.TransID

		If objAuthorize.ErrorCode = 0 Then
			 ' Communication was successful.
			 ' Examine Results
			 If objAuthorize.ResponseCode <> 1 then
				strError = "Sorry, but there is the following problem with the card you entered:<br><font size='+1'>" & objAuthorize.ResponseReasonText & "</font><br><br>Make sure you have the correct address, card number, and expiration date."
			end if
		Else
			Select Case objAuthorize.ErrorCode
				Case -1
					strError = "Sorry, a connection could not be established with the authorization network.  Please try again."
				Case -2
					strError = "Sorry, a connection could not be established with the authorization network.  Please try again."
				Case Else
					strError = "Sorry, an unknown error occured with the authorization network.  Please try again, and notify support@keist.com if this keeps happening."
			End Select
		End If

		Set objAuthorize = Nothing


		'We had a problem
		if strError <> "" then
			Set rsUpdate = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode(strError) )
		end if

		rsUpdate("TransID") = strTransID

	end if

	Query = "SELECT * FROM Customers WHERE ID = " & CustomerID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		Set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("FirstName") = Format( strFirstName )
	rsUpdate("LastName") = Format( strLastName )
	rsUpdate("Organization") = Format( strOrganization )
	rsUpdate("EMail") = strNewEMail
	rsUpdate("Street1") = strStreet1
	rsUpdate("Street1") = strStreet2
	rsUpdate("City") = Format( strCity )
	rsUpdate("State") = strState
	rsUpdate("Phone") = strPhone
	rsUpdate("Zip") = strZip

	strSubDir = rsUpdate("SubDirectory")

	if strNewCCNumber <> "" then
		rsUpdate("CCNumber") = strNewCCNumber
		rsUpdate("CCFirstName") = strNewCCFirstName
		rsUpdate("CCLastName") = strNewCCLastName
		rsUpdate("CCType") = strNewCCType
		rsUpdate("CCExpDate") = dateNewCCExp
		rsUpdate("BillingStreet1") = strCCStreet1
		rsUpdate("BillingStreet2") = strCCStreet2
		rsUpdate("BillingCity") = strCCCity
		rsUpdate("BillingState") = strCCState
		rsUpdate("BillingPhone") = strPhone
		rsUpdate("BillingZip") = strCCZip
		rsUpdate("BillingCountry") = strCCCountry

	end if

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>
	Your account has been updated.  To return to your site, <a href="http://www.GroupLoop.com/<%=strSubDir%>/write_header_footer.asp">click here</a>.  Thanks!
	</p>
<%
'-----------------------Begin Code---------------------------
else
	Query = "SELECT * FROM Customers WHERE ID = " & CustomerID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if rsUpdate.EOF then
		Set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
%>
	<script language="JavaScript">
		function submit_page(form) {
			//Error message variable
			var strError = "";

			if(form.FirstName.value == "")
				strError += "          You forgot your First Name. \n";
			if(form.LastName.value == "")
				strError += "          You forgot your Last Name. \n";
			if(form.EMail.value == "")
				strError += "          You forgot your EMail. \n";
			else{
				if ((getFront(form.EMail.value,"@") == null) || (getEnd(form.EMail.value,"@") == ""))
					strError += "          Please enter a valid e-mail address, such as JoesPizza@aol.com. \n";
			}
			if(form.Street1.value == "")
				strError += "          You forgot your Street Address. \n";
			if(form.City.value == "")
				strError += "          You forgot your City. \n";
			if(form.Zip.value == "")
				strError += "          You forgot your Zip Code. \n";
			if(form.Phone.value == "")
				strError += "          You forgot your Phone Number. \n";

			if(form.NewCCNumber.value != ""){
				if(form.CCFirstName.value == "" && form.CCLastName.value == "" && form.CCCompany.value == "")
					strError += "          You forgot your Credit Card Name or Company. \n";




			}




			if(!form.Agree.checked)
				strError += "You forgot to check the authorize box.\n";
			//They didn't enter a name or company
			//They entered a first name, but not a last
			else if( form.CCFirstName.value != "" && form.CCLastName.value == "" )
				strError += "          You forgot your Credit Card Last Name. \n";
			//They entered a last name, but not a first
			else if( form.CCFirstName.value == "" && form.CCLastName.value != "" )
				strError += "          You forgot your Credit Card First Name. \n";

			if(form.CCNumber.value == "")
				strError += "          You forgot your Credit Card Number. \n";
			if(form.Address1.value == "")
				strError += "          You forgot your Street Address. \n";
			if(form.City.value == "")
				strError += "          You forgot your City. \n";
			if(form.Zip.value == "")
				strError += "          You forgot your Zip Code. \n";
			if(form.State.value == "" && (form.Country.value == "USA" || form.Country.value == "CAN"))
				strError += "          You forgot your State. \n";

			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can sign up: \n" + strError;
				alert (strError);
				return false;
			}   
		}

		function getFront(mainStr,searchStr){
			foundOffset = mainStr.indexOf(searchStr)
			if (foundOffset <= 0) {
				return null // if the @ symbol is missing the value is -1
							// if the @ symbol is the first char the value is 0
			} 
			else {
				return mainStr.substring(0,foundOffset)
			}
		}
    
		function getEnd(mainStr,searchStr) {
			foundOffset = mainStr.indexOf(searchStr)
			if (foundOffset <= 0) {
				return ""   // if the @ symbol is missing the value is -1
							// if the @ symbol is the first char the value is 0
			}
			else {
				return mainStr.substring(foundOffset+searchStr.length,mainStr.length)
			}
		}
	</script>

	* indicated required information<br>
	<form METHOD="POST" ACTION="https://www.OurClubPage.com/<%=SubDirectory%>/admin_account_edit.asp">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Account Information
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* First Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="FirstName" size="55" value="<%=rsUpdate("FirstName")%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Last Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="LastName" size="55" value="<%=rsUpdate("LastName")%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* E-Mail Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="EMail" size="55" value="<%=rsUpdate("EMail")%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Organization</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Organization" size="15" value="<%=rsUpdate("Organization")%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Address1" size="55" value="<%=rsUpdate("Street1")%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Address 2nd Line</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Address2" size="55" value="<%=rsUpdate("Street2")%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* City</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="City" size="55" value="<%=rsUpdate("City")%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* State</td>
      		<td class="<% PrintTDMain %>"> 
				<% PrintStates "State", rsUpdate("State") %>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Zip Code</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Zip" size="10" value="<%=rsUpdate("Zip")%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Phone Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Phone" size="15" value="<%=rsUpdate("Phone")%>">
     		</td>
		</tr>


		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Billing Information.  You may change the credit/check card that is being billed.  If 
				the billing address for the credit card is different than above, please 
				enter it below.  <b>Keep this section blank if you don't want to change your billing 
				information.  This is a totally secure connection, so have no worries about entering your vital information.</b>
     		</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" align="right">Card Type</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<select name="NewCCType" size=1>
					<option value="VISA">VISA</option>
					<option value="MasterCard">MasterCard</option>
					<option value="AmEx">American Express</option>
				</select>
			</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Name on Credit Card (first then last)</td>
      		<td class="<% PrintTDMain %>"> 
				<input type="text" name="NewCCFirstName" size="20">&nbsp;
				<input type="text" name="NewCCLastName" size="20">
     		</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Company Name (optional)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<input type="text" name="CCCompany" size="40">
			</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Credit Card Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="NewCCNumber" size="20">
     		</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Expiration Date
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
			<select name="CCExpMonth">
				<option value="1">January</option>
				<option value="2">February</option>
				<option value="3">March</option>
				<option value="4">April</option>
				<option value="5">May</option>
				<option value="6">June</option>
				<option value="7">July</option>
				<option value="8">August</option>
				<option value="9">September</option>
				<option value="10">October</option>
				<option value="11">November</option>
				<option value="12">December</option>
			</select>
			<select name="CCExpYear">
				<option value="2001">2001</option>
				<option value="2002">2002</option>
				<option value="2003">2003</option>
				<option value="2004">2004</option>
				<option value="2005">2005</option>
				<option value="2006">2006</option>
				<option value="2007">2007</option>
				<option value="2008">2008</option>
				<option value="2009">2009</option>
				<option value="2010">2010</option>
			</select>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street Address For Card (leave the address stuff blank if it is the same address 
				you entered above)
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<input type="text" name="CCStreet1" size="40">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street Address Line 2
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<input type="text" name="CCStreet2" size="40">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<input type="text" name="CCCity" size="40">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMainSwitch %>">
				<% PrintStates "CCState", "" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMainSwitch %>" align="left">
				<input type="text" name="CCZip" size="8">
			</td>
		</tr>

		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Update My Account">
	   		</td>
		</tr>
  	</table>
	</form>

<%
end if
%>
