<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\dsn.asp" -->
<!-- #include file="..\sourcegroup\functions.asp" -->

<p align="center"><span class=Heading>Change Account Information</span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("CustomerID") = "" then Redirect("message.asp?Message=" & Server.URLEncode("You are missing your Customer ID.  Please go back to the Modify Account menu and use the links there."))
intCustomerID = CInt(Request("CustomerID"))

strSubmit = Request("Submit")

if strSubmit = "Update My Account" then
	if Request("CCName") = "" or Request("CCNumber") = "" or Request("EMail") = "" _
		or Request("FirstName") = "" or Request("LastName") = "" or Request("NewEMail") = "" _
		or Request("Address1") = "" or Request("City") = "" or Request("State") = "" _
		or Request("Zip") = "" or Request("Phone") = "" then Redirect("incomplete.asp")

	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")

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
	strNewCCName = Request("NewCCName")
	strNewCCType = Request("NewCCType")
	dateNewCCExp = CDate(Request("CCExpMonth") & "/1/" & Request("CCExpYear"))
	if not IsDate( dateNewCCExp ) then Redirect("error.asp?Message=" & Server.URLEncode("The expiration date is invalid."))

	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'Check the scheme to make sure the CC info is correct
		.ActiveConnection = Connect
		.CommandText = "ValidCustomerInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@EMail") = strEMail
		.Parameters("@CCName") = strCCName
		.Parameters("@CCNumber") = strCCNumber
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Valid"))
		'Wrong info
		if not blValid then
			Set Command = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode("The information you entered did not exactly match that of the account.  Remember that the credit card being billed for the account is the only one that will work.  Please try again."))
		end if

		'HERE YOU HAVE TO CHECK THE NEW CC INFO

	End With
	Set Command = Nothing


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

	Query = "SELECT * FROM Customers WHERE ID = " & intCustomerID
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

	if strNewCCName <> "" and strNewCCNumber <> "" then
		rsUpdate("CCNumber") = strNewCCNumber
		rsUpdate("CCName") = strNewCCName
		rsUpdate("CCType") = strNewCCType
		rsUpdate("CCExpDate") = dateNewCCExp
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
elseif strSubmit = "Edit My Account" then
	if Request("CCName") = "" or Request("CCNumber") = "" or Request("EMail") = "" then Redirect("incomplete.asp")
	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")
	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'Check to make sure the CC info is correct
		.ActiveConnection = Connect
		.CommandText = "ValidCustomerInfo"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@EMail") = strEMail
		.Parameters("@CCName") = strCCName
		.Parameters("@CCNumber") = strCCNumber
		.Execute , , adExecuteNoRecords
		blValid = CBool(.Parameters("@Valid"))
		'Wrong info
		if not blValid then
			Set Command = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode("The information you entered did not exactly match that of the account.  Remember that the credit card being billed for the account is the only one that will work.  Please try again."))
		end if

		'Get the customer's shit
		.CommandText = "GetCustomerInfo"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords
		strFirstName = .Parameters("@FirstName")
		strLastName = .Parameters("@LastName")
		strOrg = .Parameters("@Organization")
		strStreet1 = .Parameters("@Street1")
		strStreet2 = .Parameters("@Street2")
		strCity = .Parameters("@City")
		strState = .Parameters("@State")
		strZip = .Parameters("@Zip")
		strPhone = .Parameters("@Phone")
		strCCType = .Parameters("@CCType")
		dateExp = .Parameters("@CCExpDate")
	End With
	Set Command = Nothing
%>
	* indicated required information<br>
	<form METHOD="POST" ACTION="account_edit.asp">
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
	<input type="hidden" name="EMail" value="<%=strEMail%>">
	<input type="hidden" name="CCName" value="<%=strCCName%>">
	<input type="hidden" name="CCNumber" value="<%=strCCNumber%>">
	<% PrintTableHeader 0 %>
		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Account Information
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* First Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="FirstName" size="55" value="<%=strFirstName%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Last Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="LastName" size="55" value="<%=strLastName%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* E-Mail Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="NewEMail" size="55" value="<%=strEMail%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Organization</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Organization" size="15" value="<%=strOrg%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Address1" size="55" value="<%=strStreet1%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Address 2nd Line</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Address2" size="55" value="<%=strStreet2%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* City</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="City" size="55" value="<%=strCity%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* State</td>
      		<td class="<% PrintTDMain %>"> 
				<% PrintStates "State", strState %>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Zip Code</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Zip" size="10" value="<%=strZip%>">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Phone Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Phone" size="15" value="<%=strPhone%>">
     		</td>
		</tr>


		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Billing Information.  You may change the credit card that is being billed.  If 
				the billing address for the credit card is different than above, please 
				enter it below.  Keep this section blank if you don't want to change your billing 
				information.
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
      		<td class="<% PrintTDMain %>" align="right">Name on Credit Card</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="NewCCName" size="55">
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
'Get their info
else
	Set Command = Server.CreateObject("ADODB.Command")

	With Command
		'Check to make sure the CC info is correct
		.ActiveConnection = Connect
		.CommandText = "GetOwnerMemberID"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords
		intOwnerID = .Parameters("@MemberID")
	End With
	Set Command = Nothing

%>
	<p>Before we can edit your account, we must validate your account information.  Please enter 
	your information <b>exactly</b> like you did when you signed up.  Otherwise, it won't work.</p>
	<form METHOD="POST" ACTION="account_edit.asp">
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
	<% PrintTableHeader 0 %>
		<tr>
      		<td class="TDHeader" colspan=2 align="center"> 
       			Verify Account Information
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Account E-Mail Address</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="EMail" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Name on Credit Card</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCName" size="55">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Credit Card Number</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="CCNumber" size="20">
     		</td>
		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Edit My Account">
	   		</td>
		</tr>
  	</table>
	</form>
<%
end if

Function GetSelected( strComp1, strComp2 )
	if strComp1 = strComp2 then
		GetSelected = " selected"
	else
		GetSelected = ""
	end if
End Function
%>

<!-- #include file="..\sourcegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->