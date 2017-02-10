<!-- #include file="header.asp" -->
<!-- #include file="..\dsn.asp" -->
<!-- #include file="..\functions.asp" -->

<p align="center"><span class=Heading>Remove A Sub-Site</span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("CustomerID") = "" then Redirect("message.asp?Message=" & Server.URLEncode("You are missing your Customer ID.  Please go back to the Modify Account menu and use the links there."))
intCustomerID = CInt(Request("CustomerID"))

strSubmit = Request("Submit")

if strSubmit = "Remove My Account" then
	if Request("RemoveID") = "" or Request("CCNumber") = "" or Request("CCName") = "" or Request("EMail") then Redirect("incomplete.asp")

	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")
	intRemoveID = Request("RemoveID")

	Set FileSystem = CreateObject("Scripting.FileSystemObject")
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


		'make sure the child site is really a child site
		.ActiveConnection = Connect
		.CommandText = "ValidChildSite"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh
		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@ChildID") = intRemoveID
		blValid = CBool(.Parameters("@Valid"))
		'Wrong info
		if not blValid then
			Set Command = Nothing
			Redirect("message.asp?Message=" & Server.URLEncode("Invalid child site ID."))
		end if

		'Get the subdirectory
		.CommandText = "GetCustomerInfo"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intRemoveID
		.Execute , , adExecuteNoRecords
		strSubDir = .Parameters("@SubDirectory")
		strFirstName = .Parameters("@FirstName")
		strLastName = .Parameters("@LastName")
		if strSubDir = "" then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("Your Subdirectory was not in our records.  Please e-mail <a href=mailto:support@grouploop.com>support@grouploop.com</a> immediately.  Include your Credit Card information and your CustomerID (" & intCustomerID & ")"))
		end if

		'Check the folder.  better be there
		strFolder = Server.MapPath("../" & strSubDir)
		if not FileSystem.FolderExists(strFolder) then
			Set Command = Nothing
			Redirect("error.asp?Message=" & Server.URLEncode("Your Subdirectory could not be found on our server.  Please e-mail <a href=mailto:support@grouploop.com>support@grouploop.com</a> immediately.  Include your Credit Card information and your CustomerID (" & intCustomerID & ")"))
		end if

		'Delete the customer in the db
		.CommandText = "DeleteCustomer"
		.Parameters.Refresh
		.Parameters("@CustomerID") = intRemoveID
		.Execute , , adExecuteNoRecords
		Response.Write "<b>Removing...</b><br>1. Customer Removed From Database<br>"
	End With
	Set Command = Nothing

	'Check just once more!
	if LCase(strFolder) = "e:\webs\websites\ourclubpage.com" then
		Redirect "error.asp"
	else
		'Delete the folder
		FileSystem.DeleteFolder strFolder, True
		Response.Write "2. Files Removed From Server<br>"
	end if

	Set FileSystem = Nothing

'------------------------End Code-----------------------------
%>
	<p>
	The sub-site account has been removed.  <a href="account_childsites_remove.asp">Click here</a> to remove another.
	</p>
<%
'-----------------------Begin Code---------------------------
elseif strSubmit = "Verify" then
	if Request("CCNumber") = "" or Request("CCName") = "" or Request("EMail") then Redirect("incomplete.asp")

	strCCName = Request("CCName")
	strCCNumber = Request("CCNumber")
	strEMail = Request("EMail")

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
	End With
	Set Command = Nothing

	Query = "SELECT ID, Date, SubDirectory FROM Customers WHERE (ParentID = " & intCustomerID & ") ORDER BY ID DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	if rsPage.EOF then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but there are no sub-sites at the moment."))

%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			var Confirmation = confirm('If you remove this sub-site, you can NEVER get it back.  Are you completely sure?');
			if (Confirmation == true){
				return true;
			}
			else
				return false;
		}
	//-->
	</SCRIPT>


<%

	Set Command = Server.CreateObject("ADODB.Command")
	With Command
		.ActiveConnection = Connect
		.CommandText = "GetSiteTitle"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Title", adVarWChar, adParamOutput, 200 )

		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Subdirectory = rsPage("Subdirectory")

		PrintTableHeader 0
		PrintTableTitle
		do until rsPage.EOF
			.Parameters("@CustomerID") = ID
			.Execute , , adExecuteNoRecords
			strTitle = .Parameters("@Title")
%>
	<form METHOD="POST" ACTION="account_childsites_remove.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
		<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
		<input type="hidden" name="RemoveID" value="<%=ID%>">
		<input type="hidden" name="EMail" value="<%=strEMail%>">
		<input type="hidden" name="CCName" value="<%=strCCName%>">
		<input type="hidden" name="CCNumber" value="<%=strCCNumber%>">
		<tr>
			<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
			<td class="<% PrintTDMain %>"><a href="http://www.GroupLoop.com/<%=Subdirectory%>"><%=strTitle%></a></td>
			<td class="<% PrintTDMainSwitch %>">
				<input type="submit" name="Submit" value="Remove">
			</td>
		</tr>
	</form>

<%
			rsPage.MoveNext
		loop
		Response.Write("</table>")

	End With
	Set Command = Nothing

	rsPage.Close
	set rsPage = Nothing

else
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			var strError = "";
			if (form.EMail.value == "")
				strError += "          You forgot your e-mail address. \n";
			if (form.CCName.value == "")
				strError += "          You forgot your credit card name. \n";
			if (form.CCNumber.value == "")
				strError += "          You forgot your credit card number. \n";

			if(strError == "") {
				//Error message variable
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

	<p>Please remember that once you remove a sub-site, it is <b>permanently</b> deleted.  Be 
	completely sure before doing this.</p>
	<p>Before your can remove a sub-site, we must validate your account information.  Please enter 
	your information <b>exactly</b> like you did when you signed up.  Otherwise, it won't work.</p>

	<form METHOD="POST" ACTION="account_remove.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
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
				<input type="submit" name="Submit" value="Verify">
	   		</td>
		</tr>
  	</table>
	</form>
<%
end if

Function GetTitle( intCustomerID )
	Set Command = Server.CreateObject("ADODB.Command")
	With Command
		.ActiveConnection = Connect
		.CommandText = "GetSiteTitle"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Title", adVarWChar, adParamOutput, 200 )

		.Parameters("@CustomerID") = intCustomerID
		.Execute , , adExecuteNoRecords
		strTitle = .Parameters("@Title")
	End With
	Set Command = Nothing

	GetTitle = strTitle

End Function
%>

<!-- #include file="..\closedsn.asp" -->

<!-- #include file="footer.asp" -->