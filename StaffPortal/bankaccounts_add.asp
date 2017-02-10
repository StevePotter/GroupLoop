<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Bank Account</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=bankaccounts_add.asp&ID=" & intID)
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

'We are going to check for errors if they are updating the profile
if Request("Submit") = "Add" then
	strNickName = Format(Request("NickName"))

	if EmployeeNickNameTaken( strNickName ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but the nickname " & strNickName & " is already taken."))

	Query = "SELECT * FROM BankAccounts"
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	rsNew.AddNew

	rsNew("DateCreated") = AssembleDate("DateCreated")

	rsNew("Type") = Request("Type")
	rsNew("AccountNumber") = Request("AccountNumber")
	rsNew("ClientNumber") = Request("ClientNumber")
	rsNew("Description") = Format(Request("Description"))
	rsNew("BankName") = Format(Request("BankName"))
	rsNew("BankStreet") = Format(Request("BankStreet"))
	rsNew("BankCity") = Format(Request("BankCity"))
	rsNew("BankState") = Format(Request("BankState"))
	rsNew("BankZip") = Format(Request("BankZip"))
	rsNew("BankPhone") = Format(Request("BankPhone"))
	rsNew("BankWebSite") = Format(Request("BankWebSite"))
	rsNew("EmployeeID") = Session("EmployeeID")

	rsNew.Update
	rsNew.Close
	set rsNew = Nothing
'------------------------End Code-----------------------------
%>
	<p>The account has been added. &nbsp;<a href="bankaccounts_add.asp">Add another.</a><br>
	<a href="bankaccounts_modify.asp">Modify accounts.</a>
	</p>
<%
'-----------------------Begin Code----------------------------

else
'------------------------End Code-----------------------------
%>
	<form METHOD="post" ACTION="bankaccounts_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Account Information
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Account Number
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="20" name="AccountNumber" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Client Number
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="20" name="ClientNumber" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Account Type
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Type" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date Created
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "DateCreated", Date, 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Account Description (account name used when listing - ex First Union Checking)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Description" >
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Bank Information
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Bank Name
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="BankName" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="BankStreet" >
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="BankCity">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "BankState", "" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="8" name="BankZip">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="BankPhone">

			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Web Site
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="BankWebSite">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Add">
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------

end if

'------------------------End Code-----------------------------
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->