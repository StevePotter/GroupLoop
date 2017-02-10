<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Employees</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=bankaccounts_modify.asp&ID=" & Request("ID") & "&Submit=" & Request("Submit"))
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

strSubmit = Request("Submit")

'We are going to check for errors if they are updating the profile
if strSubmit = "Update" then
	Query = "SELECT * FROM BankAccounts WHERE ID = " & Request("ID")
	Set rsAccount = Server.CreateObject("ADODB.Recordset")
	rsAccount.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect

	rsAccount("DateCreated") = AssembleDate("DateCreated")

	rsAccount("Type") = Request("Type")
	rsAccount("AccountNumber") = Request("AccountNumber")
	rsAccount("ClientNumber") = Request("ClientNumber")
	rsAccount("Description") = Format(Request("Description"))
	rsAccount("BankName") = Format(Request("BankName"))
	rsAccount("BankStreet") = Format(Request("BankStreet"))
	rsAccount("BankCity") = Format(Request("BankCity"))
	rsAccount("BankState") = Format(Request("BankState"))
	rsAccount("BankZip") = Format(Request("BankZip"))
	rsAccount("BankPhone") = Format(Request("BankPhone"))
	rsAccount("BankWebSite") = Format(Request("BankWebSite"))


	rsAccount.Update
	rsAccount.Close
	set rsAccount = Nothing
'------------------------End Code-----------------------------
%>
	<p>The account has been edited. &nbsp;<a href="bankaccounts_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM BankAccounts WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close

	set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The account has been deleted. &nbsp;<a href="bankaccounts_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM BankAccounts WHERE ID = " & intID
	Set rsAccount = Server.CreateObject("ADODB.Recordset")
	rsAccount.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsAccount.EOF then
		set rsAccount = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
%>

	<form METHOD="post" ACTION="bankaccounts_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=intID%>">
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
				<input type="text" size="40" name="AccountNumber"  value="<%=rsAccount("AccountNumber")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Client Number
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="ClientNumber"  value="<%=rsAccount("ClientNumber")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Account Type
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Type"  value="<%=rsAccount("Type")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date Created
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "DateCreated", rsAccount("DateCreated"), 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Account Description (account name used when listing - ex First Union Checking)
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Description"  value="<%=rsAccount("Description")%>">
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
				<input type="text" size="40" name="BankName"  value="<%=rsAccount("BankName")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Street
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="BankStreet"  value="<%=rsAccount("BankStreet")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				City
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="BankCity" value="<%=rsAccount("BankCity")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				State
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStates "BankState", rsAccount("BankState") %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Zip Code
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="BankZip" value="<%=rsAccount("BankZip")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Phone
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="15" name="BankPhone" value="<%=rsAccount("BankPhone")%>">

			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Web Site
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="BankWebSite" value="<%=rsAccount("BankWebSite")%>">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Update">
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsAccount.Close
	Set rsAccount = Nothing

else
	Query = "SELECT ID, AccountNumber, ClientNumber, Description FROM BankAccounts ORDER BY DateCreated"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsPage.EOF then
		Set rsPage = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("You have to have employees before you can modify them."))
	end if

	Set ID = rsPage("ID")
	Set AccountNumber = rsPage("AccountNumber")
	Set ClientNumber = rsPage("ClientNumber")
	Set Description = rsPage("Description")
'-----------------------End Code----------------------------
%>
	<p><a href="bankaccounts_add.asp">Add An Account</a></p>
	<form METHOD="POST" ACTION="bankaccounts_modify.asp">
<%
'-----------------------Begin Code----------------------------
	PrintPagesHeader
'-----------------------End Code----------------------------
%>
	<%PrintTableHeader 0%>
	<tr>
		<td class="TDHeader">Description</td>
		<td class="TDHeader">Account Number</td>
		<td class="TDHeader">Client Number</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
	for p = 1 to rsPage.PageSize
		if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
		<form METHOD="post" ACTION="bankaccounts_modify.asp">
		<input type="hidden" name="ID" value="<%=ID%>">
			<tr>
				<td class="<% PrintTDMain %>"><%=Description%></td>
				<td class="<% PrintTDMain %>"><%=AccountNumber%></td>
				<td class="<% PrintTDMain %>"><%=ClientNumber%></td>
				<td class="<% PrintTDMainSwitch %>">
				<input type="submit" name="Submit" value="Edit"> 
				<input type="button" value="Delete" onClick="DeleteBox('If you delete this account, there is no way to get it back.  Are you sure?', 'bankaccounts_modify.asp?Submit=Delete&ID=<%=ID%>')">				
				</td>
			</tr>
		</form>
<%
'-----------------------Begin Code----------------------------
		rsPage.MoveNext
		end if
	next
	Response.Write("</table>")
	rsPage.Close
	set rsPage = Nothing
end if

'------------------------End Code-----------------------------
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->