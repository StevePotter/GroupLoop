<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Bank Withdrawals</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=bankwithdrawals_modify.asp&ID=" & Request("ID") & "&Submit=" & Request("Submit"))
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

strSubmit = Request("Submit")

if strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, FileName FROM BankWithdrawals WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	strPath = GetPath ("posts")
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
		if FileSystem.FileExists( strPath & rsUpdate("FileName") ) then FileSystem.DeleteFile( strPath & rsUpdate("FileName") )
	Set FileSystem = Nothing

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close

	set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The withdrawal has been deleted. &nbsp;<a href="bankwithdrawals_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM BankWithdrawals WHERE ID = " & intID
	Set rsAccount = Server.CreateObject("ADODB.Recordset")
	rsAccount.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsAccount.EOF then
		set rsAccount = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
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

			form.Total.value = ConvertDollar(form.Total.value);


			if (form.Total.value == "" )
				strError += "          You forgot the total. \n";

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

	<form enctype="multipart/form-data" method="post" action="bankwithdrawals_modify_process.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="ID" value="<%=intID%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Bank Account
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintAccountsPullDown rsAccount("BankAccountID"), "BankAccountID" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Statement
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStatementsPullDown rsAccount("BankStatementID"), "BankStatementID" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date of Withdrawal
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "Date", rsAccount("Date"), 0 %>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Withdrawal(expense) Total
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="10" name="Total" value="<%=FormatCurrency(rsAccount("Total"))%>">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who was the payment to?
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="50" name="PaidTo" value="<%=rsAccount("PaidTo")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What form of payment was it?
			</td>
			<td class="<% PrintTDMain %>">
				<select name="PaymentType">
				<%
					WriteOption "Check", "Check", rsAccount("PaymentType")
					WriteOption "Credit Card", "Credit Card", rsAccount("PaymentType")
					WriteOption "ATM Card", "ATM Card", rsAccount("PaymentType")
					WriteOption "Cash", "Cash", rsAccount("PaymentType")
				%>
				</select> &nbsp; Details: <input type="text" size="50" name="PaidHow" value="<%=rsAccount("PaidHow")%>">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				If it was paid via check, what was the check number?
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="5" name="CheckNum" value="<%=rsAccount("CheckNum")%>">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				If the invoice or receipt has a number, enter it here.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="5" name="InvoiceReceivedID" value="<%=rsAccount("InvoiceReceivedID")%>">
			</td>
		</tr>



<%
	Set FileSystem = CreateObject("Scripting.FileSystemObject")

	blFile = False

	if rsAccount("FileName") <> "" then
		blFile = FileSystem.FileExists( GetPath("posts") & rsAccount("FileName") )
	end if

	if blFile then 
%>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">
			Should you keep the current file? 	<% PrintRadio 1, "UseFile" %><br>
			If you want to replace the current file with a new one, click Browse and select it.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File">
			</td>
		</tr>
<%
	else
%>	
		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				If it is stored on a file, click Browse and select it.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File">
			</td>
		</tr>

<%
	end if


	Set FileSystem = Nothing
%>
		<tr> 
     		<td class="<% PrintTDMain %>" align="right">Description</td>
     		<td class="<% PrintTDMain %>"> 
    				<textarea name="Description" cols="55" rows="2" wrap="PHYSICAL"><%=rsAccount("Description")%></textarea>
    		</td>
		</tr>

		<tr> 
     		<td class="<% PrintTDMain %>" align="right">Note</td>
     		<td class="<% PrintTDMain %>"> 
    				<textarea name="StaffNote" cols="55" rows="2" wrap="PHYSICAL"><%=rsAccount("StaffNote")%></textarea>
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
	Query = "SELECT * FROM BankWithdrawals ORDER BY Date DESC, ID DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsPage.EOF then
		Set rsPage = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("You have to have withdrawals before you can modify them."))
	end if

	Set ID = rsPage("ID")
	Set ItemDate = rsPage("Date")
	Set Description = rsPage("Description")
	Set AccountID = rsPage("BankAccountID")
	Set StatementID = rsPage("BankStatementID")
	Set Total = rsPage("Total")
	Set PaidTo = rsPage("PaidTo")
	Set PaymentType = rsPage("PaymentType")
	Set CheckNum = rsPage("CheckNum")
	Set FileName = rsPage("FileName")


	strPath = GetPath ("posts")
	Set FileSystem = CreateObject("Scripting.FileSystemObject")

'-----------------------End Code----------------------------
%>
	<p><a href="bankwithdrawals_add.asp">Add A Withdrawal</a></p>
	<form METHOD="POST" ACTION="bankwithdrawals_modify.asp">
<%
'-----------------------Begin Code----------------------------
	PrintPagesHeader
'-----------------------End Code----------------------------
%>
	<%PrintTableHeader 0%>
	<tr>
		<td class="TDHeader">&nbsp;</td>
		<td class="TDHeader">Account</td>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">Paid To</td>
		<td class="TDHeader">Paid By</td>
		<td class="TDHeader">Total</td>
		<td class="TDHeader">Description</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
	for p = 1 to rsPage.PageSize
		if not rsPage.EOF then
			blExists = FileSystem.FileExists( strPath & rsPage("FileName") )
'------------------------End Code-----------------------------
%>
		<form METHOD="post" ACTION="bankwithdrawals_modify.asp">
		<input type="hidden" name="ID" value="<%=ID%>">
			<tr>
				<td class="<% PrintTDMain %>">
<%
				if blExists then
					%><a href="posts/<%=FileName%>">View</a><%
				else
					Response.Write "&nbsp;"
				end if
%>

				</td>
				<td class="<% PrintTDMain %>"><%=GetAccountName( AccountID )%></td>
				<td class="<% PrintTDMain %>"><%=FormatDateTime(ItemDate, 2)%></td>
				<td class="<% PrintTDMain %>"><%=PaidTo%></td>
				<td class="<% PrintTDMain %>"><%=PaymentType%>
<%
				if not IsNull(CheckNum) and PaymentType = "Check" then
					if CheckNum > 0 then
						%>&nbsp; Check # <%=CheckNum%><%
					end if
				end if
%>
				
				</td>
				<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
				<td class="<% PrintTDMain %>"><%=Description%></td>
				<td class="<% PrintTDMainSwitch %>">
				<input type="submit" name="Submit" value="Edit"> 
				<input type="button" value="Delete" onClick="DeleteBox('If you delete this withdrawal, there is no way to get it back.  Are you sure?', 'bankwithdrawals_modify.asp?Submit=Delete&ID=<%=ID%>')">				
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

	Set FileSystem = Nothing

end if

'------------------------End Code-----------------------------
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->