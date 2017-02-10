<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Deposit</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=bankdeposits_add.asp")
if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
'------------------------End Code-----------------------------

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

	<form enctype="multipart/form-data" method="post" action="bankdeposits_add_process.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Deposit Total</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Total" size="5" value="$">
     		</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Bank Account
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintAccountsPullDown 0, "BankAccountID" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Statement (if any)
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStatementsPullDown 0, "BankStatementID" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date of Deposit
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "DateDeposited", Date, 0 %>
			</td>
		</tr>


		<tr> 
			<td class="<% PrintTDMain %>" align="right">Customer (if any)</td>
			<td class="<% PrintTDMain %>"> 
				<% PrintCustomerPullDown 0, 0, 0, "None", "CustomerID" %>
			</td>
		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Customer Invoice (if any)</td>
			<td class="<% PrintTDMain %>"> 
				<% PrintInvoicePullDown 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What form of payment was it?
			</td>
			<td class="<% PrintTDMain %>">
				<select name="BillingType" onChange="if (getSelectValue(this) == 'Check') show('checkNum'); else hide('checkNum');">

				<%
					WriteOption "Check", "Check", ""
					WriteOption "Credit Card", "Credit Card", ""
					WriteOption "ATM Card", "ATM Card", ""
					WriteOption "Cash", "Cash", ""
				%>
				</select> &nbsp; 
				<span id="checkNum">
						Check number (from customer) <input type="text" size="5" name="CheckNum" value=""><br>
				</span>
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				If the deposit is stored on a file, click Browse and select it.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File">
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
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Add">
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------

Sub PrintInvoicePullDown( intHighLightID )
	Set rsPulldown = Server.CreateObject("ADODB.Recordset")
	rsPulldown.CacheSize = 150


	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetInvoiceRecordset"
	cmdTemp.CommandType = adCmdStoredProc

	cmdTemp.Parameters.Refresh
	cmdTemp.Parameters("@InvoiceID") = intHighLightID
	cmdTemp.Parameters("@CustomerID") = 0


	rsPulldown.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	Set cmdTemp = Nothing

	if rsPulldown.EOF then
		Set rsPulldown = Nothing
		Exit Sub
	end if

	Set ID = rsPulldown("ID")
	Set Total = rsPulldown("Total")
	Set Description = rsPulldown("Description")

	%><select name="InvoiceID" size="1"><%
Response.Write "<option value = ''>None</option>" & vbCrlf
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