<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add a Withdrawal</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
if not LoggedStaff() then Redirect("login.asp?Source=bankwithdrawals_add.asp")
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

	<form enctype="multipart/form-data" method="post" action="bankwithdrawals_add_process.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
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
				Statement
			</td>
			<td class="<% PrintTDMain %>">
				<% PrintStatementsPullDown 0, "BankStatementID" %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Date of Withdrawal
			</td>
			<td class="<% PrintTDMain %>">
				<% DatePulldown "Date", Date, 0 %>
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Withdrawal(expense) Total
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="10" name="Total" value="$">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Who was the payment to?
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="50" name="PaidTo" value="">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				What form of payment was it?
			</td>
			<td class="<% PrintTDMain %>">
				<select name="PaymentType" onChange="if (getSelectValue(this) == 'Check') show('checkNum'); else hide('checkNum');">

				<%
					WriteOption "Check", "Check", ""
					WriteOption "Credit Card", "Credit Card", ""
					WriteOption "ATM Card", "ATM Card", ""
					WriteOption "Cash", "Cash", ""
				%>
				</select> &nbsp; 
				<span id="checkNum">
						Check number <input type="text" size="5" name="CheckNum" value=""><br>
				</span>
				 Details: <input type="text" size="50" name="PaidHow" value="">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				If the invoice or receipt is stored on a file, click Browse and select it.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				If the invoice or receipt has a number, enter it here.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="5" name="InvoiceReceivedID" value="">
			</td>
		</tr>

		<tr> 
     		<td class="<% PrintTDMain %>" align="right">Description</td>
     		<td class="<% PrintTDMain %>"> 
    				<textarea name="Description" cols="55" rows="2" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>

		<tr> 
     		<td class="<% PrintTDMain %>" align="right">Note</td>
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


'------------------------End Code-----------------------------
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->