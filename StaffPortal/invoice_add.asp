<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Add Invoice Charge</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=invoice_add.asp&CustomerID=" & Request("CustomerID"))

strSubmit = Request("Submit")

if strSubmit = "Add" then
	if Request("CustomerID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the CustomerID."))
	intCustomerID = CInt(Request("CustomerID"))

	Set cmdReviews = Server.CreateObject("ADODB.Command")
	With cmdReviews
		.ActiveConnection = Connect
		.CommandText = "AddCustomerInvoice"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CustomerID") = intCustomerID
		.Parameters("@Description") = Format(Request("Description"))
		.Parameters("@CustomerNote") = Format(Request("CustomerNote"))
		.Parameters("@StaffNote") = Format(Request("StaffNote"))
		.Parameters("@Sent") = 0

		if Request("BillingTypeOther") <> "" then
			.Parameters("@BillingType") = Request("BillingTypeOther")
		else
			.Parameters("@BillingType") = Request("BillingType")
		end if

		.Execute , , adExecuteNoRecords

		intInvoiceID = .Parameters("@InvoiceID")

	End With

	Set cmdReviews = Nothing

%>
<p>The invoice has been added.<br>
<a href="invoices_modify.asp?Submit=Edit&ID=<%=intInvoiceID%>">View the invoice.</a><br>
<a href="customer_view.asp?ID=<%=intCustomerID%>">View this customer's details.</a><br>

<a href="customers.asp">Browse the list of customers.</a>
</p>

<%
else
%>


	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";

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

<%
	Query = "SELECT CustomerID, Date, DateSent, DateReceived, InvoiceSent, Paid, BillingType, TransID, CheckNum, " & _
	"MaintenanceID, AutomaticallyCreated, CustomerNote, StaffNote, Description " & _
	"FROM CustomerInvoices WHERE ID = " & intID

%>

	* indicates required information<br>
	<form method="post" action="invoice_add.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">

	<% PrintTableHeader 0 %>
<%
	if Request("CustomerID") = "" then
%>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Customer</td>
     	<td class="<% PrintTDMain %>"> 
    		<% PrintCustomerPullDown 0, 1, 0, "", "" %>
    	</td>
	</tr>
<%
	else
		intCustomerID = CInt(Request("CustomerID"))
%>
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
<%
	end if
%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Payment Method</td>
     	<td class="<% PrintTDMain %>"> 
			<select name="BillingType" 	onChange="if (this.form.BillingType.value == 'Other') this.form.BillingTypeOther.focus();" >
<%
				BillingType = "Check"
				WriteOption "CreditCard", "Credit Card", BillingType
				WriteOption "Check", "Check", BillingType
				WriteOption "Other", "Other", BillingType

				if BillingType <> "CreditCard" and BillingType <> "Check" then
					strOther = BillingType
				else
					strOther = ""
				end if
%>
				</select>	Other <input type="text" name="BillingTypeOther" size="15" value="<%=strOther%>">
     	</td>
	</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Description/Summary</td>
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
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Add">
    		</td>
		</tr>
  	</table>
	</form>



<%
end if
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->













