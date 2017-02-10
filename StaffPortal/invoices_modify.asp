<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Customer Billing</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if


if not LoggedStaff() then Redirect("login.asp?Source=invoices_modify.asp&ID=" & Request("ID") & "&Submit=" & Request("Submit"))

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT CustomerID, Date, DateSent, DateReceived, InvoiceSent, Paid, BillingType, TransID, CheckNum, " & _
	"MaintenanceID, AutomaticallyCreated, CustomerNote, StaffNote, Description " & _
	"FROM CustomerInvoices WHERE ID = " & intID

	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	intCustomerID = rsUpdate("CustomerID")

	rsUpdate("Date") = AssembleDate("Date")
	rsUpdate("DateSent") = AssembleDate("DateSent")
	rsUpdate("DateReceived") = AssembleDate("DateReceived")
	rsUpdate("InvoiceSent") = cInt(Request("InvoiceSent"))
	rsUpdate("Paid") = cInt(Request("Paid"))
	rsUpdate("MaintenanceID") = cInt(Request("MaintenanceID"))
	rsUpdate("InvoiceSent") = cInt(Request("InvoiceSent"))
	rsUpdate("CheckNum") = cInt(Request("CheckNum"))
	rsUpdate("CustomerID") = cInt(Request("CustomerID"))

	rsUpdate("TransID") = Request("TransID")

	if Request("BillingTypeOther") <> "" then
		strBillingType = Request("BillingTypeOther")
	else
		strBillingType = Request("BillingType")
	end if
	rsUpdate("BillingType") = strBillingType


	rsUpdate("CustomerNote") = Format( Request("CustomerNote") )
	rsUpdate("StaffNote") = Format( Request("StaffNote") )
	rsUpdate("Description") = Format( Request("Description") )


	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The invoice has been edited.  
	<a href="invoice_print.asp?InvoiceID=<%=intID%>">Print the invoice.<br>
	<a href="invoices_modify.asp?Submit=Edit&ID=<%=intID%>">Edit the invoice again.<br>
	<a href="customer_view.asp?ID=<%=intCustomerID%>">View the customer's details.<br>
	<a href="customers.asp">Browse the list of customers.
	</p>

<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, CustomerID FROM CustomerInvoices WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	intCustomerID = rsUpdate("CustomerID")
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE CustomerInvoiceCharges WHERE InvoiceID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

'------------------------End Code-----------------------------
%>
	<p>The invoice has been deleted.  
	<a href="customer_view.asp?ID=<%=intCustomerID%>">View the customer's details.<br>
	<a href="customers.asp">Browse the list of customers.
	</p>

<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM CustomerInvoices WHERE ID = " & intID
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	if rsEdit.EOF then
		Set rsEdit = Nothing
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


		function PutDate(form, Field){
			var date = new Date();
			var d  = date.getDate();
			var day = (d < 10) ? '0' + d : d;
			var m = date.getMonth() + 1;
			var month = (m < 10) ? '0' + m : m;
			var yy = date.getYear();
			var year = (yy < 1000) ? yy + 1900 : yy;



			myhours = date.getHours();
			if (myhours >= 12) {
			myhours = (myhours == 12) ? 12 : myhours - 12; mm = " PM";
			}
			else {
			myhours = (myhours == 0) ? 12 : myhours; mm = " AM";
			}
			myminutes = date.getMinutes();
			if (myminutes < 10){
			myminutes = ":0" + myminutes;
			}
			else {
			myminutes = ":" + myminutes;
			};
			mysecs = date.getSeconds();
			if (mysecs < 10){
			mysecs = ":0" + mysecs;
			}
			else {
			mysecs = ":" + mysecs;
			};

			form.elements[Field + 'Month'].value = m;
			form.elements[Field + 'Day'].value = d;
			form.elements[Field + 'Year'].value = year;


			if ( form.elements[Field + 'Time'] )
				form.elements[Field + 'Time'].value = myhours+myminutes+mysecs+mm;

			return;
		}


	//-->
	</SCRIPT>

	<form method="post" action="invoices_modify.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">

	<p align="<%=HeadingAlignment%>" class=Heading>Invoice Information</p>
	<input type="button" value="Printable Version" onClick="Redirect('invoice_print.asp?InvoiceID=<%=intID%>')">&nbsp;&nbsp;&nbsp;&nbsp;

	<input type="button" value="Charge For This Invoice" onClick="Redirect('customer_charge.asp?CustomerID=<%=rsEdit("CustomerID")%>&InvoiceID=<%=intID%>')">&nbsp;&nbsp;&nbsp;&nbsp;
<%
	if rsEdit("InvoiceSent") = 0 then
%>
	<input type="button" value="Send Invoice" onClick="Redirect('invoice_send.asp?ID=<%=intID%>')">&nbsp;&nbsp;&nbsp;&nbsp;
<%
	end if
%>
	<br>
	<b>Current invoice total: <%=FormatCurrency(rsEdit("Total"))%></b><br>

	<% PrintTableHeader 100 %>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Customer</td>
     	<td class="<% PrintTDMain %>"> 
    		<% PrintCustomerPullDown rsEdit("CustomerID"), 1, 0, "", "" %>
    	</td>
	</tr>
	</table>
	<div ID="contactParent" NAME="contactParent" CLASS=parent>
	<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('contact'); return false" ID="contactIm">
		Dates</a>
		</td></tr></table>
	</div>
	<div ID="contactChild" NAME="contactChild" CLASS=child>
	<% PrintTableHeader 100 %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Date Created</td>
      	<td class="<% PrintTDMain %>"><% DatePulldown "Date", rsEdit("Date"), 1 %>&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Now" onClick="PutDate(this.form, 'Date')">
		</td>
    </tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Date Sent</td>
      	<td class="<% PrintTDMain %>"><% DatePulldown "DateSent", rsEdit("DateSent"), 1 %>&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Now" onClick="PutDate(this.form, 'DateSent')">
		</td>
    </tr>
	<tr>
      	<td class="<% PrintTDMain %>" align="right">Date Recieved</td>

       	<td class="<% PrintTDMain %>"><% DatePulldown "DateReceived", rsEdit("DateReceived"), 1 %>&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="Now" onClick="PutDate(this.form, 'DateReceived')">
		</td>
    </tr>
	</table>
	</div>

	<div ID="payParent" NAME="payParent" CLASS=parent>
	<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('pay'); return false" ID="payIm">
		Payment</a>
		</td></tr></table>
	</div>
	<div ID="payChild" NAME="payChild" CLASS=child>
	<% PrintTableHeader 100 %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Invoice Sent?</td>
      	<td class="<% PrintTDMain %>"><% PrintRadio rsEdit("InvoiceSent"), "InvoiceSent" %></td>
    </tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Paid?</td>
      	<td class="<% PrintTDMain %>"><% PrintRadio rsEdit("Paid"), "Paid" %></td>
    </tr>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Payment Method</td>
     	<td class="<% PrintTDMain %>"> 
			<select name="BillingType" 	onChange="if (this.form.BillingType.value == 'Other') this.form.BillingTypeOther.focus();" >
<%
				BillingType = rsEdit("BillingType")
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
		<td class="<% PrintTDMain %>" align="right">Credit Card Transaction ID</td>
		<td class="<% PrintTDMain %>"> 
			<input type="text" name="TransID" value="<%=rsEdit("TransID")%>" size="15">
		</td>
   	</tr>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Check Number</td>
		<td class="<% PrintTDMain %>"> 
			<input type="text" name="CheckNum" value="<%=rsEdit("CheckNum")%>" size="5">
		</td>
   	</tr>
	</table>

	</div>

	<div ID="otherParent" NAME="otherParent" CLASS=parent>
	<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('other'); return false" ID="otherIm">
		Other</a>
		</td></tr></table>
	</div>
	<div ID="otherChild" NAME="otherChild" CLASS=child>
	<% PrintTableHeader 100 %>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Maintenance ID</td>
		<td class="<% PrintTDMain %>"> 
			<input type="text" name="MaintenanceID" value="<%=rsEdit("MaintenanceID")%>" size="5">
		</td>
   	</tr>
	</table>

	</div>

	<div ID="descriptionParent" NAME="descriptionParent" CLASS=parent>
	<% PrintTableHeader 100 %>
		<tr><td class="TDHeader">
		<a class="TDHeader" HREF="javascript://" onClick="expandIt('description'); return false" ID="descriptionIm">
		Descriptions</a>
		</td></tr></table>
	</div>
	<div ID="descriptionChild" NAME="descriptionChild" CLASS=child>
	<% PrintTableHeader 100 %>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Note to Customer</td>
     	<td class="<% PrintTDMain %>"> 
    			<textarea name="CustomerNote" cols="55" rows="2" wrap="PHYSICAL"><%=FormatEdit( rsEdit("CustomerNote") )%></textarea>
    	</td>
	</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Note to Staff Only</td>
     	<td class="<% PrintTDMain %>"> 
    			<textarea name="StaffNote" cols="55" rows="2" wrap="PHYSICAL"><%=FormatEdit( rsEdit("StaffNote") )%></textarea>
    	</td>
	</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Description/Summary</td>
     	<td class="<% PrintTDMain %>"> 
    			<textarea name="Description" cols="55" rows="2" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Description") )%></textarea>
    	</td>
	</tr>
  	</table>
	</div>

	<% PrintTableHeader 100 %>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
	</tr>
  	</table>

	<p align="<%=HeadingAlignment%>" class=Heading>Invoice Charges</p>
	<input type="button" value="Add Charge" onClick="Redirect('invoicecharge_add.asp?InvoiceID=<%=intID%>')">&nbsp;&nbsp;
<%
	Query = "SELECT ID, Date, Hours, Description, StaffNote, CustomerNote, DateStarted, DateEnded, Total FROM CustomerInvoiceCharges WHERE InvoiceID = " & intID & " ORDER BY ID"
	rsEdit.Close
	rsEdit.CacheSize = PageSize
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsEdit.EOF then

		Set ID = rsEdit("ID")
		Set ItemDate = rsEdit("Date")
		Set Hours = rsEdit("Hours")
		Set Description = rsEdit("Description")
		Set StaffNote = rsEdit("StaffNote")
		Set CustomerNote = rsEdit("CustomerNote")
		Set DateStarted = rsEdit("DateStarted")
		Set DateEnded = rsEdit("DateEnded")
		Set Total = rsEdit("Total")

		PrintTableHeader 100
%>
		<tr>
			<td class="TDHeader">Total</td>
			<td class="TDHeader">Dates</td>
			<td class="TDHeader">Description</td>
			<td class="TDHeader">Notes</td>
			<td class="TDHeader">Hours Spent</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		dblRunningTotal = 0

		do until rsEdit.EOF
			dblRunningTotal = dblRunningTotal + rsEdit("Total")
			if Hours > 0 then
				strHours = Hours & " Hours"
			else
				strHours = "N/A or recorded"
			end if

			strDates = "Created on:" & FormatDateTime(ItemDate, 0)

			if Hours > 0 then
				if not IsNull(DateStarted) then strDates = strDates & "<br>Work started on:" & FormatDateTime(DateStarted, 0)
				if not IsNull(DateEnded) then strDates = strDates & "<br>Work ended on:" & FormatDateTime(DateEnded, 0)
			end if
			if CustomerNote <> "" and StaffNote <> "" then
				strNote = "<b>To customer:</b>&nbsp; " & CustomerNote & "<br><br><b>To staff:</b>&nbsp; " & StaffNote
			elseif CustomerNote <> "" then
				strNote = "<b>To customer:</b>&nbsp; " & CustomerNote
			elseif CustomerNote <> "" then
				strNote = "<b>To staff:</b>&nbsp; " & StaffNote
			else
				strNote = "&nbsp;"
			end if
'------------------------End Code-----------------------------
%>
			<tr>
				<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
				<td class="<% PrintTDMain %>"><%=strDates%></td>
				<td class="<% PrintTDMain %>"><%=Description%></td>
				<td class="<% PrintTDMain %>"><%=strNote%></td>
				<td class="<% PrintTDMain %>"><%=strHours%></td>
				<td class="<% PrintTDMainSwitch %>">
					<input type="button" value="Edit" onClick="Redirect('invoicecharges_modify.asp?Submit=Edit&ID=<%=rsEdit("ID")%>')">
					<input type="button" value="Delete" onClick="DeleteBox('If you delete this charge, there is no way to get it back.  Are you sure?', 'invoicecharges_modify.asp?Submit=Delete&ID=<%=ID%>')">
				</td>
			</tr>
<%
'-----------------------Begin Code----------------------------
			rsEdit.MoveNext
		loop
		if rsEdit.RecordCount > 1 then
%>
			<tr>
				<td class="TDHeader" colspan="6" align="left">Total - <%=FormatCurrency(dblRunningTotal)%></td>
			</tr>
<%
		end if
		Response.Write("</table>")
	else
'------------------------End Code-----------------------------
%>
		<p><b>No invoice charges.</b></p>
<%
'-----------------------Begin Code----------------------------
	end if

%>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing

else
	strType = Request("Type")
	strChecked1 = ""
	strChecked2 = ""
	strChecked3 = ""
	strChecked34= ""
	if strType = "Paid" then
		strChecked1 = "checked"
		strWhere = "WHERE Paid = 1"
	elseif strType = "Outstanding" then
		strChecked2 = "checked"
		strWhere = "WHERE Paid = 0 AND InvoiceSent = 1"
	elseif strType = "Current" then
		strChecked3 = "checked"
		strWhere = "WHERE Paid = 0 AND InvoiceSent = 0"
	else
		strChecked4 = "checked"
	end if
%>
		<p><a href="invoice_add.asp">Add New Invoice</a></p>

		<form METHOD="POST" ACTION="invoices_modify.asp">
		<p>View:<br>
		<input type="radio" name="Type" value="" <%=strChecked4%> onClick="this.form.submit();">All Invoices<br>
		<input type="radio" name="Type" value="Outstanding" <%=strChecked2%> onClick="this.form.submit();">Outstanding Invoices<br>
		<input type="radio" name="Type" value="Current" <%=strChecked3%> onClick="this.form.submit();">Current Invoices<br>
		<input type="radio" name="Type" value="Paid" <%=strChecked1%> onClick="this.form.submit();">Paid Invoices</p>
<%
	Query = "SELECT ID, CustomerID, Description, MaintenanceID, Total " & _
	"FROM CustomerInvoices " & strWHere & " ORDER BY ID Desc"

	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set CustomerID = rsPage("CustomerID")
		Set Description = rsPage("Description")
		Set MaintenanceID = rsPage("MaintenanceID")
		Set Total = rsPage("Total")


		Set cmd = Server.CreateObject("ADODB.Command")	'used for the GetCustSummary.  this way we create/destroy object once

		PrintPagesHeader
		PrintTableHeader 0

		intRunningTotal = 0
%>
		<tr>
			<td class="TDHeader">Invoice ID</td>
			<td class="TDHeader">Customer</td>
			<td class="TDHeader">Total</td>
			<td class="TDHeader">Description</td>
			<td class="TDHeader">Maintenance ID</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
				if MaintenanceID = 0 then
					strMain = ""
				else
					strMain = "<a href=maintenance_view.asp?ID=" & MaintenanceID & ">" & MaintenanceID & "</a>"
				end if
				intRunningTotal = intRunningTotal + Total
'------------------------End Code-----------------------------
%>
				<form METHOD="post" ACTION="invoices_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>"><%=ID%></td>
						<td class="<% PrintTDMain %>"><%=GetCustSummary( CustomerID )%></td>
						<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
						<td class="<% PrintTDMain %>"><%=Description%></td>
						<td class="<% PrintTDMain %>"><%=strMain%></td>
						<td class="<% PrintTDMainSwitch %>"><input type="Submit" name="Submit" value="Edit"> 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this charge, there is no way to get it back.  Are you sure?', 'invoices_modify.asp?Submit=Delete&ID=<%=ID%>')"></td>
					</tr>
				</form>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
%>
					<tr>
						<td class="TDHeader" colspan=2 align=right>Total: <%=FormatCurrency(intRunningTotal)%></td>
						<td class="TDHeader" colspan=3>&nbsp</td>
					</tr>

<%
		Response.Write("</table>")
		rsPage.Close

		Set cmd = Nothing
	else
'------------------------End Code-----------------------------
%>
		<p>You have to create invoices for you can modify them, <%=GetNickNameSession()%>.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if

%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->