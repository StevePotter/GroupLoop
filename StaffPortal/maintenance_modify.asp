<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Maintenance Runs</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if


if not LoggedStaff() then Redirect("login.asp?Source=maintenance_modify.asp&ID=" & Request("ID"))

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
	<a href="maintenance_modify.asp?Submit=Edit&ID=<%=intID%>">Edit the invoice again.<br>
	<a href="customer_view.asp?ID=<%=intCustomerID%>">View the customer's details.<br>
	<a href="customers.asp">Browse the list of customers.
	</p>

<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM NightlyMaintenance WHERE ID = " & intID
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The maintenance run has been deleted.<br>  
	<a href="maintenance_modify.asp">Back to the list.<br>
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



	//-->
	</SCRIPT>

	<form method="post" action="maintenance_modify.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">

	<p align="<%=HeadingAlignment%>" class=Heading>Invoice Information</p>
	<input type="button" value="Printable Version" onClick="Redirect('invoice_print.asp?InvoiceID=<%=intID%>')"><br>

	<b>Current invoice total: <%=FormatCurrency(rsEdit("Total"))%></b><br>

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
					<input type="button" value="Edit" onClick="Redirect('invoicecharges_modify.asp?Submit=Edit&ID=<%=ID%>')">
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

	Query = "SELECT ID, Date, CardCharges, Invoices, Errors, ExpirationWarnings, ProcessDate FROM NightlyMaintenance ORDER BY ProcessDate DESC"


	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set CardCharges = rsPage("CardCharges")
		Set Invoices = rsPage("Invoices")
		Set Errors = rsPage("Errors")
		Set ExpirationWarnings = rsPage("ExpirationWarnings")
		Set ProcessDate = rsPage("ProcessDate")
'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="maintenance_modify.asp">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">&nbsp;</td>
			<td class="TDHeader">Date Run For</td>
			<td class="TDHeader">Date Run</td>
			<td class="TDHeader">Card Charges</td>
			<td class="TDHeader">Invoices</td>
			<td class="TDHeader">Expiration Warnings</td>
			<td class="TDHeader">Errors</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<form METHOD="post" ACTION="maintenance_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="maintenance_view.asp?ID=<%=ID%>">View</a></td>
						<td class="<% PrintTDMain %>"><%=FormatDateTime( rsPage("ProcessDate"), 2)%></td>
						<td class="<% PrintTDMain %>"><%=FormatDateTime( rsPage("Date"), 2)%></td>
						<td class="<% PrintTDMain %>"><%=CardCharges%></td>
						<td class="<% PrintTDMain %>"><%=Invoices%></td>
						<td class="<% PrintTDMain %>"><%=ExpirationWarnings%></td>
						<td class="<% PrintTDMain %>"><%=Errors%></td>
						<td class="<% PrintTDMainSwitch %>"><input type="Submit" name="Submit" value="Edit"> 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this maintenance run, it cannot be brought back.  Are you sure?', 'maintenance_modify.asp?Submit=Delete&ID=<%=ID%>')"></td>
					</tr>
				</form>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
	end if

	set rsPage = Nothing
end if

%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->