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

if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
intID = CInt(Request("ID"))

if not LoggedStaff() then Redirect("login.asp?Source=customer_billing.asp&ID=" & intID)

strSubmit = Request("Submit")

if strSubmit = "Update" then


else
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = 100

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	cmdTemp.ActiveConnection = Connect
	cmdTemp.CommandText = "GetSiteInfoRecordSet"
	cmdTemp.CommandType = adCmdStoredProc

	rsPage.Open cmdTemp, , adOpenStatic, adLockReadOnly, adCmdTableDirect

	With cmdTemp
		.CommandText = "CustomerInvoicesExist"
		.Parameters.Refresh

		.Parameters("@CustomerID") = intID

		.Execute , , adExecuteNoRecords
		blInvoicesExist = CBool(.Parameters("@InvoicesExist"))
		intPast = .Parameters("@PastExist")
		intCurrent = .Parameters("@CurrentExist")
		intOutstanding = .Parameters("@OutstandingExist")

	End With

	Set cmdTemp = Nothing

	rsPage.Filter = "ID = " & intID


	if rsPage("UseDomain") = 1 then
		strAddress = rsPage("DomainName")
	else
		strAddress = "http://www.GroupLoop.com/" & rsPage("SubDirectory")
	end if
%>

	<p>
	<i>Customer Summary:</i><br>
	&nbsp;&nbsp;&nbsp;Customer ID: <%=intID%><br>
	&nbsp;&nbsp;&nbsp;Site Address: <a href="<%=strAddress%>"><%=strAddress%></a><br>
	&nbsp;&nbsp;&nbsp;Date Created: <%=FormatDateTime(rsPage("SignupDate"), 2)%><br>
	&nbsp;&nbsp;&nbsp;Owner Name: <%=rsPage("FirstName")%>&nbsp;<%=rsPage("LastName")%><br>
	&nbsp;&nbsp;&nbsp;Site Title: <%=rsPage("Title")%><br>
	&nbsp;&nbsp;&nbsp;Contact E-Mail: <a href="mailto:<%=rsPage("EMail")%>"><%=rsPage("EMail")%></a>
	</p>



	<p>
	<hr>
	<i>Billing Summary:</i><br>
	<input type="button" value="New Invoice" onClick="Redirect('invoice_add.asp?CustomerID=<%=intID%>')"><br>
<%
	Query = "SELECT ID, Date, Total, Description, BillingType, TransID, Paid, CheckNum, DateReceived, MaintenanceID, InvoiceSent, DateSent FROM CustomerInvoices WHERE CustomerID = " & intID & " ORDER BY ID DESC"

	rsPage.Close
	Set rsPage = Nothing
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = 100
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
	%>
		<b><%=rsPage.RecordCount%> total invoices.</b><br>
	<%
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Total = rsPage("Total")
		Set Description = rsPage("Description")
		Set BillingType = rsPage("BillingType")
		Set TransID = rsPage("TransID")
		Set Paid = rsPage("Paid")
		Set CheckNum = rsPage("CheckNum")
		Set DateReceived = rsPage("DateReceived")
		Set MaintenanceID = rsPage("MaintenanceID")
		Set InvoiceSent = rsPage("InvoiceSent")
		Set DateSent = rsPage("DateSent")

		rsPage.Filter = "Paid = 0 AND InvoiceSent = 1"

		if not rsPage.EOF then
	%>
			<div ID="outstandingParent" NAME="outstandingParent" CLASS=parent>
			<% PrintTableHeader 100 %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="expandIt('outstanding'); return false" ID="outstandingIm">
			<%=rsPage.RecordCount%> Outstanding Invoices</a>
			</td></tr></table>
			</div>
			<div ID="outstandingChild" NAME="outstandingChild" CLASS=child>

			<% PrintTableHeader 100 %>
			<tr>
				<td class="TDHeader">Invoice ID</td>
				<td class="TDHeader">Dates</td>
				<td class="TDHeader">Total</td>
				<td class="TDHeader">Payment Type</td>
				<td class="TDHeader">Description/Summary</td>
				<td class="TDHeader">&nbsp;</td>
			</tr>
	<%
			do until rsPage.EOF
					strDates = "Created on:" & FormatDateTime(ItemDate, 2)

					if MaintenanceID > 0 then strDates = strDates & " - <a href=maintenance_view.asp?ID=" & MaintenanceID & ">Maintenance ID #" & MaintenanceID & "</a>"

					if not IsNull(DateSent) then strDates = strDates & "<br>Sent on:" & FormatDateTime(DateSent, 2)

					if BillingType = "CreditCard" then
						if TransID = "" then
							strBillingType = "Credit Card"
						else
							strBillingType = "Credit Card<br>TransID: " & TransID
						end if
					elseif BillingType = "Check" then
						strBillingType = "Check"
					else
						strBillingType = BillingType
					end if
	'------------------------End Code-----------------------------
	%>
					<form METHOD="post" ACTION="invoices_modify.asp">
					<input type="hidden" name="ID" value="<%=ID%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="invoice_print.asp?InvoiceID=<%=ID%>"><%=ID%></a></td>
							<td class="<% PrintTDMain %>"><%=strDates%></td>
							<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
							<td class="<% PrintTDMain %>"><%=strBillingType%></td>
							<td class="<% PrintTDMain %>"><%=Description%></td>
							<td class="<% PrintTDMainSwitch %>">
								<input type="Submit" name="Submit" value="Edit">
								<input type="button" value="Delete" onClick="DeleteBox('If you delete this invoice, there is no way to get it back.  Are you sure?', 'invoices_modify.asp?Submit=Delete&ID=<%=ID%>')">			
								
							</td>
						</tr>
					</form>
	<%
	'-----------------------Begin Code----------------------------
					rsPage.MoveNext
			loop
			Response.Write("</table></div>")
		end if

		rsPage.Filter = "Paid = 0 AND InvoiceSent = 0"

		if not rsPage.EOF then
	%>
			<div ID="currentParent" NAME="currentParent" CLASS=parent>
			<% PrintTableHeader 100 %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="expandIt('current'); return false" ID="currentIm">
			<%=rsPage.RecordCount%> Current Invoices</a>
			</td></tr></table>
			</div>
			<div ID="currentChild" NAME="currentChild" CLASS=child>

			<% PrintTableHeader 100 %>
			<tr>
				<td class="TDHeader">Invoice ID</td>
				<td class="TDHeader">Created On</td>
				<td class="TDHeader">Description/Summary</td>
				<td class="TDHeader">&nbsp;</td>
			</tr>
	<%
			do until rsPage.EOF
					strDates = "Created on:" & FormatDateTime(ItemDate, 2)

					if MaintenanceID > 0 then strDates = strDates & " - <a href=maintenance_view.asp?ID=" & MaintenanceID & ">Maintenance ID #" & MaintenanceID & "</a>"


					if BillingType = "CreditCard" then
						strBillingType = "Credit Card<br>TransID: " & TransID
					elseif BillingType = "Check" then
						strBillingType = "Check<br>Check # " & CheckNum
					else
						strBillingType = BillingType
					end if
	'------------------------End Code-----------------------------
	%>
					<form METHOD="post" ACTION="invoices_modify.asp">
					<input type="hidden" name="ID" value="<%=ID%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="invoice_print.asp?InvoiceID=<%=ID%>"><%=ID%></a></td>
							<td class="<% PrintTDMain %>"><%=strDates%></td>
							<td class="<% PrintTDMain %>"><%=Description%></td>
							<td class="<% PrintTDMainSwitch %>">
								<input type="Submit" name="Submit" value="Edit">
								<input type="button" value="Delete" onClick="DeleteBox('If you delete this invoice, there is no way to get it back.  Are you sure?', 'invoices_modify.asp?Submit=Delete&ID=<%=ID%>')">			
								
							</td>
							</tr>
					</form>
	<%
	'-----------------------Begin Code----------------------------
					rsPage.MoveNext
			loop
			Response.Write("</table></div>")
		end if

		rsPage.Filter = "Paid = 1"

		if not rsPage.EOF then
	%>
			<div ID="paidParent" NAME="paidParent" CLASS=parent>
			<% PrintTableHeader 100 %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="expandIt('paid'); return false" ID="paidIm">
			<%=rsPage.RecordCount%> Paid Invoices</a>
			</td></tr></table>
			</div>
			<div ID="paidChild" NAME="paidChild" CLASS=child>

			<% PrintTableHeader 100 %>
			<tr>
				<td class="TDHeader">Invoice ID</td>
				<td class="TDHeader">Dates</td>
				<td class="TDHeader">Total</td>
				<td class="TDHeader">Description/Summary</td>
				<td class="TDHeader">Payment Type</td>
				<td class="TDHeader">&nbsp;</td>
			</tr>
	<%
			do until rsPage.EOF
					strDates = "Created on:" & FormatDateTime(ItemDate, 2)

					if MaintenanceID > 0 then strDates = strDates & " - <a href=maintenance_view.asp?ID=" & MaintenanceID & ">Maintenance ID #" & MaintenanceID & "</a>"

					if not IsNull(DateSent) then
						if FormatDateTime(DateSent, 2) <> FormatDateTime(ItemDate, 2) then strDates = strDates & "<br>Sent on:" & FormatDateTime(DateSent, 2)
					end if
					if not IsNull(DateReceived) then
						if FormatDateTime(DateReceived, 2) <> FormatDateTime(ItemDate, 2) then strDates = strDates & "<br>Recieved on:" & FormatDateTime(DateReceived, 2)
					end if

					if BillingType = "CreditCard" then
						strBillingType = "Credit Card<br>TransID: " & TransID
					elseif BillingType = "Check" then
						strBillingType = "Check<br>Check # " & CheckNum
					else
						strBillingType = BillingType
					end if
	'------------------------End Code-----------------------------
	%>
					<form METHOD="post" ACTION="invoices_modify.asp">
					<input type="hidden" name="ID" value="<%=ID%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="invoice_print.asp?InvoiceID=<%=ID%>"><%=ID%></a></td>
							<td class="<% PrintTDMain %>"><%=strDates%></td>
							<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
							<td class="<% PrintTDMain %>"><%=Description%></td>
							<td class="<% PrintTDMain %>"><%=strBillingType%></td>
							<td class="<% PrintTDMainSwitch %>">
								<input type="Submit" name="Submit" value="Edit">
								<input type="button" value="Delete" onClick="DeleteBox('If you delete this invoice, there is no way to get it back.  Are you sure?', 'invoices_modify.asp?Submit=Delete&ID=<%=ID%>')">			
								
							</td>
						</tr>
					</form>
	<%
	'-----------------------Begin Code----------------------------
					rsPage.MoveNext
			loop
			Response.Write("</table></div>")
		end if
	else
%>
		<b>There are no invoices for this customer.</b>
<%
	end if


	Response.Write "<br><br>"

	Query = "SELECT ID, Date, Total, Description FROM CustomerMonthlyCharges WHERE CustomerID = " & intID & " ORDER BY ID DESC"

	rsPage.Close
	Set rsPage = Nothing
	Set rsPage = Server.CreateObject("ADODB.Recordset")

	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
%>
	<i>Monthly recurring charges for this customer:</i><br>
	<input type="button" value="New Monthly Charge" onClick="Redirect('monthlycharge_add.asp?CustomerID=<%=intID%>')">&nbsp;&nbsp;
<%
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Total = rsPage("Total")
		Set Description = rsPage("Description")
	%>
			<div ID="monthlyParent" NAME="monthlyParent" CLASS=parent>
			<% PrintTableHeader 100 %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="expandIt('monthly'); return false" ID="monthlyIm">
			<%=rsPage.RecordCount%> Monthly Recurring Charges</a>
			</td></tr></table>
			</div>
			<div ID="monthlyChild" NAME="monthlyChild" CLASS=child>

			<% PrintTableHeader 100 %>
			<tr>
				<td class="TDHeader">Total</td>
				<td class="TDHeader">Description</td>
				<td class="TDHeader">&nbsp;</td>
			</tr>
	<%
			do until rsPage.EOF
	'------------------------End Code-----------------------------
	%>
					<form METHOD="post" ACTION="monthlycharges_modify.asp">
					<input type="hidden" name="ID" value="<%=ID%>">
						<tr>
							<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
							<td class="<% PrintTDMain %>"><%=Description%></td>
							<td class="<% PrintTDMainSwitch %>"><input type="Submit" name="Submit" value="Edit">
							<input type="button" value="Delete" onClick="DeleteBox('If you delete this charge, there is no way to get it back.  Are you sure?', 'monthlycharges_modify.asp?Submit=Delete&ID=<%=ID%>')"></td>
							 </td>
						</tr>
					</form>
	<%
	'-----------------------Begin Code----------------------------
					rsPage.MoveNext
			loop
			Response.Write("</table></div>")
	else
%>
		<br><b>There are no monthly charges.</b>
<%
	end if

	Set rsPage = Nothing
end if
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->