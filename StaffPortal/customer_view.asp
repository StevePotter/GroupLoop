<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
intID = CInt(Request("ID"))

if not LoggedStaff() then Redirect("login.asp?Source=customer_view.asp&ID=" & intID)


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


	Set ID = rsPage("ID")
	Set OwnerID = rsPage("OwnerID")
	Set SignupDate = rsPage("SignupDate")
	Set MasterID = rsPage("MasterID")
	Set ParentID = rsPage("ParentID")
	Set Version = rsPage("Version")


	Set UseDomain = rsPage("UseDomain")
	Set DomainName = rsPage("DomainName")
	Set SubDirectory = rsPage("SubDirectory")

	Set Title = rsPage("Title")

	Set Organization = rsPage("Organization")
	Set FirstName = rsPage("FirstName")
	Set LastName = rsPage("LastName")
	Set Street1 = rsPage("Street1")
	Set Street2 = rsPage("Street2")
	Set City = rsPage("City")
	Set State = rsPage("State")
	Set Zip = rsPage("Zip")
	Set Country = rsPage("Country")
	Set Phone = rsPage("Phone")

	Set BillingType = rsPage("BillingType")
	Set BillingStreet1 = rsPage("BillingStreet1")
	Set BillingStreet2 = rsPage("BillingStreet2")
	Set BillingCity = rsPage("BillingCity")
	Set BillingState = rsPage("BillingState")
	Set BillingZip = rsPage("BillingZip")
	Set BillingPhone = rsPage("BillingPhone")
	Set BillingCountry = rsPage("BillingCountry")
	Set CCCompany = rsPage("CCCompany")
	Set CCType = rsPage("CCType")
	Set TransID = rsPage("TransID")
	Set MerchantClientIDNumber = rsPage("MerchantClientIDNumber")
	Set MerchantBank = rsPage("MerchantBank")

	Set MemberFirstName = rsPage("MemberFirstName")
	Set MemberLastName = rsPage("MemberLastName")
	Set HomeStreet = rsPage("HomeStreet")
	Set HomeCity = rsPage("HomeCity")
	Set HomeState = rsPage("HomeState")
	Set HomeZip = rsPage("HomeZip")
	Set HomePhone = rsPage("HomePhone")
	Set Beeper = rsPage("Beeper")
	Set CellPhone = rsPage("CellPhone")


	Set NickName = rsPage("NickName")
	Set Password = rsPage("Password")

	Set EMail = rsPage("EMail")
	Set MemberEMail1 = rsPage("EMail1")
	Set MemberEMail2 = rsPage("EMail2")

rsPage.Filter = "ID = " & intID
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Customer Details - <%=PrintStart(Title)%></span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>


<script language="JavaScript">
<!--
function NewWind( Link )
{
	window.name = 'parentWnd';
	newWindow = window.open(Link);
	newWindow.focus();
}

//-->
</SCRIPT>

<%

if UseDomain = 1 then
	strAddress = DomainName
else
	strAddress = "http://www.GroupLoop.com/" & SubDirectory
end if

strLogin = strAddress & "/login.asp?Source=index.asp&NickName=" & NickName & "&Password=" & Password
%>
<p class="SubHeading" align="left">Options</p>
<div align="left">
	<input type="button" value="Log in as Owner" onClick="NewWind('<%=strLogin%>')">&nbsp;&nbsp;
	<input type="button" value="Edit Customer" onClick="Redirect('customer_edit.asp?ID=<%=ID%>')">&nbsp;&nbsp;
	<input type="button" value="Charge Customer" onClick="Redirect('customer_charge.asp?CustomerID=<%=ID%>')">&nbsp;&nbsp;
	<input type="button" value="Remove Site" onClick="Redirect('customer_delete.asp?Submit=Delete&ID=<%=ID%>')">			
</div>

<p class="SubHeading" align="left">Customer Info</p>
Version: <%=Version%><br>

Site Address: <a href="<%=strAddress%>"><%=strAddress%></a><br>
<%
if UseDomain = 1 then
%>
Sub-Directory: <%=SubDirectory%><br>
<%
end if
if Organization <> "" then
%>
Organization: <%=Organization%><br>
<%
end if
%>
Date Created: <%=FormatDateTime(SignupDate, 2)%><br>
Owner Name: <%=FirstName%>&nbsp;<%=LastName%><br>
Site Title: <%=Title%><br>

<p class="SubHeading" align="left">E-Mail</p>
Contact E-Mail: <a href="mailto:<%=EMail%>"><%=EMail%></a><br>
<%
if MemberEMail1 <> "" and MemberEMail1 <> EMail then
%>
Owner's E-Mail: <a href="mailto:<%=MemberEMail1%>"><%=MemberEMail1%></a><br>
<%
end if
if MemberEMail2 <> "" and MemberEMail2 <> EMail then
%>
Owner's Secondary E-Mail: <a href="mailto:<%=MemberEMail2%>"><%=MemberEMail2%></a><br>
<%
end if


%>
<p class="Heading" align="left">Billing Summary</p>

	<input type="button" value="New Invoice" onClick="Redirect('invoice_add.asp?CustomerID=<%=intID%>')"><br>
<%
	Query = "SELECT ID, Date, Total, Description, BillingType, TransID, Paid, CheckNum, DateReceived, MaintenanceID, InvoiceSent, DateSent FROM CustomerInvoices WHERE CustomerID = " & intID & " ORDER BY ID DESC"

	Set rsBilling = Server.CreateObject("ADODB.Recordset")
	rsBilling.CacheSize = 100
	rsBilling.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsBilling.EOF then
	%>
		<b><%=rsBilling.RecordCount%> total invoices.</b><br>
	<%
		Set InvoiceID = rsBilling("ID")
		Set ItemDate = rsBilling("Date")
		Set Total = rsBilling("Total")
		Set Description = rsBilling("Description")
		Set BillingType = rsBilling("BillingType")
		Set TransID = rsBilling("TransID")
		Set Paid = rsBilling("Paid")
		Set CheckNum = rsBilling("CheckNum")
		Set DateReceived = rsBilling("DateReceived")
		Set MaintenanceID = rsBilling("MaintenanceID")
		Set InvoiceSent = rsBilling("InvoiceSent")
		Set DateSent = rsBilling("DateSent")

		rsBilling.Filter = "Paid = 0 AND InvoiceSent = 1"

		if not rsBilling.EOF then
	%>
			<div ID="outstandingParent" NAME="outstandingParent" CLASS=parent>
			<% PrintTableHeader 100 %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="expandIt('outstanding'); return false" ID="outstandingIm">
			<%=rsBilling.RecordCount%> Outstanding Invoices</a>
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
			do until rsBilling.EOF
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
					<input type="hidden" name="ID" value="<%=InvoiceID%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="invoice_print.asp?InvoiceID=<%=InvoiceID%>"><%=InvoiceID%></a></td>
							<td class="<% PrintTDMain %>"><%=strDates%></td>
							<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
							<td class="<% PrintTDMain %>"><%=strBillingType%></td>
							<td class="<% PrintTDMain %>"><%=Description%></td>
							<td class="<% PrintTDMainSwitch %>">
								<input type="Submit" name="Submit" value="Edit">
								<input type="button" value="Delete" onClick="DeleteBox('If you delete this invoice, there is no way to get it back.  Are you sure?', 'invoices_modify.asp?Submit=Delete&ID=<%=InvoiceID%>')">			
								
							</td>
						</tr>
					</form>
	<%
	'-----------------------Begin Code----------------------------
					rsBilling.MoveNext
			loop
			Response.Write("</table></div>")
		end if

		rsBilling.Filter = "Paid = 0 AND InvoiceSent = 0"

		if not rsBilling.EOF then
	%>
			<div ID="currentParent" NAME="currentParent" CLASS=parent>
			<% PrintTableHeader 100 %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="expandIt('current'); return false" ID="currentIm">
			<%=rsBilling.RecordCount%> Current Invoices</a>
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
			do until rsBilling.EOF
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
					<input type="hidden" name="ID" value="<%=InvoiceID%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="invoice_print.asp?InvoiceID=<%=InvoiceID%>"><%=InvoiceID%></a></td>
							<td class="<% PrintTDMain %>"><%=strDates%></td>
							<td class="<% PrintTDMain %>"><%=Description%></td>
							<td class="<% PrintTDMainSwitch %>">
								<input type="Submit" name="Submit" value="Edit">
								<input type="button" value="Delete" onClick="DeleteBox('If you delete this invoice, there is no way to get it back.  Are you sure?', 'invoices_modify.asp?Submit=Delete&ID=<%=InvoiceID%>')">			
								
							</td>
							</tr>
					</form>
	<%
	'-----------------------Begin Code----------------------------
					rsBilling.MoveNext
			loop
			Response.Write("</table></div>")
		end if

		rsBilling.Filter = "Paid = 1"

		if not rsBilling.EOF then
	%>
			<div ID="paidParent" NAME="paidParent" CLASS=parent>
			<% PrintTableHeader 100 %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="expandIt('paid'); return false" ID="paidIm">
			<%=rsBilling.RecordCount%> Paid Invoices</a>
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
			do until rsBilling.EOF
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
					<input type="hidden" name="ID" value="<%=InvoiceID%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="invoice_print.asp?InvoiceID=<%=InvoiceID%>"><%=InvoiceID%></a></td>
							<td class="<% PrintTDMain %>"><%=strDates%></td>
							<td class="<% PrintTDMain %>"><%=FormatCurrency(Total)%></td>
							<td class="<% PrintTDMain %>"><%=Description%></td>
							<td class="<% PrintTDMain %>"><%=strBillingType%></td>
							<td class="<% PrintTDMainSwitch %>">
								<input type="Submit" name="Submit" value="Edit">
								<input type="button" value="Delete" onClick="DeleteBox('If you delete this invoice, there is no way to get it back.  Are you sure?', 'invoices_modify.asp?Submit=Delete&ID=<%=InvoiceID%>')">			
								
							</td>
						</tr>
					</form>
	<%
	'-----------------------Begin Code----------------------------
					rsBilling.MoveNext
			loop
			Response.Write("</table></div>")
		end if
	else
%>
		<b>There are no invoices for this customer.</b>
<%
	end if


	Query = "SELECT ID, Date, Total, Description FROM CustomerMonthlyCharges WHERE CustomerID = " & intID & " ORDER BY ID DESC"

	rsBilling.Close
	Set rsBilling = Nothing
	Set rsBilling = Server.CreateObject("ADODB.Recordset")

	rsBilling.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
%>
	<i>Monthly recurring charges for this customer:</i><br>
	<input type="button" value="New Monthly Charge" onClick="Redirect('monthlycharge_add.asp?CustomerID=<%=intID%>')">&nbsp;&nbsp;
<%
	if not rsBilling.EOF then
		Set ChargeID = rsBilling("ID")
		Set ItemDate = rsBilling("Date")
		Set Total = rsBilling("Total")
		Set Description = rsBilling("Description")
	%>
			<div ID="monthlyParent" NAME="monthlyParent" CLASS=parent>
			<% PrintTableHeader 100 %>
			<tr><td class="TDHeader">
			<a class="TDHeader" HREF="javascript://" onClick="expandIt('monthly'); return false" ID="monthlyIm">
			<%=rsBilling.RecordCount%> Monthly Recurring Charges</a>
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
			do until rsBilling.EOF
	'------------------------End Code-----------------------------
	%>
					<form METHOD="post" ACTION="monthlycharges_modify.asp">
					<input type="hidden" name="ID" value="<%=ChargeID%>">
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
					rsBilling.MoveNext
			loop
			Response.Write("</table></div>")
	else
%>
		<br><b>There are no monthly charges.</b>
<%
	end if
	rsBilling.Close
	Set rsBilling = Nothing

%>

<p class="SubHeading" align="left">Contact Address</p>
<%=Street1%><br>
<%
if Street2 <> "" then
%>
	<%=Street2%><br>
<%
end if
%>
<%=City%>,&nbsp;<%=State%>&nbsp;<%=Zip%>&nbsp;<%=Country%><br>
<%
if Phone <> "" then
%>
	<%=Phone%><br>
<%
end if
%>


<%
if (Street1 <> BillingStreet1 or Street2 <> BillingCity or State <> BillingState or Zip <> BillingZip or _
	Country <> BillingCountry or Phone <> BillingPhone) and _
	(BillingStreet1 <> "" and BillingCity <> "" and BillingState <> "" and BillingZip <> "") then
%>
<p class="SubHeading" align="left">Billing Address</p>
	<%=BillingStreet1%><br>
	<%
	if BillingStreet2 <> "" then
	%>
		<%=BillingStreet2%><br>
	<%
	end if
	%>
	<%=BillingCity%>,&nbsp;<%=BillingState%>&nbsp;<%=BillingZip%>&nbsp;<%=BillingCountry%><br>
	<%
	if BillingPhone <> "" then
	%>
		<%=BillingPhone%><br>
	<%
	end if

end if
%>

<p class="SubHeading" align="left">Owner's Member Information</p>
Name: <%=MemberFirstName%>&nbsp;<%=MemberLastName%><br>
Nickname: <%=NickName%><br>
Password: <%=Password%><br>
Address: <%=HomeStreet%><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=HomeCity%>,&nbsp;<%=HomeState%>&nbsp;<%=HomeZip%><br>
<%
if Beeper <> "" then
%>
	Beeper: <%=Beeper%><br>
<%
end if
%>
<%
if CellPhone <> "" then
%>
	Cellphone: <%=CellPhone%><br>
<%
end if

rsPage.Close
Set rsPage = Nothing
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->