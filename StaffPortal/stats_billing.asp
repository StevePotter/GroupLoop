<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=maintenance.asp&ID=" & intID)


Set cmd = Server.CreateObject("ADODB.Command")
With cmd
	.ActiveConnection = Connect
	.CommandText = "GetFinancialStats"
	.CommandType = adCmdStoredProc

	.Parameters.Refresh

	.Execute , , adExecuteNoRecords

	intDeposits = .Parameters("@Deposits")
	curTotalDeposits = .Parameters("@TotalDeposits")


	intInvoices = .Parameters("@Invoices")
	intInvoiceCharges = .Parameters("@InvoiceCharges")
	intCustomerMonthlyCharges = .Parameters("@CustomerMonthlyCharges")
	curTotalInvoiceCharges = .Parameters("@TotalInvoiceCharges")
	curTotalCustomerMonthlyCharges = .Parameters("@TotalCustomerMonthlyCharges")

	intInvoices = .Parameters("@Invoices")
	curTotalInvoiceCharges = .Parameters("@TotalInvoiceCharges")

	intExpenses = .Parameters("@Expenses")
	curTotalExpenses = .Parameters("@TotalExpenses")


'	intInvoicesOwed = .Parameters("@InvoicesOwed")
	curTotalInvoicesOwed = .Parameters("@TotalOwed")

End With
Set cmd = Nothing
%>
<form method="post" action="stats_billing.asp" name="MyForm">

<p align="<%=HeadingAlignment%>"><span class=Heading>Financial Stats</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<p><span class="SubHeading">Summary:</span><br>
Invoices: <%=intInvoices%><br>
Invoice charges (charges that make up each invoice): <%=intInvoiceCharges%><br>
Total for invoice charges: <%=FormatCurrency(curTotalInvoiceCharges)%><br><br>

Monthly charges to customers: <%=intCustomerMonthlyCharges%><br>
Total for monthly charges: <%=FormatCurrency(curTotalCustomerMonthlyCharges)%><br><br>

Number of deposits: <%=intDeposits%><br>
Total deposits: <%=FormatCurrency(curTotalDeposits)%><br><br>


Number of expenses: <%=intExpenses%><br>
Total expenses: <%=FormatCurrency(curTotalExpenses)%><br><br>

Money owed (current and outstanding invoices) <%=curTotalInvoicesOwed%><br><br>

Deposits - expenses = 
<%
dblNet = cDbl(curTotalDeposits)-cDbl(curTotalExpenses)
if dblNet < 0 then Response.Write "-"
%>
<%=FormatCurrency(abs(dblNet))%>




</p>




</form>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->