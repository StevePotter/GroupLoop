<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Financial Options</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=financial.asp&ID=" & intID)
%>


<p>
<b>Deposits/Withdraws</b><br>
<a href="bankdeposits_add.asp">Record Deposits</a><br>
<a href="bankdeposits_modify.asp">Modify Deposits</a><br><br>

<a href="bankwithdrawals_add.asp">Record Withdrawals</a><br>
<a href="bankwithdrawals_modify.asp">Modify Withdrawals</a>
</p>


<p>
<b>Bank Accounts</b><br>
<a href="bankstatements_add.asp">Record New Statements</a><br>
<a href="bankstatements_modify.asp">Past Statements</a><br>
<%
if Session("AccessLevel") = 3 then
%>

<a href="bankaccounts_add.asp">New Account</a><br>
<a href="bankaccounts_modify.asp">Modify Accounts</a>
<%
end if
%>
</p>

<p><a href="stats_financial.asp">Financial Summaries</a></p>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->