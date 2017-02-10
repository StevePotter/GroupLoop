<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Statistics</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=maintenance.asp&ID=" & intID)
%>

<a href="stats_homesite.asp">GroupLoop.com Home Page</a><br>
<a href="stats_customers.asp">GroupLoop.com Customers</a><br>
<a href="stats_billing.asp">GroupLoop.com Billing</a><br>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->