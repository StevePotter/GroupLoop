<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Employees</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=financial.asp&ID=" & intID)

if Session("AccessLevel") < 3 then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, you do not have access to this area."))
%>
<a href="employees_add.asp">New Employee</a><br>
<a href="employees_modify.asp">Modify Employees</a>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->