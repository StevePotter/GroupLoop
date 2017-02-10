<!-- #include file="dsn.asp" -->

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=index.asp")
%>

<!-- #include file="header.asp" -->


<p align="<%=HeadingAlignment%>"><span class=Heading>Staff</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<p>What up <%=Session("NickName")%>?</p>

<%
if Session("AccessLevel") = 2 then
%>
<b>Your options:</b><br>
<a href="employeecharge_add.asp" target="_blank"><b>Start the Clock</b></a><br>
<a href="financial.asp">Manage Finances</a><br>

<br>
<a href="mailto:president@grouploop.com">E-Mail Steve</a><br>
<%
elseif Session("AccessLevel") = 3 then
%>


<%
end if
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->