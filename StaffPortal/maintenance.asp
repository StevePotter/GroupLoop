<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Maintenance Options</span><br>
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

<a href="daily_setup.asp">Run the daily maintenance</a><br>
<a href="maintenance_modify.asp">Maintenance run list</a><br>
<a href="template_copy.asp">Distribute a template file</a><br>
<a href="upload.asp">Upload files to the server</a><br>



<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->