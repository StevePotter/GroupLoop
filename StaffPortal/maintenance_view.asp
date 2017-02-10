<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->


<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
intID = CInt(Request("ID"))

if not LoggedStaff() then Redirect("login.asp?Source=maintenance_view.asp&ID=" & intID)

	Query = "SELECT ID, Date, Output FROM NightlyMaintenance WHERE ID = " & intID


	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly

Response.Write rsPage("Output")

rsPage.Close
Set rsPage = Nothing
%>

<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->