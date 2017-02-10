<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Run Daily Maintenance</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>



<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=daily_setup.asp&ID=" & intID)

strSubmit = Request("Submit")

if strSubmit = "Run It" then
	RunDate = AssembleDate("Date")
	RunDate = FormatDateTime( RunDate, 2 )

	strLink = "nightmaint.asp?Date="&RunDate

	if Request("Override") = "1" then strLink = strLink & "&Override=YES"

	if Request("Warnings") = "0" then strLink = strLink & "&Warnings=NO"

	if Request("Charges") = "0" then strLink = strLink & "&Charges=NO"

	if Request("WriteIndex") = "0" then strLink = strLink & "&WriteIndex=NO"

	if Request("WriteConstants") = "1" then strLink = strLink & "&WriteConstants=YES"
	if Request("WriteHeaderFooter") = "1" then strLink = strLink & "&WriteHeaderFooter=YES"

	if Request("SendEMail") = "0" then strLink = strLink & "&SendEMail=NO"



'	Response.write strLink

	Redirect strLink
else
%>
<form method="post" action="daily_setup.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = true; return this.submitted">
	Date to run maintenance on: <% DatePulldown "Date", Date, 0 %><br>
	If the maintenance has already been run for the day, run it anyway? <% PrintRadio 1, "Override" %><br>
	Give expiration warnings? <% PrintRadio 0, "Warnings" %><br>
	Charge customers? <% PrintRadio 1, "Charges" %><br>
	Create index files (for each customer's home page)? <% PrintRadio 0, "WriteIndex" %><br>
	Create constants files? <% PrintRadio 0, "WriteConstants" %><br>
	Create header and footer files? <% PrintRadio 0, "WriteHeaderFooter" %><br>
	Send results via e-mail to accounts@grouploop.com? <% PrintRadio 1, "SendEMail" %><br>




	<input type="submit" name="Submit" value="Run It">


</form>
<%
end if
%>
<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->