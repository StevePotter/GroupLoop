<!-- #include file="header.asp" -->
<!-- #include file="..\homegroup\dsn.asp" -->
<!-- #include file="functions.asp" -->
<%
if not LoggedEmployee then Redirect("login.asp?Source=salesman_support.asp")
%>

<p align="center"><span class=Heading>Get Help</span><br>
<span class=LinkText><a href="login.asp">Back To Salesman Options</a></span></p>

<%
'-----------------------Begin Code----------------------------
'We are going to check for errors if they are updating the profile
strSubmit = Request("Submit")
if strSubmit = "Send" then
	Query = "SELECT * FROM Employees WHERE ID = " & Session("EmployeeID")
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenStatic, adLockOptimistic, adCmdTableDirect
	if rsMember.EOF then
		set rsMember = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("We can't find your salesman record."))
	end if



	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.ContentType = "text/html"
	Mailer.IgnoreMalformedAddress = true
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = "GroupLoop.com"
	Mailer.FromAddress = "support@grouploop.com"
	Mailer.AddRecipient "Support", "support@grouploop.com"
	Mailer.Subject    = "Salesman Support Question - " & rsMember("FirstName") & " " & rsMember("LastName")


	strBody = "<html><body><p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>" & _
	"<p>Salesman ID: <b>" & rsMember("ID") & "</b><br>" & _
	"Name: <b>" & rsMember("FirstName") & " " & rsMember("LastName") & "</b><br>" & _
	"NickName: <b>" & rsMember("NickName") & "</b><br>" & _
	"E-Mail: <b><a href=mailto:" & rsMember("EMail1") & ">" & rsMember("EMail1") & "</a></b></p>" & _
	"<p>" & Format(Request("Body")) & "</p>" & _
	"</body></html>"

	Mailer.BodyText   = strBody
	Mailer.Sendmail

	Set Mailer = Nothing

	rsMember.Close
	set rsMember = Nothing
'------------------------End Code-----------------------------
%>
	<p>Your question has been sent. You will receive a reply via e-mail soon.</p>
<%
'-----------------------Begin Code----------------------------

else
'------------------------End Code-----------------------------
%>

	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.Body.value == "")
				strError += "Sorry, but you must enter a question or problem. \n";
				
			if(strError == "") {
				return true;
			}
			else{
				alert (strError);
				return false;
			}   
		}

	//-->
	</SCRIPT>

	<form method="post" action="<%=SecurePath%>salesman_support.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">

	<table border=0 cellspacing=0 cellpadding=0>
	<tr><td align="left">

	Please enter your question or problem below:<br>
	<textarea name="Body" cols="55" rows="10" wrap="PHYSICAL"></textarea>

	</td></tr>
	<tr><td align="center">

	<input type="submit" name="Submit" value="Send">

	</td></tr>
	</table>

	</form>
<%
'-----------------------Begin Code----------------------------
end if

%>




<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->
