<!-- #include file="header.asp" -->
<!-- #include file="..\homegroup\dsn.asp" -->
<!-- #include file="functions.asp" -->
<%
if not LoggedEmployee then Redirect("login.asp?Source=salesman_refer.asp")
%>

<p align="center"><span class=Heading>Refer and Make Even More</span><br>
<span class=LinkText><a href="login.asp">Back To Salesman Options</a></span></p>

<%
'-----------------------Begin Code----------------------------
'We are going to check for errors if they are updating the profile
strSubmit = Request("Submit")
if strSubmit = "Send Referral" then
	strName = Format(Request("Name"))
	strEMail = Request("EMail")

	if strEMail = "" then Redirect("incomplete.asp")

	SendEMail
'------------------------End Code-----------------------------
%>
	<p>Your referral has been sent. &nbsp;<a href="salesman_refer.asp">Click here</a> to send another.</p>
<%
'-----------------------Begin Code----------------------------

else
'------------------------End Code-----------------------------
%>
	<p>If you refer a friend, you will get 5% from new customers, along with the 20% they 
	get!  Get everyone you know to sign up, and you will benefit even more from our sales program!</p>


	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.EMail.value == "")
				strError += "Sorry, but you forgot the e-mail address. \n";
				
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

	* indicates required information<br>
	<form method="post" action="salesman_refer.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Name & Such
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Name
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="Name">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				* E-Mail Address
			</td>
			<td class="<% PrintTDMain %>">
				<input type="text" size="40" name="EMail">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
				<input type="submit" name="Submit" value="Send Referral">
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if




Sub SendEMail
	intEmployeeID = Session("EmployeeID")

	Query = "SELECT FirstName, LastName, EMail1 FROM Employees WHERE ID = " & intEmployeeID
	Set rsMember = Server.CreateObject("ADODB.Recordset")
	rsMember.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	SalesmanName = rsMember("FirstName") & " " & rsMember("LastName")

	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.ContentType = "text/html"
	Mailer.IgnoreMalformedAddress = true
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = "GroupLoop.com"
	Mailer.FromAddress = "support@grouploop.com"
	Mailer.Subject    = SalesmanName & " has Referred you for our Rewarding Sales Program"

	strBody = "<html><body><p align=center><a href='http://www.GroupLoop.com'><img src='http://www.GroupLoop.com/sales/title.gif' border=0></a></p>"

	if strName = "" then
		strBody = strBody & "<p>Dear Potential Salesperson,<br>"
	else
		strBody = strBody & "<p>Dear " & Format(strName) & ",<br>"
	end if

	strBody = strBody & "<p>Congratulations!  " & SalesmanName & " has sent you this e-mail to inform you and hopefully get you to join <a href='http://www.GroupLoop.com'>GroupLoop.com's</a> " & _
	"Continual Commission Sales Program.  This unique sales program pays you <b>each month</b> for the customers you bring to us.</p>" & _
	"<p>This is not child's play.  This is an opportunity for you to make some real money at your leisure.  We are an established, reliable, successful " & _
	"company.  We encourage you to learn more about us and the sales program at <i><a href='http://www.GroupLoop.com/sales/index.asp?ReferralID=" & intEmployeeID & "'>www.GroupLoop.com/Sales</a></i></p>" & _
	"<p>If you choose to sign up, we urge you to enter " & SalesmanName & "'s Salesperson ID number as a referral.  You will be prompted for this during " & _
	"your signup process.  <u>" & rsMember("FirstName") & "'s ID number is <b>" &  intEmployeeID & "</b></u>.  This is very important, as " & rsMember("FirstName") & " " & _
	"will get a bonus for each of your new customers.  Don't worry, it will not detract from your earnings.</p>" & _
	"<p>Signing up is totally free of any kind of obligation.  Join now, make a difference, and make serious money!</p>" & _
	"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Thank you,<br>" & VbCrLf & _
	"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;Stephen Potter<br>" & VbCrLf & _
	"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;President, GroupLoop.com</p></body></html>" & VbCrLf

	Mailer.BodyText   = strBody
	Mailer.AddRecipient strEMail, strEMail
	Mailer.Sendmail

	Set Mailer = Nothing

End Sub


%>




<!-- #include file="..\homegroup\closedsn.asp" -->

<!-- #include file="footer.asp" -->
