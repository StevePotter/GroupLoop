<!-- #include file="dsn.asp" -->

<!-- #include file="header.asp" -->
<!-- #include file="..\sourcegroup\expandscripts.inc" -->

<p align="<%=HeadingAlignment%>"><span class=Heading>Send An E-Mail To Members</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'This logs them in
strPassword = Request("Password")
strNickName = Request("NickName")
if strPassword <> "" or strNickName <> "" then
	EmployeeLogin strPassword, strNickName
end if

if not LoggedStaff() then Redirect("login.asp?Source=monthlycharge_add.asp&CustomerID=" & Request("CustomerID"))

strSubmit = Request("Submit")

if strSubmit = "Preview" then
	if Request("Body") = "" or Request("Subject") = "" then Redirect("incomplete.asp")

	'Get the customer(s) to send it to
	intCustomerID = 0
	if Request("CustomerID") <> "" then intCustomerID = CInt(Request("CustomerID"))

	'put the pre tag in front of regular for exact preview
	if Request("HTML") = "YES" then
		strBody = "<pre>" & Request("Body") & "</pre>"
	else
		strBody = GetTextArea( Request("Body") )
	end if

	'Put the body, subject into session since request can make problems with line breaks
	Session("Body") = Request("Body")
	Session("Subject") = Request("Subject")

	'Just display the proper audience
	if Request("Type") = "MailingList" then
		strAudience = "Anyone recieving the GroupLoop newsletter "
	elseif Request("Type") = "Members" then
		strAudience = "Regular members "
	else
		strAudience = "Just administrators "
	end if

	if intCustomerID = 0 then
		strAudience = strAudience & " from all sites."
	else
		strAudience = strAudience & " from: " & GetCustSummary( intCustomerID )
	end if
%>
	<p>Below is a preview of the e-mail that will be sent.  Be sure to check it over.</p><br>
	<form method="post" action="mailing.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = true; return this.submitted">
	<input type="hidden" name="CustomerID" value="<%=intCustomerID%>">
	<input type="hidden" name="HTML" value="<%=Request("HTML")%>">
	<input type="hidden" name="Logo" value="<%=Request("Logo")%>">
	<input type="hidden" name="Type" value="<%=Request("Type")%>">

	<% PrintTableHeader 0 %>
	<tr> 
		<td><b>To:</b></td>
		<td><%=strAudience%></td>
	</tr>
	<tr> 
		<td><b>Subject:</b></td>
		<td><%=Request("Subject")%></td>
	</tr>
	</table>
	<%=strBody%>
	<p align="center">
		<input type="submit" name="Submit" value="Send">
	</p>
	</form>

<%
elseif strSubmit = "Send" then
	'Regular text or html?
	blHTML = False
	if Request("HTML") = "YES" then blHTML = True

	intCustomerID = CInt(Request("CustomerID"))
	strSubject = Session("Subject")
	strType = Request("Type")

	if blHTML then
		strBody = GetTextArea( Session("Body") )
	else
		strBody = Session("Body")
	end if

	strWhere = ""
	'This is the WHERE clause
	if strType = "MailingList" then
		strWhere = "SubscribeGroupLoopNewsletter = 1 "
	elseif strType = "Administrators" then
		strWhere = "Admin = 1 "
	end if

	'If there is a specific customer, add it to the where clause
	if intCustomerID > 0 then
		if strWhere <> "" then strWhere = strWhere & " AND "	'This puts the AND if we need it
		strWhere = strWhere & " CustomerID = " & intCustomerID
	end if

	if strWhere <> "" then strWhere = "AND " & strWhere	'This puts the WHERE if we need it

	'Get the members we are sending to
'	Query = "SELECT  EMail1, FirstName, LastName, EMail2 FROM Members WHERE EMail1 <> '' " & strWhere & " GROUP BY EMail1 ORDER BY LastName"
	Query = "SELECT DISTINCT(EMail1), FirstName, LastName, EMail2 FROM Members WHERE EMail1 <> '' GROUP BY EMail1, FirstName, LastName, EMail2 ORDER BY EMail1"
	Response.Write Query & "<br>"'GROUP BY CommonID 
	Set rsMembers = Server.CreateObject("ADODB.Recordset")
	rsMembers.CacheSize = 500
	rsMembers.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	'Set up the mailer object
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.IgnoreMalformedAddress = true
	Mailer.RemoteHost  = "mail4.burlee.com"
	Mailer.FromName    = "GroupLoop.com"
	Mailer.FromAddress = "support@grouploop.com"
	Mailer.Subject    = strSubject

	if blHTML then
		Mailer.ContentType = "text/html"

		strHeader = "<html><title>" & strSubject & "</title><body>"
		if Request("Logo") = "Yes" then strHeader = strHeader & "<p align=center><img src='http://www.GroupLoop.com/homegroup/images/title.gif' border=0></p>"
		strFooter = "</body></html>"

		strBody = strHeader & strBody & strFooter
	end if

	Mailer.BodyText = strBody

	do until rsMembers.EOF
		Response.Write "Sending mail to " & rsMembers("EMail1") & " - " & rsMembers("FirstName") & "&nbsp;" & rsMembers("LastName") & "<br>"

		Mailer.ClearRecipients

		if rsMembers("EMail1") <> "" then
			Mailer.AddRecipient rsMembers("FirstName") & " " & rsMembers("LastName"), rsMembers("EMail1")
	'		Mailer.SendMail
		end if

		rsMembers.MoveNext
	loop

	rsMembers.Close
	Set rsMembers = Nothing

	Set Mailer = Nothing
%>
	<p>The mail has been sent out.<br>
	<a href="mailing.asp">Send another.</a>
	</p>
<%
else
%>


	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";


			if (form.Subject.value == "" )
				strError += "          You forgot the subject. \n";
			if (form.Body.value == "" )
				strError += "          You forgot the body. \n";

			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can add this: \n" + strError;
				alert (strError);
				return false;
			}   
		}

	//-->
	</SCRIPT>

	<form method="post" action="mailing.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<% PrintTableHeader 0 %>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Send this e-mail to:</td>
     	<td class="<% PrintTDMain %>">
<%
				PrintRadioOption "Type", "MailingList", "Everyone signed up for the mailing list.<br>", "MailingList"
				PrintRadioOption "Type", "Members", "All members<br>", ""
				PrintRadioOption "Type", "Administrators", "Administrators<br>", ""
%>
					Send to - <% PrintCustomerPullDown 0, 1, 0, "All Sites", "" %>
			
    	</td>
	</tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Formatting</td>
			<td class="<% PrintTDMain %>"> 
<%
				PrintRadioOptionNew "HTML", "Yes", "Use HTML", "Yes", "show('HTMLDet');"
%>
					<span id="HTMLDet" <%=GetDisplay(1)%>>&nbsp; &nbsp; &nbsp; &nbsp; Include the GroupLoop logo on top <input type="checkbox" name="Logo" value="Yes" checked></span><br>	
<%
				PrintRadioOptionNew "HTML", "No", "Plain Text<br>", "", "hide('HTMLDet');"
%>


			</td>
   		</tr>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Subject</td>
			<td class="<% PrintTDMain %>"> 
				<input type="text" name="Subject" value="" size="55">
			</td>
   		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Details</td>
    		<td class="<% PrintTDMain %>"> 
				<% TextArea "Body", 55, 20, True, "" %>
    		</td>
		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Preview">
    		</td>
		</tr>
  	</table>
	</form>
<%
end if


'-------------------------------------------------------------
'This function prints yes and no radio boxes, highlighting the right one depending on the bool passed
'-------------------------------------------------------------
Sub PrintRadioOptionNew( Name, Value, Display, Selected, onClick )
	Response.Write "<input type='radio' name=" & Chr(34) & Name & Chr(34) & " value=" & Chr(34) & Value & Chr(34)
	if Value = Selected then Response.Write " checked"
	Response.Write " onClick=" & Chr(34) & onClick & Chr(34) & " > " & Display
End Sub


Function GetDisplay( blDisplay )
	blDisplay = CBool(blDisplay)
	if blDisplay then
		GetDisplay = ""
	else
		GetDisplay = " style=" & chr(34) & "display: none;" & chr(34)
	end if
End Function
%>


<!-- #include file="footer.asp" -->

<!-- #include file ="closedsn.asp" -->