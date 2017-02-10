<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_newsletter_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Send an E-Mail</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
strSubmit = Request("Submit")

if strSubmit = "Preview" then
	if Request("Body") = "" or Request("Subject") = "" then Redirect("incomplete.asp")

	'Put the body, subject into session since request can make problems with line breaks
	Session("Body") = GetTextArea(Request("Body"))
	Session("Subject") = Request("Subject")

%>
	<p>Below is a preview of the e-mail that will be sent.  Be sure to check it over.</p><br>
	<form method="post" action="mailing.asp" name="MyForm" onsubmit="if (this.submitted) return false; this.submitted = true; return this.submitted">
	<input type="hidden" name="SendTo" value="<%=Request("TargetMembers")%>">
	<% PrintTableHeader 0 %>
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

	'If there is a specific customer, add it to the where clause
	if intCustomerID > 0 then
		if strWhere <> "" then strWhere = strWhere & " AND "	'This puts the AND if we need it
		strWhere = strWhere & " CustomerID = " & intCustomerID
	end if

	if strWhere <> "" then strWhere = "AND " & strWhere	'This puts the WHERE if we need it

	'Get the members we are sending to
	Query = "SELECT  EMail1 FirstName, LastName, EMail2, CommonID FROM Members WHERE EMail1 <> '' " & strWhere & " GROUP BY EMail1 ORDER BY LastName"
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
		Response.Write "Sending mail to " & rsMembers("FirstName") & "&nbsp;" & rsMembers("LastName") & " - " & rsMembers("EMail1") & "<br>"

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
			if (form.Subject.value == "")
				strError += "          You forgot the subject. \n";
			if (form.Body.value == "")
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

		function getpulldown(form) {
			var Pulldown, Length, Result
			Pulldown = form.elements['TargetMembers'];
			Length = Pulldown.length;

			for ( TempIndex = 0; TempIndex <= NumMenus; TempIndex++){
				Result = Result + " OR " + Pulldown.options[TempIndex].value;
			}
			alert('hi'+Result);
			return Result;
		}

		function transferOption(FormObject, object1name, object2name) {
			var index = FormObject.elements[object1name].selectedIndex;
			if (index > -1) {
				var newoption = new Option(FormObject.elements[object1name].options[index].text, FormObject.elements[object1name].options[index].value, true, true);
				FormObject.elements[object2name].options[FormObject.elements[object2name].length] = newoption;
				if (!document.getElementById) history.go(0);
				FormObject.elements[object1name].options[index] = null;
				FormObject.elements[object1name].selectedIndex = 0;
			}
		}
//if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted
	//-->
	</SCRIPT>

	You can send this e-mail to every member, or just certain ones.  To send to everyone, select 'Everyone' and click the '>>' button.  Otherwise, 
	select each member you want to send to, and click the '>>' button.<br>
	<form method="post" action="members_sendmail.asp" onsubmit="alert(getpulldown(this)); return false">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" align="right">Recipients (explanation above)</td>
			<td class="<% PrintTDMain %>">
				<table>
				<tr>
				<td valign="middle"><% intSize = PrintMemberPullDown() %></td>
				<td>&nbsp;<input type="button" value=">>" onClick="transferOption(this.form, 'MemberIDPulldown', 'TargetMembers');">&nbsp;<br>
					&nbsp;<input type="button" value="<<" onClick="transferOption(this.form, 'TargetMembers', 'MemberIDPulldown');">&nbsp;
				</td>
				<td valign="middle"><select name="TargetMembers" size="<%=intSize%>"></select></td>
				</tr>
				</table>
			</td>
   		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Subject</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Subject" size="55">
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



Function PrintMemberPullDown( )

	Query = "SELECT ID, FirstName, LastName FROM Members WHERE ID <> " & Session("MemberID") & " AND EMail1 <> '' AND CustomerID = " & CustomerID & " ORDER BY LastName"
	Set rsPulldown = Server.CreateObject("ADODB.Recordset")
	rsPulldown.CacheSize = 50
	rsPulldown.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	if rsPulldown.EOF then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but there must be multiple members before you can send mail to others."))


	Set MemberID = rsPulldown("ID")
	Set FirstName = rsPulldown("FirstName")
	Set LastName = rsPulldown("LastName")

	intMembers = rsPulldown.RecordCount + 1
	if intMembers > 10 then intMembers = 10

	%><select name="MemberIDPulldown" size="<%=intMembers%>"><option value='All'>Everyone</option> <%

	do until rsPulldown.EOF
		Response.Write "<option value = '" & MemberID & "'>" & FirstName & " " & LastName & "</option>" & vbCrlf
		rsPulldown.MoveNext
	loop
	rsPulldown.Close

	set rsPulldown = Nothing
	Response.Write("</select>")

	PrintMemberPullDown = intMembers

End Function
%>