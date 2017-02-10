<%
'
'-----------------------Begin Code----------------------------
if not CBool(IncludeNewsletter) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_newsletter_add.asp")
if not (LoggedAdmin or CBool( NewsletterMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Send a New Newsletter</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.Subject.value == "")
				strError += "          You forgot the subject. \n";
			if (form.Body.value == "" && form.File.value == "")
				strError += "          You forgot the details. \n";

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
	<a href="inserts_view.asp" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>

	* indicates required information<br>
	<form enctype="multipart/form-data" method="post" action="<%=SecurePath%>members_newsletter_add_process.asp" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted" name="MyForm">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Subject</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Subject" size="55">
     		</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				If you have already created this newsletter in another program (such as Word), you can upload the file here instead of typing it below.  
				However, it is usually recommended you copy and paste into the Details box below.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File">
			</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Details (inserts allowed)</td>
    		<td class="<% PrintTDMain %>"> 
				<% TextArea "Body", 55, 4, True, "" %>
    		</td>
		</tr>
		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Add">
    		</td>
		</tr>
  	</table>
	</form>
