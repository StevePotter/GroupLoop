<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeMeetings ) or CBool( MeetingsMembers ) ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_meetings_add.asp")
if not (LoggedAdmin or CBool( MeetingsMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20


Public DisplayPrivacy

Query = "SELECT IncludePrivacyMeetings, DisplaySearchMeetings, DisplayDaysOldMeetings, InfoTextMeetings, ListTypeMeetings, DisplayDateListMeetings, DisplayAuthorListMeetings, DisplayPrivacyListMeetings  FROM Look WHERE CustomerID = " & CustomerID
Set rsNew = Server.CreateObject("ADODB.Recordset")
rsNew.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
DisplayPrivacy = CBool(rsNew("IncludePrivacyMeetings")) and not cBool(SiteMembersOnly)

rsNew.Close
Set rsNew = Nothing
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Add A Meeting</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>


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
	* indicates required information<br>
	<form enctype="multipart/form-data" method="post" action="<%=SecurePath%>members_meetings_add_process.asp" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted" name="MyForm">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
<%
	if IncludeCommittees = 1 then
%>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Category
		</td>
		<td class="<% PrintTDMain %>">
<%			PrintCommitteePullDown 0	%>
		</td>
	</tr>
<%
		end if
%>
<%
		if DisplayPrivacy then
%>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Only let members read it?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="Private" value="1">
			</td>
   		</tr>
<%
		end if
%>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Should this meeting be sent via e-mail to all the site members?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="EMail" value="1">
			</td>
   		</tr>

		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Subject</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Subject" size="55">
     		</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				If you have already created this in another program, you can upload the file here instead of typing it below.  
				However, it is usually recommended you copy and paste into the Details box below.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="file" name="File" onChange="this.form.FileLinkDirect.checked = true;"><br>
				<input type="checkbox" name="FileLinkDirect" value="1"> Link directly to this file from the list
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
				<input type="submit" name="Submit" value="Add">
    		</td>
		</tr>
  	</table>
	</form>
