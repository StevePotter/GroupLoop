<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeAnnouncements ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_sectionitems_modify.asp")
if not (LoggedAdmin or CBool( AnnouncementsMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))

Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Announcements</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
blLoggedAdmin = LoggedAdmin

if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")

	Query = "SELECT Private, Subject, Date, Body, IP, ModifiedID FROM Announcements WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if Request("Private") = "1" then 
		rsUpdate("Private") = 1
	else
		rsUpdate("Private") = 0
	end if
	rsUpdate("Subject") = Format( Request("Subject") )
	if blLoggedAdmin then rsUpdate("Date") = AssembleDate( "Date" )
	rsUpdate("Body") = GetTextArea( Request("Body") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The announcement has been edited. &nbsp;<a href="announcements.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM Announcements WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'Announcements' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The announcement has been deleted. &nbsp;<a href="announcements.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT * FROM CustomSectionItems WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if rsEdit("Private") = 1 then 
		strChecked = "checked"
	else
		strChecked = ""
	end if


	intSectionID = CInt(rsEdit("SectionID"))


	Query = "SELECT * FROM Sections WHERE ID = " & intSectionID
	Set rsSection = Server.CreateObject("ADODB.Recordset")
	rsSection.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	if rsSection.EOF then Redirect("error.asp")

	Noun = rsSection("NounSingular")
	PluralNoun = rsSection("NounPlural")

	if rsSection("ModifySecurity") = "Administrators" and not LoggedAdmin() then
		Set rsSection = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("Sorry" & GetNickNameSession() & ", but only administrators may add " & PluralNoun & "."))
	end if
%>
	<p align="<%=HeadingAlignment%>"><span class=Heading>Add <%=PrintAn(Noun)%>&nbsp;<%=PrintFirstCap(Noun)%></span><br>
	<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span><br>
	</p>

	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";


			if(strError == "") {
<%

	for i = 1 to 10
		if rsSection("FieldName"&i) <> "" and rsSection("RequireFieldInput"&i) = 1 then
%>
			if (form.Field<%=i%>.value == "")
				strError += "          You forgot to enter something for <%=rsSection("FieldName"&i)%>. \n";
<%
		end if
	next
%>

				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can update this: \n" + strError;
				alert (strError);
				return false;
			}   
		}
	//-->
	</script>

	* Indicates Required Info

	<form enctype="multipart/form-data" method="post" action="<%=SecurePath%>members_sectionitems_add_process.asp" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted" name="MyForm">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=intSectionID%>">

<%
	PrintTableHeader 100

	for FieldNums = 1 to 10
		PrintInput FieldNums

	next

%>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Add">
    	</td>
    </tr>
	</table>
	</form>
<%

end if


















































	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.Subject.value == "")
				strError += "          You forgot the subject. \n";
			if (form.Body.value == "")
				strError += "          You forgot the details. \n";

			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can update this: \n" + strError;
				alert (strError);
				return false;
			}   
		}

	//-->
	</SCRIPT>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	<a href="inserts_view.asp?Table=InfoPages" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_sectionitems_modify.asp" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>
<%	if blLoggedAdmin then %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<% DatePulldown "Date", rsEdit("Date"), 0 %>
     	</td>
    </tr>
<%	end if %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Subject</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Subject" size="55" value="<%=FormatEdit( rsEdit("Subject") )%>">
     	</td>
    </tr>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">* Details (inserts allowed)</td>
    	<td class="<% PrintTDMain %>"> 
			<% TextArea "Body", 55, 10, True, ExtractInserts( rsEdit("Body") ) %>
    	</td>
    </tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
    </tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing
else
	Redirect("announcements.asp")
end if

'------------------------End Code-----------------------------
%>