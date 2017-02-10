<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_pages_add.asp")
if not (LoggedAdmin or CBool( InfoPagesMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Add An Info Page</span><br>
<span class=LinkText><a href="<%=NonSecurePath%>members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
'Add the story
if Request("Submit") = "Add" then
	if Request("Title") = "" or Request("Body") = "" then Redirect("incomplete.asp")

	'put the content in local storage
	strText = Request("Body")

	if GetCheckedResult(Request("HTML")) = 0 then
		strText = GetTextArea( strText )
	else
		strText = AddInserts( strText )
	end if

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddInfoPage"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ID")

		'Add the menu button for the info page
		if Request("Show") <> "Nowhere" then
			.CommandText = "AddMenuButton"

			.Parameters.Refresh


			.Parameters("@Name") = "InfoPage" & intID
			.Parameters("@Align") = Request("Align")
			.Parameters("@Show") = Request("Show")
			.Parameters("@CustomerID") = CustomerID
			.Parameters("@Position") = 0

			.Execute , , adExecuteNoRecords

			intShow = 1
		else
			intShow = 0
		end if

	End With
	Set cmdTemp = Nothing





	'Add the page
	Set rsNew = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT ID, MemberID, Title, HTML, ShowButton, Body, CustomerID, IP, ModifiedID, DisplayTitle, Privacy FROM InfoPages WHERE ID = " & intID
	rsNew.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Update the fields
		rsNew("HTML") = GetCheckedResult(Request("HTML"))
		rsNew("ShowButton") = intShow

		rsNew("MemberID") = Session("MemberID")
		rsNew("ModifiedID") = Session("MemberID")
		rsNew("Title") = Format( Request("Title") )
		rsNew("DisplayTitle") = GetCheckedResult(Request("DisplayTitle"))
		rsNew("Body") = strText
		if Request("Privacy") <> "" then rsNew("Privacy") = Request("Privacy")
		rsNew("CustomerID") = CustomerID

		if Request("Show") <> "Nowhere" then
			rsNew("ShowButton") = 1
		else
			rsNew("ShowButton") = 0
		end if

		rsNew("IP") = Request.ServerVariables("REMOTE_HOST")

	rsNew.Update

	Session("ItemID") = intID
	Session("Show") = intShow

	rsNew.Close
	Set rsNew = Nothing

	Redirect("write_header_footer.asp?Source=members_pages_add.asp?Submit=Changed&ID="&intID)

elseif Request("Submit") = "Changed" then
'------------------------End Code-----------------------------
%>
	<p>The info page has been added.<br>
	<a href="members_pages_add.asp">Add another page.</a><br>
<%	if Session("Show") = 1 and LoggedAdmin() then %>
		<a href="admin_buttons_modify.asp">Change the position of the page's button.</a><br>
<%	end if %>
	<a href="pages_read.asp?ID=<%=Session("ItemID")%>">View the page.</a><br>

	</p>
<%
'-----------------------Begin Code----------------------------

else

	DisplayPrivacy = not cBool(SiteMembersOnly)

%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
			if (form.Title.value == "")
				strError += "          You forgot the page title. \n";
			if (form.Body.value == "")
				strError += "          You forgot the page contents. \n";


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
	<form method="post" action="<%=SecurePath%>members_pages_add.asp" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
<%		 if not UseWYSIWYGEdit() then %>
		<tr> 
			<td class="<% PrintTDMain %>" align="left" colspan="2">Use pure HTML, no auto-formatting (recommended only for advanced designers)&nbsp;
				<% PrintCheckBox 0, "HTML" %>
     		</td>
   		</tr>
<%		end if %>

		<tr> 
			<td class="<% PrintTDMain %>" align="left" colspan="2">Should the title be displayed when someone reads the page?&nbsp;
				<% PrintCheckBox 1, "DisplayTitle" %>
   			</td>
   		</tr>

		<tr> 
			<td class="<% PrintTDMain %>" align="left" colspan="2">Put a button (link) for this page in the 
			<select name="Show">
	<%
			WriteOption "Menu", "Main Menu Only", "MenuFooter"
			WriteOption "Footer", "Footer Only", "MenuFooter"
			WriteOption "MenuFooter", "Main Menu And Footer", "MenuFooter"
			WriteOption "Nowhere", "Do Not Show Button", "MenuFooter"
	%>
			</select> and align it to the 
			<select name="Align">
	<%

			WriteOption "Left", "Left", "Left"
			WriteOption "Right", "Right", "Left"
			WriteOption "Top", "Top", "Left"
	%>
			</select>
			</td>
   		</tr>
<%
		if DisplayPrivacy then
%>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Who can read it?</td>
      		<td class="<% PrintTDMain %>"> 
<%
				PrintRadioOption "Privacy", 0, "Anyone<br>", 0
				PrintRadioOption "Privacy", 1, "Site Members Only<br>", 0
			Response.Write "</td></tr>"
		end if
%>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Page Title</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Title" size="55">
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Page Content (inserts allowed)</td>
    		<td class="<% PrintTDMain %>"> 
				<% TextArea "Body", 55, 30, True, "" %>
    		</td>
		</tr>

		<tr>
    		<td colspan="2" align="center" class="<% PrintTDMain %>">
				<input type="submit" name="Submit" value="Add">
    		</td>
		</tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if

'------------------------End Code-----------------------------
%>