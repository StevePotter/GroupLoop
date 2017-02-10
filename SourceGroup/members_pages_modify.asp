<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_pages_modify.asp")
if not (LoggedAdmin or CBool( InfoPagesMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Info Pages</span><br>
<span class=LinkText><a href="<%=NonSecurePath%>members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
blLoggedAdmin = LoggedAdmin

strMatch = "CustomerID = " & CustomerID

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Title") = "" or Request("Body") = "" then Redirect("incomplete.asp")


	'put the content in local storage
	strText = Request("Body")

	if not UseWYSIWYGEdit() and GetCheckedResult(Request("HTML")) = 0 then
		strText = Format( strText )
		strText = AddInserts( strText )
	end if


	Query = "SELECT ID, MemberID, Title, HTML, ShowButton, Body, CustomerID, IP, ModifiedID, DisplayTitle, Privacy FROM InfoPages WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

		rsUpdate("HTML") = GetCheckedResult(Request("HTML"))
		rsUpdate("MemberID") = Session("MemberID")
		rsUpdate("ModifiedID") = Session("MemberID")
		rsUpdate("Title") = Format( Request("Title") )
		rsUpdate("Body") = strText
		rsUpdate("CustomerID") = CustomerID
		rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
		rsUpdate("DisplayTitle") = GetCheckedResult(Request("DisplayTitle"))
		if Request("Privacy") <> "" then rsUpdate("Privacy") = Request("Privacy")


		intShow = 0

		'We need to add a button
		if Request("Show") <> "" and Request("Show") <> "Nowhere" and rsUpdate("ShowButton") = 0 then
			Set cmdTemp = Server.CreateObject("ADODB.Command")
			With cmdTemp
				.ActiveConnection = Connect
				.CommandText = "AddMenuButton"
				.CommandType = adCmdStoredProc

				.Parameters.Refresh

				.Parameters("@Name") = "InfoPage" & intID
				.Parameters("@Align") = Request("Align")
				.Parameters("@Show") = Request("Show")
				.Parameters("@CustomerID") = CustomerID
				.Execute , , adExecuteNoRecords

			End With
			Set cmdTemp = Nothing

			intShow = 1

		elseif GetCheckedResult(Request("ShowButton")) = 0 then
			if rsUpdate("ShowButton") = 1 then
				'Delete the button if it's there
				Query = "DELETE MenuButtons WHERE CustomerID = " & CustomerID & " AND Name = 'InfoPage" & intID & "'"
				Connect.Execute Query, , adCmdText + adExecuteNoRecords
			end if
		elseif GetCheckedResult(Request("ShowButton")) = 1 then
			intShow = 1
		end if

		rsUpdate("ShowButton") = intShow

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	if Request("Title") = "Home Page" then
%>
		<!-- #include file="write_index.asp" -->
<%
	end if

	Session("ItemID") = intID

	Redirect("write_header_footer.asp?Source=members_pages_modify.asp?Submit=Edited&ID="& intID)

elseif strSubmit = "Edited" then
'------------------------End Code-----------------------------
%>
	<p>The page has been edited. &nbsp;	<a href="pages_read.asp?ID=<%=Session("ItemID")%>">Click here</a> to read it.</p>
	<a href="members_pages_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "DELETE InfoPages WHERE ID = " & intID & " AND " & strMatch
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	Query = "DELETE MenuButtons WHERE CustomerID = " & CustomerID & " AND Name = 'InfoPage" & intID & "'"
	Connect.Execute Query, , adCmdText + adExecuteNoRecords


	Redirect("write_header_footer.asp?Source=members_pages_modify.asp?Submit=Deleted")

elseif strSubmit = "Deleted" then
'------------------------End Code-----------------------------
%>
	<p>The page has been deleted. &nbsp;<a href="members_pages_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, HTML, Title, Body, ShowButton, DisplayTitle, Privacy FROM InfoPages WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if


	strText = rsEdit("Body")

	'If they are using a regular text box with autoformatting, set it
	if not UseWYSIWYGEdit() and rsEdit("HTML") = 0 then strText = FormatEdit( strText )

	'strText = ExtractInserts( strText )

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
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	<a href="inserts_view.asp?Table=InfoPages" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_pages_modify.asp" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
<%		 if not UseWYSIWYGEdit() then %>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Use pure HTML, no auto-formatting (recommended for advanced designers)</td>
		<td class="<% PrintTDMain %>"> 
			<% PrintCheckBox rsEdit("HTML"), "HTML" %>
     	</td>
   	</tr>
<%		end if 

		if DisplayPrivacy then
%>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Who can read it?</td>
      		<td class="<% PrintTDMain %>"> 
<%
				PrintRadioOption "Privacy", 0, "Anyone<br>", rsEdit("Privacy")
				PrintRadioOption "Privacy", 1, "Site Members Only<br>", rsEdit("Privacy")
			Response.Write "</td></tr>"
		end if
%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Should the title be shown when someone reads the page?</td>
		<td class="<% PrintTDMain %>"> 
			<% PrintCheckBox rsEdit("DisplayTitle"), "DisplayTitle" %>
     	</td>
   	</tr>
<%
	if rsEdit("ShowButton") = 1 then
%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Should there be a button (link) to it in the menu?</td>
		<td class="<% PrintTDMain %>"> 
			<% PrintCheckBox rsEdit("ShowButton"), "ShowButton" %>
     	</td>
   	</tr>
<%	else %>
			<td class="<% PrintTDMain %>" align="left" colspan="2">Put a button (link) for this page in the 
			<select name="Show">
	<%
			WriteOption "Menu", "Main Menu Only", "Nowhere"
			WriteOption "Footer", "Footer Only", "Nowhere"
			WriteOption "MenuFooter", "Main Menu And Footer", "Nowhere"
			WriteOption "Nowhere", "Do Not Show Button", "Nowhere"
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
<%	end if %>


	<tr> 
   		<td class="<% PrintTDMain %>" align="right">* Page Title</td>
   		<td class="<% PrintTDMain %>"> 
   			<input type="text" name="Title" size="55" value="<%=FormatEdit(rsEdit("Title"))%>">
   		</td>
	</tr>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">* Page Content</td>
    	<td class="<% PrintTDMain %>"> 
			<% TextArea "Body", 55, 5, True, strText %>
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
	Query = "SELECT ID, Date, MemberID, Title FROM InfoPages WHERE " & strMatch & " ORDER BY Title"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set MemberID = rsPage("MemberID")
		Set ItemTitle = rsPage("Title")
'-----------------------End Code----------------------------
%>
		<a href="members_pages_add.asp">Add a Page</a><br>
		<form METHOD="POST" ACTION="members_pages_modify.asp">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
		PrintTableHeader 0
		PrintTableTitle
		for j = 1 to rsPage.PageSize
			if not rsPage.EOF then
				PrintTableData
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
	else
'------------------------End Code-----------------------------
%>
		<p>Sorry, but there are no info pages at the moment.</p>
<%
'-----------------------Begin Code----------------------------
	end if
	rsPage.Close
	set rsPage = Nothing
end if





'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>
	<tr>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">Title</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
End Sub

'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
%>
	<form METHOD="POST" ACTION="members_pages_modify.asp">
	<input type="hidden" name="ID" value="<%=ID%>">
	<tr>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
		<td class="<% PrintTDMain %>"><a href="pages_read.asp?ID=<%=ID%>"><%=PrintTDLink(ItemTitle)%></a></td>
		<td class="<% PrintTDMainSwitch %>">
			<input type="submit" name="Submit" value="Edit">
			<input type="button" value="Delete" onClick="DeleteBox('If you delete this page, there is no way to get it back.  Are you sure?', 'members_pages_modify.asp?Submit=Delete&ID=<%=ID%>')">			
		</td>
		</tr>
	</form>
<%
End Sub
'------------------------End Code-----------------------------
%>