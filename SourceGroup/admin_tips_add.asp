<%
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_tips_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add A Tip</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Add" then
	if Request("Body") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddTip"
		.CommandType = adCmdStoredProc
		.Parameters.Refresh

		.Parameters("@CustomerID") = CustomerID
		.Parameters("@ModifiedID") = Session("MemberID")
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Body") = GetTextArea( Request("Body") )

		.Execute , , adExecuteNoRecords
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<!-- #include file="write_index.asp" -->

	<p>The tip has been added. &nbsp;<a href="admin_tips_add.asp">Click here</a> to add another.</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_tips_add.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Tip</td>
    		<td class="<% PrintTDMain %>"> 
			<% TextArea "Body", 55, 6, True, "" %>
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