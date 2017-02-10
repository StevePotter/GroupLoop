<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_news_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add A News Update</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Add" then
	if Request("Body") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddNews"
		.CommandType = adCmdStoredProc
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@ModifiedID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@IP", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@Body", adVarWChar, adParamInput, 4000 )

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

	<p>The news has been added. &nbsp;<a href="admin_news_add.asp">Click here</a> to add more news.</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	<a href="inserts_view.asp" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>

	<p>The news appears on the home page until you delete it.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_news_add.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* News (inserts allowed)</td>
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