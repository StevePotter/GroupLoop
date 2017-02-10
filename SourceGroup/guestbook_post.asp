<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeGuestbook ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Add An Entry</span><br>
<span class=LinkText><a href="javascript:history.go(-1)">Back</a></span></p>

<%
'-----------------------Begin Code----------------------------
if LoggedMember then
'------------------------End Code-----------------------------
%>
	<p><b>Sorry, but members can't add entries to the guestbook.</b></p>
<%
elseif Request("Submit") = "Add" then
	if Request("Author") = "" or Request("Body") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddGuestbook"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@IP", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@Author", adVarWChar, adParamInput, 100 )
		.Parameters.Append .CreateParameter ("@Email", adVarWChar, adParamInput, 100 )
		.Parameters.Append .CreateParameter ("@Body", adVarWChar, adParamInput, 2000 )

		.Parameters("@CustomerID") = CustomerID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Author") = FormatNonMember( Request("Author") )
		.Parameters("@Email") = FormatNonMember( Request("EMail") )
		.Parameters("@Body") = GetTextArea( Request("Body") )

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ItemID")
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<p>
	Your entry has been added. &nbsp;<a href="guestbook_read.asp?ID=<%=intID%>">Click here</a> to read it.
	</p>
<%
else
%>
	* indicates required information<br>
	<form method="post" action="guestbook_post.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<% PrintTableHeader 0 %>
		<tr> 
   			<td class="<% PrintTDMain %>" align="right">* Your Name</td>
   			<td class="<% PrintTDMain %>"> 
   				<input type="text" name="Author" size="55">
   			</td>
		</tr>
		<tr> 
   			<td class="<% PrintTDMain %>" align="right">Your Email</td>
   			<td class="<% PrintTDMain %>"> 
   				<input type="text" name="EMail" size="55">
   			</td>
		</tr>
		<tr> 
   			<td class="<% PrintTDMain %>" align="right" valign="top">* Entry</td>
   			<td class="<% PrintTDMain %>"> 
				<% TextArea "Body", 55, 10, True, "" %>
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