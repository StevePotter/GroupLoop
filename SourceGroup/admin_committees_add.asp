<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeCommittees ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedAdmin then Redirect("members.asp?Source=admin_committees_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add A New Committee</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Add" then
	if Request("Name") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddCommittees"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@ModifiedID") = Session("MemberID")
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Name") = Format( Request("Name") )

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ItemID")
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<p>The committee has been added. &nbsp;<a href="admin_committees_add.asp">Click here</a> to add another.</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	* indicates required information<br>
	<form method="post" action="<%=NonSecurePath%>admin_committees_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<% PrintTableHeader 0 %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Committee Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Name" size="50">
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