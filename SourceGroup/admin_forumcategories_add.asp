<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeForum ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_forumcategories_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add A New Topic</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Add" then
	if Request("Type") = "" or Request("Name") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddForumCategories"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@IsPrivate", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@ModifiedID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@IP", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@MembersOnly", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamInput, 200 )

		'Fully public
		if Request("Type") = "1" then 
			.Parameters("@IsPrivate") = 0
			.Parameters("@MembersOnly") = 0
		elseif Request("Type") = "2" then 
			.Parameters("@IsPrivate") = 0
			.Parameters("@MembersOnly") = 1
		elseif Request("Type") = "3" then 
			.Parameters("@IsPrivate") = 1
			.Parameters("@MembersOnly") = 0
		end if

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
	<p>The topic has been added. &nbsp;<a href="admin_forumcategories_add.asp">Click here</a> to add another.<br>
	<a href="forum.asp?ID=<%=intID%>">Click here</a> to go to it.
	</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	<p>Remember, private messages (which only members can read) can be posted in any type of topic.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_forumcategories_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">* Privacy</td>
			<td class="<% PrintTDMain %>"> 
				<input type="radio" name="Type" value="1" checked>Both members and non-members can post messages.  Non-members can read public messages.<br>
				<input type="radio" name="Type" value="2" >Only members can post messages.  Non-members can read public messages.<br>
				<input type="radio" name="Type" value="3" >Only members can post and read messages.  Non-members cannot even read the subjects.
			</td>
   		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Topic Name</td>
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