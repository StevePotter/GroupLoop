<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeMedia ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_mediacategories_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add A New Category</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Add" then
	if Request("Name") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddMediaCategories"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@IsPrivate", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@ModifiedID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@DefaultCat", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@IP", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamInput, 200 )

		'Fully public
		if Request("Private") = "1" then 
			.Parameters("@IsPrivate") = 1
		else
			.Parameters("@IsPrivate") = 0
		end if

		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@ModifiedID") = Session("MemberID")
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Name") = Format( Request("Name") )
		.Parameters("@DefaultCat") = Request("DefaultCat")

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ItemID")
	End With
	Set cmdTemp = Nothing

'------------------------End Code-----------------------------
%>
	<p>The category has been added. &nbsp;<a href="admin_mediacategories_add.asp">Click here</a> to add another.<br>
	<a href="media.asp?ID=<%=intID%>">Click here</a> to go to it.
	</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	* indicates required information<br>

	<form method="post" action="<%=SecurePath%>admin_mediacategories_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Does this category automatically come up when someone clicks on <%=MediaTitle%>?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				<% PrintRadio 0, "DefaultCat" %>
			</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Category Name</td>
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