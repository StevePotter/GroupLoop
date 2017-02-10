<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeLinks ) then Redirect("error.asp")
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_links_add.asp")
if not (LoggedAdmin() or CBool( LinksMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Add A Link</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
'Add the story
if Request("Submit") = "Add" then
	if Request("URL") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddLink"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@IsPrivate", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@ModifiedID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@IP", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@Name", adVarWChar, adParamInput, 200 )
		.Parameters.Append .CreateParameter ("@URL", adVarWChar, adParamInput, 300 )
		.Parameters.Append .CreateParameter ("@Description", adVarWChar, adParamInput, 2000 )

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
		.Parameters("@URL") = Request("URL")
		.Parameters("@Description") = GetTextArea( Request("Description") )

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ItemID")
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<p>Your link has been added. &nbsp;<a href="members_links_add.asp">Click here</a> to add another.<br>
	<a href="links_read.asp?ID=<%=intID%>">Click here</a> to read it.
	</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	<p>If you only want members to be able to see it, you should check the private box.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_links_add.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
		<tr> 
			<td class="<% PrintTDMain %>" align="right">Private?</td>
			<td class="<% PrintTDMain %>"> 
				<input type="checkbox" name="Private" value="1">
			</td>
   		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Address of link.  If linking to a web site, keep the http://. Otherwise it won't work.  For example, to link to Yahoo, make sure you enter 'http://www.yahoo.com'</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="URL" size="55" value="http://">
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Name of link.  Leaving this blank will cause only the address to be displayed as the link.</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Name" size="55">
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Description of link.</td>
    		<td class="<% PrintTDMain %>"> 
 				<% TextArea "Description", 55, 4, True, "" %>
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