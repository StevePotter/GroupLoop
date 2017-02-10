<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeQuotes ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_quotes_add.asp")
if not (LoggedAdmin() or CBool( QuotesMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Add A Quote</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
'Add the story
if Request("Submit") = "Add" then
	if Request("Author") = "" or Request("Quote") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddQuote"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		if Request("Private") = "1" then 
			.Parameters("@IsPrivate") = 1
		else
			.Parameters("@IsPrivate") = 0
		end if
		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@ModifiedID") = Session("MemberID")
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Author") = Format( Request("Author") )
		.Parameters("@Quote") = Format( Request("Quote") )
		.Parameters("@Description") = GetTextArea( Request("Description") )

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ItemID")
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<p>Your quote has been added. &nbsp;<a href="members_quotes_add.asp">Click here</a> to add another.<br>
	<a href="quotes_read.asp?ID=<%=intID%>">Click here</a> to read it.
	</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	<p>If you only want members to be able to read it, you should check the private box.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_quotes_add.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
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
      		<td class="<% PrintTDMain %>" align="right">* Person who said the quote</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Author" size="55">
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Quote (do not include quotation signs)</td>
    		<td class="<% PrintTDMain %>"> 
    			<textarea name="Quote" cols="55" rows="4" wrap="PHYSICAL"></textarea>
    		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Description of quote</td>
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