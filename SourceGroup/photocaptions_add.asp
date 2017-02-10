<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludePhotoCaptions ) then Redirect("error.asp")
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("login.asp?Source=photocaptions_add.asp&ID=" & Request("ID"))

Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add A Caption</p>

<%
'-----------------------Begin Code----------------------------
intID = Request("ID")
if intID = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
intID = CInt(intID)

if not ValidPhoto( intID ) then Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))

'Add the story
if Request("Submit") = "Add" then
	if Request("Body") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddPhotoCaption"
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
		.Parameters("@PhotoID") = intID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Caption") = GetTextArea( Request("Body") )

		.Execute , , adExecuteNoRecords
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<p>Your caption has been added. &nbsp;<a href="photocaptions_add.asp?ID=<%=intID%>">Click here</a> to add another. <br>
	 <a href="photos_view.asp?ID=<%=intID%>">Click here</a> to view the photo with your new caption.
	
	</p>
<%
'-----------------------Begin Code----------------------------

else
'------------------------End Code-----------------------------
%>
	<p>If you only want members to be able to read it, you should check the private box.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>photocaptions_add.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=intID%>">
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
    		<td class="<% PrintTDMain %>" align="right" valign="top">* Caption</td>
    		<td class="<% PrintTDMain %>"> 
				<% TextArea "Body", 55, 4, True, "" %>
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


'-------------------------------------------------------------
'This function sees if a category exists
'-------------------------------------------------------------
Function ValidPhoto( intPhotoID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "PhotosValid"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@ItemID") = intPhotoID
		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing

	ValidPhoto = CBool(blExists)
End Function
'------------------------End Code-----------------------------
%>