<!-- #include file="media_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CatsExist then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but no files can be added until the administrator creates a category."))
if not LoggedMember then Redirect("members.asp?Source=members_media_add.asp")
if not CBool( IncludeMedia ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not (LoggedAdmin or CBool( MediaMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add Media</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>
<script language="JavaScript">
<!--
	function submit_page(form) {
		//Error message variable
		var strError = "";
		if (form.MediaFile.value == "")
			strError += "Sorry, but you forgot to select a file to upload. \n";

		if(strError == "") {
			alert('Uploading files may take some time, so please be patient.  After pressing OK, your file will upload.  Please do not constantly press the Add button.');
			return true;
		}
        else{
			alert (strError);
			return false;
		}   
	}

//-->
</SCRIPT>


<p>For security purposes, no executable files can be uploaded.</p>
* indicates required information<br>

<form enctype="multipart/form-data" method="post" action="<%=SecurePath%>members_media_add_process.asp" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted" name="MyForm">
<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
<input type="hidden" name="Password" value="<%=Session("Password")%>">
<% PrintTableHeader 0 %>
<tr>
	<td class="<% PrintTDMain %>"  align="right">
		Category
	</td>
	<td class="<% PrintTDMain %>">
<%
'-----------------------Begin Code----------------------------
		PrintCategoryPullDown Request("ID"), 0, 1
'------------------------End Code-----------------------------
%>
	</td>
</tr>
<tr>
	<td class="<% PrintTDMain %>" valign="top" align="right">
		Description of File
	</td>
	<td class="<% PrintTDMain %>">
		<input type="text" size="50" name="Name">
	</td>
</tr>
<tr>
	<td class="<% PrintTDMain %>" valign="top" align="right">
		* File
	</td>
	<td class="<% PrintTDMain %>">
		<input type="file" name="MediaFile">
	</td>
</tr>
</tr>
<tr>
	<td class="<% PrintTDMain %>" align="center" colspan="2">
		<input type="submit" name="Submit" value="Add"></td>
	</td>
</tr>
</table>
</form>
<%
Function CatsExist()
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "MediaCategoriesExist"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@Exists", adInteger, adParamOutput )

		.Parameters("@CustomerID") = CustomerID

		.Execute , , adExecuteNoRecords
		blExists = .Parameters("@Exists")
	End With
	Set cmdTemp = Nothing
	CatsExist = CBool(blExists)
End Function
%>