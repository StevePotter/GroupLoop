<!-- #include file="photos_functions.asp" -->
<%
'-----------------------Begin Code----------------------------
if not CBool( IncludePhotos ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_photocategories_add.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Add A New Category</span><br>
<span class=LinkText><a href="<%=NonSecurePath%>members.asp">Back To <%=MembersTitle%></a></span><br>
<span class=LinkText><a href="admin_photocategories_modify.asp">Back To Category List</a></span>
</p>
<%
'-----------------------Begin Code----------------------------
if Request("Submit") = "Add" then
	if Request("Name") = "" then Redirect("incomplete.asp")

	'Get the parent category, if there is one
	if Request("ParentID") <> "" then
		intParentID = CInt(Request("ParentID"))
		'Check the parent ID
		if not ValidCategory( intParentID, "PhotoCategories" ) then
			Redirect("error.asp?Message=" & Server.URLEncode("The parent category is invalid."))
		end if
	else
		intParentID = 0
	end if

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddCategory"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@Table") = "PhotoCategories"
		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@ModifiedID") = Session("MemberID")
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@ParentID") = intParentID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Name") = Format( Request("Name") )
		.Parameters("@Body") = GetTextArea( Request("Body") )

		.Execute , , adExecuteNoRecords
		intID = .Parameters("@ItemID")

		.CommandText = "UpdateCategoryLongName"
		.Parameters.Refresh

		.Parameters("@ItemID") = intID
		.Parameters("@Table") = "PhotoCategories"
		.Parameters("@LongName") = GetCatHeiarchy( intID, "", "PhotoCategories", "" )

		.Execute , , adExecuteNoRecords

	End With
	Set cmdTemp = Nothing

	if intParentID = 0 then
'------------------------End Code-----------------------------
%>
	<p>The category has been added. &nbsp;<a href="admin_photocategories_add.asp">Click here</a> to add another.<br>
	<a href="admin_photocategories_add.asp?ParentID=<%=intID%>">Click here</a> to add add sub-categories to it.<br>
	<a href="members_photos_add.asp?ID=<%=intID%>">Click here</a> to add photos to it.<br>
	<a href="admin_photocategories_modify.asp">Click here</a> to view the list of categories.<br>
	</p>
<%
'-----------------------Begin Code----------------------------
	else
'------------------------End Code-----------------------------
%>
	<p>The category has been added. &nbsp;<a href="admin_photocategories_add.asp?ParentID=<%=intParentID%>">Click here</a> to add another sub-category to <%=GetCategoryName(intParentID, "PhotoCategories")%>.<br>
	<a href="admin_photocategories_add.asp?ParentID=<%=intID%>">Click here</a> to a sub-category to this sub-category (don't get too carried away though).<br>
	<a href="admin_photocategories_add.asp">Click here</a> to add another category (not a sub-category).<br>
	<a href="members_photos_add.asp?ID=<%=intID%>">Click here</a> to add photos to it.<br>
	<a href="admin_photocategories_modify.asp">Click here</a> to view the list of categories.<br>
	</p>
<%
'-----------------------Begin Code----------------------------
	end if
else
	intParentID = Request("ParentID")
'------------------------End Code-----------------------------
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			if (form.Name.value == "" ){
				strError = "You must enter a category name."
				alert (strError);
				return false;
			}
			else{
				return true;
			}
		}
	//-->
	</SCRIPT>
	<a href="inserts_view.asp" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>


	* indicates required information<br>

	<form method="post" action="admin_photocategories_add.asp" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
<%	if intParentID <> "" then %>
	<input type="hidden" name="ParentID" value="<%=intParentID%>">
<%
	end if
	PrintTableHeader 0
	if intParentID = "" then
%>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">Make This a Sub-Category of</td>
      		<td class="<% PrintTDMain %>"> 
			<% PrintCategoryPullDown 0, 1, 0, 1, 1, "PhotoCategories", "ParentID", "" %>
     		</td>
		</tr>
<%	end if %>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Category Name</td>
      		<td class="<% PrintTDMain %>"> 
       			<input type="text" name="Name" size="50">
     		</td>
		</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Details (inserts allowed)</td>
    		<td class="<% PrintTDMain %>"> 
				<% TextArea "Body", 55, 20, True, "" %>
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