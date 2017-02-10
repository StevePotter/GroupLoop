<!-- #include file="photos_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludePhotos ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_photocategories_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Categories</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
strMatch = "CustomerID = " & CustomerID

strSubmit = Request("Submit")


Table = "PhotoCategories"

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Date") = "" or Request("Name") = "" then Redirect("incomplete.asp")

	Query = "SELECT Private, Name, Date, IP, ModifiedID, ParentID, Body FROM PhotoCategories WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if Request("Private") = "1" then
		rsUpdate("Private") = 1
	else
		rsUpdate("Private") = 0
	end if

	intParentID = 0
	if Request("ParentID") <> "" then intParentID = CInt(Request("ParentID"))

	rsUpdate("ParentID") = intParentID

	rsUpdate("Body") = GetTextArea( Request("Body") )
	rsUpdate("Date") = Request("Date")
	rsUpdate("Name") = Format( Request("Name") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	MakeLongNames("PhotoCategories")
'------------------------End Code-----------------------------
%>
	<p>The category has been edited. &nbsp;<a href="admin_photocategories_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")

	strPath = GetPath ("photos")

	intOutput = RecursiveDelete( intID, Table )


	Set FileSystem = Nothing
	Set rsUpdate = Nothing

'------------------------End Code-----------------------------
%>
	<p>The category has been deleted. &nbsp;<a href="admin_photocategories_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Name, LongName, Private, ParentID, Body FROM PhotoCategories WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	strChecked = ""
	if rsEdit("Private") = 1 then strChecked = "checked"

	Query = "SELECT Name, LongName, Private, ID FROM " & Table & " WHERE (ID <> " & intID & " AND CustomerID = " & CustomerID & " AND LEFT([LongName], " & Len(rsEdit("LongName")) & ") <> '" & Format(rsEdit("LongName")) & "' ) ORDER BY LongName"
'------------------------End Code-----------------------------
%>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_photocategories_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>

	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Sub-Category of</td>
      	<td class="<% PrintTDMain %>"> 
			<% PrintCategoryPullDown rsEdit("ParentID"), 1, 0, 1, 1, Table, "ParentID", Query %>
     	</td>
    </tr>


	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Category Name</td>
     	<td class="<% PrintTDMain %>"> 
      		<input type="text" name="Name" size="50" value="<%=FormatEdit( rsEdit("Name") )%>">
    	</td>
	</tr>
		<tr> 
    		<td class="<% PrintTDMain %>" align="right" valign="top">Details</td>
    		<td class="<% PrintTDMain %>"> 
				<% TextArea "Body", 55, 20, True, rsEdit("Body") %>
    		</td>
		</tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
	</tr>

  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing

else

	'Get the Category category, if there is one
	if Request("CID") <> "" then
		intParentID = CInt(Request("CID"))
		'Check the Category ID
		if not ValidCategory( intParentID, Table ) then
			Redirect("error.asp?Message=" & Server.URLEncode("The category is invalid."))
		end if
		Query = "SELECT ID, Date, LongName, Name FROM " & Table & " WHERE " & strMatch & " AND ParentID = " & intParentID & " ORDER BY LongName"
	else
		intParentID = 0
		Query = "SELECT ID, Date, LongName, Name FROM " & Table & " WHERE " & strMatch & " ORDER BY LongName"
	end if


	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	Response.Write "</p>"

	if intParentID = 0 then
%>
		<p class="LinkText"><a href="admin_photocategories_add.asp">Add A Category</a></p>
<%
	else
		Response.Write "<p>" & GetCatHeiarchy(intCategoryID, "photos.asp", Table, PhotosTitle) & "</p>"
%>
		<p class="LinkText"><a href="admin_photocategories_add.asp?ParentID=<%=intParentID%>">Add Another Sub-Category To <%=GetCategoryName(intParentID, Table)%></a></p>
<%
	end if
	if not rsPage.EOF then

		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Name = rsPage("LongName")
'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="admin_photocategories_modify.asp"  name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">Category</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<tr>
					<form METHOD="post" ACTION="admin_photocategories_modify.asp"  name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
					<input type="hidden" name="CID" value="<%=ID%>">
						<td class="<% PrintTDMain %>"><a href="photos.asp?ID=<%=ID%>"><%=PrintTDLink( Name )%></a></td>
						<td class="<% PrintTDMainSwitch %>">
						<input type="button" value="Edit" onClick="Redirect('admin_photocategories_modify.asp?Submit=Edit&ID=<%=ID%>')" > 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this category, there is no way to get its photos.  Are you sure?', 'admin_photocategories_modify.asp?Submit=Delete&ID=<%=ID%>')">
						<input type="button" value="Add Sub-Categories" onClick="Redirect('admin_photocategories_add.asp?ParentID=<%=ID%>')" >  
						</td>
					</form>
				</tr>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
	else
'------------------------End Code-----------------------------
%>
		<p>Sorry, but there are no categories at the moment.</p>
<%
'-----------------------Begin Code----------------------------
	end if
	Set rsPage = Nothing
end if


Sub DeletePhotoCategory( intID )
	Query = "SELECT ID FROM PhotoCategories WHERE ID = " & intID & " AND " & strMatch
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close


	Query = "SELECT ID, Ext, Thumbnail, ThumbnailExt FROM Photos WHERE CategoryID = " & intID
	rsUpdate.CacheSize = 100
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	do until rsUpdate.EOF
		'First delete the photo
		strFileName = rsUpdate("ID") & "." & rsUpdate("Ext")
		if FileSystem.FileExists(strPath & "/" & strFileName) then FileSystem.DeleteFile(strPath & "/" & strFileName)
		'Now delete the thumbnail
		if rsUpdate("Thumbnail") = 1 then
			strFileName = rsUpdate("ID") & "t." & rsUpdate("ThumbnailExt")
			if FileSystem.FileExists(strPath & "/" & strFileName) then FileSystem.DeleteFile(strPath & "/" & strFileName)
		end if

		Query = "DELETE PhotoCaptions WHERE PhotoID = " & rsUpdate("ID")
		Connect.Execute Query, , adCmdText + adExecuteNoRecords

		rsUpdate.MoveNext
	loop

	rsUpdate.Close

	Query = "DELETE Photos WHERE CategoryID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords


End Sub


Function RecursiveDelete( intParentID, strTable )
	if intParentID = 0 then
		RecursiveDelete = 0
		Exit Function
	end if

	'If we have kids, keep traversing the heiarchy
	do until not CategoryHasChild( intParentID, strTable )
		intFirstChild = GetCatChildID( intParentID, strTable )
		RecursiveDelete = RecursiveDelete( intFirstChild, strTable )
	loop
	'No kids, kill this category and back on out

	DeletePhotoCategory intParentID
	RecursiveDelete = 0

End Function

'-------------------------------------------------------------
'This function gets the Child's ID
'-------------------------------------------------------------
Function GetCatChildID( intCategoryID, strTable )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetCatChildID"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CategoryID") = intCategoryID
		.Parameters("@Table") = strTable

		.Execute , , adExecuteNoRecords

		intChildID = .Parameters("@ChildID")
	End With
	Set cmdTemp = Nothing

	GetCatChildID = intChildID
End Function




'------------------------End Code-----------------------------
%>
