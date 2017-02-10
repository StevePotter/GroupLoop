<!-- #include file="category_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
'if Request("Section") = "" or Request("AddLink") = "" then Redirect("message.asp?Message=" & Server.URLEncode("No section or add link was passed."))
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp")
Session.Timeout = 20

strSection = Request("Section")
strSectionTitle = Request("SectionTitle")
strAddLink = Request("AddLink")
strModifyLink = Request("ModifyLink")
strSectionLink = Request("SectionLink")
strItemTable = Request("ItemTable")
strItemNoun = Request("ItemNoun")
strShowPrivate = Request("ShowPrivate")

strLink = "Section=" & strSection & "&SectionTitle=" & strSectionTitle & "&AddLink=" & strAddLink & "&ModifyLink=" & strModifyLink & "&SectionLink=" & strSectionLink & "&ItemTable=" & strItemTable & "&ItemNoun=" & strItemNoun & "&ShowPrivate=" & strShowPrivate
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Categories</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
strMatch = "CustomerID = " & CustomerID & " AND Section = '" & strSection & "'"

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Date") = "" or Request("Name") = "" then Redirect("incomplete.asp")

	Query = "SELECT Private, Name, Date, IP, ModifiedID, ParentID, Body FROM Categories WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if strShowPrivate <> "No" then
		if Request("Private") = "1" then
			rsUpdate("Private") = 1
		else
			rsUpdate("Private") = 0
		end if
	end if

	intParentID = 0
	if Request("ParentID") <> "" then intParentID = CInt(Request("ParentID"))

	rsUpdate("ParentID") = intParentID

	strBody = GetTextArea( Request("Body") )
	rsUpdate("Body") = strBody

	rsUpdate("Date") = Request("Date")
	rsUpdate("Name") = Format( Request("Name") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	MakeLongNames(strSection)
'------------------------End Code-----------------------------
%>
	<p>The category has been edited. &nbsp;<a href="admin_categories_modify.asp?<%=strLink%>">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")

	intOutput = RecursiveDelete( intID, strSection )

	Set FileSystem = Nothing
	Set rsUpdate = Nothing

'------------------------End Code-----------------------------
%>
	<p>The category has been deleted. &nbsp;<a href="admin_categories_modify.asp?<%=strLink%>">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Name, LongName, Private, ParentID, Body FROM Categories WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	strChecked = ""
	if rsEdit("Private") = 1 then strChecked = "checked"

	Query = "SELECT Name, LongName, Private, ID FROM Categories WHERE (ID <> " & intID & " AND " & strMatch & " AND LEFT([LongName], " & Len(rsEdit("LongName")) & ") <> '" & Format(rsEdit("LongName")) & "' ) ORDER BY LongName"
'------------------------End Code-----------------------------
%>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_categories_modify.asp?<%=strLink%>" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
<%
	if strShowPrivate <> "No" then
%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>
<%
	end if
%>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Sub-Category of</td>
      	<td class="<% PrintTDMain %>"> 
			<% PrintCategoryPullDown rsEdit("ParentID"), 1, 0, 1, 1, strSection, "ParentID", Query %>
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
   		<td class="<% PrintTDMain %>" align="right" valign="top">Details (inserts allowed)</td>
   		<td class="<% PrintTDMain %>"> 
			<% TextArea "Body", 55, 20, True, ExtractInserts( rsEdit("Body") ) %>
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
		if not ValidCategory( intParentID, strSection ) then
			Redirect("error.asp?Message=" & Server.URLEncode("The category is invalid."))
		end if
		Query = "SELECT ID, Date, LongName, Name FROM Categories WHERE " & strMatch & " AND ParentID = " & intParentID & " ORDER BY LongName"
	else
		intParentID = 0
		Query = "SELECT ID, Date, LongName, Name FROM Categories WHERE " & strMatch & " ORDER BY LongName"
	end if


	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

	Response.Write "</p>"

	if intParentID = 0 then
%>
		<p class="LinkText"><a href="admin_categories_add.asp?<%=strLink%>">Add A Category</a></p>
<%
	else
		Response.Write "<p>" & GetCatHeiarchy(intCategoryID, strSectionLink, strSection, strSectionTitle) & "</p>"
%>
		<p class="LinkText"><a href="admin_categories_add.asp?ParentID=<%=intParentID%>&<%=strLink%>">Add Another Sub-Category To <%=GetCategoryName(intParentID, strSection)%></a></p>
<%
	end if
	if not rsPage.EOF then

		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Name = rsPage("LongName")
'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="admin_categories_modify.asp?<%=strLink%>"  name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
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
					<form METHOD="post" ACTION="admin_categories_modify.asp?<%=strLink%>"  name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
					<input type="hidden" name="CID" value="<%=ID%>">
						<td class="<% PrintTDMain %>"><a href="<%=strSectionLink%>?ID=<%=ID%>"><%=PrintTDLink( Name )%></a></td>
						<td class="<% PrintTDMainSwitch %>">
						<input type="button" value="Edit" onClick="Redirect('admin_categories_modify.asp?Submit=Edit&ID=<%=ID%>&<%=strLink%>')" > 
						<input type="button" value="Delete" onClick="DeleteBox('Are you sure you want to permanently delete this category and its contents?', 'admin_categories_modify.asp?Submit=Delete&ID=<%=ID%>&<%=strLink%>')">
						<input type="button" value="Add Sub-Categories" onClick="Redirect('admin_categories_add.asp?ParentID=<%=ID%>&<%=strLink%>')" >  
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


Sub DeleteCategory( intID )
	Query = "SELECT ID FROM Categories WHERE ID = " & intID & " AND " & strMatch
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close

	if strItemTable <> "" then
		Query = "DELETE " & strItemTable & " WHERE CategoryID = " & intID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
	end if

End Sub


Function RecursiveDelete( intParentID, strSection )
	if intParentID = 0 then
		RecursiveDelete = 0
		Exit Function
	end if

	'If we have kids, keep traversing the heiarchy
	do until not CategoryHasChild( intParentID, strSection )
		intFirstChild = GetCatChildID( intParentID )
		RecursiveDelete = RecursiveDelete( intFirstChild, strSection )
	loop
	'No kids, kill this category and back on out

	DeleteCategory intParentID
	RecursiveDelete = 0

End Function

'-------------------------------------------------------------
'This function gets the Child's ID
'-------------------------------------------------------------
Function GetCatChildID( intCategoryID )
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "GetCatChildID"
		.CommandType = adCmdStoredProc

		.Parameters.Refresh

		.Parameters("@CategoryID") = intCategoryID

		.Execute , , adExecuteNoRecords

		intChildID = .Parameters("@ChildID")
	End With
	Set cmdTemp = Nothing

	GetCatChildID = intChildID
End Function




'------------------------End Code-----------------------------
%>
