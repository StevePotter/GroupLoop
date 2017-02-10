<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeMedia ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_mediacategories_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Categories</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
strMatch = "CustomerID = " & CustomerID

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Date") = "" or Request("Name") = "" then Redirect("incomplete.asp")

	Query = "SELECT Name, Date, IP, DefaultCat, ModifiedID FROM MediaCategories WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if Request("DefaultCat") = "1" then
		Query = "UPDATE MediaCategories SET DefaultCat = 0 WHERE CustomerID = " & CustomerID
		Connect.Execute Query, , adCmdText + adExecuteNoRecords
	end if

	rsUpdate("Date") = Request("Date")
	rsUpdate("Name") = Format( Request("Name") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate("DefaultCat") = Request("DefaultCat")

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing


'------------------------End Code-----------------------------
%>
	<p>The category has been edited. &nbsp;<a href="admin_mediacategories_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM MediaCategories WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close

	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	strPath = GetPath ("media")

	Query = "SELECT FileName FROM Media WHERE CategoryID = " & intID
	rsUpdate.CacheSize = 100
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	do until rsUpdate.EOF
		strFileName = rsUpdate("FileName")

		if FileSystem.FileExists(strPath & "/" & strFileName) then FileSystem.DeleteFile(strPath & "/" & strFileName)

		rsUpdate.MoveNext
	loop
	Set FileSystem = Nothing
	Set rsUpdate = Nothing

	Query = "DELETE Media WHERE CategoryID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The category has been deleted. &nbsp;<a href="admin_mediacategories_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Name, Private, DefaultCat FROM MediaCategories WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
%>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_mediacategories_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="middle" align="right">
			Does this category automatically come up when someone clicks on <%=MediaTitle%>?
		</td>
		<td class="<% PrintTDMain %>" align="left">
			<% PrintRadio rsEdit("DefaultCat"), "DefaultCat" %>
		</td>
	</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Category Name</td>
     	<td class="<% PrintTDMain %>"> 
      		<input type="text" name="Name" size="50" value="<%=FormatEdit( rsEdit("Name") )%>">
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

	Query = "SELECT ID, Date, Name FROM MediaCategories WHERE " & strMatch & " ORDER BY Name"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Name = rsPage("Name")
'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="admin_mediacategories_modify.asp">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">&nbsp;</td>
			<td class="TDHeader">Category</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<form METHOD="post" ACTION="admin_mediacategories_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="media.asp?ID=<%=ID%>">View</a></td>
						<td class="<% PrintTDMain %>"><%=Name%></td>
						<td class="<% PrintTDMainSwitch %>"><input type="Submit" name="Submit" value="Edit"> 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this category, there is no way to get it or its files back.  Are you sure?', 'admin_mediacategories_modify.asp?Submit=Delete&ID=<%=ID%>')"></td>
					</tr>
				</form>
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
		<p>You have to create categories before you can modify them, <%=GetNickNameSession()%>.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if
'------------------------End Code-----------------------------
%>
