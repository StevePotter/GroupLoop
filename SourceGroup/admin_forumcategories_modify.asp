<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeForum ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_forumcategories_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Topics</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
strMatch = "CustomerID = " & CustomerID

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Date") = "" or Request("Type") = "" or Request("Name") = "" then Redirect("incomplete.asp")

	Query = "SELECT Private, MembersOnly, Name, Date, IP, ModifiedID FROM ForumCategories WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if Request("Type") = "3" then
		intPrivate = 1
	else
		intPrivate = 0
	end if
	if Request("Type") = "2" then
		intMembersOnly = 1
	else
		intMembersOnly = 0
	end if

	rsUpdate("Private") = intPrivate
	rsUpdate("Date") = Request("Date")
	rsUpdate("MembersOnly") = intMembersOnly
	rsUpdate("Name") = Format( Request("Name") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The topic has been edited. &nbsp;<a href="admin_forumcategories_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM ForumCategories WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE ForumMessages WHERE CategoryID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The topic has been deleted. &nbsp;<a href="admin_forumcategories_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Name, MembersOnly, Private FROM ForumCategories WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	str1 = str2 = str3 = ""
	if rsEdit("MembersOnly") = 1 then 
		str2 = "checked"
	elseif rsEdit("Private") = 1 then
		str3 = "checked"
	else
		str1 = "checked"
	end if
'------------------------End Code-----------------------------
%>
	<p>Remember, private messages (which only members can read) can be posted in any type of topic.</p>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_forumcategories_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">* Privacy</td>
		<td class="<% PrintTDMain %>"> 
			<input type="radio" name="Type" value="1" <%=str1%>>Both members and non-members can post messages.  Non-members can read public messages.<br>
			<input type="radio" name="Type" value="2" <%=str2%>>Only members can post messages.  Non-members can read public messages.<br>
			<input type="radio" name="Type" value="3" <%=str3%>>Only members can post and read messages.  Non-members cannot even read the subjects.
		</td>
   	</tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Topic Name</td>
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

	Query = "SELECT ID, Date, Name FROM ForumCategories WHERE " & strMatch & " ORDER BY Name"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Name = rsPage("Name")
'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="admin_forumcategories_modify.asp">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">&nbsp;</td>
			<td class="TDHeader">Topic</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<form METHOD="post" ACTION="admin_forumcategories_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="forum.asp?ID=<%=ID%>">View</a></td>
						<td class="<% PrintTDMain %>"><%=Name%></td>
						<td class="<% PrintTDMainSwitch %>"><input type="Submit" name="Submit" value="Edit"> 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this topic, there is no way to get it or its messages back.  Are you sure?', 'admin_forumcategories_modify.asp?Submit=Delete&ID=<%=ID%>')"></td>
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
		<p>You have to create topics for you can modify them, <%=GetNickNameSession()%>.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if


'------------------------End Code-----------------------------
%>