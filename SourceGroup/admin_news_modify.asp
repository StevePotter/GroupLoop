<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedAdmin then Redirect("members.asp?Source=admin_news_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify News</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
blLoggedAdmin = LoggedAdmin

strMatch = "CustomerID = " & CustomerID

strSubmit = Request("Submit")

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Body") = "" then Redirect("incomplete.asp")

	Query = "SELECT Date, Body, IP, ModifiedID FROM News WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("Body") = GetTextArea( Request("Body") )
	rsUpdate("Date") = AssembleDate( "Date" )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<!-- #include file="write_index.asp" -->

	<p>The news has been edited. &nbsp;<a href="admin_news_modify.asp">Click here</a> to modify more news.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM News WHERE ID = " & intID & " AND " & strMatch
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
'------------------------End Code-----------------------------
%>
	<!-- #include file="write_index.asp" -->

	<p>The news has been deleted. &nbsp;<a href="admin_news_modify.asp">Click here</a> to modify more news.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Body FROM News WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	<a href="inserts_view.asp?Table=InfoPages" target="_blank">Click here</a> for page inserts.<br>
	<a href="formatting_view.asp" target="_blank">Click here</a> for formatting tips.<br>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>admin_news_modify.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<% DatePulldown "Date", rsEdit("Date"), 0 %>
     	</td>
    </tr>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">* News (inserts allowed)</td>
    	<td class="<% PrintTDMain %>"> 
			<% TextArea "Body", 55, 10, True, ExtractInserts( rsEdit("Body") ) %>
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
	Query = "SELECT ID, Date, Body FROM News WHERE (CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Body = rsPage("Body")
'-----------------------End Code----------------------------
%>
		<form METHOD="POST" ACTION="admin_news_modify.asp">
<%
'-----------------------Begin Code----------------------------
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">Date</td>
			<td class="TDHeader">News</td>
			<td class="TDHeader">&nbsp;</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
			<tr>
				<form METHOD="post" ACTION="admin_news_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
						<td class="<% PrintTDMain %>"><%=FormatDateTime(ItemDate, 2)%></td>
						<td class="<% PrintTDMain %>"><%=Body%></td>
						<td class="<% PrintTDMain %>"><input type="submit" name="Submit" value="Edit"></td>
						<td class="<% PrintTDMainSwitch %>"><input type="button" value="Delete" onClick="DeleteBox('If you delete this news, there is no way to get it back.  Are you sure?', 'admin_news_modify.asp?Submit=Delete&ID=<%=ID%>')"></td>
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
		<p>You have to create news before you can modify it, <%=GetNickNameSession()%>.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if
%>