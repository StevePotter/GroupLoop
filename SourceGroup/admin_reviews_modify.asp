<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_reviews_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Reviews</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
if Request("TargetID") = "" or Request("TargetTable") = "" or Request("Source") = "" then Redirect("error.asp")
intTargetID = CInt(Request("TargetID"))
strTable = Request("TargetTable")
strSource = Request("Source")

strMatch = "CustomerID = " & CustomerID

'update info
if Request("Submit") = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if Request("Date") = "" or Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT Subject, Author, EMail, Date, Body, IP, ModifiedID FROM Reviews WHERE ID = " & intID & " AND " & strMatch 
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("Subject") = Format( Request("Subject") )
	rsUpdate("Author") = Request("Author")
	rsUpdate("EMail") = Request("EMail")
	rsUpdate("Date") = Request("Date")
	rsUpdate("Body") = GetTextArea( Request("Body") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The review has been edited. &nbsp;<a href="admin_reviews_modify.asp?Source=<%=strSource%>&TargetID=<%=intTargetID%>&TargetTable=<%=strTable%>">Click here</a> to modify another.<br>
		<a href="<%=strSource%>">Click here</a> to go back.
	</p>
<%
'-----------------------Begin Code----------------------------
elseif Request("Submit") = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "DELETE Reviews WHERE ID = " & intID & " AND " & strMatch 
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The review has been deleted. &nbsp;<a href="admin_reviews_modify.asp?Source=<%=strSource%>&TargetID=<%=intTargetID%>&TargetTable=<%=strTable%>">Click here</a> to modify another.  <br>
	<a href="<%=strSource%>">Click here</a> to go back.
	</p>
<%
'-----------------------Begin Code----------------------------

elseif Request("Submit") = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, Date, Author, EMail, MemberID, Subject, Body FROM Reviews WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	* indicates required information<br>
	<form method="post" action="admin_reviews_modify.asp" name="MyForm" onSubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<input type="hidden" name="TargetID" value="<%=intTargetID%>">
	<input type="hidden" name="TargetTable" value="<%=strTable%>">
	<input type="hidden" name="Source" value="<%=strSource%>">
	<%PrintTableHeader 0%>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
	<%
	if rsEdit("MemberID") = 0 then
	%>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			* Author's Name
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="25" name="Author" value="<%=rsEdit("Author")%>">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			Author's E-Mail
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="25" name="EMail" value="<%=rsEdit("EMail")%>">
		</td>
	</tr>
	<% end if %>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			* Headline for the review
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="25" name="Subject" value="<%=FormatEdit( rsEdit("Subject") )%>">
		</td>
	</tr>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">* Review</td>
		<td class="<% PrintTDMain %>"> 
			<% TextArea "Body", 55, 8, True, rsEdit("Body") %>
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
%>
	<p><a href="<%=strSource%>">Click here</a> to go back.</p>
<%
	Query = "SELECT ID, Date, Author, EMail, MemberID, Subject, Body FROM Reviews WHERE (TargetID = " & intTargetID & " AND TargetTable = '" & strTable & "' AND CustomerID = " & CustomerID & ") ORDER BY Date DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsPage.EOF then
		Set rsPage = Nothing
		Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but there are no reviews for that item."))
	end if

	Set ID = rsPage("ID")
	Set ItemDate = rsPage("Date")
	Set Author = rsPage("Author")
	Set EMail = rsPage("EMail")
	Set MemberID = rsPage("MemberID")
	Set Subject = rsPage("Subject")
	Set Body = rsPage("Body")
%>
	<form METHOD="POST" ACTION="admin_reviews_modify.asp">
	<input type="hidden" name="TargetID" value="<%=intTargetID%>">
	<input type="hidden" name="TargetTable" value="<%=strTable%>">
	<input type="hidden" name="Source" value="<%=strSource%>">
<%
	PrintPagesHeader
	PrintTableHeader 0
%>
	<tr>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">Author</td>
		<td class="TDHeader">Headline</td>
		<td class="TDHeader">Review</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
	for j = 1 to rsPage.PageSize
		if not rsPage.EOF then
			if MemberID > 0 then
				strAuthor = GetNickNameLink( MemberID )
			elseif InStr( EMail, "@" ) then
				strAuthor = "<a href='mailto:" & EMail & "'>" & Author & "</a>"
			else
				strAuthor = Author
			end if
'------------------------End Code-----------------------------
%>
			<tr>
			<form METHOD="POST" ACTION="admin_reviews_modify.asp">
			<input type="hidden" name="ID" value="<%=ID%>">
			<input type="hidden" name="TargetID" value="<%=intTargetID%>">
			<input type="hidden" name="TargetTable" value="<%=strTable%>">
			<input type="hidden" name="Source" value="<%=strSource%>">
				<td class="<% PrintTDMain %>"><%=FormatDateTime(ItemDate, 2)%></td>
				<td class="<% PrintTDMain %>"><%=strAuthor%></td>
				<td class="<% PrintTDMain %>"><%=Subject%></td>
				<td class="<% PrintTDMain %>"><%=Body%></td>
				<td class="<% PrintTDMainSwitch %>"><input type="submit" name="Submit" value="Edit"> 
				<input type="button" value="Delete" onClick="DeleteBox('If you delete this review, there is no way to get it back.  Are you sure?', 'admin_reviews_modify.asp?Source=<%=strSource%>&Submit=Delete&ID=<%=ID%>&TargetID=<%=intTargetID%>&TargetTable=<%=strTable%>')"></td>
			</form>
			</tr>
<%
'-----------------------Begin Code----------------------------
			rsPage.MoveNext
		end if
	next
	Response.Write("</table>")
	rsPage.Close
	set rsPage = Nothing
end if

'------------------------End Code-----------------------------
%>