<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeNewsletter ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=NewsletterTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if

'Open up the item
Query = "SELECT ID, Date, MemberID, Subject, Body, FileName FROM Newsletters WHERE ID = " & intID & " AND CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The newsletter does not exist.  If you pasted a link, there may be a typo, or the newsletter may have been deleted.  Please refer to the newsletter list to find the desired newsletter, if it still exists."))
end if

IncrementHits intID, "Newsletters"
%>
<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
	<table align=<%=HeadingAlignment%>>
	<tr>
	<td align=right width="50%" class="LinkText"><a href="members_newsletter_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
	<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
	<a href="javascript:DeleteBox('If you delete this newsletter, there is no way to get it back.  Are you sure?', 'members_newsletter_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a>
	</td>
	</tr>
	</table>
<%
end if
'------------------------End Code-----------------------------
%>
</p>
<% PrintTableHeader 100 %>
<tr>
	<td colspan="2" class="<% PrintTDMain %>">
	<table width=100% cellspacing=0 cellpadding=0>
	<tr>
		<td class="<% PrintTDMain %>" align="left">Author: <%=PrintTDLink(GetNickNameLink(rsItem("MemberID")))%></td>
		<td class="<% PrintTDMainSwitch %>" align="right">Date Written: <%=FormatDateTime(rsItem("Date"), 2)%></td>
	</tr>
	</table>

	</td>
</tr>
<tr>
	<td class="<% PrintTDMainSwitch %>" align="left" colspan="2">Subject: <%=rsItem("Subject")%></td>
</tr>
</table>
<br>
<%
	'If there is a file, we will include it heres
	if rsItem("FileName") <> "" then
		strFileName = rsItem("FileName")
		strFullPath = GetPath("posts") & strFileName
		strLink = NonSecurePath & "posts/" & strFileName

		Set FileSystem = CreateObject("Scripting.FileSystemObject")

		if FileSystem.FileExists(strFullPath) then
%>
		<p><a href="<%=strLink%>">Click here to download and view the newsletter.</a></p>

<%
		end if
		Set FileSystem = Nothing
	end if


%>

<%=rsItem("Body")%>

<br>
<br>
<%
'-----------------------Begin Code----------------------------
rsItem.Close
set rsItem = Nothing
'------------------------End Code-----------------------------
%>