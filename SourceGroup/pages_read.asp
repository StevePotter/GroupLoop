<%
'
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if

'Open up the item
Query = "SELECT Title, Body, DisplayTitle, MemberID, Privacy FROM InfoPages WHERE ID = " & intID & " AND CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The page does not exist.  If you pasted a link, there may be a typo, or the page may have been deleted."))
end if


if rsItem("Privacy") = 1 AND not LoggedMember() then
	set rsItem = Nothing
	Redirect( "login.asp?Source=pages_read.asp&ID=" & intID & "&Submit=Read" )
end if

if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
	<table align=<%=HeadingAlignment%>>
	<tr>
	<td align=right width="50%" class="LinkText"><a href="members_pages_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
	<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
	<a href="javascript:DeleteBox('If you delete this page, there is no way to get it back.  Are you sure?', 'members_pages_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a>
	</td>
	</tr>
	</table>
<%
end if

if rsItem("DisplayTitle") = 1 then
%>

	<p class="Heading" align="<%=HeadingAlignment%>"><%=rsItem("Title")%></p>

<%
end if
%>

<%=rsItem("Body")%>

<%
rsItem.Close
set rsItem = Nothing

IncrementHits intID, "InfoPages"
'------------------------End Code-----------------------------
%>