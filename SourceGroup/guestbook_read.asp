<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeGuestbook ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=GuestbookTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if

'Open up the item
Query = "SELECT ID, Date, Author, Email, Body FROM Guestbook WHERE ID = " & intID & " AND CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The entry does not exist.  If you pasted a link, there may be a typo, or the entry may have been deleted.  Please refer to the entry list to find the desired entry, if it still exists."))
end if

if Request("Rating") <> "" and RateGuestbook = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Entry</a></p>
<%
	AddRating intID, "Guestbook"
else
	IncrementStat "GuestbookEntriesRead"
	IncrementHits intID, "Guestbook"

	if InStr( rsItem("Email"), "@" ) then
		strAuthor = "<a href='mailto:" & rsItem("Email") & "'>" & PrintTDLink(rsItem("Author")) & "</a>"
	else
		strAuthor = rsItem("Author")
	end if

%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	if LoggedAdmin then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="admin_guestbook_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;<a href="javascript:DeleteBox('If you delete this entry, there is no way to get it back.  Are you sure?', 'admin_guestbook_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a></td>
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
		<td class="<% PrintTDMain %>" align="left">Author: <%=strAuthor%></td>
		<td class="<% PrintTDMainSwitch %>" align="right">Date Written: <%=FormatDateTime(rsItem("Date"), 2)%></td>
		</tr>
		</table>

		</td>
	</tr>
	</table>
	<br>
	<%=rsItem("Body")%>

	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RateGuestbook = 1 then
		PrintRatingPulldown intID, "", "Guestbook", "guestbook_read.asp", "entry"
	end if
	if ReviewGuestbook = 1 then
%>
		<a href="review.asp?Source=guestbook_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Guestbook">Add A Review</a><br>
<%
		if ReviewsExist( "Guestbook", intID ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=guestbook_read.asp?ID=<%=intID%>&TargetTable=Guestbook&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "guestbook_read.asp", "Guestbook", intID
			Set rsPage = Nothing
		end if
	end if
end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>