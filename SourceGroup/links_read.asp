<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeLinks ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=LinksTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if

Public ListType, DisplayDate, DisplayAuthor, DisplaySubject

Query = "SELECT DisplayDateItemLinks, DisplayAuthorItemLinks, DisplaySubjectItemLinks  FROM Look WHERE CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
	DisplayDate = CBool(rsItem("DisplayDateItemLinks"))
	DisplayAuthor = CBool(rsItem("DisplayAuthorItemLinks"))
	DisplaySubject = CBool(rsItem("DisplaySubjectItemLinks"))
rsItem.Close

'Open up the item
Query = "SELECT ID, Date, MemberID, URL, Name, Description, Private FROM Links WHERE ID = " & intID & " AND CustomerID = " & CustomerID
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly


'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The link does not exist.  If you pasted a link, there may be a typo, or the link may have been deleted.  Please refer to the link list to find the desired link, if it still exists."))
end if

if rsItem("Private") = 1 AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=links_read.asp&ID=" & intID & "&Submit=Read" )
end if

if Request("Rating") <> "" and RateLinks = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Link</a></p>
<%
	AddRating intID, "Links"
else
	IncrementStat "LinksRead"
	IncrementHits intID, "Links"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_links_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;<a href="javascript:DeleteBox('If you delete this link, there is no way to get it back.  Are you sure?', 'members_links_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a></td>
		</tr>
		</table>
<%
	end if

'------------------------End Code-----------------------------
%>
	</p>
<%
	if DisplayDate or DisplayAuthor or DisplaySubject  then
		PrintTableHeader 100
		if DisplayDate or DisplayAuthor then %>

		<tr>
			<td colspan="2" class="<% PrintTDMain %>">
			<table width=100% cellspacing=0 cellpadding=0>
			<tr>
	
			<% if DisplayAuthor then %>
			<td class="<% PrintTDMain %>" align="left">Author: <%=PrintTDLink(GetNickNameLink(rsItem("MemberID")))%></td>
			<% end if %>	
			<% if DisplayDate then
				strAlign = "left"
				if DisplayAuthor then strAlign = "right"
			%>
			<td class="<% PrintTDMainSwitch %>" align="<%=strAlign%>">Date Written: <%=FormatDateTime(rsItem("Date"), 2)%></td>
			<% end if %>		
			</tr>	
			</table>
			</td>
		</tr>
<%
		end if
		Response.Write "</table>"
	end if
%>
	<%
		if rsItem("Name") = "" then %>
			<p><a href="<%=rsItem("URL")%>" target="_blank"><%=rsItem("URL")%></a></p>
	<%	else %>
			<p><a href="<%=rsItem("URL")%>" target="_blank"><%=rsItem("Name")%></a></p>
	<%	end if%>

	<%=rsItem("Description")%>

	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RateLinks = 1 then
		PrintRatingPulldown intID, "", "Links", "links_read.asp", "link"
	end if
	if ReviewLinks = 1 then
%>
		<a href="review.asp?Source=links_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Links">Add A Review</a><br>
<%
		if ReviewsExist( "Links", intID ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=links_read.asp?ID=<%=intID%>&TargetTable=Links&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "links_read.asp", "Links", intID
			Set rsPage = Nothing
		end if
	end if
end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>