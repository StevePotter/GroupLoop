<p class="Heading" align="<%=HeadingAlignment%>">Dick Moves</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back</a></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
intID =  Request("ID")

if intID = "" then Redirect("error.asp")

'Open up the item
Query = "SELECT * FROM MemberStories WHERE ID = " & intID
Set rsStory = Server.CreateObject("ADODB.Recordset")
rsStory.Open Query, Connect, adOpenStatic, adLockReadOnly


'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsStory.EOF OR rsStory("CustomerID") <> CustomerID then Redirect("error.asp")


if rsStory("Private") = 1 AND not LoggedMember then Redirect( "login.asp?Source=dickmoves_read.asp&ID=" & intID & "&Submit=Read" )


if Request("Rating") <> "" then
	AddRating rsStory("ID"), "MemberStories"
else
	IncrementHits rsStory("ID"), "MemberStories"
'------------------------End Code-----------------------------
%>
	<% PrintTableHeader 100 %>
	<tr>
		<td colspan="2" class="<% PrintTDMain %>">
		<table width=100% cellspacing=0 cellpadding=0>
		<tr>
			<td class="<% PrintTDMain %>" align="left">Author: <%=GetNickNameLink(rsStory("MemberID"))%></td>
			<td class="<% PrintTDMain %>" align="right">Date Written: <%=FormatDateTime(rsStory("Date"), 2)%></td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" align="left">Dick: <%=GetNickNameLink(rsStory("TargetID"))%></td>
			<td class="<% PrintTDMainSwitch %>" align="right">Dick Points: <%=rsStory("Points")%></td>
		</tr>
		</table>

		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMainSwitch %>" align="left" colspan="2">Subject: <%=rsStory("Subject")%></td>
	</tr>
	</table>
	<br>
	<%=rsStory("Body")%>

	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
		PrintRatingPulldown rsStory("ID"), "", "MemberStories", "dickmoves_read.asp", "dick move"
%>
		<a href="review.asp?Source=dickmoves_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=MemberStories">Add a review</a><br>
<%
		if ReviewsExist( "MemberStories", intID ) then
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "dickmoves_read.asp", "MemberStories", intID
			Set rsPage = Nothing
		end if
end if

rsStory.Close
set rsStory = Nothing
'------------------------End Code-----------------------------
%>