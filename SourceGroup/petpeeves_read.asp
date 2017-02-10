
<p class="Heading" align="<%=HeadingAlignment%>">Pet Peeves</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back</a></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
intID =  Request("ID")

if intID = "" then Redirect("error.asp")

'Open up the item
Query = "SELECT * FROM PetPeeves WHERE ID = " & intID
Set rsStory = Server.CreateObject("ADODB.Recordset")
rsStory.Open Query, Connect, adOpenStatic, adLockReadOnly


'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsStory.EOF OR rsStory("CustomerID") <> CustomerID then Redirect("error.asp")


if rsStory("Private") = 1 AND not LoggedMember then Redirect( "login.asp?Source=petpeeves_read.asp&ID=" & intID & "&Submit=Read" )


if Request("Rating") <> "" and RateStories = 1 then
	AddRating rsStory("ID"), "PetPeeves"
	%><a href="petpeeves.asp">Click here</a> to go back to the pet peeve list.<%
else
	IncrementHits rsStory("ID"), "PetPeeves"
'------------------------End Code-----------------------------
%>
	<% PrintTableHeader 100 %>
	<tr>
		<td colspan="2" class="<% PrintTDMain %>">
		<table width=100% cellspacing=0 cellpadding=0>
		<tr>
		<td class="<% PrintTDMain %>" align="left">Author: <%=GetNickNameLink(rsStory("MemberID"))%></td>
		<td class="<% PrintTDMainSwitch %>" align="right">Date Written: <%=FormatDateTime(rsStory("Date"), 2)%></td>
		</tr>
		</table>

		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMainSwitch %>" align="left" colspan="2">Pet Peeve: <%=rsStory("Subject")%></td>
	</tr>
	</table>
	<br>
	<%=rsStory("Body")%>

	<br>
<%
'-----------------------Begin Code----------------------------
	PrintRatingPulldown rsStory("ID"), "", "PetPeeves", "petpeeves_read.asp", "pet peeve"
%>
	<a href="review.asp?Source=petpeeves_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=PetPeeves">Add a review</a><br>
<%
	if ReviewsExist( "PetPeeves", intID ) then
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		PrintReviews "petpeeves_read.asp", "PetPeeves", intID
		Set rsPage = Nothing
	end if
end if

rsStory.Close
set rsStory = Nothing
'------------------------End Code-----------------------------
%>