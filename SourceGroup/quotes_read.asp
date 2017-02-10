<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeQuotes ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=QuotesTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if

'Open up the item
Query = "SELECT ID, Date, MemberID, Quote, Author, Description, Private FROM Quotes WHERE ID = " & intID & " AND CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly


'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The quote does not exist.  If you pasted a link, there may be a typo, or the quote may have been deleted.  Please refer to the quote list to find the desired quote, if it still exists."))
end if

if rsItem("Private") = 1 AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=quotes_read.asp&ID=" & intID & "&Submit=Read" )
end if

if Request("Rating") <> "" and RateQuotes = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Quote</a></p>
<%
	AddRating intID, "Quotes"
else
	IncrementStat "QuotesRead"
	IncrementHits intID, "Quotes"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_quotes_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;<a href="javascript:DeleteBox('If you delete this quote, there is no way to get it back.  Are you sure?', 'members_quotes_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a></td>
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
	</table>
	<br>
	<p>&quot;<%=rsItem("Quote")%>&quot; - <%=rsItem("Author")%></p>	
	<p><%=rsItem("Description")%></p>
	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RateQuotes = 1 then
		PrintRatingPulldown intID, "", "Quotes", "quotes_read.asp", "quote"
	end if
	if ReviewQuotes = 1 then
%>
		<a href="review.asp?Source=quotes_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Quotes">Add A Review</a><br>
<%
		if ReviewsExist( "Quotes", intID ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=quotes_read.asp?ID=<%=intID%>&TargetTable=Quotes&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "quotes_read.asp", "Quotes", intID
			Set rsPage = Nothing
		end if
	end if
end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>