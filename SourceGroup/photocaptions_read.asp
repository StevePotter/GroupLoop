<!-- #include file="photos_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludePhotoCaptions ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=PhotoCaptionsTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if

'Open up the item

Query = "SELECT ID, Date, Caption, MemberID, Private, TimesRated, TotalRating, PhotoID FROM PhotoCaptions WHERE ID = " & intID & " AND CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly


'Make sure the item is valid
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The photo does not exist.  If you pasted a link, there may be a typo, or the photo may have been deleted.  Please refer to the photo list to find the desired photo, if it still exists."))
end if

Query = "SELECT ID, Name, Ext, CategoryID, MemberID FROM Photos WHERE ID = " & rsItem("PhotoID") & " AND CustomerID = " & CustomerID
Set rsPhoto = Server.CreateObject("ADODB.Recordset")
rsPhoto.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

GetCategoryInfo rsPhoto("CategoryID"), strName, blPrivate, strBody

if rsItem("Private") = 1 AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=photocaptions_read.asp&ID=" & intID & "&Submit=View" )
end if

if Request("Rating") <> "" and RatePhotos = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Caption</a></p>
<%
	AddRating intID, "PhotoCaptions"
else
	IncrementStat "PhotoCaptionsRead"
	IncrementHits intID, "PhotoCaptions"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_photos_modify.asp?Submit=Edit&CaptionID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
		<a href="javascript:DeleteBox('If you delete this caption, there is no way to get it back.  Are you sure?', 'members_photos_modify.asp?Submit=Delete&CaptionID=<%=intID%>')">Delete</a>
		</td>
		</tr>
		</table>
<%
	end if

'------------------------End Code-----------------------------
%>
	</p>
	<p align="center"><a href="photos_view.asp?ID=<%=rsPhoto("ID")%>"><img src="photos/<%=rsPhoto("ID")%>.<%=rsPhoto("Ext")%>" border="0">
		<br><%=rsPhoto("Name")%></a>
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
	<p><%=rsItem("Caption")%></p>	
	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RatePhotoCaptions = 1 then
		PrintRatingPulldown intID, "", "PhotoCaptions", "photocaptions_read.asp", "caption"
	end if

	if ReviewPhotoCaptions = 1 then
%>
		<a href="review.asp?Source=photocaptions_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=PhotoCaptions">Add a review</a><br>
<%
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=photocaptions_read.asp?ID=<%=intID%>&TargetTable=PhotoCaptions&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if

		if ReviewsExist( "PhotoCaptions", rsItem("ID") ) then
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "photocaptions_read.asp", "PhotoCaptions", rsItem("ID")
			Set rsPage = Nothing
		end if
	end if

end if

rsPhoto.Close
set rsPhoto = Nothing

set rsItem = Nothing
'------------------------End Code-----------------------------
%>