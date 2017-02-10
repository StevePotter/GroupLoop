<!-- #include file="photos_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludePhotos ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class=Heading align="<%=HeadingAlignment%>"><%=PhotosTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if

'Open up the item
Query = "SELECT ID, Name, Ext, CategoryID, MemberID FROM Photos WHERE ID = " & intID & " AND CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenForwardOnly, adLockReadOnly

'Make sure the item is valid
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The photo does not exist.  If you pasted a link, there may be a typo, or the photo may have been deleted.  Please refer to the photo list to find the desired photo, if it still exists."))
end if

GetCategoryInfo rsItem("CategoryID"), strName, blPrivate, strBody

if blPrivate AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=photos_view.asp&ID=" & intID & "&Submit=View" )
end if

if Request("Rating") <> "" and RatePhotos = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To Photo</a></p>
<%
	AddRating intID, "Photos"
else
	IncrementStat "PhotosViewed"
	IncrementHits intID, "Photos"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_photos_modify.asp?Submit=Edit&PhotoID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
		<a href="javascript:DeleteBox('If you delete this photo, there is no way to get it back.  Are you sure?', 'members_photos_modify.asp?Submit=Delete&PhotoID=<%=intID%>')">Delete</a>
		</td>
		</tr>
		</table>
<%
	end if

'------------------------End Code-----------------------------
%>
	</p>
	<p align="center"><img src="photos/<%=rsItem("ID")%>.<%=rsItem("Ext")%>" border="0">
		<br><%=rsItem("Name")%>
	</p>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RatePhotos = 1 then
		PrintRatingPulldown intID, "", "Photos", "photos_view.asp", "photo"
	end if

	if IncludePhotoCaptions = 1 then

		Query = "SELECT ID, Caption, MemberID, Private, TimesRated, TotalRating FROM PhotoCaptions WHERE (CustomerID = " & CustomerID & " AND PhotoID = " & intID & ") ORDER BY Date DESC"
		Set rsCaptions = Server.CreateObject("ADODB.Recordset")
		rsCaptions.CacheSize = 20
		rsCaptions.Open Query, Connect, adOpenStatic, adLockReadOnly

		if not rsCaptions.EOF then
			intRateCaptions = RatePhotoCaptions
			intReviewCaptions = ReviewPhotoCaptions

			do until rsCaptions.EOF
				if rsCaptions("Private") = 1 and not LoggedMember then
%>
					<p><%=GetNickNameLink(rsCaptions("MemberID"))%> wrote a private caption.  If you are a member, <a href="login.asp?Source=photos_view.asp&ID=<%=intID%>&Submit=Read">click here</a> to log in and view the caption.</p>
<%
				else
					IncrementStat "PhotoCaptionsRead"
					IncrementHits rsCaptions("ID"), "PhotoCaptions"

					if intRateCaptions = 1 then
						strRating = ""
						if rsCaptions("TimesRated") > 0 then strRating = "Rating: " & GetRating( rsCaptions("TotalRating"), rsCaptions("TimesRated") ) & "&nbsp;&nbsp;"
%>
						<p><%=rsCaptions("Caption")%> - <%=GetNickNameLink(rsCaptions("MemberID"))%> &nbsp;&nbsp;&nbsp;<font size="-2"><%=strRating%></font> 
<%
					else
%>
						<p><%=rsCaptions("Caption")%> - <%=GetNickNameLink(rsCaptions("MemberID"))%> 
<%
					end if
					if intRateCaptions = 1 and intReviewCaptions = 0 then
						%><a href="photocaptions_read.asp?ID=<%=rsCaptions("ID")%>">Rate This Caption</a><%
					elseif intRateCaptions = 0 and intReviewCaptions = 1 then
						if ReviewsExist( "PhotoCaptions", rsCaptions("ID") ) then
							%><a href="photocaptions_read.asp?ID=<%=rsCaptions("ID")%>">Read/Add Reviews</a><%
						else
							%><a href="photocaptions_read.asp?ID=<%=rsCaptions("ID")%>">Add A Review</a><%
						end if
					elseif intRateCaptions = 1 and intReviewCaptions = 1 then
						if ReviewsExist( "PhotoCaptions", rsCaptions("ID") ) then
							%><a href="photocaptions_read.asp?ID=<%=rsCaptions("ID")%>">Rate This Caption and Read/Add Reviews</a><%
						else
							%><a href="photocaptions_read.asp?ID=<%=rsCaptions("ID")%>">Rate/Review This Caption</a><%
						end if
					end if
					Response.Write "</p>"

				end if
				rsCaptions.MoveNext
			loop
			rsCaptions.Close
		end if

		set rsCaptions = Nothing
		
		if IncludeAddButtons = 1 or LoggedMember() then

%>
		<p class="LinkText"><a href="photocaptions_add.asp?ID=<%=rsItem("ID")%>">Add A Caption</p>
<%
		end if

	end if

end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>