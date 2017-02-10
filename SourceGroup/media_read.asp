<!-- #include file="media_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeMedia ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p class="Heading" align="<%=HeadingAlignment%>"><%=MediaTitle%></p>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
intID = Request("ID")
if intID = "" then Redirect("error.asp")

'Open up the item
Query = "SELECT ID, CategoryID, Date, MemberID, FileName, Description FROM Media WHERE ID = " & intID & " AND CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

'Make sure it is valid
'If the customer ID is wrong, or it is deleted and the person isn't an administrator (admins can read deleted shit), send them away
if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The file does not exist.  If you pasted a link, there may be a typo, or the file may have been deleted.  Please refer to the file list to find the desired file, if it still exists."))
end if

GetCategoryInfo rsItem("CategoryID"), strName, blPrivate

if blPrivate AND not LoggedMember then
	set rsItem = Nothing
	Redirect( "login.asp?Source=media_read.asp&ID=" & intID & "&Submit=View" )
end if


if Request("Rating") <> "" and RateMedia = 1 then
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a><br>
	<a href="javascript:history.go(-1)">Back To File</a></p>
<%
	AddRating rsItem("ID"), "Media"
else
	IncrementStat "PhotosViewed"
	IncrementHits rsItem("ID"), "Media"
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-1)">Back</a>
<%
	if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
		<table align=<%=HeadingAlignment%>>
		<tr>
		<td align=right width="50%" class="LinkText"><a href="members_media_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
		<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
		<a href="javascript:DeleteBox('If you delete this file, there is no way to get it back.  Are you sure?', 'members_media_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a>
		</td>
		</tr>
		</table>
<%
	end if

	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	strPath = GetPath ("media")
	strFileName = strPath & "/" & rsItem("FileName")
	if not FileSystem.FileExists (strFileName) then Redirect("error.asp?Source=media.asp&Message=" & Server.URLEncode("The file does not exist.  Please notify the author so they can reupload/delete this item.") )
	Set TestFile = FileSystem.GetFile( strFileName )
	dblSize = Round((TestFile.Size / 1000000), 2 )
	Set TestFile = Nothing
	Set FileSystem = Nothing

'------------------------End Code-----------------------------
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="media.asp?ID=<%=rsItem("CategoryID")%>">All Files In <%=strName%></a></p>
	<% if IncludeAuthor + IncludeDate > 0 then %>
	<% PrintTableHeader 100 %>
	<tr>
		<td colspan="2" class="<% PrintTDMain %>">
		<table width=100% cellspacing=0 cellpadding=0>
		<tr>
		<% if IncludeAuthor = 1 then %>
		<td class="<% PrintTDMain %>" align="left">Author: <%=PrintTDLink(GetNickNameLink(rsItem("MemberID")))%></td>
		<% end if %>	
		<% if IncludeDate = 1 then %>
		<td class="<% PrintTDMainSwitch %>" align="right">Date Written: <%=FormatDateTime(rsItem("Date"), 2)%></td>
		<% end if %>	
		</tr>
		</table>

		</td>
	</tr>
	</table>
	<% end if %>	

	<p><a href="media/<%=rsItem("FileName")%>"><%=rsItem("FileName")%></a>
		 &nbsp;<font size=-2>(<%=dblSize%> Megs)</font></p>
						
	<%=rsItem("Description")%>


	<br>
	<br>
<%
'-----------------------Begin Code----------------------------
	if RateMedia = 1 then
		PrintRatingPulldown rsItem("ID"), "", "Media", "media_read.asp", "file"
	end if
	if ReviewMedia = 1 then
%>
		<a href="review.asp?Source=media_read.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Media">Add A Review</a><br>
<%
		if ReviewsExist( "Media", rsItem("ID") ) then
			if LoggedAdmin then
%>
				<a href="admin_reviews_modify.asp?Source=media_read.asp?ID=<%=intID%>&TargetTable=Media&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
			end if
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "media_read.asp", "Media", rsItem("ID")
			Set rsPage = Nothing
		end if
	end if
end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>