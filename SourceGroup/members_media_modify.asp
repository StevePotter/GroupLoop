<!-- #include file="media_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not CatsExist then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but no files can be modified until the administrator creates a category."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not CBool( IncludeMedia ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember then Redirect("members.asp?Source=members_media_modify.asp")
if not (LoggedAdmin or CBool( MediaMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Media</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
blLoggedAdmin = LoggedAdmin

if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

strSubmit = Request("Submit")

if strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT FileName FROM Media WHERE ID = " & intID & " AND " & strMatch
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	strPath = GetPath ("media")
	strFileName = strPath & "/" & rsUpdate("FileName")

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'Media' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	'Delete the file
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	if FileSystem.FileExists(strFileName) then FileSystem.DeleteFile(strFileName)
	Set FileSystem = Nothing

'------------------------End Code-----------------------------
%>
	<p>The file has been deleted. &nbsp;<a href="members_media_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT Date, FileName, Description, CategoryID FROM Media WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

'------------------------End Code-----------------------------
%>
<script language="JavaScript">
<!--
	function submit_page(form) {
		//Error message variable
		var strError = "";
<%	if blLoggedAdmin then	%>
		if (form.Date.value == ""){
			strError += "Sorry, but you forgot to enter a date. \n";
			alert (strError);
			return false;
		}
<%	end if	%>
		if(form.MediaFile.value == "") {
			return true;
		}
		else{
			alert ('Uploading your file may take some time, so please be patient and dont constantly click the Update button, because that wont speed anything up.');
			return true;
		}
	}

//-->
</SCRIPT>

	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	* indicates required information<br>
	<form enctype="multipart/form-data" method="post" action="<%=SecurePath%>members_media_modify_process.asp" onsubmit="if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted" name="MyForm">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="MediaID" value="<%=intID%>">
	<%PrintTableHeader 0%>
<%	if blLoggedAdmin then %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
<%	end if %>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">* Category</td>
		<td class="<% PrintTDMain %>"> 
			<% PrintCategoryPulldown rsEdit("CategoryID"), 0, 1 %>
     	</td>
   	</tr>

	<tr> 
      	<td class="<% PrintTDMain %>" align="right">Description of File</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Name" size="55" value="<%=FormatEdit( rsEdit("Description") )%>">
     	</td>
    </tr>
	<tr>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			File.  If you leave this blank, the original file will be kept.  Otherwise, the new file will erase the old one.
		</td>
		<td class="<% PrintTDMain %>">
			<input type="file" name="MediaFile">
		</td>
	</tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update"></td>
    	</td>
    </tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing
else
	intCategoryID = Request("ID")
	if intCategoryID <> "" then intCategoryID = CInt(intCategoryID)

	'Get the searchID from the last page.  May be blank.
	intSearchID = Request("SearchID")
	if intSearchID <> "" then intSearchID = CInt(intSearchID)

	strPath = GetPath("media")

	'Start them off in the first category if none is specified
	if intCategoryID = "" then

		'They entered text to search for, so we are going to get matches and put them into the SectionSearch
		if Request("Keywords") <> "" AND intSearchID = "" then
			Query = "SELECT ID, Date, FileName, Description, MemberID FROM Media WHERE " & strMatch & " ORDER BY Date DESC"
			Set rsList = Server.CreateObject("ADODB.Recordset")
			rsList.CacheSize = 100
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
				Set ID = rsList("ID")
				Set ItemDate = rsList("Date")
				Set FileName = rsList("FileName")
				Set Description = rsList("Description")
				Set MemberID = rsList("MemberID")
			intSearchID = SingleSearch()
			rsList.Close
			set rsList = Nothing
		end if

		if intSearchID = "" then
	%>
			<form METHOD="POST" ACTION="members_media_modify.asp">
				Search For <input type="text" name="Keywords" size="15">
				<input type="submit" name="Submit" value="Go"><br>
			</form>
	<%
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintCategoryMenu "members_media_modify.asp"
			Set rsPage = Nothing
		else
			'Their search came up empty
			if intSearchID = 0 then
				if Session("MemberID") <> "" then
	'-----------------------End Code----------------------------
	%>
					<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
					Try again, or <a href="members_media_modify.asp">click here</a> to go back to the category list.</p>
	<%
	'-----------------------Begin Code----------------------------
				else
	'-----------------------End Code----------------------------
	%>
					<p>Sorry, but your search came up empty.<br>
					Try again, or <a href="members_media_modify.asp">click here</a> to go back to the category list.</p>
	<%
	'-----------------------Begin Code----------------------------
				end if
			else
				'They have search results, so lets list their results
				Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
				Set rsPage = Server.CreateObject("ADODB.Recordset")
				rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
				rsPage.CacheSize = PageSize
				Set TargetID = rsPage("TargetID")
		%>
				<form METHOD="POST" ACTION="members_media_modify.asp">
				<input type="hidden" name="SearchID" value="<%=intSearchID%>">
		<%
				PrintPagesHeader
				PrintTableHeader 0
				PrintTableTitle

				'Instantiate the recordset for the output
				Set rsList = Server.CreateObject("ADODB.Recordset")
				Query = "SELECT ID, Date, MemberID, FileName, Description, TotalRating, TimesRated FROM Media WHERE " & strMatch
				rsList.CacheSize = PageSize
				rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

				Set ID = rsList("ID")
				Set ItemDate = rsList("Date")
				Set MemberID = rsList("MemberID")
				Set TotalRating = rsList("TotalRating")
				Set TimesRated = rsList("TimesRated")
				Set FileName = rsList("FileName")
				Set Description = rsList("Description")

				Set FileSystem = CreateObject("Scripting.FileSystemObject")

				for p = 1 to rsPage.PageSize
					if not rsPage.EOF then
						rsList.Filter = "ID = " & TargetID

						PrintTableData

						rsPage.MoveNext
					end if
				next
				Response.Write("</table>")
				rsPage.Close
				set rsPage = Nothing
				set rsList = Nothing

				Set FileSystem = Nothing
			end if
		end if

	'They have a category selected
	else
		if not ValidCategory(intCategoryID) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but that is not a valid category."))

		GetCategoryInfo intCategoryID, strName, blPrivate

		if blPrivate AND not LoggedMember then Redirect( "login.asp?Source=members_media_modify.asp&ID=" & intCategoryID & "&Submit=Go" )

		'Keep track of shit
		IncrementHits intCategoryID, "MediaCategories"

		'------------------------End Code-----------------------------
		%>

		<form METHOD="POST" ACTION="members_media_modify.asp">
			<input type="hidden" name="ID" value="<%=intCategoryID%>">
			Search <%=strName%> For <input type="text" name="Keywords" size="15">
			<input type="submit" name="Submit" value="Go"><br>
		</form>

		<table width="100%">
			<tr>
				<td align="left">
					<span class="Heading">Category: <%=strName%></span>
				</td>

	<%			if NeedCategoryMenu("MediaCategories") then	%>
				<td align="right">
					<font size="-1">Change Category To:</font><br>
					<form action="members_media_modify.asp" method="post">
						<% PrintCategoryPullDown intCategoryID, 0, 0 %>
						<input type="Submit" value="Switch">
					</form>
				</td>
	<%			end if %>
			</tr>
		</table>

	<%
	'-----------------------Begin Code----------------------------
		'They entered text to search for, so we are going to get matches and put them into the SectionSearch
		if Request("Keywords") <> "" AND intSearchID = "" then
			Query = "SELECT ID, Date, FileName, Description, MemberID FROM Media WHERE (CategoryID = " & intCategoryID & " AND " & strMatch & ") ORDER BY Date DESC"
			Set rsList = Server.CreateObject("ADODB.Recordset")
			rsList.CacheSize = 100
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
				Set ID = rsList("ID")
				Set ItemDate = rsList("Date")
				Set FileName = rsList("FileName")
				Set Description = rsList("Description")
				Set MemberID = rsList("MemberID")
			intSearchID = SingleSearch()
			rsList.Close
			set rsList = Nothing
		end if

		if intSearchID <> "" then
			'Their search came up empty
			if intSearchID = 0 then
				if Session("MemberID") <> "" then
		'-----------------------End Code----------------------------
		%>
					<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
					Try again, or <a href="members_media_modify.asp?ID=<%=intCategoryID%>">click here</a> to view all files in this category.</p>
		<%
		'-----------------------Begin Code----------------------------
				else
		'-----------------------End Code----------------------------
		%>
					<p>Sorry, but your search came up empty.<br>
					Try again, or <a href="members_media_modify.asp?ID=<%=intCategoryID%>">click here</a> to view all files in this category.</p>
		<%
		'-----------------------Begin Code----------------------------
				end if
			else
				'They have search results, so lets list their results
				Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
				Set rsPage = Server.CreateObject("ADODB.Recordset")
				rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
				rsPage.CacheSize = PageSize
				Set TargetID = rsPage("TargetID")
		%>
				<form METHOD="POST" ACTION="members_media_modify.asp">
				<input type="hidden" name="SearchID" value="<%=intSearchID%>">
				<input type="hidden" name="ID" value="<%=intCategoryID%>">
		<%
				PrintPagesHeader
				PrintTableHeader 0
				PrintTableTitle

				'Instantiate the recordset for the output
				Set rsList = Server.CreateObject("ADODB.Recordset")
				Query = "SELECT ID, Date, MemberID, FileName, Description, TotalRating, TimesRated FROM Media WHERE CategoryID = " & intCategoryID & " AND " & strMatch
				rsList.CacheSize = PageSize
				rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

				Set ID = rsList("ID")
				Set ItemDate = rsList("Date")
				Set MemberID = rsList("MemberID")
				Set TotalRating = rsList("TotalRating")
				Set TimesRated = rsList("TimesRated")
				Set FileName = rsList("FileName")
				Set Description = rsList("Description")

				Set FileSystem = CreateObject("Scripting.FileSystemObject")

				for p = 1 to rsPage.PageSize
					if not rsPage.EOF then
						rsList.Filter = "ID = " & TargetID

						PrintTableData

						rsPage.MoveNext
					end if
				next
				Response.Write("</table>")
				rsPage.Close
				set rsPage = Nothing
				set rsList = Nothing

				Set FileSystem = Nothing
			end if

		'They are just cycling through the Media.  No searching.
		else
			'Instantiate the recordset for the output
			Query = "SELECT ID, Date, MemberID, FileName, Description, TotalRating, TimesRated FROM Media WHERE CategoryID = " & intCategoryID & " AND " & strMatch & " ORDER BY Date DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.CacheSize = PageSize
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			'Don't navigate if it's empty
			if not rsPage.EOF then
	%>
				<form METHOD="POST" ACTION="members_media_modify.asp">
				<input type="hidden" name="ID" value="<%=intCategoryID%>">
	<%
				Set ID = rsPage("ID")
				Set ItemDate = rsPage("Date")
				Set MemberID = rsPage("MemberID")
				Set TotalRating = rsPage("TotalRating")
				Set TimesRated = rsPage("TimesRated")
				Set FileName = rsPage("FileName")
				Set Description = rsPage("Description")

				PrintPagesHeader
				PrintTableHeader 100
				PrintTableTitle

				strPath = GetPath ("media")
				Set FileSystem = CreateObject("Scripting.FileSystemObject")

				for j = 1 to rsPage.PageSize
					if not rsPage.EOF then
						PrintTableData

						rsPage.MoveNext
					end if
				next

				Response.Write("</table>")

				Set FileSystem = Nothing
				set rsPage = Nothing
			else
				'If there are no available Media
				Response.Write "<p>Sorry, but there are no files in this category.</p>"
			end if

			set rsPage = Nothing
		end if

	end if

end if


'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(FileName & Description & GetNickName(MemberID) )
End Function


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader">Date</td>
<%		if blLoggedAdmin then %>
		<td class="TDHeader">Author</td>
<%		end if %>
		<td class="TDHeader">File</td>
		<td class="TDHeader">Description</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
End Sub

'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
	strFileName = strPath & "/" & FileName
	if FileSystem.FileExists (strFileName) then
		Set TestFile = FileSystem.GetFile( strFileName )
		dblSize = Round((TestFile.Size / 1000000), 2 )
		strLink = "<a href='media/" & FileName & "'>" & FileName & "</a> &nbsp;<font size=-2>(" & dblSize & " Megs)</font>"
		Set TestFile = Nothing
	else
		strLink = "File Does Not Exist"
	end if
%>
	<form METHOD="POST" ACTION="members_media_modify.asp">
	<input type="hidden" name="ID" value="<%=ID%>">
	<tr>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
<%		if blLoggedAdmin then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
<%		end if %>
		<td class="<% PrintTDMain %>"><%=strLink%></td>
		<td class="<% PrintTDMain %>"><a href="media_read.asp?ID=<%=ID%>"><%=PrintTDLink(PrintStart(Description))%></a></td>
		<td class="<% PrintTDMainSwitch %>">
			<input type="submit" name="Submit" value="Edit">
			<input type="button" value="Delete" onClick="DeleteBox('If you delete this file, there is no way to get it back.  Are you sure?', 'members_media_modify.asp?Submit=Delete&ID=<%=ID%>')">			
			<%if ReviewsExist( "Media", ID ) AND blLoggedAdmin then%>
				<input type="button" value="Modify Reviews" onClick="Redirect('admin_reviews_modify.asp?Source=members_media_modify.asp&TargetTable=Media&TargetID=<%=ID%>')">
			<%end if%>	
		</td>
	</tr>
	</form>
<%
End Sub
'------------------------End Code-----------------------------
%>