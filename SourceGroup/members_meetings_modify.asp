<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeMeetings ) or CBool( MeetingsMembers ) ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_meetings_modify.asp")
if not (LoggedAdmin or CBool( MeetingsMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Meetings</span><br>
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

if strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if (blLoggedAdmin and Request("Date") = "") or Request("Subject") = "" then Redirect("incomplete.asp")

	Query = "SELECT Private, Subject, Date, Body, IP, ModifiedID, CommitteeID, FileLinkDirect FROM Meetings WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic

	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("Private") = GetCheckedResult(Request("Private"))
	rsUpdate("FileLinkDirect") = GetCheckedResult(Request("FileLinkDirect"))


	rsUpdate("Subject") = Format( Request("Subject") )
	if blLoggedAdmin then rsUpdate("Date") = Request("Date")
	rsUpdate("Body") = GetTextArea( Request("Body") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")

	if IncludeCommittees = 1 then rsUpdate("CommitteeID") = CInt(Request("CommitteeID"))

	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The meeting has been edited. &nbsp;<a href="members_meetings_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID, FileName FROM Meetings WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	'Delete the file
	if rsUpdate("FileName") <> "" then
		Set FileSystem = CreateObject("Scripting.FileSystemObject")
		strFileName = rsUpdate("FileName")
		strFullPath = GetPath("posts") & strFileName
		if FileSystem.FileExists(strFullPath) then FileSystem.DeleteFile strFullPath, True
		Set FileSystem = Nothing
	end if

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'Meetings' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The meeting has been deleted. &nbsp;<a href="members_meetings_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Public DisplayPrivacy

	Query = "SELECT IncludePrivacyMeetings, DisplaySearchMeetings, DisplayDaysOldMeetings, InfoTextMeetings, ListTypeMeetings, DisplayDateListMeetings, DisplayAuthorListMeetings, DisplayPrivacyListMeetings  FROM Look WHERE CustomerID = " & CustomerID
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect

	'show the privacy if they've included it in the section and chose to list it.  don't display if the site is members only
	DisplayPrivacy = CBool(rsEdit("IncludePrivacyMeetings")) and not cBool(SiteMembersOnly)

	rsEdit.Close

	Query = "SELECT ID, Date, Subject, Body, Private, CommitteeID, FileName, FileLinkDirect FROM Meetings WHERE ID = " & intID & " AND " & strMatch
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsEdit.EOF then
		Set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if rsEdit("Private") = 1 then 
		strChecked = "checked"
	else
		strChecked = ""
	end if
'------------------------End Code-----------------------------
%>
	<script language="JavaScript">
	<!--
		function submit_page(form) {
			//Error message variable
			var strError = "";
<%	if blLoggedAdmin then %>
			if (form.Date.value == "")
				strError += "          You forgot the date. \n";
<%	end if %>
			if (form.Subject.value == "")
				strError += "          You forgot the subject. \n";


			if(strError == "") {
				return true;
			}
			else{
				strError = "Sorry, but you must go back and fix the following errors before you can update this: \n" + strError;
				alert (strError);
				return false;
			}   
		}

	//-->
	</SCRIPT>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_meetings_modify.asp" name="MyForm" onsubmit="<%=GetOnAESubmit()%>if (this.submitted) return false; this.submitted = submit_page(this); return this.submitted">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
<%
	if IncludeCommittees = 1 then
%>
	<tr>
		<td class="<% PrintTDMain %>"  align="right">
			Category
		</td>
		<td class="<% PrintTDMain %>">
<%			PrintCommitteePullDown rsEdit("CommitteeID")	%>
		</td>
	</tr>
<%
		end if
%>
<%
	if DisplayPrivacy then
%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>
<%	end if %>
<%	if blLoggedAdmin then %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
<%	end if %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Subject</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Subject" size="55" value="<%=FormatEdit( rsEdit("Subject") )%>">
     	</td>
    </tr>

<%
	'Delete the file
	if rsEdit("FileName") <> "" then
		Set FileSystem = CreateObject("Scripting.FileSystemObject")
		strFileName = rsEdit("FileName")
		strFullPath = GetPath("posts") & strFileName
		if FileSystem.FileExists(strFullPath) then
%>
		<tr>
			<td class="<% PrintTDMain %>" valign="top" align="right">
				Edit file.  Click the button here to update/overwrite the file.
			</td>
			<td class="<% PrintTDMain %>">
				<input type="button" value="Edit File" onClick="Redirect('members_files_modify.asp?Path=posts&FileName=<%=rsEdit("FileName")%>')" ><br>
				<%	PrintCheckBox rsEdit("FileLinkDirect"), "FileLinkDirect" %> Link directly to this file from the list

			</td>
		</tr>
<%
		end if
		Set FileSystem = Nothing
	end if
%>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">* Details</td>
    	<td class="<% PrintTDMain %>"> 
			<% TextArea "Body", 55, 4, True, ExtractInserts( rsEdit("Body") ) %>
    	</td>
    </tr>
	<tr>
    	<td colspan="2" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Update">
    	</td>
    </tr>
  	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing
else
'------------------------End Code-----------------------------
%>
	<form METHOD="POST" ACTION="members_meetings_modify.asp">
		View Meetings In The Last  <% PrintDaysOld %>
		<br>
		Or Search For <input type="text" name="Keywords" size="25">
		<input type="submit" name="Submit" value="Go"><br>
	</form>
<%
'-----------------------Begin Code----------------------------
	'Get the searchID from the last page.  May be blank.
	intSearchID = Request("SearchID")


	'They entered text to search for, so we are going to get matches and put them into the SectionSearch
	if Request("Keywords") <> "" then
		Query = "SELECT ID, Date, MemberID, Subject, Body FROM Meetings WHERE (" & strMatch & ") ORDER BY Date DESC"
		Set rsList = Server.CreateObject("ADODB.Recordset")
		rsList.CacheSize = 100
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set MemberID = rsList("MemberID")
			Set Body = rsList("Body")
			Set Subject = rsList("Subject")
		intSearchID = SingleSearch()
		Session("SearchID") = intSearchID
		rsList.Close
	end if

	if intSearchID <> "" then
		'Their search came up empty
		if intSearchID = 0 then
			if Session("MemberID") <> "" then
'-----------------------End Code----------------------------
%>
				<p>Sorry <%=GetNickNameSession()%>.  Your search came up empty.<br>
				Try again, or <a href="members_meetings_modify.asp">click here</a> to view all meetings.</p>
<%
'-----------------------Begin Code----------------------------
			else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="members_meetings_modify.asp">click here</a> to view all meetings.</p>
<%
'-----------------------Begin Code----------------------------
			end if
		else
			'They have search results, so lets list their results
			Query = "SELECT TargetID FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect
			rsPage.CacheSize = PageSize
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="members_meetings_modify.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
'-----------------------Begin Code----------------------------
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle

			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			Query = "SELECT ID, Date, MemberID, Subject, Private FROM Meetings WHERE " & strMatch
			rsList.CacheSize = PageSize
			rsList.Open Query, Connect, adOpenStatic, adLockReadOnly, adCmdTableDirect

			Set ID = rsList("ID")
			Set ItemDate = rsList("Date")
			Set MemberID = rsList("MemberID")
			Set Subject = rsList("Subject")
			Set IsPrivate = rsList("Private")

			for p = 1 to rsPage.PageSize
				if not rsPage.EOF then
					rsList.Filter = "ID = " & rsPage("TargetID")

					PrintTableData

					rsPage.MoveNext
				end if
			next
			Response.Write("</table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing
		end if
	'They are just cycling through the Meetings.  No searching.
	else
		'This is if they requested Meetings written in a time period
		if Request("DaysOld") <> "" then
			CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
			Query = "SELECT ID, Date, MemberID, Subject, Private FROM Meetings WHERE (" & strMatch & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
		else
			Query = "SELECT ID, Date, MemberID, Subject, Private FROM Meetings WHERE (" & strMatch & ") ORDER BY Date DESC"
		end if
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.CacheSize = PageSize
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsPage.EOF then
			Set ID = rsPage("ID")
			Set ItemDate = rsPage("Date")
			Set MemberID = rsPage("MemberID")
			Set Subject = rsPage("Subject")
			Set IsPrivate = rsPage("Private")
'-----------------------End Code----------------------------
%>
			<form METHOD="POST" ACTION="members_meetings_modify.asp">
			<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%
'-----------------------Begin Code----------------------------
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle
					for j = 1 to rsPage.PageSize
					if not rsPage.EOF then
						PrintTableData
						rsPage.MoveNext
					end if
				next
				Response.Write("</table>")
		else
			if Request("DaysOld") <> "" then
'------------------------End Code-----------------------------
%>
				<p>Sorry, but there have been no meetings added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<p>Sorry, but there are no meetings at the moment.</p>
<%
'-----------------------Begin Code----------------------------
			end if
		end if
		rsPage.Close
		set rsPage = Nothing
	end if
end if


'-------------------------------------------------------------
'This function returns the search description of an object to match with
'Must have the recordset rsList open
'-------------------------------------------------------------
Function GetDesc
	GetDesc = UCASE(Subject & Body & ItemDate & GetNickName(MemberID) )
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
		<td class="TDHeader">Subject</td>
		<td class="TDHeader">Public?</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
End Sub

'-------------------------------------------------------------
'This prints the data for a row
'-------------------------------------------------------------
Sub PrintTableData
%>
	<form METHOD="POST" ACTION="members_meetings_modify.asp">
	<input type="hidden" name="ID" value="<%=ID%>">
	<tr>
		<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><%=FormatDateTime(ItemDate, 2)%></td>
<%		if blLoggedAdmin then %>
		<td class="<% PrintTDMain %>"><%=PrintTDLink(GetNickNameLink(MemberID))%></td>
<%		end if %>
		<td class="<% PrintTDMain %>"><a href="meetings_read.asp?ID=<%=ID%>"><%=PrintTDLink(Subject)%></a></td>
		<td class="<% PrintTDMain %>"><%=PrintPublic(IsPrivate)%></td>
		<td class="<% PrintTDMainSwitch %>">
			<input type="submit" name="Submit" value="Edit">
			<input type="button" value="Delete" onClick="DeleteBox('If you delete this meeting, there is no way to get it back.  Are you sure?', 'members_meetings_modify.asp?Submit=Delete&ID=<%=ID%>')">			
			<%if ReviewsExist( "Meetings", ID ) AND blLoggedAdmin then%>
				<input type="button" value="Modify Reviews" onClick="Redirect('admin_reviews_modify.asp?Source=members_meetings_modify.asp&TargetTable=Meetings&TargetID=<%=ID%>')">
			<%end if%>	
		</td>
		</tr>
	</form>
<%
End Sub
'------------------------End Code-----------------------------
%>