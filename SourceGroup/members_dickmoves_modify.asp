<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember then Redirect("members.asp?Source=members_dickmoves_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Modify Dick Moves</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
'-------------------------------------------------------------
'This function writes a pulldown menu for members
'-------------------------------------------------------------
Sub PrintMemberPullDown( intMemberID )
	intMemberID = CInt(intMemberID)
	'Now we are going to get the group names to list in the pull-down menu
	Query = "SELECT ID, NickName FROM Members WHERE (CustomerID = " & CustomerID & ")"
	Set rsTempMembers = Server.CreateObject("ADODB.Recordset")
	rsTempMembers.Open Query, Connect, adOpenStatic, adLockReadOnly
	
	'Make the size 3 if there are many members
	if rsTempMembers.RecordCount <= 30 then
		%><select name="MemberID" size="1"><%
	else
		%><select name="MemberID" size="3"><%
	end if

	'We have passed a 0, which is non member
	if intMemberID = 0 then Response.Write "<option value = '0' SELECTED>Non-Member</option>" & vbCrlf

	do While not rsTempMembers.EOF
		'Highlight the current section
		if rsTempMembers("ID") = intMemberID then
			Response.Write "<option value = '" & rsTempMembers("ID") & "' SELECTED>" & rsTempMembers("NickName") & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & rsTempMembers("ID") & "'>" & rsTempMembers("NickName") & "</option>" & vbCrlf
		end if
		rsTempMembers.MoveNext
	loop
	rsTempMembers.Close
	set rsTempMembers = Nothing
	Response.Write("</select>")
End Sub

'update info
if Request("Submit") = "Update" then
	if Request("ID") = "" or Request("Date") = "" or Request("Subject") = "" or Request("Body") = "" then Redirect("incomplete.asp")
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM MemberStories WHERE ID = " & Request("ID")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Let's be sure they aren't some asshole who is trying to fuck shit up
	if rsUpdate.EOF or not rsUpdate("MemberID") = Session("MemberID") then Redirect("error.asp")

	if Request("Private") = "1" then 
		intPrivate = 1
	else
		intPrivate = 0
	end if
	rsUpdate("Private") = intPrivate
	rsUpdate("Subject") = Format( Request("Subject") )
	rsUpdate("Date") = Request("Date")
	rsUpdate("TargetID") = Request("MemberID")
	rsUpdate("Points") = Request("Points")
	rsUpdate("Body") = Format( Request("Body") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The dick move has been edited. &nbsp;<a href="members_dickmoves_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif Request("Submit") = "Delete" then
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM MemberStories WHERE ID = " & Request("ID")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Let's be sure they aren't some asshole who is trying to fuck shit up
	if not rsUpdate("MemberID") = Session("MemberID") then Redirect("error.asp")

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'MemberStories' AND TargetID = " & Request("ID")
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The dick move has been deleted. &nbsp;<a href="members_dickmoves_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif Request("Submit") = "Edit" then
	Query = "SELECT * FROM MemberStories WHERE ID = " & Request("ID")
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockReadOnly

	if not rsEdit("MemberID") = Session("MemberID") then Redirect("error.asp")

	if rsEdit("Private") = 1 then 
		strChecked = "checked"
	else
		strChecked = ""
	end if
'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To List</a></p>

	* indicates required information<br>
	<form method="post" action="members_dickmoves_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="ID" value="<%=rsEdit("ID")%>">
	<%PrintTableHeader 0%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
	<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Dick Points</td>
      		<td class="<% PrintTDMain %>"> 
				<select name="Points">
<%
				for i = 1 to 10
					strSelected = ""
					if i = rsEdit("Points") then strSelected = " selected"
					%><option value="<%=i%>" <%=strSelected%>><%=i%></option><%
				next
%>
				</select>
     		</td>
		</tr>
		<tr> 
      		<td class="<% PrintTDMain %>" align="right">* Dick</td>
      		<td class="<% PrintTDMain %>"> 
     			<% PrintMemberPullDown rsEdit("TargetID") %>
   		</td>
	</tr>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Subject</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Subject" size="55" value="<%=FormatEdit( rsEdit("Subject") )%>">
     	</td>
    </tr>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">* Story</td>
    	<td class="<% PrintTDMain %>"> 
    		<textarea name="Body" cols="55" rows="20" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Body") )%></textarea>
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
	<form METHOD="POST" ACTION="members_dickmoves_modify.asp">
		View Dick Moves In The Last <% PrintDaysOld %>
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
		Set rsList = Server.CreateObject("ADODB.Recordset")
		Query = "SELECT * FROM MemberStories WHERE (MemberID = " & Session("MemberID") & ") ORDER BY Date DESC"
		rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
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
				Try again, or <a href="members_dickmoves_modify.asp">click here</a> to view all dick moves.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="members_dickmoves_modify.asp">click here</a> to view all dick moves.</p>
<%
'-----------------------Begin Code----------------------------
			end if
		else
			'They have search results, so lets list their results
			Query = "SELECT * FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
%>
			<form METHOD="POST" ACTION="members_dickmoves_modify.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle
			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			for p = 1 to rsPage.PageSize
				if not rsPage.EOF then
					Query = "SELECT * FROM MemberStories WHERE ID = " & rsPage("TargetID")
					rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
'------------------------End Code-----------------------------
%>
					<form METHOD="POST" ACTION="members_dickmoves_modify.asp">
					<input type="hidden" name="ID" value="<%=rsList("ID")%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(rsList("Date")) %><a href="dickmoves_read.asp?ID=<%=rsList("ID")%>">Read</a></td>
						<td class="<% PrintTDMain %>"><%=FormatDateTime(rsList("Date"), 2)%></td>
						<td class="<% PrintTDMain %>"><%=GetNickNameLink(rsList("TargetID"))%></td>
						<td class="<% PrintTDMain %>" align=center><%=rsList("Points")%></td>

						<td class="<% PrintTDMain %>"><%=rsList("Subject") %></td>
						<td class="<% PrintTDMain %>"><%=PrintPublic(rsList("Private"))%></td>
						<td class="<% PrintTDMain %>"><input type="submit" name="Submit" value="Edit"></td>
						<td class="<% PrintTDMainSwitch %>"><input type="button" value="Delete" onClick="DeleteBox('If you delete this dick move, there is no way to get it back.  Are you sure?', 'members_dickmoves_modify.asp?Submit=Delete&ID=<%=rsList("ID")%>')"></td>
					</tr>
					</form>
<%
'-----------------------Begin Code----------------------------
					rsList.Close
					rsPage.MoveNext
				end if
			next
			Response.Write("</table>")
			rsPage.Close
			set rsPage = Nothing
			set rsList = Nothing
		end if
	'They are just cycling through the dickmoves.  No searching.
	else
		'This is if they requested dickmoves written in a time period
		if Request("DaysOld") <> "" then
			CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
			Query = "SELECT * FROM MemberStories WHERE (MemberID = " & Session("MemberID") & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
		else
			Query = "SELECT * FROM MemberStories WHERE (MemberID = " & Session("MemberID") & ") ORDER BY Date DESC"
		end if
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsPage.EOF then
%>
			<form METHOD="POST" ACTION="members_dickmoves_modify.asp">
			<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle
				for j = 1 to rsPage.PageSize
					if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
						<form METHOD="POST" ACTION="members_dickmoves_modify.asp">
						<input type="hidden" name="ID" value="<%=rsPage("ID")%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(rsPage("Date")) %><a href="dickmoves_read.asp?ID=<%=rsPage("ID")%>">Read</a></td>
							<td class="<% PrintTDMain %>"><%=FormatDateTime(rsPage("Date"), 2)%></td>
							<td class="<% PrintTDMain %>"><%=GetNickNameLink(rsPage("TargetID"))%></td>
							<td class="<% PrintTDMain %>" align=center><%=rsPage("Points")%></td>
							<td class="<% PrintTDMain %>"><%=(rsPage("Subject")) %></td>
							<td class="<% PrintTDMain %>"><%=PrintPublic(rsPage("Private"))%></td>
							<td class="<% PrintTDMain %>"><input type="submit" name="Submit" value="Edit"></td>
							<td class="<% PrintTDMainSwitch %>"><input type="button" value="Delete" onClick="DeleteBox('If you delete this dick move, there is no way to get it back.  Are you sure?', 'members_dickmoves_modify.asp?Submit=Delete&ID=<%=rsPage("ID")%>')"></td>
						</tr>
						</form>
<%
'-----------------------Begin Code----------------------------
						rsPage.MoveNext
					end if
				next
				Response.Write("</table>")
		else
			if Request("DaysOld") <> "" then
'------------------------End Code-----------------------------
%>
				<p>Sorry, but there have been no dick moves added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<p>You have to add a story before you can modify it, <%=GetNickNameSession()%>.</p>
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
	GetDesc = UCASE(rsList("Subject") & " " & rsList("Body") & " " & rsList("ID") & " " & rsList("Date") & " " & GetNickName(rsList("MemberID")) )
End Function


'-------------------------------------------------------------
'This prints the top row of the table listing the items
'-------------------------------------------------------------
Sub PrintTableTitle
%>		
	<tr>
		<td class="TDHeader">&nbsp;</td>
		<td class="TDHeader">Date</td>
		<td class="TDHeader">Dick</td>
		<td class="TDHeader">Dick Points</td>
		<td class="TDHeader">Subject</td>
		<td class="TDHeader">Public?</td>
		<td class="TDHeader">&nbsp;</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
End Sub
'------------------------End Code-----------------------------
%>