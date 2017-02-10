<%
'
'-----------------------Begin Code----------------------------
if not LoggedMember then Redirect("members.asp?Source=members_petpeeves_modify.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Modify Your Pet Peeves</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
'update info
if Request("Submit") = "Update" then
	if Request("ID") = "" or Request("Date") = "" or Request("Subject") = "" then Redirect("incomplete.asp")
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM PetPeeves WHERE ID = " & Request("ID")
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
	rsUpdate("Body") = Format( Request("Body") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The pet peeve has been edited. &nbsp;<a href="members_petpeeves_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif Request("Submit") = "Delete" then
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	Query = "SELECT * FROM PetPeeves WHERE ID = " & Request("ID")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	'Let's be sure they aren't some asshole who is trying to fuck shit up
	if not rsUpdate("MemberID") = Session("MemberID") then Redirect("error.asp")

	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE Reviews WHERE TargetTable = 'PetPeeves' AND TargetID = " & Request("ID")
	Connect.Execute Query, , adCmdText + adExecuteNoRecords
'------------------------End Code-----------------------------
%>
	<p>The pet peeve has been deleted. &nbsp;<a href="members_petpeeves_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif Request("Submit") = "Edit" then
	Query = "SELECT * FROM PetPeeves WHERE ID = " & Request("ID")
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
	<form method="post" action="members_petpeeves_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
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
      	<td class="<% PrintTDMain %>" align="right">* Pet Peeve</td>
      	<td class="<% PrintTDMain %>"> 
    		<textarea name="Subject" cols="55" rows="4" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Subject") )%></textarea>
     	</td>
    </tr>
	<tr> 
    	<td class="<% PrintTDMain %>" align="right" valign="top">Details (story behind it, whatever)</td>
    	<td class="<% PrintTDMain %>"> 
    		<textarea name="Body" cols="55" rows="10" wrap="PHYSICAL"><%=FormatEdit( rsEdit("Body") )%></textarea>
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
	<form METHOD="POST" ACTION="members_petpeeves_modify.asp">
		View Pet Peeves In The Last <% PrintDaysOld %>
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
		Query = "SELECT * FROM PetPeeves WHERE (MemberID = " & Session("MemberID") & ") ORDER BY Date DESC"
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
				Try again, or <a href="members_petpeeves_modify.asp">click here</a> to view all pet peeves.</p>
<%
'-----------------------Begin Code----------------------------
		else
'-----------------------End Code----------------------------
%>
				<p>Sorry, but your search came up empty.<br>
				Try again, or <a href="members_petpeeves_modify.asp">click here</a> to view all pet peeves.</p>
<%
'-----------------------Begin Code----------------------------
			end if
		else
			'They have search results, so lets list their results
			Query = "SELECT * FROM SectionSearch WHERE SearchID = " & intSearchID & " ORDER BY Score DESC"
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
%>
			<form METHOD="POST" ACTION="members_petpeeves_modify.asp">
			<input type="hidden" name="SearchID" value="<%=intSearchID%>">
<%
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle
			'Instantiate the recordset for the output
			Set rsList = Server.CreateObject("ADODB.Recordset")
			for p = 1 to rsPage.PageSize
				if not rsPage.EOF then
					Query = "SELECT * FROM PetPeeves WHERE ID = " & rsPage("TargetID")
					rsList.Open Query, Connect, adOpenStatic, adLockReadOnly
'------------------------End Code-----------------------------
%>
					<form METHOD="POST" ACTION="members_petpeeves_modify.asp">
					<input type="hidden" name="ID" value="<%=rsList("ID")%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(rsList("Date")) %><a href="petpeeves_read.asp?ID=<%=rsList("ID")%>">Read</a></td>
						<td class="<% PrintTDMain %>"><%=FormatDateTime(rsList("Date"), 2)%></td>
						<td class="<% PrintTDMain %>"><%=rsList("Subject")%></td>
						<td class="<% PrintTDMain %>"><%=PrintPublic(rsList("Private"))%></td>
						<td class="<% PrintTDMain %>"><input type="submit" name="Submit" value="Edit"></td>
						<td class="<% PrintTDMainSwitch %>"><input type="button" value="Delete" onClick="DeleteBox('If you delete this pet peeve, there is no way to get it back.  Are you sure?', 'members_petpeeves_modify.asp?Submit=Delete&ID=<%=rsList("ID")%>')"></td>
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
	'They are just cycling through the pet peeves.  No searching.
	else
		'This is if they requested pet peeves written in a time period
		if Request("DaysOld") <> "" then
			CutoffDate = DateAdd("d", (-1*Request("DaysOld") ), Date)
			Query = "SELECT * FROM PetPeeves WHERE (MemberID = " & Session("MemberID") & " AND Date >= '" & CutoffDate & "') ORDER BY Date DESC"
		else
			Query = "SELECT * FROM PetPeeves WHERE (MemberID = " & Session("MemberID") & ") ORDER BY Date DESC"
		end if
		Set rsPage = Server.CreateObject("ADODB.Recordset")
		rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
		if not rsPage.EOF then
%>
			<form METHOD="POST" ACTION="members_petpeeves_modify.asp">
			<input type="hidden" name="DaysOld" value="<%=Request("DaysOld")%>">
<%
			PrintPagesHeader
			PrintTableHeader 0
			PrintTableTitle
				for j = 1 to rsPage.PageSize
					if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
						<form METHOD="POST" ACTION="members_petpeeves_modify.asp">
						<input type="hidden" name="ID" value="<%=rsPage("ID")%>">
						<tr>
							<td class="<% PrintTDMain %>" align="center"><% PrintNew(rsPage("Date")) %><a href="petpeeves_read.asp?ID=<%=rsPage("ID")%>">Read</a></td>
							<td class="<% PrintTDMain %>"><%=FormatDateTime(rsPage("Date"), 2)%></td>
							<td class="<% PrintTDMain %>"><%=rsPage("Subject")%></td>
							<td class="<% PrintTDMain %>"><%=PrintPublic(rsPage("Private"))%></td>
							<td class="<% PrintTDMainSwitch %>"><input type="submit" name="Submit" value="Edit"></td>
							<td class="<% PrintTDMainSwitch %>"><input type="button" value="Delete" onClick="DeleteBox('If you delete this pet peeve, there is no way to get it back.  Are you sure?', 'members_petpeeves_modify.asp?Submit=Delete&ID=<%=rsPage("ID")%>')"></td>
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
				<p>Sorry, but there have been no pet peeves added in that time period. <a href="javascript:history.back(1)">Click here</a> to go back</p>
<%
'-----------------------Begin Code----------------------------
			else
'------------------------End Code-----------------------------
%>
				<p>You have to add a pet peeve before you can modify it, <%=GetNickNameSession()%>.</p>
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
		<td class="TDHeader">Pet Peeve</td>
		<td class="TDHeader">Public?</td>
		<td class="TDHeader">&nbsp;</td>
		<td class="TDHeader">&nbsp;</td>
	</tr>
<%
End Sub
'------------------------End Code-----------------------------
%>