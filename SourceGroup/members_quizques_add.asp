<!-- #include file="quizzes_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeQuizzes ) or CBool( QuizzesMembers ) ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_quizques_add.asp")
if not (LoggedAdmin or CBool( QuizzesMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add New Questions To A Quiz</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
if LoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
intID = CInt(Request("ID"))

if not ValidQuiz(intID) then Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))

if Request("Submit") = "Add" then
	if Request("Question") = "" or Request("A") = "" or Request("Answer") = "" then Redirect("incomplete.asp")

	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddQuizQuestion"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@ModifiedID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@QuizID", adInteger, adParamInput )
		.Parameters.Append .CreateParameter ("@IP", adVarWChar, adParamInput, 20 )
		.Parameters.Append .CreateParameter ("@Question", adVarWChar, adParamInput, 400 )
		.Parameters.Append .CreateParameter ("@A", adVarWChar, adParamInput, 400 )
		.Parameters.Append .CreateParameter ("@B", adVarWChar, adParamInput, 400 )
		.Parameters.Append .CreateParameter ("@C", adVarWChar, adParamInput, 400 )
		.Parameters.Append .CreateParameter ("@D", adVarWChar, adParamInput, 400 )
		.Parameters.Append .CreateParameter ("@E", adVarWChar, adParamInput, 400 )
		.Parameters.Append .CreateParameter ("@F", adVarWChar, adParamInput, 400 )
		.Parameters.Append .CreateParameter ("@Answer", adVarWChar, adParamInput, 1 )

		.Parameters("@MemberID") = Session("MemberID")
		.Parameters("@ModifiedID") = Session("MemberID")
		.Parameters("@CustomerID") = CustomerID
		.Parameters("@QuizID") = intID
		.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
		.Parameters("@Question") = GetTextArea( Request("Question") )
		.Parameters("@Answer") = Request("Answer")
		.Parameters("@A") = Format( Request("A") )
		.Parameters("@B") = Format( Request("B") )
		.Parameters("@C") = Format( Request("C") )
		.Parameters("@D") = Format( Request("D") )
		.Parameters("@E") = Format( Request("E") )
		.Parameters("@F") = Format( Request("F") )

		.Execute , , adExecuteNoRecords
	End With
	Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
	<p>The question has been added. &nbsp;<a href="members_quizques_add.asp?ID=<%=intID%>">Click here</a> to add another.</p>
<%
'-----------------------Begin Code----------------------------
else
'------------------------End Code-----------------------------
%>
	<p>You must add you questions one at a time.  Whatever answer options you do not need, leave blank.  Remember to click the answer box to indicate the correct answer.</p>
	* indicates required information<br>

	<form METHOD="post" ACTION="<%=SecurePath%>members_quizques_add.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>
	<tr>
		<td class="<% PrintTDMain %>">&nbsp;</td>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Quiz Name
		</td>
		<td class="<% PrintTDMain %>">
<%
'-----------------------Begin Code----------------------------
		PrintQuizPullDown intID, Session("MemberID")
'------------------------End Code-----------------------------
%>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>">
			Answer?
		</td>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			* Question
		</td>
		<td class="<% PrintTDMain %>">
			<% TextArea "Question", 55, 4, True, "" %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="center">
			<input type="radio" name="Answer" value="A" checked>
		</td>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			* Option A
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="A">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="center">
			<input type="radio" name="Answer" value="B">
		</td>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Option B
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="B">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="center">
			<input type="radio" name="Answer" value="C">
		</td>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Option C
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="C">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="center">
			<input type="radio" name="Answer" value="D">
		</td>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Option D
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="D">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="center">
			<input type="radio" name="Answer" value="E">
		</td>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Option E
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="E">
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>" align="center">
			<input type="radio" name="Answer" value="F">
		</td>
		<td class="<% PrintTDMain %>" valign="top" align="right">
			Option F
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="F">
		</td>
	</tr>
	<tr>
		<td colspan="3" align="center" class="<% PrintTDMain %>">
			<input type="submit" name="Submit" value="Add"><br>
		</td>
	</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
end if

'-------------------------------------------------------------
'This function writes a pulldown menu for members
'-------------------------------------------------------------
Sub PrintQuizPullDown( intQuizID, intMemberID )
	intQuizID = CInt(intQuizID)
	'Now we are going to get the group names to list in the pull-down menu
	if LoggedAdmin then
		Query = "SELECT ID, Subject FROM Quizzes WHERE (CustomerID = " & CustomerID & ")"
	else
		Query = "SELECT ID, Subject FROM Quizzes WHERE (MemberID = " & intMemberID & " AND CustomerID = " & CustomerID & ")"
	end if
	Set rsTempQuizzes = Server.CreateObject("ADODB.Recordset")
	rsTempQuizzes.CacheSize = PageSize
	rsTempQuizzes.Open Query, Connect, adOpenStatic, adLockReadOnly
	
	'Make the size 3 if there are many members
	if rsTempQuizzes.RecordCount <= 30 then
		%><select name="ID" size="1"><%
	else
		%><select name="ID" size="3"><%
	end if

	do until rsTempQuizzes.EOF
		'Highlight the current section
		if rsTempQuizzes("ID") = intQuizID then
			Response.Write "<option value = '" & rsTempQuizzes("ID") & "' SELECTED>" & rsTempQuizzes("Subject") & "</option>" & vbCrlf
		else
			Response.Write "<option value = '" & rsTempQuizzes("ID") & "'>" & rsTempQuizzes("Subject") & "</option>" & vbCrlf
		end if
		rsTempQuizzes.MoveNext
	loop
	rsTempQuizzes.Close
	set rsTempQuizzes = Nothing
	Response.Write("</select>")
End Sub

'------------------------End Code-----------------------------
%>
