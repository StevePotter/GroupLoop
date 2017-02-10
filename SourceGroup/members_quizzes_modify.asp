<%
'
'-----------------------Begin Code----------------------------
if not ( CBool( IncludeQuizzes ) or CBool( QuizzesMembers ) ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
if not LoggedMember and Request("MemberID") <> "" and Request("Password") <> "" then Relog Request("MemberID"), Request("Password")
if not LoggedMember then Redirect("members.asp?Source=members_quizzes_modify.asp")
blLoggedAdmin = LoggedAdmin
if not (blLoggedAdmin or CBool( QuizzesMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Quizzes</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>


<%
'-----------------------Begin Code----------------------------
if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

strSubmit = Request("Submit")


'Update the question
if strSubmit = "Update" and Request("QuestionID") <> "" then
	intID = CInt(Request("QuestionID"))

	if Request("Question") = "" then Redirect("incomplete.asp")

	Query = "SELECT QuizID, Question, Answer, A, B, C, D, E, F, IP, ModifiedID FROM QuizQuestions WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	rsUpdate("Question") = GetTextArea( Request("Question") )
	rsUpdate("A") = Format( Request("A") )
	rsUpdate("B") = Format( Request("B") )
	rsUpdate("C") = Format( Request("C") )
	rsUpdate("D") = Format( Request("D") )
	rsUpdate("E") = Format( Request("E") )
	rsUpdate("F") = Format( Request("F") )
	rsUpdate("Answer") = Request("Answer")
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	intQuizID = rsUpdate("QuizID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The question has been edited. &nbsp;<a href="members_quizzes_modify.asp?Submit=Edit&ID=<%=intQuizID%>">Click here</a> to modify another question in this quiz.<br>
	<a href="members_quizzes_modify.asp">Click here</a> to modify a different quiz.
	</p>
<%
'-----------------------Begin Code----------------------------
'Delete the question
elseif strSubmit = "Delete" and Request("QuestionID") <> "" then
	intID = CInt(Request("QuestionID"))

	Query = "SELECT QuizID FROM QuizQuestions WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	intQuizID = rsUpdate("QuizID")
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

'------------------------End Code-----------------------------
%>
	<p>The question has been deleted. &nbsp;<a href="members_quizzes_modify.asp?Submit=Edit&ID=<%=intQuizID%>">Click here</a> to modify another question in this quiz.<br>
	<a href="members_quizzes_modify.asp">Click here</a> to modify a different quiz.
	</p>
<%
'-----------------------Begin Code----------------------------
'This is for editing a single question
elseif strSubmit = "Edit" and Request("QuestionID") <> "" then
	intID = CInt(Request("QuestionID"))

	Query = "SELECT Question,Answer,A,B,C,D,E,F FROM QuizQuestions WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsEdit.EOF then
		set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

'------------------------End Code-----------------------------
%>
	<p class=LinkText align=<%=HeadingAlignment%>><a href="javascript:history.back(1)">Back To Question List</a></p>
	<p>Whatever options you don't want to use, leave blank.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_quizzes_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0%>
	<input type="hidden" name="QuestionID" value="<%=intID%>">
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			* Question
		</td>
		<td class="<% PrintTDMain %>">
			<% TextArea "Question", 50, 2, True, rsEdit("Question") %>
		</td>
	</tr>
	<tr>
		<td class="<% PrintTDMain %>">
			Answer?
		</td>
		<td class="<% PrintTDMain %>">
			Option
		</td>
	</tr>
<%
'-----------------------Begin Code----------------------------
	for i = 1 to 5
	next

	if rsEdit("Answer") = "A" then	strSelected = "checked"
'------------------------End Code-----------------------------
%>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			<input type="radio" name="Answer" value="A" <%=strSelected%>>A
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="A" value="<%=FormatEdit( rsEdit("A") )%>">
		</td>
	</tr>
<%
'-----------------------Begin Code----------------------------
	strSelected = ""
	if rsEdit("Answer") = "B" then	strSelected = "checked"
'------------------------End Code-----------------------------
%>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			<input type="radio" name="Answer" value="B" <%=strSelected%>>B
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="B" value="<%=FormatEdit( rsEdit("B") )%>">
		</td>
	</tr>
<%
'-----------------------Begin Code----------------------------
	strSelected = ""
	if rsEdit("Answer") = "C" then	strSelected = "checked"
'------------------------End Code-----------------------------
%>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			<input type="radio" name="Answer" value="C" <%=strSelected%>>C
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="C"value="<%=FormatEdit( rsEdit("C") )%>">
		</td>
	</tr>
<%
'-----------------------Begin Code----------------------------
	strSelected = ""
	if rsEdit("Answer") = "D" then	strSelected = "checked"
'------------------------End Code-----------------------------
%>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			<input type="radio" name="Answer" value="D" <%=strSelected%>>D
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="D" value="<%=FormatEdit( rsEdit("D") )%>">
		</td>
	</tr>
<%
'-----------------------Begin Code----------------------------
	strSelected = ""
	if rsEdit("Answer") = "E" then	strSelected = "checked"
'------------------------End Code-----------------------------
%>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			<input type="radio" name="Answer" value="E" <%=strSelected%>>E
		</td>
		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="E" value="<%=FormatEdit( rsEdit("E") )%>">
		</td>
	</tr>
<%
'-----------------------Begin Code----------------------------
	strSelected = ""
	if rsEdit("Answer") = "F" then	strSelected = "checked"
'------------------------End Code-----------------------------
%>
	<tr>
		<td class="<% PrintTDMain %>" align="right">
			<input type="radio" name="Answer" value="F" <%=strSelected%>>F
		</td>

		<td class="<% PrintTDMain %>">
			<input type="text" size="50" name="F" value="<%=FormatEdit( rsEdit("F") )%>">
		</td>
	</tr>
		<tr>
			<td colspan="2" class="<% PrintTDMain %>" align="center">
				<input type="submit" name="Submit" value="Update">
			</td>
		</tr>
	</table>
	</form>
<%
'-----------------------Begin Code----------------------------
	rsEdit.Close
	set rsEdit = Nothing

'Update the quiz
elseif strSubmit = "Update" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	if (blLoggedAdmin and Request("Date") = "") or Request("Subject") = "" then Redirect("incomplete.asp")

	Query = "SELECT Private, Subject, Date, IP, ModifiedID, Description FROM Quizzes WHERE ID = " & intID & " AND " & strMatch 
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if Request("Private") = "1" then 
		rsUpdate("Private") = 1
	else
		rsUpdate("Private") = 0
	end if
	if blLoggedAdmin then rsUpdate("Date") = Request("Date")
	rsUpdate("Subject") = Format( Request("Subject") )
	rsUpdate("Description") = GetTextArea( Request("Description") )
	rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
	rsUpdate("ModifiedID") = Session("MemberID")
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing
'------------------------End Code-----------------------------
%>
	<p>The quiz has been edited. &nbsp;<a href="members_quizzes_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------
elseif strSubmit = "Delete" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT ID FROM Quizzes WHERE ID = " & intID & " AND " & strMatch
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsUpdate.EOF then
		set rsUpdate = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if
	rsUpdate.Delete
	rsUpdate.Update
	rsUpdate.Close
	Set rsUpdate = Nothing

	Query = "DELETE QuizQuestions WHERE QuizID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

	Query = "DELETE Reviews WHERE TargetTable = 'Quizzes' AND TargetID = " & intID
	Connect.Execute Query, , adCmdText + adExecuteNoRecords

'------------------------End Code-----------------------------
%>
	<p>The quiz has been deleted. &nbsp;<a href="members_quizzes_modify.asp">Click here</a> to modify another.</p>
<%
'-----------------------Begin Code----------------------------

elseif strSubmit = "Edit" then
	if Request("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the ID."))
	intID = CInt(Request("ID"))

	Query = "SELECT Private, Subject, Date, Description FROM Quizzes WHERE ID = " & intID & " AND " & strMatch
	Set rsEdit = Server.CreateObject("ADODB.Recordset")
	rsEdit.Open Query, Connect, adOpenStatic, adLockOptimistic
	if rsEdit.EOF then
		set rsEdit = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
	end if

	if rsEdit("Private") = 1 then 
		strChecked = "checked"
	else
		strChecked = ""
	end if

'------------------------End Code-----------------------------
%>
	<p><a href="members_quizques_add.asp?ID=<%=intID%>">Click here</a> to add new questions to this quiz.</p>
	* indicates required information<br>
	<form method="post" action="<%=SecurePath%>members_quizzes_modify.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<input type="hidden" name="ID" value="<%=intID%>">
	<%PrintTableHeader 0%>
	<tr> 
		<td class="<% PrintTDMain %>" align="right">Private?</td>
		<td class="<% PrintTDMain %>"> 
			<input type="checkbox" name="Private" value="1" <%=strChecked%>>
     	</td>
   	</tr>
<%	if blLoggedAdmin then %>
	<tr> 
      	<td class="<% PrintTDMain %>" align="right">* Date Posted</td>
      	<td class="<% PrintTDMain %>"> 
       		<input type="text" name="Date" size="15" value="<%=FormatDateTime(rsEdit("Date"), 2)%>">
     	</td>
    </tr>
<%	end if %>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">* Quiz Name</td>
     	<td class="<% PrintTDMain %>"> 
			<input type="text" size="50" name="Subject" value="<%=rsEdit("Subject")%>">
    	</td>
	</tr>
	<tr> 
     	<td class="<% PrintTDMain %>" align="right">Description of this quiz</td>
     	<td class="<% PrintTDMain %>"> 
			<% TextArea "Description", 50, 2, True, rsEdit("Description") %>
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
	Query = "SELECT ID, Question FROM QuizQuestions WHERE QuizID = " & intID & " AND " & strMatch & " ORDER BY ID"
	Set rsQuestions = Server.CreateObject("ADODB.Recordset")
	rsQuestions.CacheSize = 20
	rsQuestions.Open Query, Connect, adOpenForwardOnly, adLockOptimistic
	if not rsQuestions.EOF then
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">Question</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		intQuesNum = 0
		do until rsQuestions.EOF
			intQuesNum = intQuesNum + 1
%>
			<form method="post" action="<%=SecurePath%>members_quizzes_modify.asp">
			<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
			<input type="hidden" name="Password" value="<%=Session("Password")%>">
			<input type="hidden" name="QuestionID" value="<%=rsQuestions("ID")%>">
			<tr> 
     			<td class="<% PrintTDMain %>" align="left"><%=intQuesNum%>. <%=rsQuestions("Question")%> </td>
     			<td class="<% PrintTDMainSwitch %>"> 
				<input type="submit" name="Submit" value="Edit">
				<input type="button" value="Delete" onClick="DeleteBox('If you delete this question, there is no way to get it back.  Are you sure?', 'members_quizzes_modify.asp?Submit=Delete&QuestionID=<%=rsQuestions("ID")%>')">
				</td>
			</tr>
			</form>
<%
			rsQuestions.MoveNext
		loop
		rsQuestions.Close
		Response.Write "</table>"
	end if
	set rsQuestions = Nothing
	rsEdit.Close
	set rsEdit = Nothing

else
	Query = "SELECT ID, Date, Subject FROM Quizzes WHERE (CustomerID = " & CustomerID & " AND " & strMatch & ") ORDER BY Date DESC"
	Set rsPage = Server.CreateObject("ADODB.Recordset")
	rsPage.CacheSize = PageSize
	rsPage.Open Query, Connect, adOpenStatic, adLockReadOnly
	if not rsPage.EOF then
		Set ID = rsPage("ID")
		Set ItemDate = rsPage("Date")
		Set Subject = rsPage("Subject")
%>
		<form METHOD="POST" ACTION="members_quizzes_modify.asp">
<%
		PrintPagesHeader
		PrintTableHeader 0
%>
		<tr>
			<td class="TDHeader">&nbsp;</td>
			<td class="TDHeader">Subject</td>
			<td class="TDHeader">&nbsp;</td>
		</tr>
<%
		for i = 1 to rsPage.PageSize
			if not rsPage.EOF then
'------------------------End Code-----------------------------
%>
				<form METHOD="post" ACTION="members_quizzes_modify.asp">
				<input type="hidden" name="ID" value="<%=ID%>">
					<tr>
						<td class="<% PrintTDMain %>" align="center"><% PrintNew(ItemDate) %><a href="quizzes_rate.asp?ID=<%=ID%>">View</a></td>
						<td class="<% PrintTDMain %>"><%=Subject%></td>
						<td class="<% PrintTDMainSwitch %>">
						<input type="submit" name="Submit" value="Edit"> 
						<input type="button" value="Delete" onClick="DeleteBox('If you delete this quiz, there is no way to get it back.  Are you sure?', 'members_quizzes_modify.asp?Submit=Delete&ID=<%=ID%>')">
						<%if ReviewsExist( "Quizzes", ID ) AND blLoggedAdmin then%>
							<input type="button" value="Modify Reviews" onClick="Redirect('admin_reviews_modify.asp?Source=members_quizzes_modify.asp&TargetTable=Quizzes&TargetID=<%=ID%>')">
						<%end if%>
						</td>
					</tr>
				</form>
<%
'-----------------------Begin Code----------------------------
				rsPage.MoveNext
			end if
		next
		Response.Write("</table>")
		rsPage.Close
	else
'------------------------End Code-----------------------------
%>
		<p>Sorry, but there are no quizzes to modify.</p>
<%
'-----------------------Begin Code----------------------------
	end if

	set rsPage = Nothing
end if
'------------------------End Code-----------------------------
%>