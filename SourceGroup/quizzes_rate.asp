<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeQuizzes ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading><%=QuizzesTitle%></span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span><br>

<%
'-----------------------Begin Code----------------------------
'Get the ID of the item
if Request("ID") <> "" then
	intID = CInt(Request("ID"))
else
	Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
end if


'Open up the item
Query = "SELECT Date, Private, MemberID, Subject, TimesTaken, TotalScore, ID FROM Quizzes WHERE ID = " & intID  & " AND CustomerID = " & CustomerID
Set rsItem = Server.CreateObject("ADODB.Recordset")
rsItem.Open Query, Connect, adOpenStatic, adLockReadOnly

if rsItem.EOF then
	Set rsItem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The quiz does not exist.  If you pasted a link, there may be a typo, or the quiz may have been deleted.  Please refer to the quiz list to find the desired quiz, if it still exists."))
end if

if rsItem("Private") = 1 AND not LoggedMember then
	rsItem.Close
	Set rsItem = Nothing
	Redirect( "login.asp?Source=quizzes_rate.asp&ID=" & intID & "&Submit=View" )
end if

if LoggedAdmin or (LoggedMember and Session("MemberID") = rsItem("MemberID"))  then
%>
	<table align=<%=HeadingAlignment%>>
	<tr>
	<td align=right width="50%" class="LinkText"><a href="members_quizzes_modify.asp?Submit=Edit&ID=<%=intID%>">Edit</a>&nbsp;&nbsp;</td>
	<td align=left width="50%" class="LinkText">&nbsp;&nbsp;
		<a href="javascript:DeleteBox('If you delete this quiz, there is no way to get it back.  Are you sure?', 'members_quizzes_modify.asp?Submit=Delete&ID=<%=intID%>')">Delete</a>
	</td>
	</tr>
	</table>
<%
end if

Response.Write "</p>"

if Request("Rating") <> "" and RateQuizzes = 1 then
%>	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a></p>
<%
	AddRating intID, "Quizzes"
else
	IncrementHits intID, "Quizzes"
'------------------------End Code-----------------------------
%>
	<p align="<%=HeadingAlignment%>" class=LinkText><a href="quizzes_take.asp?ID=<%=intID%>">Take This Quiz</a></p>

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

	<p><strong><b>Subject: <%=rsItem("Subject")%></b></strong></p>
<%
'-----------------------Begin Code----------------------------
	Query = "SELECT Question FROM QuizQuestions WHERE QuizID = " & intID & " ORDER BY ID"
	Set rsQuestions = Server.CreateObject("ADODB.Recordset")
	rsQuestions.CacheSize = 20
	rsQuestions.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsQuestions.EOF then
%>
		<p>Sorry, but there are no questions in this quiz.</p>
<%
	else
		intQuesNum = 0
		Set Question = rsQuestions("Question")
		'Get results on each question
		do until rsQuestions.EOF
			intQuesNum = intQuesNum + 1
	'------------------------End Code-----------------------------
	%>
			<%=intQuesNum%>. <%=Question%><br>
	<%
	'-----------------------Begin Code----------------------------
			rsQuestions.MoveNext
		loop

	end if

	Set rsQuestions = Nothing

	Response.Write "<br><br>"

	if RateQuizzes = 1 then
		PrintRatingPulldown intID, "", "Quizzes", "quizzes_rate.asp", "quiz"
	end if
	if ReviewQuizzes = 1 then
		if LoggedAdmin then
%>
			<a href="admin_reviews_modify.asp?Source=quizzes_rate.asp?ID=<%=intID%>&TargetTable=Quizzes&TargetID=<%=intID%>">Modify Reviews</a><br>
<%
		end if
%>
		<a href="review.asp?Source=quizzes_rate.asp?ID=<%=intID%>&TargetID=<%=intID%>&Table=Quizzes">Add a review</a><br>
<%
		if ReviewsExist( "Quizzes", intID ) then
			Set rsPage = Server.CreateObject("ADODB.Recordset")
			PrintReviews "quizzes_rate.asp", "Quizzes", intID
			Set rsPage = Nothing
		end if
	end if
end if

set rsItem = Nothing
'------------------------End Code-----------------------------
%>