<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeQuizzes ) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but this section has been deactivated. An administrator can reactivate it."))
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading><%=QuizzesTitle%></span><br>
<span class=LinkText><a href="javascript:history.back(1)">Back</a></span>

<%
'-----------------------Begin Code----------------------------
intID = Request("ID")
if intID = "" then Redirect("error.asp?Message=" & Server.URLEncode("No ID was specified."))
intID = CInt(intID)

'Open up the item
Query = "SELECT ID, MemberID, Date, Subject, Private, TimesTaken, TotalScore, Description FROM Quizzes WHERE ID = " & intID & " AND CustomerID = " & CustomerID
Set rsQuiz = Server.CreateObject("ADODB.Recordset")
rsQuiz.Open Query, Connect, adOpenStatic, adLockReadOnly

if rsQuiz.EOF then
	Set rsQuiz = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The quiz does not exist.  If you pasted a link, there may be a typo, or the quiz may have been deleted.  Please refer to the quiz list to find the desired quiz, if it still exists."))
end if

if rsQuiz("Private") = 1 AND not LoggedMember then
	rsQuiz.Close
	Set rsQuiz = Nothing
	Redirect( "login.asp?Source=quizzes_take.asp&ID=" & intID & "&Submit=Take" )
end if

if LoggedAdmin or (LoggedMember and Session("MemberID") = rsQuiz("MemberID"))  then
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
%>
	<p class="LinkText" align="<%=HeadingAlignment%>"><a href="javascript:history.go(-2)">Back To List</a></p>
<%
	AddRating intID, "Quizzes"
'Process their answers
elseif Request("Submit") = "I'm Done" then
	'Open it for writing
	rsQuiz.Close
	rsQuiz.Open Query, Connect, adOpenStatic, adLockOptimistic

	Query = "SELECT ID, Question, A, B, C, D, E, F, Answer FROM QuizQuestions WHERE QuizID = " & intID & " AND CustomerID = " & CustomerID & " ORDER BY ID"
	Set rsQuestions = Server.CreateObject("ADODB.Recordset")
	rsQuestions.CacheSize = 20
	rsQuestions.Open Query, Connect, adOpenStatic, adLockReadOnly

	if rsQuiz.EOF or rsQuestions.EOF then
		set rsQuiz = Nothing
		set rsQuestions = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("There are no questions on this quiz, even though you took it.  Wow.  You must have nothing to do."))
	end if

	Set ID = rsQuestions("ID")
	Set Question = rsQuestions("Question")
	Set Answer = rsQuestions("Answer")

	intQuesNum = 0
	intNumCorrect = 0
%>
	<p class="Heading" align="left">Results</p>
<%
	'Get results on each question
	do until rsQuestions.EOF
		intQuesNum = intQuesNum + 1
'------------------------End Code-----------------------------
%>
		<%=intQuesNum%>. <%=Question%><br>
<%
'-----------------------Begin Code----------------------------
		QuesID = ID
		if Request(QuesID) = Answer then
			IncrementHits ID, "QuizQuestions"
			intNumCorrect = intNumCorrect + 1
'------------------------End Code-----------------------------
%>
			Correct!<br><br>
<%
'-----------------------Begin Code----------------------------
		else
			strAnswer = Answer
'------------------------End Code-----------------------------
%>
			Wrong! (correct answer was: <%=rsQuestions(strAnswer)%>)<br><br>
<%
'-----------------------Begin Code----------------------------
		end if
		rsQuestions.MoveNext
	loop
	Percnt = Round( ( intNumCorrect / intQuesNum ) * 100 )
	if Percnt > 89 then
		Rating = QuizResult90
	elseif Percnt > 60 and Percnt < 90 then
		Rating = QuizResult60
	else
		Rating = QuizResult0

	end if

'------------------------End Code-----------------------------
%>
	<p>You got <%=intNumCorrect%> out of <%=intQuesNum%> correct.  That's a <b><%=Percnt%>%</b>. <%=Rating%><br>
<%
'-----------------------Begin Code----------------------------
	intNumQuestions = ((rsQuiz("TimesTaken") + 1) * intQuesNum)
	AvgPercnt = Round( ( (rsQuiz("TotalScore") + intNumCorrect) / intNumQuestions ) * 100 )
	rsQuiz("TimesTaken") = rsQuiz("TimesTaken") + 1
	rsQuiz("TotalScore") = rsQuiz("TotalScore") + intNumCorrect
	rsQuiz.Update

'------------------------End Code-----------------------------
%>
	<p>This quiz has been taken <%=rsQuiz("TimesTaken")%> times.  The average score is a <%=AvgPercnt%>%.</p>
<%
'-----------------------Begin Code----------------------------
	if RateQuizzes = 1 then
		PrintRatingPulldown rsQuiz("ID"), "", "Quizzes", "quizzes_take.asp", "quiz"
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
			PrintReviews "quizzes_rate.asp?", "Quizzes", intID
			Set rsPage = Nothing
		end if
	end if

	set rsQuestions = Nothing
'List the quiz
else
	Query = "SELECT ID, Question, A, B, C, D, E, F, Answer FROM QuizQuestions WHERE QuizID = " & intID & " AND CustomerID = " & CustomerID & " ORDER BY ID"
	Set rsQuestions = Server.CreateObject("ADODB.Recordset")
	rsQuestions.CacheSize = 20
	rsQuestions.Open Query, Connect, adOpenStatic, adLockReadOnly
	if rsQuestions.EOF then
'------------------------End Code-----------------------------
%>
		<p>Sorry, but there are no questions in this quiz, so you cannot take it.</p>
<%
'-----------------------Begin Code----------------------------
	else
		Set ID = rsQuestions("ID")
		Set Question = rsQuestions("Question")
		Set Answer = rsQuestions("Answer")
		Set A = rsQuestions("A")
		Set B = rsQuestions("B")
		Set C = rsQuestions("C")
		Set D = rsQuestions("D")
		Set E = rsQuestions("E")
		Set F = rsQuestions("F")


		IncrementStat "QuizzesTaken"
		IncrementHits rsQuiz("ID"), "Quizzes"

		strDes = rsQuiz("Description")
		if not IsNull(strDes) then
			if strDes <> "" then strDes = "<p>" & strDes & "</p>"

		end if
'------------------------End Code-----------------------------
%>
		<p align="center" class=Heading><%=rsQuiz("Subject")%></p>
		<%=strDes%>
		<form method="post" action="quizzes_take.asp">
		<input type="hidden" name="ID" value="<%=intID%>">
<%
'-----------------------Begin Code----------------------------
		do until rsQuestions.EOF
			intintQuesNum = intintQuesNum + 1
			'Now display the question number and question
'------------------------End Code-----------------------------
%>
			<p><b><%=intintQuesNum%>. <%=Question%></b>
<%
'-----------------------Begin Code----------------------------	
			if not A = "" then
'------------------------End Code-----------------------------
%>
				<BR>&nbsp;&nbsp;&nbsp;
 	 			<input type="radio" name="<%=ID%>" value="A"><%=A%>
<%
'-----------------------Begin Code----------------------------	
			end if
			if not B = "" then
'------------------------End Code-----------------------------
%>
				<BR>&nbsp;&nbsp;&nbsp;
 		 		<input type="radio" name="<%=ID%>" value="B"><%=B%>
<%
'-----------------------Begin Code----------------------------
			end if
			if not C = "" then
'------------------------End Code-----------------------------
%>
				<BR>&nbsp;&nbsp;&nbsp;
 		 		<input type="radio" name="<%=ID%>" value="C"><%=C%>
<%
'-----------------------Begin Code----------------------------
			end if
			if not D = "" then
'------------------------End Code-----------------------------
%>
				<BR>&nbsp;&nbsp;&nbsp;
 	 			<input type="radio" name="<%=ID%>" value="D"><%=D%>
<%
'-----------------------Begin Code----------------------------
			end if
			if not E = "" then
'------------------------End Code-----------------------------
%>
				<BR>&nbsp;&nbsp;&nbsp;
 	 			<input type="radio" name="<%=ID%>" value="E"><%=E%>
<%
'-----------------------Begin Code----------------------------
			end if
			if not F = "" then
'------------------------End Code-----------------------------
%>
			<BR>&nbsp;&nbsp;&nbsp;
 	 		<input type="radio" name="<%=ID%>" value="F"><%=F%>
<%
'-----------------------Begin Code----------------------------
			end if
			Response.Write("</p>")
			rsQuestions.MoveNext
		loop
		rsQuestions.Close
'------------------------End Code-----------------------------
%>
		<p align="center"><input type="submit" name="Submit" value="I'm Done"></p>
		</form>
<%
'-----------------------Begin Code----------------------------
	end if
	set rsQuestions = Nothing
end if

rsQuiz.Close
set rsQuiz = Nothing

'------------------------End Code-----------------------------
%>