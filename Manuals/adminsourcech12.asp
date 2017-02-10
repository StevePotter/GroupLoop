<% intChapter = 12 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: THE QUIZ SECTION</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 QUIZZES AND THE ADMINISTRATOR: </span>&nbsp;
Quizzes are great way to inject a refreshing brand of comedy and conversation into your site.  This type of addition 
changes the pace of your site by using a new exciting format unlike the standard site additions.  Just like voting polls, 
you can either allow or prevent non-administrators from creating quizzes.</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Specifically, a quiz is a set of several multiple-choice questions.  Each question has one correct answer.  After a 
	site-goer takes a quiz, your site will evaluate the quiz and show the quiz taker which questions he/she got 
	wrong.  Every quiz has three properties:

	<blockquote>
	<b><i>Name: </i></b>
	The name is just the title of the quiz.  A good quiz name typically sums up the content of the quiz's questions.
	<br>
	<b><i>Privacy Status: </i></b>
	This property determines whether or not guests can participate in a quiz.  Guests cannot take or even view a quiz marked with privacy.
	<br>
	<b><i>Questions: </i></b>
	The quiz questions are just that: the individual questions, one correct answer choice, and up to five incorrect choices.
	</blockquote>

<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 USING QUIZZES: </span>&nbsp;
Manipulating quizzes can sometimes get a little difficult because each question is a separate entity.  Keep in mind 
that there is the concept of the quiz and then it's individual questions.  You manipulate questions associated with 
the quiz.

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Creating a New Quiz - </span>
	Before creating a quiz, it's best to first plan and write it out on paper or in a word processing program.  This 
	will cut down on the time you spend because it's changes are easier and faster.  When your done with the offline 
	work, you can upload the new quiz to the site:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under the subheading Quizzes, click Create a New Quiz.
		<br>	
		<b><i>2 <%PrintSymb "Create", "add a new quiz1.gif"%>: </i></b> 
		Enter the quiz name and select the quiz's privacy status.
		<br>
		<b><i>3 <%PrintSymb "Confirmation", "add a new quiz2.gif"%>: </i></b> 
		Click the Add button to move on to the questions.
		<br>
		<b><i>4 <%PrintSymb "Create", "add question1.gif"%>: </i></b> 
		Enter the question and up to six choices.
		<blockquote>
			<%PrintArrow%><b><i>Correct Answer: </i></b> 
			Select the option bubble to the left of the correct choice's checkbox.
			<br>
			<%PrintArrow%><b><i>Blank Textboxes:: </i></b> 
			If you have less than six choices for the question, it's ok.  Just leave the additional choice textboxes blank.
		</blockquote>
		<br>
		<b><i>5: </i></b> 
		Click the Add button to include the question with your quiz.
		<br>
		<b><i>6 <%PrintSymb "Confirmation", "question added.gif"%>: </i></b> 
		Select the appropriate link. To add more questions, click the Additional Questions link and 
		repeat steps 4, 5 and 6.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing an Existing Quiz - </span>
	The editing procedure for a quiz is a little different than the creation because all of the quizzes questions 
	 will appear on the same page.  To edit an existing quiz:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under the subheading Quizzes, click Modify Quizzes.
		<br>	
		<b><i>2 <%PrintSymb "List", "choose quiz to modify.gif"%>: </i></b> 
		Find the appropriate quiz and click the Edit button to its right.
		<br>
		<b><i>3 <%PrintSymb "Edit", "edit quiz.gif"%>: </i></b> 
		Make the necessary changes to the quiz.
		<blockquote>
			<%PrintArrow%><b><i>Adding New Questions: </i></b> 
			To add new questions to your quiz, click the Add A Question link at the top of the page.  You can then 
			use steps 4, 5 and 6 from the previous Creating a New Quiz procedure.
			<br>
			<%PrintArrow%><b><i>Deleting Questions: </i></b> 
			To completely remove a question (and all of its choices), simply press the Delete button next to 
			the question.
			<br>
			<%PrintArrow%><b><i>Editing Questions: </i></b> 
			To edit the question or add/edit/delete the question's answers, just click on the Edit button right next to 
			the question.
		</blockquote>
		<b><i>4: </i></b> 
		Click the Update button located below the final option.
		<br>
		<b><i>5 <%PrintSymb "Confirmation", "quiz has been edited.gif"%>: </i></b> 
		Using the links, reload either the quiz list or the admin menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Quiz - </span>
	When you delete a quiz, you are removing it completely from your site.  Each question, it's answer choices, 
	and it's statistical results will be lost forever.  To delete a quiz:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under the subheading Quizzes, click Modify Quizzes.
		<br>	
		<b><i>2 <%PrintSymb "List", "choose quiz to modify.gif"%>: </i></b> 
		Find the desired quiz to delete.
		<br>
		<b><i>3 <%PrintSymb "Delete", "none"%>: </i></b> 
		Click the Delete button to the right of the desired quiz.
		<br>
		<b><i>4 <%PrintSymb "PopUp", "delete quiz warning box.gif"%>: </i></b> 
		If you're sure, click the OK button.  If not, click Cancel.
		<br>	
		<b><i>5 <%PrintSymb "Confirmation", "quiz has been deleted.gif"%>: </i></b> 
		Using the links, reload either the quiz list or the admin menu.

	</blockquote>

