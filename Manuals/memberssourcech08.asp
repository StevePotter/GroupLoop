<% intChapter = 8 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: THE VOTING SECTION</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 VOTING POLLS: </span>&nbsp;
A voting poll is a single multiple-choice question you ask of your website's audience.  Site-goers participate in the 
poll and select the poll option of their choice.  Keep in mind that voting polls usually have no definite answer, just 
opinions.  If your question or subject has a specific answer, you may want to make it part of a quiz instead. The 
resulting percentages for each option are displayed in a bar graph.  Each time a site-goer votes in a poll, the poll's 
results are retabulated and a graph of the new percentages is created.  Each voting poll has six properties:
	<blockquote>
	<b><i>Poll Question: </i></b>
	This is just the question (or subject) asked of the audience.  The poll question also acts as the poll title.
	<br>
	<b><i>Poll Status: </i></b>
	The poll status determines whether or not a poll is open to voting.  A poll is automatically open when created.  When 
	a poll is closed, site-goers can no longer vote in it.  However, the final results of a closed poll are still 
	available.  A closed poll can be reopened at any time.
	<br>
	<b><i>Privacy Status: </i></b>
	Privacy status determines whether or not guests can take a particular voting poll.  If a poll is marked 
	private, guests cannot vote in the poll.
	<br>
	<b><i>Results Privacy: </i></b>
	Result privacy determines whether or not guests can view the current result percentages of a poll.  If a poll is 
	marked with results privacy only members can view its results graph.
	<br>
	<b><i>Vote Limit: </i></b>
	This decides how many times a site-goer can vote in a poll.  For instance, the member can make it so each 
	individual site-goer can only vote in a particular poll one time.  Remember the limit is optional, and can be set 
	to a daily limit or a total limit.
	<br>
	<b><i>Options: </i></b>
	The poll options are the multiple-choice answers for the poll question.  You can have anywhere from 1 to a million options.  
	It's up to you (as usual).
	</blockquote>

	<p align=left class=BodyText><b>Note: </b> 
	Remember that the member can restrict things in each section.  They may choose to not allow you to create 
	voting polls.  So if the option isn't there, we didn't do it.
	</p>

<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 USING VOTING POLLS: </span>&nbsp;
The voting poll section is a fun way to get a feel for what your site-goers are thinking.  Set up voting polls to find 
out your audience's opinion on the quality your site, the best sports team ever, the movie pick of the month....(you get 
the picture.).

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Creating a New New Voting Poll - </span>
	Use the following procedure to create a voting poll:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Add A Poll link under the subheading Voting.
		<br>	
		<b><i>2 <%PrintSymb "Create", "add a new voting poll.gif"%>: </i></b> 
		Make the proper data field selections and entries (see above).
		<blockquote>
			<%PrintArrow%><b><i>Six Options: </i></b> 
			You don't have to create exactly six poll options.  If your poll has less than six options, just leave the 
			option textboxes you don't need blank.  If you have more than six option, you will need to finish creating 
			your poll with just the first six options and then, add the additional ones in the poll's edit page (next 
			procedure).
		</blockquote>
		<b><i>3: </i></b> 
		Click the Add button.
		<br>
		<b><i>4 <%PrintSymb "Confirmation", "poll has been added.gi"%>: </i></b> 
		Using the supplied links, either add more options to the poll, add another, or return to the main member page.
		<br><br>
		<b>Note:</b> Installing a voting limit isn't always 100% accurate.  This is because your site can only remember certain 
		voters.  Although member voters are distinctly remembered, guest voters can only be identified by their IP 
		addresses (unique Internet ID).  Many internet providers assign each computer with a new IP address every time that computer logs 
		on to the internet. In other words, certain guests may be able to vote an infinite number of times.  So if you 
		really want to create solid limits, make the poll private.  Sorry, but we can't do anything about this one.
	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing an Existing Voting Poll - </span>
	To edit an existing poll use the following procedure:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Polls link under the subheading Voting.
		<br>	
		<b><i>2 <%PrintSymb "List", "choose poll to modify.gif"%>: </i></b> 
		Find the desired poll and click the Edit button to its right.
		<br>
		<b><i>3 <%PrintSymb "Edit", "editing poll.gif"%>: </i></b> 
		Make the necessary changes and click the Update button.
		<blockquote>
			<%PrintArrow%><b><i>More Options: </i></b> 
			If you wish to add more options, click the Add More Options link located below your last option.  Your 
			browser will load a new page with six textboxes for your additional options.  Enter your new options and 
			click the Add button (you don't have to add exactly six options; leave any additional textboxes blank).  
			If you want to add <b>more</b> than six options you will have to repeat this process.
			<br>
			<%PrintArrow%><b><i>Removing Options: </i></b> 
			To remove a poll option, mark the checkbox to the left of the doomed option.
			<br>
			<%PrintArrow%><b><i>Changing Existing Options: </i></b> 
			Realize that when you simply change the information in an option textbox, that option will maintain the same 
			status in the results graph.  In other words, if you change an option's text from Red to Blue after 
			site-goers have already voted in the poll, all votes for the Red option will carry over to the Blue 
			option (as will be reflected in the results graph).
		</blockquote>
		<b><i>4: </i></b> 
		Click the Update button located below the final option.
		<br>
		<b><i>5 <%PrintSymb "Confirmation", "poll has been edited.gif"%>: </i></b> 
		Using the links, reload either the voting poll list or the member menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Voting Poll - </span>
	This procedure is <b>not</b> for closing a poll to voting.  Using this procedure will remove a poll and its corresponding 
	results graph entirely from your site.  To delete a voting poll:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Polls link under the subheading Voting.
		<br>	
		<b><i>2 <%PrintSymb "List", "choose poll to modify.gif"%>: </i></b> 
		Find the desired poll to delete.
		<br>
		<b><i>3 <%PrintSymb "Delete", "none"%>: </i></b> 
		Click the Delete button to the right of the desired poll.
		<br>
		<b><i>4 <%PrintSymb "PopUp", "delete poll warning box.gif"%>: </i></b> 
		If you're sure, click the OK button.  If not, click Cancel.
		<br>	
		<b><i>5 <%PrintSymb "Confirmation", "poll has been deleted.gif"%>: </i></b> 
		Using the links, reload either the voting poll list or the member menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Closing a Poll to Voting - </span>
	Closing a poll to voting simply means site-goers can no longer vote in it; the results graph will still be intact.  
	Although a poll can be closed from a poll's edit page, there is an faster way.  Since we figured you may have 
	several polls going at once, we devoted an entire member option to closing opened polls:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under the Voting subheading, click the Close a Poll link.
		<br>	
		<b><i>2 <%PrintSymb "List", ""%>: </i></b> 
		Find the poll you wish to close and click the Close button.
	</blockquote>

	<p><b>Note:</b> Changing the poll status property in a poll's edit page to Closed will accomplish the same goal 
	as the above procedure.  Also, you can re-open a closed poll by editing it (as shown above).
	</p>