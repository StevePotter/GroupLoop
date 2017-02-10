<% intChapter = 6 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: THE MESSAGE FORUM</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 THE MESSAGE FORUM: </span>&nbsp;
The message forum is without a doubt going to be your site's busiest and most popular section.  The forum's user 
friendly design allows both members and guests to post and respond to messages with ease.  As you start getting 
site regulars, you can look forward to seeing dozens of message posts a day.  With this activity though, you may  
find a fair amount of forum upkeep.
</p>

<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 FORUM TOPICS: </span>&nbsp;
The message forum can be separated into two distinct tiers: forum topics and message posts.  Topics are the main 
categories into which all message posts fit.  A forum topic has two data fields: the topic's name and its privacy level.
</p>

	<p>
	<b><i>Name: </i></b>
	This property is simply the name you choose for a topic.  It's best to use topics with self-explanatory names to make message forum navigation easy for your site-goers.
	</p>

	<p>
	<b><i>Privacy Level: </i></b>
	The privacy feature decides the exclusivity of the forum topic.  Each forum topic can fit into one of three 
	privacy levels: fully public, semi-public, and private.
	<blockquote>

		<b><i>Fully Public: </i></b> 
		Message posts in a fully public topic can be both read and written by members and guests alike.  However, 
		you should remember that members can still post private messages, which guests <b>cannot</b> read.  So 
		remember that this option still has plenty of security.
		<br>	
		<b><i>Semi-Public: </i></b> 
		The messages of a semi-public forum topic can only be written by members.  However, site guests can still 
		read these messages, unless they are private (just remember that anything, anywhere that is private cannot be 
		read by visitors).
		<br>
		<b><i>Private: </i></b> 
		Finally, only members can enter a private topic; site guests can neither read nor write messages in a topic 
		of this type.  The only thing that can ever be viewed by guests is the subject line in the Latest Additions section.
	
	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Creating a New Forum Topic - </span>
	Only administrators can create and modify forum topics.  If you have an idea for a new topic, post your idea in 
	the forum and see if they add it!


<a name="3"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.3 MESSAGE POST PROPERTIES: </span>&nbsp;
Every message post has seven properties: topic, type, date, author, privacy status, subject, and body.
</p>
	
	<blockquote>

		<b><i>Topic: </i></b>
		The topic property determines which forum topic a message post is located in.
		<br>

		<b><i>Type: </i></b>
		Message posts can be separated into two types: post heads and post threads.  A head post is a primary 
		conversational subject directly inserted into a forum topic.  Thread posts are replies to this head.  To 
		save space, threads can be collapsed into their head and basically hidden from view.  The head then acts 
		as the folder for all its threads.  A head and all of its threads is called a message unit.
		<br>

		<b><i>Date: </i></b>
		The date of a head message post determines where the post is displayed in its particular forum topic.  The 
		most recent head posts appear at the top their topic.  In turn, the date of a thread determines the 
		position within its head.
		<br>

		<b><i>Author: </i></b>
		This field is the name of the member or guest that originally wrote the message post.  A message written 
		by a member is automatically tagged with his/her nickname.  Guests have the ability use any author name 
		they choose.
		<br>

		<b><i>Privacy Status: </i></b>
		The privacy status property of a message post determines whether or not guests can view the message.  All 
		messages within a private topic are automatically marked private.
		<br>

		<b><i>Subject: </i></b>
		The subject is simply the title of the post.
		<br>

		<b><i>Body: </i></b>
		The body is the main information of a message post.

	</blockquote>

<a name="4"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.4 NAVIGATING THE MESSAGE FORUM: </span>&nbsp;
Before you begin editing, moving, or deleting message posts, you will need to first know how to find them.  First, 
enter the message forum by clicking either Message Forum in the section menu or Modify Messages in the member menu.  Now 
you may use one of the two main ways for finding a specific message post: the search method or the browse method.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Search Method - </span>
	Instead of manually browsing each topic to find your desired message post, you may want to use the search 
	feature at the top of the topic list page.  If you know a keyword(s) that appears in the desired post's subject 
	or author data field, type it in the Search For: textbox.  Click the Go button to the textbox's right.  The 
	search results page displayed may contain more than one match.  The items most likely to be a perfect match 
	with your search keyword(s), will appear at the beginning of the search results list.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Browse Method - </span>
	The browse method is the more inexact of the two navigation methods.  Basically, you use this method when you 
	have an idea where your message post is but you're not quite sure what it's called.  The browse method narrows 
	down your post list to a smaller number of items so you don't have to look through every single message post 
	in the forum.
	</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	You must first know the forum topic that the desired message is in.  Enter the appropriate topic from the forum 
	topic list.  If the message you're looking for is a thread, you must first find its corresponding head post.  To 
	find the desired head post by navigating through the pages.  Or, if the message has been added 
	recently, we recommend using the Latest Additions section on the Home page to find it.
	</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	If the message post you're looking for is a thread but is not visible below its head, it means the message 
	unit is collapsed.  Click on the plus symbol to the heads left and the message unit will expand.  Your desired 
	thread post should now be visible.
	</p>

	<p align=left class=BodyText><b>Note: </b> 
	Anytime while navigating the message forum, you can click on the name of the message post.  This will allow 
	you to view the message body.
	</p>

<a name="5"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.5 ALTERING YOUR EXISTING MESSAGE POSTS: </span>&nbsp;
You can easily modify one of your previously created messages.  However, if the below options are not available to you, 
it's because the administrator has disabled this option.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing a Message Post - </span>
	When you find it necessary to eit the content of a message post, use the following procedure:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Messages link under the Message Forum subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", "find message browse method.gif"%>: </i></b> 
		Find the desired post.
		<br>
		<b><i>4 <%PrintSymb "Edit", "modifying message.gif"%>: </i></b> 
		Make the necessary changes.
		<br>
		<b><i>5: </i></b> 
		Click the Update button below the body textbox.
		<br>	
		<b><i>6 <%PrintSymb "Confirmation", "message has been edited.gif"%>: </i></b> 
		Use the links to return to the member menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Message Post - </span>
	In case you decide to delete a message you recently wrote:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Messages link under the Message Forum subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", "find message browse method.gif"%>: </i></b> 
		Find the desired post.
		<br>	
		<b><i>3 <%PrintSymb "Delete", "none"%>: </i></b> 
		Click the Delete button.
		<br>
		<b><i>4 <%PrintSymb "PopUp", "delete message warning box.gif"%>: </i></b> 
		Click the OK button if you're sure.  If not, click Cancel.
		<br>
		<b><i>5: </i></b> 
		Use the links to return to the member menu.
	</blockquote>

