<% intChapter = 3 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: NAVIGATION</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 SECTORS AND SECTIONS: </span>&nbsp;
Your entire site consists of a series of webpages.  Every single one of these hundreds 
(and possibly thousands) of pages have four distinct sectors: the title, the section menu, 
the body, and the footer.  The title, section menu, and footer sectors appear identically on 
each individual page; the only variation comes in the body sector.  By default, the title 
sector is at the page ceiling, the section menu runs down the left wall, the body sector runs 
down the right wall, and the footer runs along the page floor.
</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Each one of these pages is contained within a specific section (such as Announcements, Quizzes, and 
	the Message Forum).  These sections are the backbone of your site and contain specific formats and 
	site items.  You will learn much more about the individual items and sections later on in this manual.
	</p>


<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 NAVIGATING BASIC POST LISTS: </span>&nbsp;
Basic posts are additions in the most basic sections: Announcements, Stories, Links, Quotes, and the Guestbook. 
Every item addition made to your site is saved and catalogued within its corresponding section.  
As the months progress, you will begin to notice the enormous volume of posts contained in each 
section.  Before editing or deleting an existing post, you will have to know how to find it. 
There a two ways to navigate a basic post list: the browse method and the search method.  These are very simple, 
and chances are you have already used them.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Browse Method - </span>
	The browse method is the more inexact of the two navigation methods.  
	Basically, you use this method to see everyone's additions or when you have an idea where your desired post 
	is but not exactly.  The browse method narrows down your post list to a smaller number of items so you don't 
	have to look through every single item in the section.  There are two features involved in the browse method:	
	<blockquote>
	<p align=left class=BodyText>
	<b><i>1 - Item Dates:</i></b>  
	If you think your desired post was added to your site recently, you may want to use the recent additions 
	feature at the top of the post list page.  Use the View... pull down menu to select the date at which you 
	wish your post list to begin.  Then, click the Go button.  The new post list will only contain items added 
	to your site within the specified number of days.
	</p>
	<p align=left class=BodyText>
	<b><i>2 - Multiple Pages:</i></b>  
	If a section contains more items than the fixed page maximum number, the post list will automatically create 
	additional pages to hold the extra items (default maximum items per page is 40 but, it can be changed in section 
	customization: see Chapter 5).  To navigate multiple pages either, use the next and previous buttons, or 
	enter an exact page number and click the Go button.
	</p>

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Search Method - </span>
	Instead of manually browsing the post lists to look for your desired post, you may want to use the search feature 
	at the top of the page.  If you know a keyword(s) that appears in the desired post's subject, author, 
	or extra description data field, type it in the Search For: textbox.  Click the Go button to the 
	textbox's right.  The search results page displayed may contain more than one match.  The items most 
	likely to be a perfect match with your search keyword(s), will appear at the beginning of the search results list.
	</p>

<a name="3"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.3 NAVIGATING THE CALENDAR: </span>&nbsp;
The calendar uses the browse and search methods discussed above.  The search method is the exact same, so it will not 
be discussed.  However, the browse method is a bit different.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Browse Method In The Calendar - </span>
	When you click on the Calendar menu button, you are automatically viewing events for this month.  Instead of pages, 
	the calendar items are listed per month.  You can naviage through the months using the Next and Last month buttons, 
	or by specifying a month.
	</p>

<a name="4"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.4 NAVIGATING THE MESSAGE FORUM: </span>&nbsp;
The message forum is much like the basic posts, except there are topics and the messages have replies.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Search Method In The Message Forum - </span>
	When you click on the Message Forum menu button, you are given a menu of the current topics.  To search through all 
	topics at once, simple enter your keywords in the search box above the topic list.  Click 'Go' and wait for your results.  
	You may search just the messages in a certain topic by entering it and then using the search box given.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Browse Method In The Message Forum - </span>
	When you click on the Message Forum menu button, you are given a menu of the current topics.  Click on the desired 
	topic, and you will be taken to its messages.  If a message has replies, you can click on the Plus sign next to it to 
	view it's replies.  Just like the basic posts, there may be multiple pages of results.
	</p>

<a name="5"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.5 NAVIGATING THE PHOTOS SECTION: </span>&nbsp;
The photos section is structurally pretty simple.  Photos are simply put into their respective category.  Something 
that makes the photos section unique is the way the items are layed out.  Each item in other lists have their own row.  
In the photos section, there rows are also divided into columns, each photo in its own column.  The default number 
of photos per row is 5, although that can be changed by you.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Search Method In The Photos Section - </span>
	When you click on the Photos menu button, you are given a menu of the current categories.  To search through all 
	categories at once, simple enter your keywords in the search box above the category list.  Click 'Go' and wait for your results.  
	You may search just the photos in a certain category by entering it and then using the search box given.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Browse Method In The Photos Section - </span>
	When you click on the Photos menu button, you are given a menu of the current categories.  Click on the desired 
	category, and you will be taken to its photos.  Just like the basic posts, there may be multiple pages of results.
	</p>
