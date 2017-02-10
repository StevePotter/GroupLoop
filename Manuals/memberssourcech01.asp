<% intChapter = 1 %>

<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: OVERVIEW</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 WHAT IS GROUPLOOP.COM?: </span>&nbsp;
If you are reading this, you are probably new to being a member, and this may be the first time 
you've seen or even heard of GroupLoop.com.  So here's an overview.
</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	A few years ago a site was made for a group of friends.  Everyone in and around the group was loving the site, so 
	the designer figured out a way to create a site just like it for any other group of people.  A while later, GroupLoop.com 
	was released.  The response was better than anyone could have imagined.  The sites could be created instantly, and 
	some sites had over a hundred additions the first day!  Since then nobody has been able 
	to offer the power, flexibility, convienence, and price that GroupLoop.com prides itself in.  We are looking forward 
	to you falling in love with your site like so many already have.
	</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Your site is not just a regular web page.  It is what we call an Interactive Community.  As a member, you can 
	participate in many different sections.  You can add stories, announcements, quizzes, voting polls, calendar events, 
	photos, files, links, quotes, and participate in a powerful message forum.  You can choose to make your additions 
	public for anyone to view, or just for members' eyes.  Your site's look and feel can be completely customized.  
	This may seem a bit overwhelming, but don't be scared.  Most people get the hang after an hour or so on the site, 
	and never look back.
	</p>


<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 WHAT IS AN ADMINISTRATOR: </span>&nbsp;
As a member of your particular site, you have certain privalages.  You can edit and delete everything you post, but 
you can't change what other members do.  An administrator can.
</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	An administrator is a member just like you, but has additional privalages. They can change any member's addition, 
	change the way the site looks, configure the site's sections, and modify members.  In fact, you were added by an 
	administrator, and if they so chose, they could remove you.  Just remember that if a problem arises that you can't 
	solve, the administrator can.
	</p>


<a name="3"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.3 THIS MANUAL: </span>&nbsp;
In this manual, we do our best to ease you into your role as member.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>The Layout - </span>
	Each chapter of the Member's Manual focuses on a specific section(s) of your site.  The 
	chapter's first subchapter (designated by a decimal and all caps) typically explains your 
	responsibilities for that particular site section. 
	</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Each additional subchapter usually consists of either definitions and/or several procedures.  
	The definitions describe features to be used in the upcoming procedures.  Each procedure is a 
	numbered listing of steps and diagrams that instruct you on how to complete a specific task.
	</p>

	<p align=left class=BodyText>
	<b>Note:</b> There are often breaks between steps (signified by <%PrintArrow%>).  These procedure 
	breaks are used when a step has multiple options, each needing its own explanation.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Page Types - </span>
	Each step of a procedure consists of first, completing a task on one webpage and then, loading a 
	new page for the next step (unless otherwise specified).  Since your site has the possibility of 
	having an infinite number of individual webpages, we have decided not to label each one separately 
	in this manual.  Instead, we have grouped all of your pages into types.
	</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	At the start of each step there is a symbol referring to the specific page type where that step begins.  
	If you click on the page symbol, this manual will automatically direct you to a diagram depicting the actual 
	page where that step takes place.  There are ten page types all with their own corresponding symbols:
	</p>

	<blockquote>

		<p align=left class=BodyText>
		<b><i>Addition:</i></b> 
		When you direct your web browser to your new site, the home page is the first webpage you will see.  There 
		is only one home page and it will be far and away your most commonly visited page.  Your home page will have 
		your site title, links to all your section areas, news updates, and a listing of the site's most recent member 
		and guest additions.  Within a procedure, the home page is symbolized by <%PrintSymb "Home", ""%>.  (The home page 
		is described in much greater detail in the Member's Manual.  PLEASE READ THAT FIRST).
		</p>

		<p align=left class=BodyText>
		<b><i>The Main Member Page:</i></b> 
		This page can only be accessed by site members because it requires a sign-in process.  The main member page 
		has a list of options called the member menu.  You will be accessing this section a lot, so take the time to 
		get to know it well.  The main member page is symbolized by <%PrintSymb "Member", ""%>.
		</p>

		<p align=left class=BodyText>
		<b><i>Viewing Pages:</i></b> 
		A viewing page is just that: a page for viewing a particular member or guest post.  Every single item on your 
		site has its own viewing page.  These pages are symbolized by <%PrintSymb "Viewing", "none"%>.
		</p>

		<p align=left class=BodyText>
		<b><i>Property Pages:</i></b> 
		Each item on your site has a specific set of properties.  When a site-goer creates a new item, he/she must define 
		those properties.  Furthermore, when a site-goer wishes to edit an existing item (that he/she created originally), 
		he/she must change initial properties.  The pages that allow one to manipulate an item's properties are called 
		property pages and have the symbol <%PrintSymb "Property", "none"%>.
		</p>

		<p align=left class=BodyText>
		<b><i>Confirmation Pages:</i></b> 
		These pages load after a site-goer has completed an action on the site.  For example, when a guest adds a 
		message post, a confirmation page will load to verify that his/her message has been added.  The symbol for 
		confirmation pages is <%PrintSymb "Confirmation", "none"%>.
		</p>

		<p align=left class=BodyText>
		<b><i>List Pages:</i></b> 
		A list page is just a chronological catalogue of items such as posts, photos, categories, members, etc.  
		When a site-goer first enters a section area, he/she is shown a list page that contains all of that section's 
		items.  The symbol for a list page is <%PrintSymb "List", "none"%>.
		</p>


		<p align=left class=BodyText>
		<b><i>Search Result Pages:</i></b> 
		Similar to a list page, this type of page is an index of all items returned by a search.  The items of a 
		search result page however, are ordered slightly differently than a list page.  Items most closely matching 
		your search request are listed first.  These pages are symbolized by <%PrintSymb "Search", "none"%>.
		</p>

		<p align=left class=BodyText>
		<b><i>Log-In Page:</i></b> 
		This page will load when a site-goer attempts to reach a section or perform an action that is designated 
		members only.  The symbol for the log-in page is <%PrintSymb "Login", ""%>.
		</p>

	</blockquote>


	<p align=left class=BodyText>
	<b>Note:</b> 
	Warning or notice boxes pop up on your screen after you click certain action buttons.  These are not considered 
	true webpages but still have a symbol: <%PrintSymb "Popup", "none"%>.  Similarly, clicking a <b>Browse</b> button 
	(you will learn about these later) will open a directory box denoted by <%PrintSymb "Browse", "none"%>.
	</p>
