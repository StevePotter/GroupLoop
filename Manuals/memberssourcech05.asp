<% intChapter = 5 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: BASIC POST SECTIONS - ANNOUNCEMENTS, STORIES, QUOTES, LINKS THE CALENDAR AND THE GUESTBOOK </p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 BASIC POST PROPERTIES: </span>&nbsp;
Announcements, stories, quotes, links, calendar events and guestbook entries are all considered basic posts because 
each respective section is simply a chronological listing (newest to oldest) of individual items called a post list.  
Furthermore, all basic posts consist of up to five basic data fields: a <a href="#"><b>date</b></a>, 
<a href="#"><b>author</b></a>, <a href="#"><b>subject</b></a>, <a href="#"><b>body</b></a> and up to two 
of the following special properties:
</p>

	<p align=left class=BodyText>
	<b><i>Privacy: </i></b>
	This special property is specific to those posts that can be created exclusively by members.  A post marked 
	private cannot be viewed by site guests.
	</p>

	<p align=left class=BodyText>
	<b><i>Starting/Ending Date: </i></b>
	This special property is unique to calendar events.  It allows an author to specify multiple days for an event.  A 
	calendar event is posted in every date box between and including the starting and ending dates.
	</p>

	<p align=left class=BodyText>
	<b><i>Extra Description: </i></b>
	The extra description property is simply an additional textbox data field for information that belongs in 
	neither the subject/author field nor the body field.
	</p>


<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 CREATING BASIC POSTS: </span>&nbsp;
Creating Basic Posts is the easiest thing on the site.  
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Creating an Existing Basic Post - </span>
	To avoid the redundancy, we have combined the individual procedures of each basic post section into one 
	universal procedure.  The only difference between a specific section's procedure is the explicit wording 
	of the particular section.  Therefore, in the following steps, we will substitute the exact basic post 
	wording with the symbol <%PrintSymb "BasicPost", "none"%>.  For instance, if you wish to add an announcement, use the following 
	procedure replacing every <%PrintSymb "BasicPost", "none"%> symbol with the word announcement(s):
	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		There are two links on the site to create a basic post.<br>
		1. Under <%PrintSymb "BasicPost", "none"%> subheading, click the Add <%PrintSymb "BasicPost", "none"%> link.
		<br>&nbsp;&nbsp;&nbsp;&nbsp;<b>OR</b>
		1. In the <%PrintSymb "BasicPost", "none"%> section , click the Add <%PrintSymb "BasicPost", "none"%> link, 
		right below the title.
		<br>	
		<b><i>2 <%PrintSymb "Create", "none"%>: </i></b> 
		Type int your new post and click the Add button.
		<br>
		<b><i>3: </i></b> 
		Use the links to add another or to return to the member menu.
	</blockquote>


<a name="3"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.3 ALTERING BASIC POSTS: </span>&nbsp;
You can edit and delete your existing posts rather easily.
</p>


	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing an Existing Basic Post - </span>
	To avoid the redundancy, we have combined the individual procedures of each basic post section into one 
	universal procedure.  The only difference between a specific section's procedure is the explicit wording 
	of the particular section.  Therefore, in the following steps, we will substitute the exact basic post 
	wording with the symbol <%PrintSymb "BasicPost", "none"%>.  For instance, if you wish to edit an announcement, use the following 
	procedure replacing every <%PrintSymb "BasicPost", "none"%> symbol with the word announcement(s):

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under <%PrintSymb "BasicPost", "none"%> subheading, click the Modify <%PrintSymb "BasicPost", "none"%> link.
		<br>	
		<b><i>2 <%PrintSymb "List", "modify basicpost.gif"%>: </i></b> 
		Find your desired <%PrintSymb "BasicPost", "none"%> file using the browse or search method.
		<br>
		<b><i>3: </i></b> 
		Click the Edit button to the desired post's right.
		<br>
		<b><i>4 <%PrintSymb "Edit", "modify page.gif"%>: </i></b> 
		Make the necessary changes.
		<blockquote>
			<%PrintArrow%><b><i>Dates (excluding calendar events): </i></b> 
			If you wish for the modified post to appear in the latest additions, change the date field to today's date.
		</blockquote>
		<b><i>5: </i></b> 
		Click the Update button below the body textbox.
		<br>
		<b><i>6: </i></b> 
		Using the links provided, either reload the <%PrintSymb "BasicPost", "none"%> list page, or return to the main member page.
	</blockquote>


	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Basic Post - </span>
	When you wish to remove a basic post entirely from your site:
	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under <%PrintSymb "BasicPost", "none"%> subheading, click the Modify <%PrintSymb "BasicPost", "none"%> link.
		<br>	
		<b><i>2 <%PrintSymb "List", "modify basicpost.gif"%>: </i></b> 
		Find your desired <%PrintSymb "BasicPost", "none"%> file using the browse or search method.
		<br>
		<b><i>3 <%PrintSymb "Delete", "none"%>: </i></b> 
		Click the Delete button to the desired post's right.
		<br>	
		<b><i>4 <%PrintSymb "Confirmation", "basicpost has been deleted.gif"%>: </i></b> 
		If you're sure, click the OK button.  If not, click Cancel.
		<br>	
		<b><i>5: </i></b> 
		Using the links provided, either reload the <%PrintSymb "BasicPost", "none"%> list page, or return to the main member page.

	</blockquote>
