<% intChapter = 7 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: NEWS UPDATES</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 USING NEWS: </span>&nbsp;
The news update section is intended to inform members and guests alike of the newest happenings on the site.  News 
updates are the only type of posts that are fully displayed omnn the home page.  In other words, a visitor is able to 
read the news updates from the home page without having to click a link to a specific section; the entire news 
section is displayed on the home page.  News is also the first information site-goers are going to see (just below 
the title of your site).  Therefore, it's a good idea to always keep changing the news updates to keep your site 
active and exciting.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Creating a News Update - </span>
	Each news update is made up of only two data fields: the date and the body.  The simplicity of news updates makes 
	manipulating them pretty straight forward.  To add a news update to your home page: 
	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under the News subheading, click on Add A News Update.
		<br>	
		<b><i>2 <%PrintSymb "Create", "add news update.gif"%>: </i></b> 
		Type your news in the textbox and click the Add button.
		<br>
		<b><i>3: </i></b> 
		Use the links to add another or to return to the admin menu.

	</blockquote>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Your new news update should now appear on the home page with the current date and it will stay there until you remove it.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing a Current News Update - </span>
	If you wish to edit a news update: 
	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under the News subheading, click the Modify News link.
		<br>	
		<b><i>2 <%PrintSymb "List", "modify news.gif"%>: </i></b> 
		Find the desired update and click the Edit button to its right.
		<br>
		<b><i>3 <%PrintSymb "Edit", "modify news2.gif"%>: </i></b> 
		Make the necessary changes to the date and/or body.
		<br>	
		<b><i>4: </i></b> 
		Click the Update button beneath the body textbox.
		<br>	
		<b><i>5 <%PrintSymb "Confirmation", "news has been edited.gif"%>: </i></b> 
		Using the links provided, either reload the news update list page or return to the admin menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Current News Update - </span>
	Finally, to permanently remove a news update from the home page:
	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Under the News subheading, click the Modify News link.
		<br>	
		<b><i>2 <%PrintSymb "List", "modify news.gif"%>: </i></b> 
		Find the desired update and click the Delete button to its right.
		<br>
		<b><i>3 <%PrintSymb "PopUp", "delete news warning box.gif"%>: </i></b> 
		If you're sure, click the OK button.  If not, click Cancel.
		<br>	
		<b><i>4 <%PrintSymb "Confirmation", "news has been deleted.gif"%>: </i></b> 
		Using the links provided, either reload the news update list page or return to the admin menu.

	</blockquote>