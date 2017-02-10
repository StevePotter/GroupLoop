<% intChapter = 13 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: THE MEDIA SECTION</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 THE ADMINISTRATOR AND MEDIA: </span>&nbsp;
The media section is a great way for users to share movies, sounds, Word documents, songs, etc.  All types of files can 
be uploaded, except executable files (for security's sake).  Customers always find this a fun, useful section.  The 
media section is structured much like the photos section because the items are categorized. 

	<img src="../images/structure_media.gif">

<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 MEDIA CATEGORIES: </span>&nbsp;
Media files are grouped into categories.  Each category has two properties or data fields:

	<blockquote>
	<b><i>Privacy Status: </i></b>
	When a media category is marked private, only members can enter and view/download its files.
	<br>
	<b><i>Name: </i></b>
	The second field is simply the name of the category.
	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Creating a New Media Category - </span>
	Manipulating categories is pretty straightforward.  To create a new category:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click Create a New Category under the subheading Media.
		<br>	
		<b><i>2 <%PrintSymb "Create", ""%>: </i></b> 
		Make the proper data field selections and entries (see above).
		<br>
		<b><i>3: </i></b> 
		Click the Add button.
		<br>
		<b><i>4 <%PrintSymb "Confirmation", ""%>: </i></b> 
		Using the supplied links, either add another or return to the main administrator page.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing a Media Category - </span>
	Altering media categories is just as easy as creating them.  To edit the properties of a media category:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click Modify Categories under the Media subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", ""%>: </i></b> 
		Browse the category list and find the desired category.
		<blockquote>
			<%PrintArrow%><b><i>Category Link: </i></b> 
			Clicking on the name of a category will take you to that category's file list.  Use this feature to 
			double check your category selection.
		</blockquote>
		<b><i>3: </i></b> 
		Click the Edit button next to desired category .
		<br>
		<b><i>4 <%PrintSymb "Edit", ""%>: </i></b> 
		Make the necessary changes and click the Update button.
		<br>
		<b><i>5 <%PrintSymb "Confirmation", ""%>: </i></b> 
		Using the links, reload either the category list or the admin menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Media Category - </span>
	When you delete a media category, remember that you are also deleting <b>all the files contained 
	therein</b>.  To remove a media category entirely:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click Modify Categories under the Media subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", ""%>: </i></b> 
		Browse the category list and find the desired category.
		<blockquote>
			<%PrintArrow%><b><i>Category Link: </i></b> 
			Clicking on the name of a category will take you to that category's file list.  Use this feature to 
			double check your category selection.
		</blockquote>
		<b><i>3 <%PrintSymb "Delete", ""%>: </i></b> 
		Click the Delete button to the right of the desired topic.
		<br>
		<b><i>4 <%PrintSymb "PopUp", ""%>: </i></b> 
		If you're sure, click the OK button.  If not, click Cancel.
		<br>	
		<b><i>5 <%PrintSymb "Confirmation", ""%>: </i></b> 
		Using the links, reload either the category list or the admin menu.

	</blockquote>

<a name="3"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.3 MEDIA FILES: </span>&nbsp;
A file is made up of six properties:

	<blockquote>

		<b><i>Date: </i></b>
		This field is simply the date the file is posted.  Since files are organized within categories 
		chronologically (almost everything on your site is), the date property dictates their order.  Also, if a 
		file has a date within a certain amount (exact number can be changed) of days from today's current date, 
		it will appear on the home page's latest additions.
		<br>

		<b><i>Category: </i></b>
		This field is simply the category where the file is located.
		<br>

		<b><i>Description: </i></b>
		The description field is what is sounds like, a description of the file.  It is a good idea to always 
		leave a description of the file so people know what they are downloading.
		<br>

		<b><i>File: </i></b>
		This is the actual file saved on your computer.
		<br>

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Adding a New File - </span>
	Once you have a file to add:
	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click Add a New File under the Media subheading.
		<br>	
		<b><i>2 <%PrintSymb "Create", ""%>: </i></b> 
		Assign the appropriate properties to the media file.
		<br>
		<b><i>3: </i></b> 
		Click the Add (click once) button one time.
		<blockquote>
			<%PrintArrow%><b><i>Click ONCE: </i></b> 
			After you click the Add (click once) button, it may take a while for the next page to load.  This is 
			because the file you named has to upload from your computer onto our server.  Each time you click 
			the button, the upload process has to start from the beginning.  Point being: <b>only click once</b>.
		</blockquote>	
		<b><i>4 <%PrintSymb "Confirmation", ""%>: </i></b> 
		Using the links, add another or return to the admin menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing an Existing File - </span>
	To edit the properties of an existing file, use the following procedure:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Files link under the Media subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", ""%>: </i></b> 
		Find the desired file using the browse or search method. 
		<br>
		<b><i>3 <%PrintSymb "List", ""%>: </i></b> 
		Click the Edit button next to the desired file.
		<br>
		<b><i>4 <%PrintSymb "Edit", ""%>: </i></b> 
		Make the necessary changes to the file's properties.
		<blockquote>
			<%PrintArrow%><b><i>Empty Textbox: </i></b> 
			When you first enter the edit page, the file textbox is <b>supposed to be empty</b>.  Don't 
			worry, we didn't lose your file; it's still saved on our server. Only type information 
			into the box if you wish to overwrite the old file with a new one.  Otherwise, <b>leave it blank.</b>
		</blockquote>

		<b><i>5: </i></b> 
		Click the Update (click once) button one time.
		<blockquote>
			<%PrintArrow%><b><i>Click ONCE: </i></b> 
			If you changed the file, the confirmation page may take a while to load.  Every time you click the 
			button, the upload process has to restart.  Do yourself a favor and <b>only click once</b>.
		</blockquote>
		<br>	
		<b><i>6 <%PrintSymb "Confirmation", ""%>: </i></b> 
		Using the links provided, reload the category list or return to the administrator menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a File - </span>
	If instead of editing, you wish to remove a file entirely from your site, you must:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Files link under the Media subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", ""%>: </i></b> 
		Find the desired file using the browse or search method. 
		<br>
		<b><i>3 <%PrintSymb "List", ""%>: </i></b> 
		Click the Delete button next to the desired file.
		<br>
		<b><i>4 <%PrintSymb "PopUp", ""%>: </i></b> 
		If you're sure, click the Yes button.  To abort, click No.
		<br>	
		<b><i>5 <%PrintSymb "Confirmation", ""%>: </i></b> 
		Using the links provided, reload to the category list or return to the administrator menu.
	</blockquote>