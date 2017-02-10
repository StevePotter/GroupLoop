<% intChapter = 10 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: THE MEDIA SECTION</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 OVERVIEW: </span>&nbsp;
The media section is a great way for users to share movies, sounds, Word documents, songs, etc.  All types of files can 
be uploaded, except executable files (for security's sake).  Customers always find this a fun, useful section.  The 
media section is structured much like the photos section because the items are categorized. 

	<img src="../images/structure_media.gif">

	<p align=left class=BodyText><b>Note: </b> 
	Remember that the member can restrict things in each section.  They may choose to not allow you to upload 
	files.  So if the option isn't there, we didn't do it.
	</p>

<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 MEDIA CATEGORIES: </span>&nbsp;
Media files are grouped into categories.  The categories can only be created by administrators.

	<blockquote>
	<b><i>Privacy Status: </i></b>
	When a media category is marked private, only members can enter and view/download its files.
	<br>
	<b><i>Name: </i></b>
	The second field is simply the name of the category.
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
		<b><i>2 <%PrintSymb "Create", "none"%>: </i></b> 
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
		<b><i>4 <%PrintSymb "Confirmation", "none"%>: </i></b> 
		Using the links, add another or return to the member menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing an Existing File - </span>
	To edit the properties of an existing file, use the following procedure:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Files link under the Media subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", "none"%>: </i></b> 
		Find the desired file using the browse or search method. 
		<br>
		<b><i>3 <%PrintSymb "List", "none"%>: </i></b> 
		Click the Edit button next to the desired file.
		<br>
		<b><i>4 <%PrintSymb "Edit", "none"%>: </i></b> 
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
		<b><i>6 <%PrintSymb "Confirmation", "none"%>: </i></b> 
		Using the links provided, reload the category list or return to the member menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a File - </span>
	If instead of editing, you wish to remove a file entirely from your site, you must:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Files link under the Media subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", "none"%>: </i></b> 
		Find the desired file using the browse or search method. 
		<br>
		<b><i>3 <%PrintSymb "List", "none"%>: </i></b> 
		Click the Delete button next to the desired file.
		<br>
		<b><i>4 <%PrintSymb "PopUp", "none"%>: </i></b> 
		If you're sure, click the Yes button.  To abort, click No.
		<br>	
		<b><i>5 <%PrintSymb "Confirmation", "none"%>: </i></b> 
		Using the links provided, reload to the category list or return to the member menu.
	</blockquote>