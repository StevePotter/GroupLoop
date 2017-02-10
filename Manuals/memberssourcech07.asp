<% intChapter = 7 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: THE PHOTOS SECTION</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 PHOTOS - AN OVERVIEW: </span>&nbsp;
The photo section is consistently one of GroupLoop's most asked about and most popular sections.  It's pretty simple.  
Photos are organized into categories (made by the administrator).  Photos are uploaded by members.  You then may  
rate the photos, write captions for each photo, and rate/review the captions.  Some people spend the majority of their 
time on the site writing funny captions.

	<img src="../images/structure_photos.gif">

	<p align=left class=BodyText><b>Note: </b> 
	Remember that the administrator can restrict things in each section.  They may choose to not allow you to upload 
	photos, rate photos, or even write captions.  Please understand that we do not set these restrictions.  It's the 
	administrator.  So please direct angry e-mails to them.  Thanks!
	</p>


<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 PHOTO CATEGORIES: </span>&nbsp;
Online photos are grouped into categories.  Only the administrator can create/modify them.  Each category has two properties or data fields:

	<blockquote>
	<b><i>Privacy Status: </i></b>
	When a photo category is marked private, only members can enter and view its photos and captions.
	<br>
	<b><i>Name: </i></b>
	The second field is simply the name of the category.
	</blockquote>


<a name="3"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.3 PHOTOS: </span>&nbsp;
A site photo is made up of six properties:

	<blockquote>

		<b><i>Category: </i></b>
		This field is simply the category where the photo is located.
		<br>

		<b><i>Description: </i></b>
		The description field is a short phrase used to label a photo.  This field is important for finding a 
		particular photo when you don't have a thumbnail for it (see below).
		<br>

		<b><i>Photo Image File: </i></b>
		This is the actual photo file saved on your computer.
		<br>

		<b><i>Thumbnail: </i></b>
		A thumbnail is a minimized version of your photo.  This small version allows people to get a quick preview of 
		the photo.  Although not required, thumbnails make navigating your photo section much easier.  Don't worry about 
		creating thumbnails, it is done automatically.  <b>Please note:</b> if you are uploading this to a private 
		section, we recommend <b>not</b> creating a thumbnail.  Although the photo can't be viewed elsewhere, the 
		thumbnail can.
		<br>

		<b><i>Captions: </i></b>
		Captions can be added to photos by members.  Since each photo caption has its own sub-properties, we save most information 
		concerning them for the next subchapter.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Adding a New Photo - </span>
	If you have access to a scanner and can convert physical photographs into digital image files (saved to disk), you can easily add them to 
	the photo section yourself.  To add a photo to your photo section:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click Add a new photo under the Photos subheading.
		<br>	
		<b><i>2 <%PrintSymb "Create", "adding photo.gif"%>: </i></b> 
		Assign the appropriate properties to the photo file.
		<blockquote>
			<%PrintArrow%><b><i>Description: </i></b> 
			Remember this is only required if you choose not to create a thumbnail.  If you do have a thumbnail, you 
			can still have a description.  It will show up right under the photo when people view it!
			<br>
			<%PrintArrow%><b><i>Image Files: </i></b> 
			If you know the exact name (dos directory protocol) of the image file (photo or thumbnail) on your hard-drive 
			or disk, you can type it in the textbox provided.  Otherwise, click the Browse... button to search for the 
			file.  A directory box will pop-up.  Once you have found the desired image file, select it.  The proper 
			label will then appear in the appropriate textbox.
			<br>
			<%PrintArrow%><b><i>Thumbnail: </i></b> 
			As explained above, a thumbnail can automatically be created of your photo.  This is recommended by us, since 
			people will get a visual preview instead of a description (yes, pictures say a thousand words).
			<br>
			<%PrintArrow%><b><i>Optimization: </i></b> 
			Your photo can automatically be optimized (because we love and care about you).  Many times when you scan an 
			image, it may be very big, which will make many people with smaller screen sizes unable to see the whole thing.  Also the acutal download size of 
			the photo may be unnecessarily large, which will cause people to wait very long.  When you select to have the 
			photo optimized, this is all fixed.  The photo is set at a size that people with all screen sizes can view, 
			and the best possible download time.  Needless to say, we strongly recommend this.
		</blockquote>
		<b><i>5 <%PrintSymb "Popup", "add photo warning box.gif"%>: </i></b> 
		Click the Add (click once) button one time.
		<blockquote>
			<%PrintArrow%><b><i>Click ONCE: </i></b> 
			After you click the Add (click once) button, it may take a while for the next page to load.  This is 
			because the file you named has to upload from your computer onto our server.  Each time you click 
			the button, the upload process has to start from the beginning.  Point being: <b>only click once</b>.
		</blockquote>	
		<b><i>6 <%PrintSymb "Confirmation", "photo added.gif"%>: </i></b> 
		Using the links, add another or return to the admin menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing an Existing Photo (Excluding Captions) - </span>
	To edit the properties (excluding captions) of an existing photo, use the following procedure:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Photos link under the Photos subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", "modify choose photo.gif"%>: </i></b> 
		Find the desired photo using the browse or search method. 
		<br>
		<b><i>3 <%PrintSymb "Edit", "edit photo.gif"%>: </i></b> 
		Make the necessary changes to the photo's properties.
		<blockquote>
			<%PrintArrow%><b><i>Empty Textbox: </i></b> 
			When you first enter the photo edit page, the photo file textbox is <b>supposed to be empty</b>.  Don't 
			worry, we didn't lose your photo; it's still saved on our server. Only type information 
			into the box if you wish to overwrite the photo with a new file.  Otherwise, <b>leave it blank.</b>
		</blockquote>

		<b><i>4: </i></b> 
		Click the Update (click once) button one time.
		<blockquote>
			<%PrintArrow%><b><i>Click ONCE: </i></b> 
			If you changed the photo file, the confirmation page may take a while to load.  Every time you click the 
			button, the upload process has to restart.  Do yourself a favor and <b>only click once</b>.
		</blockquote>
		<br>	
		<b><i>5 <%PrintSymb "Confirmation", "photo has been edited.gif"%>: </i></b> 
		Using the links provided, reload the category list or return to the administrator menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Photo - </span>
	If instead of editing, you wish to remove a photo entirely from your site, you must:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Photos link under the Photos subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", "modify choose photo.gif"%>: </i></b> 
		Find the desired photo using the browse or search method. 
		<br>
		<b><i>3: </i></b> 
		Click the Delete Photo button below the photo's thumbnail/description.
		<br>
		<b><i>4 <%PrintSymb "PopUp", "delete photo warning box.gif"%>: </i></b> 
		If you're sure, click the Yes button.  To abort, click No.
		<br>	
		<b><i>5 <%PrintSymb "Confirmation", "photo has been deleted.gif"%>: </i></b> 
		Using the links provided, reload to the category list or return to the administrator menu.
	</blockquote>

<a name="4"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.4 PHOTO CAPTIONS: </span>&nbsp;
Members have the ability to insert captions that appear below a photo.  A photo caption has four sub-properties:

	<blockquote>

		<b><i>Author: </i></b>
		This is simply the member who created the photo caption.  You as administrator have no access to this data field.  
		Once a member authors a caption, it is automatically affixed with his/her nickname.
		<br>

		<b><i>Privacy Status: </i></b>
		When a caption is marked private, only members can read it.  Captions for photos contained in a private category are automatically marked private.
		<br>

		<b><i>Body: </i></b>
		The body is the actual comment the member makes about photo.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Adding a Photo Caption - </span>
	Adding photo captions is very simple.

	<blockquote>

		<b><i>1: </i></b> 
		View the photo you would like to add a caption to.
		<br>	
		<b><i>2: </i></b> 
		Click the 'Add a Caption' link below the photo.  If this link is not there, the administrator has disabled 
		photo captions.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing a Photo Caption - </span>
	Editing photo captions is very similar to editing photo properties.  Use the following procedure to edit a photo 
	caption:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Photos link under the Photos subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", "choose photo with caption to modify.gif"%>: </i></b> 
		Find the photo containing the desired caption using the browse or search method.
		<br>
		<b><i>3: </i></b> 
		Click the Modify Captions button below the photo's thumbnail/description.
		<br>
		<b><i>4 <%PrintSymb "List", "choose edit or delete caption.gif"%>: </i></b> 
		Find the desired caption and click the Edit button to its right.
		<br>
		<b><i>5 <%PrintSymb "Edit", "edit caption.gif"%>: </i></b> 
		Make the necessary change and click the Update button.
		<br>
		<b><i>6 <%PrintSymb "Confirmation", "caption edited.gif"%>: </i></b> 
		Using the links provided, reload the category list or return to the administrator menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Photo Caption - </span>
	To completely remove a caption:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b> 
		Click the Modify Photos link under the Photos subheading.
		<br>	
		<b><i>2 <%PrintSymb "List", "choose photo with caption to modify.gif"%>: </i></b> 
		Find the photo containing the desired caption using the browse or search method.
		<br>
		<b><i>3: </i></b> 
		Click the Modify Captions button below the photo's thumbnail/description.
		<br>
		<b><i>4 <%PrintSymb "List", "choose edit or delete caption.gif"%>: </i></b> 
		Find the desired caption and click the Delete button to its right.
		<br>
		<b><i>5 <%PrintSymb "PopUp", "delete caption warning box.gif"%>: </i></b> 
		If you're sure, click the OK button.  To abort, click Cancel.
		<br>
		<b><i>6 <%PrintSymb "Confirmation", "caption has been deleted.gif"%>: </i></b> 
		Either reload the category list or return to the admin menu.

	</blockquote>
