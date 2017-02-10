<% intChapter = 5 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: VISUAL CUSTOMIZATION</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 BASIC VISUALS: </span>&nbsp;
Nearly every property of your site's appearance can be altered.  Basic visuals are the simple properties 
of your site (ie. the text style, the main background color, the location of the title, etc.).  If you 
don't know much about spatial graphic editing, it's a good idea to stick with altering basic visuals for 
a while.  As you learn more about your customization capabilities, you can move on to the more advanced 
visual properties.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Sectors and Common Objects - </span>
	Although your site may now look like one unit because of the solid background color, there 
	are distinct areas common to every page.  Each page can be divided into three adjacent 
	rectangular sectors:  the site title sector, the section menu sector, and the body sector.  
	Within these sectors are objects that are common to many pages on your site.  Let's see a brief 
	description of each of these three sectors and their common objects before we begin altering their properties:
	<blockquote>
		<p align=left class=BodyText>
		<span class=SubHeading>1. Site Title Sector: </span> 
		The title sectors of every page of your site are identical.  Depending on your site's configuration, 
		the title sector can either span the entire page ceiling (default config.) or share a portion of the 
		ceiling with the section menu sector.  When you first acquire your page, your site name appears in 
		text in the center of the title sector.  However, using basic visuals, you can change the site title's 
		text, font, style, color, size and position.
		</p>
		<p align=left class=BodyText>
		<span class=SubHeading>2. Section Menu Sector: </span> 
		The section menu sector is also exactly the same on every page of your site.  Depending on your site's 
		configuration, it can occupy one of four areas: spanning the entire left wall, spanning the entire right 
		wall, sitting below the title sector base against the left wall (default), or sitting below the title 
		sector against the right wall.  The common objects within the menu sector are links to each the sections 
		in your site.  Using basic visuals, you are able to change the section links text size and position 
		(the text color is dependant on the link colors dictated by the body sector common objects: see below).
		</p>
		<p align=left class=BodyText>
		<span class=SubHeading>3. Body Sector: </span> 
		The body sector fits below the title sector on either the left or right (default) side of your page.   
		The body sector contains the main information of each page so every page's body sector is slightly 
		different.  Although varying in detail, body sectors contain the same common objects:
		<blockquote>
			<p align=left class=BodyText>
			<b><i>Headings: </i></b>  
			Headings appear at the extreme top of the body sector on every page except the home page.  These objects 
			act as subtitles by labeling each particular page.  For instance, if you enter the message forum, the 
			heading for every page within that section is "Message Forum".  You can change the headings' font, 
			style, color and position.
			<br>
		
			<b><i>Body Text (does not include tables): </i></b>  
			Body text usually appears on view pages and is printed directly on a page's main background.  By 
			manipulating basic visuals, you are able to change the main background color and body text's font, 
			style, color, size, and alignment.
			<br>
		
			<b><i>Tables: </i></b>  
			Listings of properties or items are organized into tables. These tables appear in the home page and in 
			nearly all creation, edit, list, and search result pages.  Tables can be broken into two areas: the 
			table title and table data.  The separate background colors of both the table areas are different from 
			the main background color so as to make a table appear as a distinct rectangle.  As administrator, you 
			are able to change every table's structure, both background colors, and its text's font, style, color, 
			and size.
			<br>
		
			<b><i>Links: </i></b>  
			Links appear throughout the body sector of every page (they also appear in the section menu sector).  These 
			common objects, when clicked on, will load a new page in the web browser.  When a link is clicked on from 
			one particular computer, it's color changes permanently.  This occurs so that the site-goer can tell 
			what links he/she has or has not already visited.  As site administrator, you have the ability to 
			change the original and the visited color of links.
		</blockquote>
		<p align=left class=BodyText>
		<span class=SubHeading>4. Footer Sector: </span> 
		Footer spans the floor of every page and is a duplicate of the section menu spread horizontally.  A footer 
		provides a way for site-goers to navigate sections without having to scroll back up to the section menu.  
		You can also hide certain section links by having them only in the inconspicuous footer and not the main 
		section menu. 
		</p>
	</blockquote>
	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Altering Basic Visual Properties - </span>
	We don't expect you to remember all of your common objects from the previous lesson, nor do we believe 
	you'll need to change them all immediately.  The best way to learn about your common object properties 
	is to experiment.  Alter a few properties and then return to the home page to check out the change.  The 
	point of your customization capabilities is not that you'll necessarily need all of them, but if you do, 
	the option is always there.
	</p>

		<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
		Most of the features for common object properties were either explained earlier in the chapter or are 
		self-explanatory.  Therefore, in the following procedure, we will only detail those features which are 
		not inherently obvious.  To alter the basic visual properties of your site:

		<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>:</i></b>  
		Click the Change Basic Visuals link under Visual Customization.
		<br>
		
		<b><i>2 <%PrintSymb "Edit", "change basic visuals.GIF"%>:</i></b>  
		Make the desired changes.<br>
		<blockquote>
			<%PrintArrow%><b><i>Colors: </i></b> 
			Certain web browsers may not show the color choices of the color pull-down menus properly.  If this 
			is the case, click on the Color Problems? link at the top of the page.  This will load a similar page 
			with a different color selection process.  If this does happen, we recommend updating to the latest version 
			of your browser.  That should correct the problem.
			<br>
			<%PrintArrow%><b><i>Table Spacing: </i></b> 
			If you're not familiar with table language, you might want to leave the cell spacing, padding, and 
			border properties alone.  If you do decide to experiment with these features, remember that the 
			defaults are 1, 3, and 0 (in that order).
			<br>
			<%PrintArrow%><b><i>Alternating Background Colors of Table Data: </i></b> 
			Under table structure, you are able to alternate the background colors for each data listing of 
			tables.  If you wish to keep a solid color for the table data background, simply select the 
			same color for both pull-down menus.
		</blockquote>

		<b><i>3:</i></b>  
		Scroll down to the page bottom and click the Update button.
		<br>
		
		<b><i>4:</i></b>  
		Using the provided links, to either reload the basic visuals page or to return to the main administrator page.
		</p>

		</blockquote>

<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 BASIC GRAPHICS: </span>&nbsp;
A graphic is an image file used to replace one of four types of common objects within your site: the site title, 
the section menu links, the new item symbol, or the message forum's collapse/expand symbols.  For instance, instead 
of using the plain text from the basic visuals for your title, you can substitute a graphic.  Since a graphic can 
be any image file and are easy to create, visual customization using graphics offers infinite possibilities. This 
is the reason every GroupLoop site looks so different.
</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Although we have a few community graphics to choose from, we encourage you to create your own using a graphic 
	editing program.  A	good way to get started on graphic customization is to create just a title graphic.  Once you become proficient 
	in manipulating the title graphic on your site, inserting  buttons and symbol graphics can come next.  Creating 
	graphics can be extremely involved, and since you must use a third party program (such as Photoshop, Paintbrush, etc.), 
	we cannot provide support or instructions for creating the graphic.  We realize that many people would just rather 
	have someone else to create graphics for them.  That's why we offer our Custom Graphics Service.  We have skilled 
	graphics artists ready to make your site look amazing (or ugly, whatever you want).  It usually isn't too 
	expensive, and may be the best thing you ever do for your site.  If you are interested in the Custom Graphics 
	Service, please e-mail <a href="mailto:support@grouploop.com">support@grouploop.com</a>.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Graphic Types - </span>
	There are three types of customization graphics: standard graphics, rollover graphics, and background 
	graphics.  Standard graphics are image files that appear naturally as objects on your screen.  When a 
	site-goer positions the cursor over an object with rollover, the standard graphic will change to the 
	rollover graphic.  The standard graphic returns when the site-goer moves the cursor away from the 
	object.  Background graphics are image files that act like wallpaper tiled behind all other site objects.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Inserting and Removing Basic Graphics - </span>
	When inserting graphics, you first need to have them saved to disk as image files.  Use the following 
	procedure to insert or remove basic graphics:
	</p>

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>:</i></b>  
		Click the Change Basic Graphics link under the Visual Customization subheading.
		<br>
			
		<b><i>2 <%PrintSymb "Edit", "change graphics.GIF"%>:</i></b>  
		Make the appropriate property selections and entries.<br>
		<blockquote>
			<%PrintArrow%><b><i>Graphic Status Option Bubbles: </i></b> 
			The selection in the Use an image? property just below each subheading dictates whether or 
			not an object has an associated graphic.  If you wish to associate an image file with an object, 
			mark the Yes option bubble.  If you wish to <b>remove</b> an existing graphic (not replace it with a new 
			one) and return the chosen object to plain text or the default symbol, mark the No option bubble.  
			Otherwise, <b>leave this setting as is</b>.
			<br>
			<%PrintArrow%><b><i>Blank Textboxes: </i></b> 
			We didn't lose your currently inserted graphics.  When the basic graphic edit page first loads, 
			all the textboxes <b>should</b> be blank.  You should only enter information into these textboxes when 
			you wish to insert a new graphic.
			<br>
			<%PrintArrow%><b><i>Browse Feature: </i></b> 
			If you do not know the exact label (ie. c:\images\title.jpg - you probably don't know it) of the image file 
			you wish to insert, click the Browse... button to the right of the appropriate textbox.  In the directory 
			box that pops up, you can search your computer's drives and folders for the desired image file.  When 
			you find it, double-click on it and the image file's proper label will appear in the appropriate textbox.
		</blockquote>

		<b><i>3:</i></b>  
		Click the Update (click once) button at the bottom of the page.  As the popup box will say, don't keep pressing Update.  
		You are sending files to our server, and this sometimes takes time.  Just be patient.  Constantly clicking Update will 
		just reset the upload, which will waste more of your time.
		<br>

		<b><i>4:</i></b>  
		Using the links provided, either reload the graphic edit page or return to the main administrator page.

	</blockquote>

<a name="3"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.3 ADVANCED VISUALS: </span>&nbsp;
The advanced visual options are intended for site administrators who are extremely proficient in graphic editing.  
Specifically, advanced visual options allow the administrator to integrate graphics with different sector 
backgrounds.  Since, properly combining advanced visuals with the rest of your page's objects is very 
time-consuming and somewhat trial and error, mastery of advanced visuals takes a lot of practice.
</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Outline for Creating Advanced Graphics - </span>
	The best way to go about inserting an aesthetically pleasing site title and section menu with advanced graphics 
	is to first create your entire  page layout in a graphic editing program.  Divide the page into three adjacent 
	rectangles called sectors: the site title, the section menu, and the body.  There are five configurations in 
	which the three sectors can possibly be situated:

	<blockquote>

		<b><i>1 Alfa Configuration: </i></b>  
		The menu sector occupies the upper left corner of the page and extends down to the page floor (spans the 
		entire left wall).  The title sector is to the menu's right and spans the remainder of the page ceiling.  The 
		body sector fits against the title and menu sectors and extends to the right wall and page floor.
		<br>
		
		<b><i>2 Beta Configuration: </i></b>  
		The vertically mirrored image of the alfa configuration (menu sector on right wall).
		<br>
		
		<b><i>3 Chi Configuration: </i></b>  
		The title sector occupies the upper left corner and spans the entire page ceiling.  The menu sector fits 
		below the left side of the title and extends down the left wall to the page floor.  The remainder of the 
		page is the body sector.
		<br>
		
		<b><i>4 Delta Configuration: </i></b>  
		The vertically mirrored image of the chi configuration (menu sector on right side).
		<br>
		
		<b><i>5 Sigma Configuration: </i></b>  
		The title sector spans the entire page ceiling.  The menu sector fits below the title and extends horizontally 
		from the left wall to the right wall.  We do not recommend using this configuration with graphics because the 
		way individual web browsers will display it is unpredictable.

	</blockquote>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Still in the graphic editing program, you must now further divide each sector into subdivisions.  You can 
	use the same processes for either an alpha, beta, chi, or delta configuration.  For the menu sector, divide 
	the link texts into vertically stacked rectangular buttons each of identical horizontal width and roughly equal 
	vertical height.  Each button must span the width of your menu sector.  If there is artistry above the link blocks 
	that you don't wish to include in the first link button, make it into its own rectangular menu header (shape does 
	not have to be similar to the link subdivisions).  If there is non-symmetrical artistry below the link buttons, 
	make it into its own rectangle (menu footer).  
	</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Finally, cut out a piece of the remainder of the menu sector that you want tiled behind the buttons (any 
	artistry contained in the link buttons and header/footer will sit on top of these tiles).  These tiles must 
	span the width of your menu sector.  The background tiles will later be repeated vertically down the length 
	of the menu sector and should thus, be horizontally symmetrical so they can line up properly.  Any portion of 
	the menu sector that remains to be subdivided is unimportant because it should just be redundant background.
	</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	For the title sector, divide the title text and any asymmetrical artistry into one subdivision that spans the 
	height of the sector.  Cut out a piece of the vertically symmetrical remainder into a background tile (this 
	will later be repeated horizontally).  This tile must also span the height of the sector.  What remains in 
	the title sector is unimportant.  The body sector can also be discarded (this background is dictated by basic 
	visuals or basic graphics).
	</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	Save each subdivision as an appropriately named image file (ie. title text subdivision = c:\graphics\ttext.jpg).  If 
	you would like to use rollovers for a menu header and/or footer, repeat this entire outline with the rollover in 
	place and save their subdivisions to disk.  The header and/or footer rollover subdivisions must be of the exact 
	same dimensions as their corresponding standard graphics.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Manipulating Advanced Visuals - </span>
	Once you have all of your graphics saved to disk as image files, use the basic graphic insertion procedure to 
	introduce the new title graphic and each section link graphic.  Don't worry if they're not positioned properly 
	on your home page yet.  You can now use the following procedure to integrate your background tiles and/or menu 
	header and footer:

	<blockquote>

		<p align=left class=BodyText>
		<b><i>1 <%PrintSymb "Member", ""%>:</i></b>  
		Click the Change Basic Graphics link under the Visual Customization subheading.<br>
		<br>
		<b><i>2:</i></b>  
		Click the Advanced Graphics link below the heading.<br>
		<br>		
		<b><i>3 <%PrintSymb "Edit", "advanced graphics.GIF"%>:</i></b>  
		Make the appropriate data field selections and/or entries.<br>
		<blockquote>
			<%PrintArrow%><b><i>Configuration: </i></b> 
			If you used an Alpha or Beta configuration, mark the Yes option bubble, in the first property field.  If 
			you used a Chi or Delta configuration, mark the No option bubble.
			<br>
			<%PrintArrow%><b><i>Title Spacing: </i></b> 
			If you either have a title background tile or the artistry of the title graphic lines up with artistry 
			from the menu sector, this must be marked No.  Only mark this Yes when you want the title sector slightly 
			larger than the title graphic or when you're not using a title graphic at all (plain text).
			<br>
			<%PrintArrow%><b><i>Graphic Status Option Bubbles: </i></b> 
			The selection in the Use a ... image? property just below each subheading dictates whether or not a sector 
			subdivision has an associated graphic.  If you wish to associate an image file with a subdivision, mark 
			the Yes option bubble.  If you wish to <b>remove</b> (not replace) an existing graphic and return the object to 
			plain text, mark the No option bubble.  Otherwise, <b>leave these settings as is</b>.
			<br>
			<%PrintArrow%><b><i>Blank Textboxes: </i></b> 
			As with basic graphics, when the advanced visuals edit page first loads, all the textboxes <b>should</b> be blank.  You 
			only need to enter information when you wish to insert a new graphic.
			<br>
			<%PrintArrow%><b><i>Browse Feature: </i></b> 
			Explained in the basic graphic procedure.
			<br>
			<%PrintArrow%><b><i>Menu Width Percentage: </i></b> 
			If you are using a menu background tile, this data field must always be set to 1 so the body will begin 
			to the immediate right/left (depending on configuration type) of the link buttons.  If you are not using a 
			background tile and wish for their to be space between your section links and the body, set the width 
			percentage accordingly (you may have to play with this number until the look of the menu meets your standards).
		</blockquote>

		<b><i>4:</i></b>  
		Scroll to the bottom and click the Update (click once) button.
		<blockquote>
			<%PrintArrow%><b><i>Click Once: </i></b> 
			If you have inserted any image files, the confirmation page may take several seconds to load.  This is 
			because your computer has to upload the image file(s) onto our server.  Each time you click on the Update 
			(click once) button, the upload process starts over.  Point being, click the button just one time and be 
			patient.
		</blockquote>

		<b><i>5:</i></b>  
	Use the links provided to either reload the advanced visuals edit page or to return to the main administrator page.

	</blockquote>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	If you have made the correct data field selections, your pages should now look identical to the page layout 
	you created in the graphic editing program.  If not, it's ok.  Getting advanced graphics to fit together properly 
	takes a lot of practice, so don't be discouraged if your page doesn't look right on the first try.  You may have 
	to play with the basic visual, basic graphic, and advanced visual settings until you get your  page looking the 
	way you want.
	</p>

<a name="4"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.4 SCHEMES: </span>&nbsp;
A scheme is a saved set of visual and/or graphic property settings.  Schemes allow the administrator to load a entire 
group of visual and/or graphic settings in one step instead of having to change them individually.  We offer a few 
standard schemes but we encourage you to develop and experiment with your own.
</p>

	<p align=left class=BodyText><b>Tip: </b>
	A number of GroupLoop site administrators rotate through a set of different custom schemes on a regular basis 
	to keep their sites always new and exciting.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Scheme Types - </span>
	There are three types of schemes: basic schemes, graphic schemes, and master schemes. A basic scheme is a saved 
	set of basic visual settings.  A graphic scheme is a set of basic graphics.  A master scheme is a combination of 
	all visual properties; it contains basic visuals, basic graphics, and advanced visual settings.  Using the scheme 
	manager, you can save a scheme of with current visual properties, load an existing scheme, or mix and match basic 
	with graphic schemes.
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Saving the Current Visual Properties into a Scheme - </span>
	Once you find a combination of visual settings that you like, you should save them as a scheme.  To save a scheme 
	containing your current settings:
	
	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b>  
		Click the Scheme Manager link under the Visual Customization subheading.
		<br>
		<b><i>2 <%PrintSymb "Member", "scheme manager.gif"%>: </i></b>  
		Click the Save Your Current Scheme link.
		<br>

		<b><i>3 <%PrintSymb "Create", "save scheme.gif"%>: </i></b>  
		Make the proper data entries and selections.
		<br>

		<blockquote>
			<%PrintArrow%><b><i>Scheme Type: </i></b> 
			With this field, you can choose to save your basic visuals and basic graphics separately (basic or 
			graphic option) or you can save all the visual settings at once.  Please note that saving the look and graphics 
			will result in two saved schemes.  This presents no problem to you, and also allows you to mix and match 
			looks and graphics.
		</blockquote>

		<b><i>4: </i></b>  
		Click the Add button below the Description textbox.
		<br>

		<b><i>5: </i></b>  
		Using the links provided, either reload the scheme creation page or return to the administrator menu.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Loading an Existing Scheme - </span>
	To load either one of your custom schemes or an GroupLoop standard scheme use the following procedure:

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b>  
		Click the Scheme Manager link under the Visual Customization subheading.
		<br>
		<b><i>2 <%PrintSymb "Member", "scheme manager.gif"%>: </i></b>  
		Click the Load One of Your Saved Schemes link.
		<br>

		<b><i>3 <%PrintSymb "List", "load scheme.gif"%>: </i></b>  
		Find the scheme you wish to load and click the Load button.  If you saved both the look and graphics together, you 
		need to load the look and graphics separately (only takes 5 seconds).
		<br>

		<blockquote>
			<%PrintArrow%><b><i>GroupLoop Community Schemes: </i></b> 
			If you wish to use one of the master schemes we have developed, click the GroupLoop schemes link near 
			the top of the body sector.  A new scheme list will load.  You can view a sample screen shot of each by 
			clicking the View link to the scheme's left.  Once you find the scheme you want, click the Load button 
			to its right.
		</blockquote>

		<b><i>4: </i></b>  
		Make the appropriate decision.
		<br>

		<blockquote>
			<%PrintArrow%><b><i>Warning! </i></b> 
			Remember that when you load a new scheme, the graphics/look will be changes, and your current setting will 
			be lost.  We suggest saving your current look/graphics as a scheme so you can always load it again later.
		</blockquote>

		<b><i>5 <%PrintSymb "Confirmation", "new scheme loaded.gif"%>: </i></b>  
		Using the links provided, either reload the scheme list page or the return to the main administrator page.
		<br>

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Custom Scheme - </span>
	If you wish to remove one of your saved schemes from your site completely, use the following procedure: 

	<blockquote>

		<b><i>1 <%PrintSymb "Member", ""%>: </i></b>  
		Click the Scheme Manager link  under the Visual Customization subheading.
		<br>
		<b><i>2 <%PrintSymb "Member", "scheme manager.gif"%>: </i></b>  
		Click the Modify one of your Saved Schemes link.
		<br>

		<b><i>3 <%PrintSymb "List", "edit-delete scheme.gif"%>: </i></b>  
		Find the scheme you wish to delete.
		<br>

		<b><i>4: </i></b>  
		Click the Delete button to the schemes right.
		<br>

		<b><i>5 <%PrintSymb "PopUp", "delete warning box.gif"%>: </i></b>  
		If you're sure click the OK button.  If not, click Cancel.
		<br>

		<b><i>6 <%PrintSymb "Confirmation", "deleted scheme.gif"%>: </i></b>  
		Using the links provided, either reload the custom scheme list page or the return to the main administrator page.
		<br>

	</blockquote>
