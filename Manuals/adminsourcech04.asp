<% intChapter = 4 %>
<a href="default.asp"><img src="../images/toc.gif" border="0" alt="Table Of Contents"></a>
<a href="ch0<%=intChapter - 1%>.asp"><img src="../images/previous.gif" border="0"></a>
<a href="ch0<%=intChapter + 1%>.asp"><img src="../images/next.gif" border="0"></a>

<p class=Title align=center>CHAPTER <%=intChapter%>: MEMBERS AND THE PUBLIC</p>

<a name="1"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.1 WHAT'S A MEMBER?: </span>&nbsp;
Visitors to your site can be divided into two distinct categories.  The first is non-members.  This group has 
very limited capabilities on your site.  Non-members can only view public information and participate in 
public sections such as the message board, the guestbook, voting polls, and quizzes.  Don't discount them 
yet, though.  The majority of our longest running sites typically have hundreds (in a few cases thousands) 
of non-members that visit on a regular basis.  Often times, a non-member who is a regular visitor to one 
particular GroupLoop site, has a membership with his own GroupLoop site.  You could, however, make your site 
completely private, never allowing non-members to read or write anything on your site.  It's up to you.
</p>

	<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
	The other group of site visitors is members.  Members are able to enter and participate in any section on 
	the site except the administrators section (as the administrator, you already know that you are the only 
	person able to enter that section).  For a more complete guide to member capabilities read the Member's Manual.
	</p>

	<p align=left class=BodyText>
	<b>Note:</b> Based on the customer feedback we get, site members typically split the monthly GroupLoop 
	bill with the administrator.  For most sites, members and administrators end up paying about two bucks a 
	piece per month.  So, find some more members and fatten your wallet!
	</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Member Properties - </span>
	Every membership has six data fields that are originally created by the administrator but can be changed 
	later by the member or the administrator (except admin access; only the administrator can change that).
	<blockquote>
		<p align=left class=BodyText>
		<b><i>1 - First Name:</i></b>  
		This is the legal first name of the site member.
		<br>
		<b><i>2 - Last Name:</i></b>  
		This is the legal last name of the site member.
		<br>
		<b><i>3 - Nickname:</i></b>  
		The nickname is the label that will be displayed as author on a member's posts.  It's also the screen 
		name that your member will need to sign in to the members section.  Most members and/or administrators 
		choose nicknames that they commonly hear from the other members.  For instance, your group of friends 
		might call Robert just Bob for short.  This field is not case-sensitive meaning whether you use capital 
		letters or lower-case ones makes no difference (A is equal to a).
		<br>
		<b><i>4 - Password:</i></b>  
		This is the series of letters and/or numbers that must be entered after the nickname in the 
		log-in page.  It's a good idea for the member to use a password that isn't too obvious to others, 
		but one that he/she still won't forget.  This field is also not case-sensitive.
		<br>
		<b><i>5 - Email Address:</i></b>  
		An entry in this field, though not required, is strongly recommended.  If other members or guests 
		wish to contact the author of a particular post, his/her email can be easily accessed.  Also, the 
		email notification feature requires a valid email address.  When adding a new member, selecting email 
		notification will automatically inform the new member of his/her new privileges at your site.
		<br>
		<b><i>6 - Administrator Access:</i></b>  
		This property can only be enacted the administrator and decides whether or not a member will have 
		administrator capabilities.  Realize that once you enact this feature, you will give that member 
		all of the same administrator options that you have (including the ability to remove YOUR 
		administrator access!).  Therefore, think very carefully about to whom you give these powers.  
		We recommend you only grant this property to a responsible member if you are completely certain 
		that you cannot do without a second administrator.
		</p>
	</blockquote>
	


<a name="2"></a>
<p align=left class=BodyText><span class=Heading><%=intChapter%>.2 THE ADMINISTRATOR'S OPTIONS REGARDING MEMBERS: </span>&nbsp;
The administrator is able to add new members, edit a membership's properties, edit a member's personal information, 
and terminate a membership.
</p>

	<%PrintBullet%><span class=SubHeading>Adding a New Member - </span>
	The administrator is the only person able to add site members.  Use the following procedure to add a new member:

	<blockquote>
		<p align=left class=BodyText>
		<b><i>1 <%PrintSymb "Member", ""%>:</i></b>  
		Click Add a New Member under the Members subheading.
		<br>

		<b><i>2 <%PrintSymb "Create", "create member.gif"%>:</i></b>  
		Enter the proper info and properties for the new membership.<br>
		<blockquote>
			<%PrintArrow%><b><i>Email Notification:</i></b> 
			This feature, when selected, will automatically send your new member an email containing information about 
			his/her new membership to your site.  The notification email will include the member's nickname and password 
			and a brief explanation of your GroupLoop site.<br>
		</blockquote>

		<b><i>3:</i></b>  
		Click the Add button.
		<br>

		<b><i>4 <%PrintSymb "Confirmation", "member added.gif"%>:</i></b>  
		The member has been added, and if you so chose, an e-mail has been sent giving then their name, password, and links to 
		the site.  Using the links provided, either reload the member creation page to add more members or return to the main administrator page.
	</blockquote>

		<p align=left class=BodyText>&nbsp;&nbsp;&nbsp;
		If you decided not to use the email notification feature or you did not enter an email address, you should 
		contact your new member by some other means.  Tell him/her the membership nickname and password, and give a 
		quick explanation of your site.  Don't forget to tell him to read the Member's Manual.  If you have for some 
		reason forgotten the nickname and/or password, you can find it by using the first step in the next procedure.
		</p>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Editing an Existing Membership and/or Member Personal Info - </span>
	At some time you may find it necessary to alter a member's information, additions or status.  To change a 
	membership's properties and/or a member's personal info:

	<blockquote>

		<p align=left class=BodyText>
		<b><i>1 <%PrintSymb "Member", ""%>:</i></b>  
		Click Modify Members under the Members subheading.
		<br>
		
		<b><i>2 <%PrintSymb "List", "choose member.gif"%>:</i></b>  
		Find the membership you wish to change.
		<br>
		
		<b><i>3 <%PrintSymb "Edit", "modify member.GIF"%>:</i></b>  
		Make the necessary changes.<br>
		<blockquote>
			<%PrintArrow%><b><i>Member's Personal Info:</i></b> 
			You may notice that there are many additional data fields (other than those from the membership 
			creation page) such as home address, personal phone number, etc.  These extra fields should originally 
			be entered by the member, but the administrator can change them at any time using this procedure.
			<br>
			<%PrintArrow%><b><i>Password:</i></b> 
			This is the only place on the site where you can access other members' passwords.  If someone loses their password 
			and asks you for it, this is where to find it!

		</blockquote>

		<b><i>5:</i></b>  
		Scroll to the bottom of the page and click the Update button.
		<br>
		
		<b><i>6:</i></b>  
		The member has been modified.

	</blockquote>

	<p align=left class=BodyText>
	<%PrintBullet%><span class=SubHeading>Deleting a Membership - </span>
	There may come a time when you need to terminate a site membership.  To do so, use the following procedure:

	<blockquote>
		<b><i>1 <%PrintSymb "Member", ""%>:</i></b>  
		Click Modify Members under the Members subheading.
		<br>
		
		<b><i>2 <%PrintSymb "List", "choose member.gif"%>:</i></b>  
		Find the membership you wish to terminate.
		<br>
		
		<b><i>3:</i></b>  
		Click the appropriate deletion button.<br>
		<blockquote>
			<%PrintArrow%><b><i>Deletion Buttons:</i></b> 
			Before removing a membership from your site's member list, it is important to first decide if you want 
			to remove all of his/her additions from your site also (ie. stories, photo comments, etc.).  This will 
			be your only chance to delete <b>all</b> of a member's additions in one step (doing them individually could 
			take forever!).  If you wish to keep the member's additions, click the Delete Membership Only button.  
			If you wish to remove all of the member's additions along with his/her membership, click the Complete 
			Delete button.
		</blockquote>

		<b><i>4 <%PrintSymb "PopUp", "delete member box.gif"%>:</i></b>  
		If you're sure, click the OK button.  If not, click Cancel.
		<br>
		
		<b><i>4 <%PrintSymb "Confirmation", "member deleted.gif"%>:</i></b>  
		<b><i>5:</i></b>  
		The member has been deleted.
	</blockquote>