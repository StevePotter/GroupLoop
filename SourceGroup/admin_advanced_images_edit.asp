<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_advanced_images_edit.asp")
Session.Timeout = 20
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change The Advanced Graphics of Your Site</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
if Request("Submit") = "Changed" then
'------------------------End Code-----------------------------
%>
		<p>The advanced look settings have been changed.  &nbsp;<a href="admin_advanced_images_edit.asp">Click here</a> to change them again. If you added new images (such as a background) and can't see the changes, simply press the Reload or Refresh button on your browser.</p>
<%
'-----------------------Begin Code----------------------------
else

	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	strImagePath = GetPath("images")
'------------------------End Code-----------------------------
%>
	<p>If you already have graphics and would like to keep them, just leave the button at 'Yes' and leave the file box blank. 
	If you have a graphic you would like to delete, just click the 'No' button, and leave the file box blank.  
	A rollover images is a graphic that gets switched with the original once someone puts their mouse pointer 
	over it.  That's how you can get the buttons on your site to 'light up' when someone goes to click on them.  
	Rollover images are optional, so if you don't have any, don't worry.
	</p> 
	<p>Remember that only image files may be uploaded.  Anything else will be rejected.</p>
	<form enctype="multipart/form-data" method="post" ACTION="<%=SecurePath%>admin_advanced_images_edit_process.asp" name="MyForm" onSubmit="if (this.submitted) return false; this.submitted = true; return true">
	<input type="hidden" name="MemberID" value="<%=Session("MemberID")%>">
	<input type="hidden" name="Password" value="<%=Session("Password")%>">
	<% PrintTableHeader 0 %>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Title Background
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				If you use a title background, it will be shown across the entire span of the page, vertically aligned with the title.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a title background image? 	<% PrintRadio ImageExistsInt("TitleBackgroundImage"), "TitleBackgroundImage" %><br>
				Image File <input type="file" name="UpTitleBackgroundImage">
			</td>
		</tr>
		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Menu Header and Footer
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu header image?  (This will appear before the buttons)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu header image? 	<% PrintRadio ImageExistsInt("MenuTopImage"), "MenuTopImage" %><br>
				Image <input type="file" name="UpMenuTopImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu header rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu header rollover image? 	<% PrintRadio ImageExistsInt("MenuTopRolloverImage"), "MenuTopRolloverImage" %><br>
				Image File<input type="file" name="UpMenuTopRolloverImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu footer image?  (This will appear after the buttons)
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu footer image? 	<% PrintRadio ImageExistsInt("MenuBottomImage"), "MenuBottomImage" %><br>
				Image File <input type="file" name="UpMenuBottomImage">
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				Do you want to use a menu footer rollover image?
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu footer rollover image? 	<% PrintRadio ImageExistsInt("MenuBottomRolloverImage"), "MenuBottomRolloverImage" %><br>
				Image File <input type="file" name="UpMenuBottomRolloverImage">
			</td>
		</tr>

		<tr>
			<td class="TDHeader" valign="middle" align="center" colspan="2">
				Menu Background
			</td>
		</tr>
		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="right">
				If you use a menu background on a right or left positioned menu, the background will extend all the way down to the end of the page, 
				only taking up the width of the menu.  If the menu is positioned at the top, the background will extend the width of the page, 
				only taking up the height of the menu.
			</td>
			<td class="<% PrintTDMain %>" align="left">
				Use a menu background image? 	<% PrintRadio ImageExistsInt("MenuBackgroundImage"), "MenuBackgroundImage" %><br>
				Image File <input type="file" name="UpMenuBackgroundImage">
			</td>
		</tr>

		<tr>
			<td class="<% PrintTDMain %>" valign="middle" align="center" colspan="2">
			<input type="submit" name="Submit" value="Update (click once)"  onClick="alert('If you are not uploading any files, just click okay and dont worry about this message.  You may be uploading many images, so please wait as long as it takes.  After pressing OK, your files will upload.  Please do not constantly press the Update button.')"></td>
			</td>
		</tr>
	</table>
	</form>
<%
'----------------------Begin Code----------------------------
	Set FileSystem = Nothing
end if

Function ImageExistsInt( strImage )
	strExt = ""
	if ImageExists( strImage, strExt) then
		ImageExistsInt = 1
	else
		ImageExistsInt = 0
	end if

End Function
'------------------------End Code-----------------------------
%>