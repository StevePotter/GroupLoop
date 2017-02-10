<!-- #include file="photos_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludePhotos ) or not CBool( PhotosMembers ) then Redirect("error.asp")
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Modify Your Photos</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
'update info
Set upl = Server.CreateObject("SoftArtisans.FileUp")
upl.Path = GetPath ("photos")

if not LoggedMember and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
if not LoggedMember then Redirect("members.asp?Source=members_photos_modify.asp")

intPhotoID = CInt(upl.Form("PhotoID"))

blLoggedAdmin = LoggedAdmin

if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

Set rsUpdate = Server.CreateObject("ADODB.Recordset")
Query = "SELECT * FROM Photos WHERE ID = " & intPhotoID & " AND " & strMatch
rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
if rsUpdate.EOF then
	set upl = Nothing
	set rsUpdate = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
end if

rsUpdate("CategoryID") = upl.Form("ID")
rsUpdate("Name") = Format( upl.Form("Name") )
if blLoggedAdmin then rsUpdate("Date") = upl.Form("Date")
rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
rsUpdate("ModifiedID") = Session("MemberID")

strError = ""
if blLoggedAdmin and upl.Form("Date") = "" then strError = "You forgot to enter the date.<br>"

if not upl.Form("PhotoFile").IsEmpty then
	'--- Retrieve the file's content type and assign it to a variable
	FTYPE = upl.Form("PhotoFile").ContentType

	'--- Restrict the file types saved using a Select condition
	if FTYPE = "image/gif" then
		upl.Form("PhotoFile").SaveAs intPhotoID&".gif"
		rsUpdate("Ext") = "gif"
	elseif FTYPE = "image/pjpeg" or FTYPE = "image/jpeg" then
		upl.Form("PhotoFile").SaveAs intPhotoID&".jpg"
		rsUpdate("Ext") = "jpg"
	elseif FTYPE = "image/bmp" then
		upl.Form("PhotoFile").SaveAs intPhotoID&".bmp"
		rsUpdate("Ext") = "bmp"
	else
		upl.Form("PhotoFile").delete
		blProceed = false
		strError = strError & "You can only upload gif and jpeg images.  Your photo was a banned type of file.<br>"
	end if
end if

if strError = "" then

	strFileName = GetPath("photos") & rsUpdate("ID") & "." & rsUpdate("Ext")
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	if not FileSystem.FileExists(strFileName) then
		strError = strError & "Cannot change because photo does not exist - " & strFileName
	else
		Set Image = Server.CreateObject("AspImage.Image")
		'Create the object and load up the big image
		Image.LoadImage(strFileName)
		'Rotate the image
		if upl.Form("Rotate") <> "0" and upl.Form("Rotate") <> "" then
			intRotate = CInt(upl.Form("Rotate"))
			Image.RotateImage intRotate
		end if
		'Brighten or darken the image
		if upl.Form("Brightness") <> "0" and upl.Form("Brightness") <> "" then
			intBrighten = CInt(upl.Form("Brightness"))
			if intBrighten > 0 then
				Image.BrightenImage intBrighten
			else
				Image.DarkenImage intBrighten
			end if
		end if
		'New contrast
		if upl.Form("Contrast") <> "0" and upl.Form("Contrast") <> "" then
			intContrast = CInt(upl.Form("Contrast"))
			if intContrast > 100 then intContrast = 100
			if intContrast < -100 then intContrast = -100

			Image.Contrast intContrast
		end if
		'New Sharpen
		if upl.Form("Sharpen") <> "0" and upl.Form("Sharpen") <> "" then
			intSharpen = CInt(upl.Form("Sharpen"))
			Image.Sharpen intSharpen
		end if

		Image.FileName = strFileName
		Image.SaveImage

		Set Image = Nothing

		'Optimize the image
		if upl.Form("Optimize") = "1" then
			OptimizeImage GetPath("photos") & rsUpdate("ID"), rsUpdate("Ext")
			rsUpdate("Ext") = "jpg"
		end if

		'Create the new thumbnail
		if upl.Form("Thumbnail") = "1" then
			strThumbFileName = GetPath("photos") & rsUpdate("ID") & "t." & rsUpdate("ThumbnailExt")
			if FileSystem.FileExists(strThumbFileName) then FileSystem.DeleteFile(strThumbFileName)
			CreateThumbnail strFileName, strThumbFileName
			rsUpdate("ThumbnailExt") = "jpg"
			rsUpdate("Thumbnail") = 1
		end if
	end if
	Set FileSystem = Nothing
end if

'Now make sure we still don't have a problem
if strError <> "" then
	Set rsUpdate = Nothing
	Set upl = Nothing
	Redirect "incomplete.asp?Message=" & Server.URLEncode(strError)
else
	intCategoryID = rsUpdate("CategoryID")
	rsUpdate.Update
	Set rsUpdate = Nothing
	Set upl = Nothing
%>
	<p>The photo has been edited. &nbsp;<a href="members_photos_modify.asp">Click here</a> to modify another.<br>
		<a href="members_photos_modify.asp?ID=<%=intCategoryID%>">Click here</a> to modify another photo in <%=GetCategoryName(intCategoryID, "PhotoCategories")%>.<br>
		<a href="photos_view.asp?ID=<%=intPhotoID%>">Click here</a> to view the modified photo.
	</p>
<%
end if
%>
