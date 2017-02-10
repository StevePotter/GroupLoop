<!-- #include file="photos_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludePhotos ) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but you can't add photos right now.  The option has been disabled.") )
Session.Timeout = 20
Server.ScriptTimeout = 5400
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add A Photo</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
	Set upl = Server.CreateObject("SoftArtisans.FileUp")
	upl.Path = GetPath ("photos")

	if not LoggedMember and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
	if not LoggedMember then Redirect("members.asp?Source=members_photos_add.asp")
	if not (LoggedAdmin or CBool( PhotosMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))

	'Create the new photo
	Set cmdTemp = Server.CreateObject("ADODB.Command")
	With cmdTemp
		.ActiveConnection = Connect
		.CommandText = "AddPhoto"
		.CommandType = adCmdStoredProc

		.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamOutput )
		.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )

		.Parameters("@MemberID") = Session("MemberID")

		.Execute , , adExecuteNoRecords
		intPhotoID = .Parameters("@ItemID")
	End With

	blProceed = true
	strError = ""

	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	Set TestFolder = FileSystem.GetFolder( GetPath("photos") )
	dblAvailable = (PhotosMegs * 1000000) - TestFolder.Size
	Set TestFolder = Nothing

	if upl.Form("PhotoFile").IsEmpty then
		strError = strError & "You did not specify a photo to upload.  Please try again.<br>"
	elseif upl.Form("Name") = "" and upl.Form("Thumbnail") = "0" then
		strError = strError & "You did not give a description of the photo.<br>"
	elseif upl.TotalBytes > dblAvailable then
		strError = strError & "Sorry, but there is not enough free space in the " & PhotosTitle & "	section for your file.  An administrator may purchase more space by clicking 'Modify Account' in the Members Only section.<br>"
	else
		'--- Retrieve the file's content type and assign it to a variable
		FTYPE = upl.Form("PhotoFile").ContentType

		strFileName = ""
		'--- Restrict the file types saved using a Select condition
		if FTYPE = "image/gif" then
			strFileName = intPhotoID&".gif"
			upl.Form("PhotoFile").SaveAs strFileName
			strExt = "gif"
		elseif FTYPE = "image/pjpeg" or FTYPE = "image/jpeg" then
			strFileName = intPhotoID&".jpg"
			upl.Form("PhotoFile").SaveAs strFileName
			strExt = "jpg"
		elseif FTYPE = "image/bmp" then
			strFileName = intPhotoID&".bmp"
			upl.Form("PhotoFile").SaveAs strFileName
			strExt = "bmp"
		else
			upl.Form("PhotoFile").delete
			blProceed = false
			strError = strError & "You can only upload gif, bitmap (bmp), and jpeg images.  Your photo was a banned type of file.<br>"
		end if
	end if

	intCategoryID = CInt(upl.form("ID"))
	strName = Format( upl.form("Name") )


	if strError = "" then
		intThumbnail = 0
		strThumbnailExt = ""

		'Process the thumbnail
		if upl.Form("Thumbnail") = "1" then
			if strExt <> "gif" then

				strFileName = GetPath("photos") & strFileName
				strThumbFileName = GetPath("photos") & intPhotoID&"t.jpg"

				'Make sure we have an image to get
				if FileSystem.FileExists (strFileName) then
					CreateThumbnail strFileName, strThumbFileName
					strThumbnailExt = "jpg"
					intThumbnail = 1
				else
					strError = strError & "The photo has been lost and a thumbnail could not be created.  Please try again.<br>"
				end if
			end if
		end if

		'Optimize the image
		if upl.Form("Optimize") = "1" then
			OptimizeImage GetPath("photos") & intPhotoID, strExt
			strExt = "jpg"
		end if

	end if

	Set FileSytem = Nothing
	Set upl = Nothing

	'Now make sure we still don't have a problem
	if strError <> "" then
		With cmdTemp
			.CommandText = "DeletePhoto"

			.Parameters.Delete ("@ItemID")
			.Parameters.Delete ("@MemberID")
			.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamInput )
			.Parameters("@ItemID") = intPhotoID

			.Execute , , adExecuteNoRecords
		End With
		Set cmdTemp = Nothing
		Redirect "incomplete.asp?Message=" & Server.URLEncode(strError)
	else
		With cmdTemp
			.CommandText = "UpdatePhoto"

			.Parameters.Refresh

			.Parameters("@ItemID") = intPhotoID
			.Parameters("@MemberID") = Session("MemberID")
			.Parameters("@CategoryID") = intCategoryID
			.Parameters("@ModifiedID") = Session("MemberID")
			.Parameters("@CustomerID") = CustomerID
			.Parameters("@IP") = intPhotoID
			.Parameters("@Name") = CStr(strName)
			.Parameters("@Thumbnail") = intThumbnail
			.Parameters("@Ext") = strExt
			.Parameters("@ThumbnailExt") = strThumbnailExt

			.Execute , , adExecuteNoRecords
		End With
		Set cmdTemp = Nothing
'------------------------End Code-----------------------------
%>
		<p>The photo has been added. &nbsp;<a href="members_photos_add.asp?ID=<%=intCategoryID%>">Click here</a> to add another.<br>
			<a href="photos_view.asp?ID=<%=intPhotoID%>">Click here</a> to view the photo you just uploaded.
		</p>
<%
'-----------------------Begin Code----------------------------
	end if
%>