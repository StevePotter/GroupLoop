<!-- #include file="media_functions.asp" -->
<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeMedia ) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but you can't add files right now.  The option has been disabled.") )
Session.Timeout = 20
Server.ScriptTimeout = 5400
'------------------------End Code-----------------------------
%>

<p align="<%=HeadingAlignment%>"><span class=Heading>Modify Media</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>

<%
'-----------------------Begin Code----------------------------
'update info
Set upl = Server.CreateObject("SoftArtisans.FileUp")
upl.Path = GetPath ("media")

'Make sure they can do this
if not LoggedMember and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
if not LoggedMember then Redirect("members.asp?Source=members_media_modify.asp")
if not (LoggedAdmin or CBool( MediaMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))

blLoggedAdmin = LoggedAdmin

'If they are a member, they can't change someone else's shit
if blLoggedAdmin then
	strMatch = "CustomerID = " & CustomerID
else
	strMatch = "MemberID = " & Session("MemberID")
end if

'Get the category, item ID and check it
if upl.form("ID") = "" or upl.form("MediaID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the category ID."))
intCategoryID = CInt(upl.form("ID"))
intMediaID = CInt(upl.Form("MediaID"))
if not ValidCategory( intCategoryID ) then Redirect("error.asp?Message=" & Server.URLEncode("You are uploading to an invalid category."))

'Open up the item in a recordset
Set rsUpdate = Server.CreateObject("ADODB.Recordset")
Query = "SELECT Date, CategoryID, Description, IP, ModifiedID, FileName FROM Media WHERE ID = " & intMediaID & " AND " & strMatch
rsUpdate.Open Query, Connect, adOpenStatic, adLockOptimistic
if rsUpdate.EOF then
	Set upl = Nothing
	Set rsUpdate = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The item you are trying to access does not exist.  It may have been deleted or you may have misspelled the link.  Try looking in the list for the item."))
end if

rsUpdate("CategoryID") = intCategoryID
rsUpdate("Description") = Format( upl.Form("Name") )
if blLoggedAdmin then rsUpdate("Date") = upl.Form("Date")
rsUpdate("IP") = Request.ServerVariables("REMOTE_HOST")
rsUpdate("ModifiedID") = Session("MemberID")

'They uploaded a new file
if not upl.Form("MediaFile").IsEmpty then

	'Check available size in the media folder (subtracting the file's size who we are overwriting... bad english)
	Set FileSystem = CreateObject("Scripting.FileSystemObject")
	Set TestFolder = FileSystem.GetFolder( GetPath("media") )
	Set TestFile = FileSystem.GetFile( GetPath("media") & rsUpdate("FileName") )
	dblAvailable = (MediaMegs * 1000000) - TestFolder.Size - TestFile.Size
	Set TestFile = Nothing
	Set TestFolder = Nothing

	'Not enough space
	if upl.TotalBytes > dblAvailable then strError = "Sorry, but there is not enough free space in the " & MediaTitle & " section for your file.  There is " & Round( dblAvailable / 1000000, 2 ) & " megabytes available and your file takes up " & Round( upl.TotalBytes / 1000000, 2 ) & " megabytes.  An administrator may purchase more space by clicking on 'Modify Account' in the Members Only section. "

	strFileName = FormatFileName(Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1))
	strExt = GetExtension(strFileName)

	'We can't have duplicate file names in the folder, so keep adding numbers to the end
	intNum = 1
	do until not MediaFileExists( strFileName )
		strFileName = GetJustFile( strFileName ) & intNum & "." & GetExtension( strFileName )
		intNum = intNum + 1
	loop

	'invalid file types
	if lcase(strExt) = ".exe" or lcase(strExt) = ".asp" or lcase(strExt) = ".com" or lcase(strExt) = ".bat" then
		strError = strError & "You are trying to update an invalid type of file.<br>"
	end if

	'Now make sure we still don't have a problem
	if strError <> "" then
		Set rsUpdate = Nothing
		Set upl = Nothing
		Redirect "error.asp?Message=" & Server.URLEncode(strError)
	end if

	'Delete the original, save the new file and update the filename in the database
	if FileSystem.FileExists( GetPath("media") & rsUpdate("FileName") ) then FileSystem.DeleteFile( GetPath("media") & rsUpdate("FileName") )
	upl.SaveAs strFileName
	rsUpdate("FileName") = strFileName
	Set FileSystem = Nothing
end if

rsUpdate.Update
rsUpdate.Close
Set rsUpdate = Nothing
Set upl = Nothing
%>
<p>The file has been edited. &nbsp;<a href="members_media_modify.asp">Click here</a> to modify another.
	<a href="media_read.asp?ID=<%=intMediaID%>">Click here</a> to view the modified file.
</p>

