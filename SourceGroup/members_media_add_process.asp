<!-- #include file="media_functions.asp" -->

<%
'
'-----------------------Begin Code----------------------------
if not CBool( IncludeMedia ) then Redirect("error.asp?Message=" & Server.URLEncode("Sorry, but you can't add files right now.  The option has been disabled.") )
Session.Timeout = 20
Server.ScriptTimeout = 5400
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Add Media</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
strError = ""

Set upl = Server.CreateObject("SoftArtisans.FileUp")
upl.Path = GetPath ("media")

'Make sure they can do this
if not LoggedMember and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
if not LoggedMember then Redirect("members.asp?Source=members_media_add.asp")
if not (LoggedAdmin or CBool( MediaMembers )) then Redirect("message.asp?Message=" & Server.URLEncode("Sorry, but you can not access this section."))

'Check on the file
if upl.IsEmpty then
	Set upl = Nothing
	Redirect "incomplete.asp?Message=" & Server.URLEncode("You did not specify a file to upload.  Please try again.")
end if

'Get the category and check it
if upl.form("ID") = "" then Redirect("error.asp?Message=" & Server.URLEncode("You are missing the category ID."))
intCategoryID = CInt(upl.form("ID"))
if not ValidCategory( intCategoryID ) then Redirect("error.asp?Message=" & Server.URLEncode("You are uploading to an invalid category."))

'Get rid of the directories and stuff, and get the extension
strFileName = FormatFileName(Mid(upl.UserFilename, InstrRev(upl.UserFilename, "\") + 1))
strExt = GetExtension(strFileName)

'Make sure it isn't executable
if lcase(strExt) = ".exe" or lcase(strExt) = ".asp" or lcase(strExt) = ".com" or lcase(strExt) = ".bat" then
	strError = strError & "You are trying to update an invalid type of file."
end if

'Check available size in the media folder
Set FileSystem = CreateObject("Scripting.FileSystemObject")
Set TestFolder = FileSystem.GetFolder( GetPath("media") )
dblAvailable = (MediaMegs * 1000000) - TestFolder.Size
Set TestFolder = Nothing
Set FileSystem = Nothing

'Too much sauce!
if upl.TotalBytes > dblAvailable then strError = "Sorry, but there is not enough free space in the " & MediaTitle & " section for your file.  There is " & Round( dblAvailable / 1000000, 2 ) & " megabytes available and your file takes up " & Round( upl.TotalBytes / 1000000, 2 ) & " megabytes.  An administrator may purchase more space by clicking on 'Modify Account' in the Members Only section. "

'Now make sure we still don't have a problem
if strError <> "" then
	Set upl = Nothing
	Redirect "message.asp?Message=" & Server.URLEncode(strError)
end if

'We can't have duplicate file names in the folder, so keep adding numbers to the end
intNum = 1
do until not MediaFileExists( strFileName )
	strFileName = GetJustFile( strFileName ) & intNum & "." & GetExtension( strFileName )
	intNum = intNum + 1
loop

'Save this badboy file
upl.SaveAs strFileName

'Now just add it to the database
Set cmdTemp = Server.CreateObject("ADODB.Command")
With cmdTemp
	.ActiveConnection = Connect
	.CommandText = "AddMedia"
	.CommandType = adCmdStoredProc

	.Parameters.Append .CreateParameter ("@ItemID", adInteger, adParamOutput )
	.Parameters.Append .CreateParameter ("@CategoryID", adInteger, adParamInput )
	.Parameters.Append .CreateParameter ("@MemberID", adInteger, adParamInput )
	.Parameters.Append .CreateParameter ("@ModifiedID", adInteger, adParamInput )
	.Parameters.Append .CreateParameter ("@CustomerID", adInteger, adParamInput )
	.Parameters.Append .CreateParameter ("@IP", adVarWChar, adParamInput, 20 )
	.Parameters.Append .CreateParameter ("@FileName", adVarWChar, adParamInput, 400 )
	.Parameters.Append .CreateParameter ("@Description", adVarWChar, adParamInput, 4000 )

	.Parameters("@CategoryID") = intCategoryID
	.Parameters("@MemberID") = Session("MemberID")
	.Parameters("@ModifiedID") = Session("MemberID")
	.Parameters("@CustomerID") = CustomerID
	.Parameters("@IP") = Request.ServerVariables("REMOTE_HOST")
	.Parameters("@FileName") = strFileName
	.Parameters("@Description") = Format( upl.form("Name") )

	.Execute , , adExecuteNoRecords
	intFileID = .Parameters("@ItemID")
End With

Set cmdTemp = Nothing
Set upl = Nothing
%>
<p>The file has been added. &nbsp;<a href="members_media_add.asp?ID=<%=intCategoryID%>">Click here</a> to add another.<br>
<a href="media_read.asp?ID=<%=intFileID%>">Click here</a> to view the file you just uploaded.
</p>
