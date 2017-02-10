<%
'
'-----------------------Begin Code----------------------------
Session.Timeout = 20
'------------------------End Code-----------------------------
%>

<p class="Heading" align="<%=HeadingAlignment%>">Edit A File</p>
<p class=LinkText align=<%=HeadingAlignment%>><a href="<%=NonSecurePath%>members.asp">Back To <%=MembersTitle%></a></p>

<%
'-----------------------Begin Code----------------------------
'update info
Set upl = Server.CreateObject("SoftArtisans.FileUp")
upl.Path = GetPath("")

if not LoggedAdmin and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
if not LoggedAdmin then Redirect("members.asp?Source=members_images_modify_process.asp")


strFileName = upl.Form("FileName")
strPath = upl.Form("Path")

if not ( strPath = "inserts" or strPath = "storeitems" or strPath = "posts" or strPath = "images" or strPath = "storegroups" or strPath = "media" or strPath = "photos" or strPath = "schemes" ) then Redirect("error.asp")

strSource= upl.Form("Source")

if strFileName = "" or strPath = "" then Redirect("incomplete.asp")

strPath = GetPath( strPath )

'Make sure the image exists
Set FileSystem = CreateObject("Scripting.FileSystemObject")
strFullFileName = strPath & strFileName

if not FileSystem.FileExists( strFullFileName ) then
	Set FileSystem = Nothing
	Redirect("error.asp?Message=" & Server.URLEncode("The file does not exist."))
end if

if not upl.Form("File").IsEmpty then
	'--- Retrieve the file's content type and assign it to a variable
	FTYPE = upl.Form("File").ContentType

	'--- Restrict the file types saved
	if FTYPE = "text/asp" then
		Set FileSystem = Nothing
		Set upl = Nothing
		Redirect("error.asp?Message=" & Server.URLEncode("You cannot upload ASP files.<br>"))
	else
		upl.Form("File").SaveAs strPath & strFileName
	end if
end if

Set upl = Nothing
Set FileSystem = Nothing

if strSource = "" then
	strLink = "javascript:history.go(-2)"
else
	strLink = strSource
end if

%>
<p>The file has been edited. &nbsp;<a href="<%=strLink%>">Click here</a> to go back.</p>

