<%
'
'-----------------------Begin Code----------------------------
if not LoggedAdmin then Redirect("members.asp?Source=admin_advanced_images_edit.asp")
Session.Timeout = 20
Server.ScriptTimeout = 5400
'------------------------End Code-----------------------------
%>
<p align="<%=HeadingAlignment%>"><span class=Heading>Change The Advanced Graphics of Your Site</span><br>
<span class=LinkText><a href="members.asp">Back To <%=MembersTitle%></a></span></p>
<%
'-----------------------Begin Code----------------------------
'update info
Set upl = Server.CreateObject("SoftArtisans.FileUp")
strPath = GetPath ("images")
upl.Path = strPath
Set FileSystem = CreateObject("Scripting.FileSystemObject")

if not LoggedAdmin and upl.Form("MemberID") <> "" and upl.Form("Password") <> "" then Relog upl.Form("MemberID"), upl.Form("Password")
if not LoggedAdmin then
	Set upl = Nothing
	Set FileSystem = Nothing
	Redirect("members.asp?Source=admin_advanced_images_edit.asp")
end if

Query = "SELECT CustomHeader FROM Look WHERE CustomerID = " & CustomerID
Set rsLook = Server.CreateObject("ADODB.Recordset")
rsLook.Open Query, Connect, adOpenStatic, adLockOptimistic
rsLook("CustomHeader") = upl.Form("CustomHeader")



rsLook.Update
set rsLook = Nothing

SetImage "TitleBackgroundImage"
SetImage "MenuTopImage"
SetImage "MenuTopRolloverImage"
SetImage "MenuBottomImage"
SetImage "MenuBottomRolloverImage"
SetImage "MenuBackgroundImage"

Set upl = Nothing
Set FileSystem = Nothing

Redirect("write_header_footer.asp?Source=admin_advanced_images_edit.asp?Submit=Changed")


'-------------------------------------------------------------
'Take in a field name and either upload its image, delete it, or do nothing
'-------------------------------------------------------------
Sub SetImage( strField )
	'Get the image
	if not upl.Form("Up"&strField).IsEmpty then
		'Delete the old one if there is one...
		if FileSystem.FileExists( strPath & strField & ".jpg" ) then FileSystem.DeleteFile( strPath & strField & ".jpg" )
		if FileSystem.FileExists( strPath & strField & ".gif" ) then FileSystem.DeleteFile( strPath & strField & ".gif" )
		if FileSystem.FileExists( strPath & strField & ".bmp" ) then FileSystem.DeleteFile( strPath & strField & ".bmp" )

		'--- Retrieve the file's content type and assign it to a variable
		FTYPE = upl.Form("Up"&strField).ContentType

		'--- Restrict the file types saved using a Select condition
		if FTYPE = "image/gif" then
			upl.Form("Up"&strField).SaveAs strField&".gif"
		elseif FTYPE = "image/pjpeg" or FTYPE = "image/jpeg" then
			upl.Form("Up"&strField).SaveAs strField&".jpg"
		elseif FTYPE = "image/bmp" then
			upl.Form("Up"&strField).SaveAs strField&".bmp"
		else
			upl.Form("Up"&strField).delete
			strError = strError & "You can only upload gif, jpeg, and bitmap (bmp) images.  For " & strField & " you uploaded an invalid type of image.<br>"
		end if

	'Erase the image
	elseif upl.Form(strField) = "0" then
		if FileSystem.FileExists( strPath & strField & ".jpg" ) then FileSystem.DeleteFile( strPath & strField & ".jpg" )
		if FileSystem.FileExists( strPath & strField & ".gif" ) then FileSystem.DeleteFile( strPath & strField & ".gif" )
		if FileSystem.FileExists( strPath & strField & ".bmp" ) then FileSystem.DeleteFile( strPath & strField & ".bmp" )
	end if
End Sub
'------------------------End Code-----------------------------
%>
